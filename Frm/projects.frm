VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form Projects 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Projects"
   ClientHeight    =   9810
   ClientLeft      =   -2490
   ClientTop       =   435
   ClientWidth     =   18780
   Icon            =   "projects.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   18780
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text9 
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
      Left            =   19560
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   241
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.TextBox Text8 
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
      Left            =   19560
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   240
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.TextBox Text14 
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
      Left            =   19560
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   239
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.TextBox Text11 
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
      Left            =   21120
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   234
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.TextBox txt_project_id 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   23160
      TabIndex        =   225
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      Caption         =   "⁄„«·Â"
      Height          =   195
      Left            =   21360
      RightToLeft     =   -1  'True
      TabIndex        =   222
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "»‰Êœ"
      Height          =   195
      Left            =   20400
      RightToLeft     =   -1  'True
      TabIndex        =   220
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      Caption         =   "«’‰«›"
      Height          =   195
      Left            =   21120
      RightToLeft     =   -1  'True
      TabIndex        =   219
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame10 
      Caption         =   "«”„«¡ «·⁄«„·Ì‰ ›Ì «·„‘—Ê⁄"
      Height          =   3615
      Left            =   20760
      RightToLeft     =   -1  'True
      TabIndex        =   200
      Top             =   6360
      Visible         =   0   'False
      Width           =   19095
      Begin VB.TextBox txt_employee_count 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9960
         TabIndex        =   214
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox txt_emp_salary 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   6360
         TabIndex        =   213
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   201
         Top             =   240
         Width           =   18855
         Begin VB.TextBox TxtEmpcount 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   7800
            TabIndex        =   206
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "«œ—«Ã"
            Height          =   255
            Left            =   960
            TabIndex        =   205
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox TxtCount 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5280
            TabIndex        =   204
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton Option5 
            Alignment       =   1  'Right Justify
            Caption         =   " Œ’Ì’ ›⁄·Ì"
            Height          =   255
            Left            =   10200
            RightToLeft     =   -1  'True
            TabIndex        =   203
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton Option4 
            Alignment       =   1  'Right Justify
            Caption         =   " ﬁœÌ—Ì"
            Height          =   255
            Left            =   11640
            RightToLeft     =   -1  'True
            TabIndex        =   202
            Top             =   120
            Value           =   -1  'True
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo dcJobTypeName 
            Height          =   315
            Left            =   13080
            TabIndex        =   207
            Top             =   120
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   2520
            TabIndex        =   208
            Top             =   120
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   93519873
            CurrentDate     =   38784
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "⁄œœ «·«Ì«„"
            Height          =   255
            Left            =   6840
            TabIndex        =   212
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   " «—ÌŒ «· Œ’Ì’"
            Height          =   255
            Left            =   3960
            TabIndex        =   211
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "⁄œœ «·⁄„«·"
            Height          =   255
            Left            =   9120
            TabIndex        =   210
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "«Œ — «·„Â‰… «·„ÿ·Ê»…"
            Height          =   255
            Left            =   16440
            TabIndex        =   209
            Top             =   120
            Width           =   1575
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   1860
         Left            =   120
         TabIndex        =   215
         Top             =   960
         Width           =   18840
         _cx             =   33232
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
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"projects.frx":000C
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
      Begin ALLButtonS.ALLButton opr_emplyees_name 
         Height          =   375
         Left            =   120
         TabIndex        =   216
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "—ÃÊ⁄ ··⁄„·Ì« "
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "projects.frx":01A0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label27 
         Caption         =   "«Ã„«·Ì ⁄œœ «·⁄„·"
         Height          =   255
         Left            =   11640
         TabIndex        =   218
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label29 
         Caption         =   "ﬁÌ„… «ÃÊ— «·⁄„«·"
         Height          =   255
         Left            =   8040
         TabIndex        =   217
         Top             =   3120
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "„Ê«œ «·⁄„·Ì… —ﬁ„"
      Height          =   3615
      Left            =   20520
      RightToLeft     =   -1  'True
      TabIndex        =   165
      Top             =   6120
      Visible         =   0   'False
      Width           =   19095
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   480
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   173
         Top             =   480
         Width           =   1530
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   480
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   172
         Top             =   840
         Width           =   1530
      End
      Begin VB.TextBox XPTxtSum 
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
         Left            =   480
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   171
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1530
      End
      Begin VB.TextBox TxtFillData 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   375
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   170
         Top             =   0
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox XPTxtBillID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   360
         Left            =   720
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   169
         Top             =   0
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   480
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   168
         Top             =   1320
         Width           =   1530
      End
      Begin VB.ComboBox XPCboDiscountType 
         Height          =   288
         Left            =   10560
         TabIndex        =   167
         Text            =   "Combo2"
         Top             =   3240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox XPTxtDiscountVal 
         Height          =   375
         Left            =   11400
         TabIndex        =   166
         Text            =   "Text7"
         Top             =   3240
         Visible         =   0   'False
         Width           =   255
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   690
         Index           =   2
         Left            =   3600
         TabIndex        =   174
         TabStop         =   0   'False
         Top             =   240
         Width           =   15435
         _cx             =   27226
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
         Appearance      =   0
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
         Begin VB.ComboBox CboItemCase 
            Height          =   288
            Left            =   7290
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   180
            Top             =   -420
            Width           =   2010
         End
         Begin VB.TextBox TxtQuantity 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2715
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   179
            Top             =   300
            Width           =   1875
         End
         Begin VB.TextBox TxtSerial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Left            =   4590
            MaxLength       =   20
            RightToLeft     =   -1  'True
            TabIndex        =   178
            Top             =   -300
            Width           =   2640
         End
         Begin VB.TextBox TxtPrice 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   795
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   177
            Top             =   300
            Width           =   1860
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4890
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   176
            Top             =   300
            Width           =   1860
         End
         Begin VB.TextBox Text16 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   6975
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   175
            Top             =   300
            Width           =   1875
         End
         Begin MSDataListLib.DataCombo DCboItemsName 
            Height          =   315
            Left            =   9300
            TabIndex        =   181
            Top             =   300
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboItemsCode 
            Height          =   315
            Left            =   12240
            TabIndex        =   182
            Top             =   300
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdAdd 
            Height          =   375
            Left            =   105
            TabIndex        =   183
            Top             =   270
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            ButtonImage     =   "projects.frx":01BC
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   255
            Index           =   31
            Left            =   12720
            RightToLeft     =   -1  'True
            TabIndex        =   191
            Top             =   0
            Width           =   2685
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈”„ «·’‰›"
            Height          =   255
            Index           =   30
            Left            =   9615
            RightToLeft     =   -1  'True
            TabIndex        =   190
            Top             =   0
            Width           =   2625
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·’‰›"
            Height          =   255
            Index           =   29
            Left            =   7470
            RightToLeft     =   -1  'True
            TabIndex        =   189
            Top             =   -720
            Width           =   1830
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ì—Ì«·"
            Height          =   255
            Index           =   28
            Left            =   4995
            RightToLeft     =   -1  'True
            TabIndex        =   188
            Top             =   -600
            Width           =   2265
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ„Ì… «·›⁄·Ì…"
            Height          =   255
            Index           =   27
            Left            =   3195
            RightToLeft     =   -1  'True
            TabIndex        =   187
            Top             =   0
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄— «·›⁄·Ì"
            Height          =   255
            Index           =   26
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   186
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄—  ﬁœÌ—Ì"
            Height          =   255
            Index           =   12
            Left            =   5475
            RightToLeft     =   -1  'True
            TabIndex        =   185
            Top             =   0
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ„Ì…  ﬁœÌ—Ì"
            Height          =   255
            Index           =   19
            Left            =   7470
            RightToLeft     =   -1  'True
            TabIndex        =   184
            Top             =   0
            Width           =   1320
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid FG 
         Height          =   2145
         Left            =   3600
         TabIndex        =   192
         Top             =   960
         Width           =   15435
         _cx             =   27226
         _cy             =   3784
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
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"projects.frx":0556
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
         WallPaperAlignment=   0
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin ALLButtonS.ALLButton opr_items 
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   193
         Top             =   2640
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "—ÃÊ⁄ ··⁄„·Ì« "
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "projects.frx":07A5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ Ê›—"
         Height          =   255
         Index           =   0
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   199
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ÕÃÊ“"
         Height          =   255
         Index           =   1
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   198
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì ﬁÌ„… «·«’‰«›"
         Height          =   255
         Index           =   2
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   197
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label LblItemsCount 
         Caption         =   "Label27"
         Height          =   135
         Left            =   240
         TabIndex        =   196
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ÿ·Ê»"
         Height          =   255
         Index           =   3
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   195
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label LblTotalQty 
         Caption         =   "Label38"
         Height          =   135
         Left            =   12120
         TabIndex        =   194
         Top             =   3240
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "«·„’—Ê›« "
      Height          =   3615
      Left            =   20280
      RightToLeft     =   -1  'True
      TabIndex        =   160
      Top             =   5760
      Visible         =   0   'False
      Width           =   19215
      Begin VB.TextBox txt_expenses_total 
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
         Left            =   6960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   161
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1530
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
         Height          =   2340
         Left            =   120
         TabIndex        =   162
         Top             =   360
         Width           =   18960
         _cx             =   33443
         _cy             =   4128
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
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"projects.frx":07C1
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
      Begin ALLButtonS.ALLButton opr_Expenses 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   163
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "—ÃÊ⁄ ··⁄„·Ì« "
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "projects.frx":08F2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì ﬁÌ„… «·„’—Ê›« "
         Height          =   255
         Index           =   6
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   164
         Top             =   2760
         Width           =   2535
      End
   End
   Begin C1SizerLibCtl.C1Elastic CEMain 
      Height          =   9810
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   18780
      _cx             =   33126
      _cy             =   17304
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
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Õ  «· ‰›Ì–"
         Height          =   240
         Left            =   12120
         TabIndex        =   297
         Top             =   225
         Width           =   1440
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   600
         Left            =   0
         TabIndex        =   253
         TabStop         =   0   'False
         Top             =   9210
         Width           =   18780
         _cx             =   33126
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
         BackColor       =   -2147483633
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
         Begin ALLButtonS.ALLButton Command1 
            Height          =   420
            Index           =   1
            Left            =   16320
            TabIndex        =   254
            Top             =   120
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   741
            BTYPE           =   3
            TX              =   "Õ›Ÿ"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":090E
            PICN            =   "projects.frx":092A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   405
            Index           =   2
            Left            =   11775
            TabIndex        =   255
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "«·„—›ﬁ« "
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   255
            BCOLO           =   192
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":718C
            PICN            =   "projects.frx":71A8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   405
            Index           =   0
            Left            =   17925
            TabIndex        =   256
            Top             =   135
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "ÃœÌœ"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   8454016
            BCOLO           =   8454016
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":DA0A
            PICN            =   "projects.frx":DA26
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   405
            Index           =   5
            Left            =   7635
            TabIndex        =   257
            Top             =   120
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "»ÕÀ"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":14288
            PICN            =   "projects.frx":142A4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command3 
            Height          =   405
            Left            =   30
            TabIndex        =   258
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   " ’œÌ— «·Ï «ﬂ”·"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":1AB06
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   -1  'True
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   405
            Index           =   3
            Left            =   17100
            TabIndex        =   259
            Top             =   135
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   " ⁄œÌ·"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":1AB22
            PICN            =   "projects.frx":1AB3E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   405
            Index           =   6
            Left            =   15420
            TabIndex        =   260
            Top             =   135
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   " —«Ã⁄"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":213A0
            PICN            =   "projects.frx":213BC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   405
            Index           =   7
            Left            =   14460
            TabIndex        =   261
            Top             =   135
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "Õ–›"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":27C1E
            PICN            =   "projects.frx":27C3A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   405
            Index           =   4
            Left            =   10020
            TabIndex        =   262
            Top             =   120
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "«·«—’œ… «·«›  «ÕÌ…"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":2E49C
            PICN            =   "projects.frx":2E4B8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton Command1 
            Height          =   405
            Index           =   8
            Left            =   8565
            TabIndex        =   263
            Top             =   135
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "»Ì«‰«  «·œ›⁄« "
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":34D1A
            PICN            =   "projects.frx":34D36
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton1 
            Height          =   405
            Left            =   4725
            TabIndex        =   264
            Top             =   135
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "ÿ»«⁄Â «Ã„«·Ì"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":5E958
            PICN            =   "projects.frx":5E974
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   -1  'True
         End
         Begin ALLButtonS.ALLButton ALLButton2 
            Height          =   405
            Left            =   3315
            TabIndex        =   265
            Top             =   135
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "ÿ»«⁄Â  Õ·Ì·Ì"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":651D6
            PICN            =   "projects.frx":651F2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   -1  'True
         End
         Begin ALLButtonS.ALLButton BtnSalary 
            Height          =   405
            Left            =   12975
            TabIndex        =   313
            Top             =   135
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "—Ê« » «·⁄«„·Ì‰"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":6BA54
            PICN            =   "projects.frx":6BA70
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton3 
            Height          =   405
            Left            =   1440
            TabIndex        =   325
            Top             =   135
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   " ’œÌ—  ›’Ì·Ì «·Ï «ﬂ”·"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":722D2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   -1  'True
         End
         Begin ALLButtonS.ALLButton ALLButton4 
            Height          =   405
            Left            =   6120
            TabIndex        =   339
            Top             =   120
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "«·„” Œœ„Ì‰"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":722EE
            PICN            =   "projects.frx":7230A
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
      Begin C1SizerLibCtl.C1Elastic Frame8 
         Height          =   450
         Left            =   8235
         TabIndex        =   248
         TabStop         =   0   'False
         Top             =   120
         Width           =   3555
         _cx             =   6271
         _cy             =   794
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
         BackColor       =   12648447
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "‰Ê⁄ «·„‘—Ê⁄"
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
         Begin VB.OptionButton Ptype 
            BackColor       =   &H00C0FFFF&
            Caption         =   "«›  «ÕÌ"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   250
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton Ptype 
            BackColor       =   &H00C0FFFF&
            Caption         =   "ÃœÌœ"
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   249
            Top             =   120
            Width           =   735
         End
      End
      Begin C1SizerLibCtl.C1Elastic Frame7 
         Height          =   600
         Left            =   360
         TabIndex        =   159
         TabStop         =   0   'False
         Top             =   9195
         Visible         =   0   'False
         Width           =   18345
         _cx             =   32359
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
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   495
         Left            =   19770
         TabIndex        =   156
         Top             =   3435
         Visible         =   0   'False
         Width           =   1200
      End
      Begin C1SizerLibCtl.C1Elastic Frame13 
         Height          =   3000
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   810
         Width           =   18705
         _cx             =   32994
         _cy             =   5292
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
         Begin VB.Frame Frame21 
            Height          =   2415
            Left            =   0
            TabIndex        =   330
            Top             =   480
            Width           =   7695
            Begin VB.ListBox ListUserSelect 
               BackColor       =   &H0080FFFF&
               Height          =   1620
               ItemData        =   "projects.frx":78B6C
               Left            =   240
               List            =   "projects.frx":78B73
               RightToLeft     =   -1  'True
               TabIndex        =   333
               Top             =   600
               Width           =   3135
            End
            Begin VB.ListBox ListAllUser 
               Height          =   1620
               ItemData        =   "projects.frx":78B87
               Left            =   4320
               List            =   "projects.frx":78B8E
               RightToLeft     =   -1  'True
               TabIndex        =   332
               Top             =   600
               Width           =   3015
            End
            Begin VB.CommandButton Command7 
               BackColor       =   &H000000FF&
               Caption         =   "X"
               Height          =   255
               Left            =   7200
               Style           =   1  'Graphical
               TabIndex        =   331
               Top             =   120
               Width           =   375
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«·„” Œœ„Ì‰"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   42
               Left            =   3240
               RightToLeft     =   -1  'True
               TabIndex        =   338
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label47 
               Alignment       =   2  'Center
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   337
               Top             =   1680
               Width           =   495
            End
            Begin VB.Label Label46 
               Alignment       =   2  'Center
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   336
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   335
               Top             =   960
               Width           =   495
            End
            Begin VB.Label LblSelect 
               Alignment       =   2  'Center
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   334
               Top             =   600
               Width           =   495
            End
         End
         Begin VB.TextBox txtFile 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   323
            Top             =   600
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Frame Frame20 
            Height          =   2415
            Left            =   7680
            TabIndex        =   276
            Top             =   600
            Width           =   6135
            Begin VB.CommandButton Command5 
               BackColor       =   &H000000FF&
               Caption         =   "X"
               Height          =   255
               Left            =   5640
               Style           =   1  'Graphical
               TabIndex        =   277
               Top             =   120
               Width           =   375
            End
            Begin VSFlex8Ctl.VSFlexGrid Grid 
               Height          =   1620
               Left            =   120
               TabIndex        =   285
               Top             =   480
               Width           =   5880
               _cx             =   10372
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
               FormatString    =   $"projects.frx":78B9F
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
               Left            =   4320
               TabIndex        =   286
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
               ButtonImage     =   "projects.frx":78C5A
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label Label64 
               Alignment       =   2  'Center
               Caption         =   "»Ì«‰«  «·„›—œ«  ÿ»ﬁ« ··„‘—Ê⁄"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1320
               TabIndex        =   278
               Top             =   120
               Width           =   3255
            End
         End
         Begin C1SizerLibCtl.C1Elastic Frame14 
            Height          =   555
            Left            =   105
            TabIndex        =   243
            TabStop         =   0   'False
            Top             =   0
            Width           =   11910
            _cx             =   21008
            _cy             =   979
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
            Begin VB.TextBox TXTOrDer_no 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   8280
               TabIndex        =   274
               Top             =   120
               Width           =   1350
            End
            Begin VB.ComboBox CBoBasedON 
               Height          =   315
               ItemData        =   "projects.frx":791F4
               Left            =   9780
               List            =   "projects.frx":791F6
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   273
               Top             =   120
               Width           =   1110
            End
            Begin MSDataListLib.DataCombo DcCurrency 
               Height          =   315
               Left            =   240
               TabIndex        =   244
               Top             =   120
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo Dcbranch 
               Height          =   315
               Left            =   2100
               TabIndex        =   245
               Top             =   120
               Width           =   5565
               _ExtentX        =   9816
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "»‰«¡ ⁄·Ï"
               Height          =   255
               Index           =   56
               Left            =   10920
               TabIndex        =   275
               Top             =   120
               Width           =   810
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               Caption         =   "«·⁄„·Â"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1395
               RightToLeft     =   -1  'True
               TabIndex        =   247
               Top             =   120
               Width           =   735
            End
            Begin VB.Label Label26 
               Alignment       =   2  'Center
               Caption         =   "«·›—⁄"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   7530
               RightToLeft     =   -1  'True
               TabIndex        =   246
               Top             =   120
               Width           =   885
            End
         End
         Begin C1SizerLibCtl.C1Elastic Frame16 
            Height          =   2325
            Left            =   2985
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   555
            Width           =   3105
            _cx             =   5477
            _cy             =   4101
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
            Begin VB.TextBox Text20 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   5
               Left            =   195
               TabIndex        =   319
               Top             =   120
               Width           =   1770
            End
            Begin VB.TextBox Text20 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   4
               Left            =   195
               TabIndex        =   318
               Top             =   510
               Width           =   1770
            End
            Begin VB.TextBox Text20 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   3
               Left            =   195
               TabIndex        =   317
               Top             =   840
               Width           =   1770
            End
            Begin VB.TextBox Text20 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   2
               Left            =   195
               TabIndex        =   316
               Top             =   1230
               Width           =   1770
            End
            Begin VB.TextBox Text20 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   1
               Left            =   195
               TabIndex        =   315
               Top             =   1560
               Width           =   1770
            End
            Begin VB.TextBox Text20 
               Alignment       =   2  'Center
               Height          =   270
               Index           =   0
               Left            =   195
               TabIndex        =   52
               Top             =   1950
               Width           =   1770
            End
            Begin VB.Label Label49 
               Alignment       =   2  'Center
               Caption         =   "«·Êﬁ  «·„ »ﬁÌ"
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   4
               Left            =   2025
               TabIndex        =   58
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label Label49 
               Alignment       =   2  'Center
               Caption         =   "«·„Õ’·"
               ForeColor       =   &H00000000&
               Height          =   345
               Index           =   2
               Left            =   2145
               TabIndex        =   57
               Top             =   855
               Width           =   855
            End
            Begin VB.Label Label49 
               Alignment       =   2  'Center
               Caption         =   "«·»«ﬁÌ"
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   3
               Left            =   2145
               TabIndex        =   56
               Top             =   1200
               Width           =   855
            End
            Begin VB.Label Label49 
               Alignment       =   2  'Center
               Caption         =   "«·Êﬁ "
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   5
               Left            =   2145
               TabIndex        =   55
               Top             =   1950
               Width           =   855
            End
            Begin VB.Label Label49 
               Alignment       =   2  'Center
               Caption         =   "«·„‰›–"
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   0
               Left            =   2145
               TabIndex        =   54
               Top             =   135
               Width           =   855
            End
            Begin VB.Label Label49 
               Alignment       =   2  'Center
               Caption         =   "‰”»… «·„‰›–"
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   1
               Left            =   2160
               TabIndex        =   53
               Top             =   480
               Width           =   855
            End
         End
         Begin C1SizerLibCtl.C1Elastic Frame9 
            Height          =   2445
            Left            =   6075
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   555
            Width           =   5955
            _cx             =   10504
            _cy             =   4313
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
            Begin VB.TextBox TxtCustCode2 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   3675
               TabIndex        =   39
               Top             =   900
               Width           =   1080
            End
            Begin VB.TextBox Text10 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   3675
               TabIndex        =   38
               Top             =   1260
               Width           =   1080
            End
            Begin VB.TextBox TxtCustCode1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3675
               TabIndex        =   37
               Top             =   420
               Width           =   1080
            End
            Begin VB.TextBox TxtCustCode 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   3675
               TabIndex        =   36
               Top             =   30
               Width           =   1080
            End
            Begin VB.TextBox TxtRemarks 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   35
               Top             =   2055
               Width           =   4635
            End
            Begin MSDataListLib.DataCombo DcAccount2 
               Height          =   315
               Left            =   120
               TabIndex        =   40
               Top             =   30
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcAccount4 
               Height          =   315
               Left            =   120
               TabIndex        =   41
               Top             =   420
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcEmp 
               Height          =   315
               Left            =   120
               TabIndex        =   42
               Top             =   1260
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcEmp1 
               Height          =   315
               Left            =   120
               TabIndex        =   43
               Top             =   900
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbDept 
               Height          =   315
               Left            =   120
               TabIndex        =   44
               Top             =   1635
               Width           =   4635
               _ExtentX        =   8176
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label43 
               Alignment       =   2  'Center
               Caption         =   "«·«œ«—…"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   4860
               TabIndex        =   50
               Top             =   1635
               Width           =   975
            End
            Begin VB.Label Label41 
               Alignment       =   2  'Center
               Caption         =   "«·„‰œÊ»"
               ForeColor       =   &H00000000&
               Height          =   360
               Left            =   4860
               TabIndex        =   49
               Top             =   900
               Width           =   975
            End
            Begin VB.Label Label35 
               Alignment       =   2  'Center
               Caption         =   "„œÌ— «·„Êﬁ⁄"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   4860
               TabIndex        =   48
               Top             =   1260
               Width           =   975
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               Caption         =   "„·«ÕŸ« "
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4860
               TabIndex        =   47
               Top             =   2055
               Width           =   975
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               Caption         =   "«·⁄„Ì· «·‰Â«∆Ì"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4740
               TabIndex        =   46
               Top             =   30
               Width           =   1095
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               Caption         =   "«·⁄„Ì· «·»«ÿ‰"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   4740
               TabIndex        =   45
               Top             =   420
               Width           =   1095
            End
         End
         Begin C1SizerLibCtl.C1Elastic Frame15 
            Height          =   3120
            Left            =   12015
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   6690
            _cx             =   11800
            _cy             =   5503
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
            Begin VB.TextBox TxtContractNo 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               TabIndex        =   291
               Top             =   1215
               Width           =   2085
            End
            Begin VB.TextBox TXTprojectnamee 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   120
               TabIndex        =   288
               Top             =   870
               Width           =   4635
            End
            Begin VB.TextBox TXTprojectname 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Left            =   120
               TabIndex        =   17
               Top             =   480
               Width           =   4635
            End
            Begin VB.TextBox TxtDiscountPercentage 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   495
               TabIndex        =   16
               Top             =   1875
               Width           =   1710
            End
            Begin VB.TextBox txt_total_discount 
               Alignment       =   1  'Right Justify
               Height          =   225
               Left            =   120
               TabIndex        =   15
               Top             =   1575
               Width           =   2085
            End
            Begin VB.TextBox txtid 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3510
               TabIndex        =   14
               Top             =   105
               Width           =   1245
            End
            Begin VB.TextBox TxtProjectCosts 
               Alignment       =   1  'Right Justify
               Height          =   225
               Left            =   3165
               TabIndex        =   13
               Top             =   1575
               Width           =   1590
            End
            Begin VB.TextBox total_after_discount 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3165
               Locked          =   -1  'True
               TabIndex        =   12
               Top             =   1875
               Width           =   1590
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   3165
               TabIndex        =   11
               Top             =   2610
               Width           =   1590
            End
            Begin MSDataListLib.DataCombo DCPreFix 
               Height          =   315
               Left            =   2550
               TabIndex        =   18
               Top             =   120
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               Height          =   315
               Left            =   120
               TabIndex        =   19
               Top             =   105
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker DTEnddate 
               Height          =   330
               Left            =   120
               TabIndex        =   20
               Top             =   2220
               Width           =   2085
               _ExtentX        =   3678
               _ExtentY        =   582
               _Version        =   393216
               Format          =   93519873
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DpNearEndDate 
               Height          =   300
               Left            =   120
               TabIndex        =   21
               Top             =   2625
               Width           =   2085
               _ExtentX        =   3678
               _ExtentY        =   529
               _Version        =   393216
               Format          =   93519873
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTStartDate 
               Height          =   330
               Left            =   3165
               TabIndex        =   289
               Top             =   2220
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               Format          =   93519873
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DataCombo5 
               Height          =   315
               Left            =   3165
               TabIndex        =   290
               Top             =   1215
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label66 
               Alignment       =   2  'Center
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   270
               Left            =   120
               TabIndex        =   294
               Top             =   1920
               Width           =   390
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "—ﬁ„ «·⁄ﬁœ"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2190
               TabIndex        =   292
               Top             =   1230
               Width           =   870
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "‰Ê⁄ «·⁄ﬁœ"
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   37
               Left            =   4740
               RightToLeft     =   -1  'True
               TabIndex        =   287
               Top             =   1215
               Width           =   1830
            End
            Begin VB.Label Label42 
               Alignment       =   2  'Center
               Caption         =   "‰”»…«·Œ’„"
               ForeColor       =   &H00000000&
               Height          =   270
               Left            =   2190
               TabIndex        =   33
               Top             =   1875
               Width           =   870
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«”„ «·„‘—Ê⁄ «‰Ã·Ì“Ì"
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   36
               Left            =   4740
               TabIndex        =   32
               Top             =   870
               Width           =   1830
            End
            Begin VB.Label Label36 
               Alignment       =   2  'Center
               Caption         =   "«ﬁ—» ‰Â«Ì…"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2190
               TabIndex        =   31
               Top             =   2625
               Width           =   870
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "„œ… «·„‘—Ê⁄"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   41
               Left            =   4740
               TabIndex        =   30
               Top             =   2625
               Width           =   1830
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               Caption         =   "«·«‰ Â«¡ "
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   2310
               TabIndex        =   29
               Top             =   2220
               Width           =   750
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   " «—ÌŒ «·»œ«Ì…"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   40
               Left            =   4740
               TabIndex        =   28
               Top             =   2220
               Width           =   1830
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               Caption         =   "Õ«·… «·„‘—Ê⁄"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   1470
               TabIndex        =   27
               Top             =   105
               Width           =   975
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               Caption         =   "ﬂÊœ «·„‘—Ê⁄"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   4740
               TabIndex        =   26
               Top             =   105
               Width           =   1830
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«”„ «·„‘—Ê⁄ ⁄—»Ì"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   35
               Left            =   4740
               TabIndex        =   25
               Top             =   480
               Width           =   1830
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "ﬁÌ„… «·„‘—Ê⁄"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   38
               Left            =   4740
               TabIndex        =   24
               Top             =   1575
               Width           =   1830
            End
            Begin VB.Label Label32 
               Alignment       =   2  'Center
               Caption         =   "ﬁÌ„… «·Œ’„"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2190
               TabIndex        =   23
               Top             =   1575
               Width           =   870
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "ﬁÌ„… «·„‘—Ê⁄  »⁄œ «·Œ’„"
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   39
               Left            =   4740
               TabIndex        =   22
               Top             =   1875
               Width           =   1830
            End
         End
         Begin MSDataListLib.DataCombo AmanhNames 
            Height          =   315
            Left            =   240
            TabIndex        =   59
            Top             =   2175
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo MunicipalityNames 
            Height          =   315
            Left            =   240
            TabIndex        =   60
            Top             =   2550
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   1260
            Left            =   120
            TabIndex        =   279
            TabStop         =   0   'False
            Top             =   2520
            Visible         =   0   'False
            Width           =   5625
            _cx             =   9922
            _cy             =   2223
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
            Begin VB.TextBox TxtEmpSalary 
               Alignment       =   2  'Center
               Height          =   330
               Left            =   255
               TabIndex        =   281
               Top             =   405
               Width           =   3285
            End
            Begin VB.TextBox TxtMangSalary 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   255
               TabIndex        =   280
               Top             =   810
               Width           =   3285
            End
            Begin VB.Label Label61 
               Alignment       =   2  'Center
               Caption         =   "—« » «·„ÊŸ›"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3660
               TabIndex        =   284
               Top             =   405
               Width           =   1830
            End
            Begin VB.Label Label62 
               Alignment       =   2  'Center
               Caption         =   "—« » «·„‘—›"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3660
               TabIndex        =   283
               Top             =   810
               Width           =   1830
            End
            Begin VB.Label Label63 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "X"
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
               Height          =   420
               Left            =   5145
               RightToLeft     =   -1  'True
               TabIndex        =   282
               Top             =   135
               Width           =   285
            End
         End
         Begin XtremeSuiteControls.RadioButton RdTyp 
            Height          =   315
            Index           =   1
            Left            =   930
            TabIndex        =   320
            Top             =   600
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "„‰ „·›"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   315
            Left            =   810
            TabIndex        =   321
            Top             =   1320
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            Caption         =   "«” Ì—«œ «·„·›"
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
            ButtonImage     =   "projects.frx":791F8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   315
            Left            =   810
            TabIndex        =   322
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   960
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            Caption         =   "Õœœ «·„”«—"
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
            ButtonImage     =   "projects.frx":7FA5A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   0
            Top             =   600
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin XtremeSuiteControls.RadioButton RdTyp 
            Height          =   315
            Index           =   0
            Left            =   2040
            TabIndex        =   324
            Top             =   600
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "ÌœÊÌ"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
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
            Height          =   240
            Index           =   4
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   570
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label Label58 
            Alignment       =   2  'Center
            Caption         =   "«·«„«‰…"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1665
            TabIndex        =   62
            Top             =   2190
            Width           =   855
         End
         Begin VB.Label Label59 
            Alignment       =   2  'Center
            Caption         =   "«·»·œÌ…"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1665
            TabIndex        =   61
            Top             =   2670
            Width           =   855
         End
      End
      Begin VB.TextBox TxtModFlg 
         Height          =   270
         Left            =   11280
         TabIndex        =   3
         Text            =   "txtmodflag"
         Top             =   0
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0FFFF&
         Caption         =   " ﬁœÌ—Ì"
         Height          =   240
         Left            =   14895
         TabIndex        =   2
         Top             =   225
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "›⁄·Ì"
         Height          =   240
         Left            =   13710
         TabIndex        =   1
         Top             =   225
         Width           =   1080
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   330
         Index           =   0
         Left            =   1410
         TabIndex        =   4
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "projects.frx":862BC
         ColorButton     =   12648447
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
         Height          =   330
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "projects.frx":86656
         ColorButton     =   12648447
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
         Height          =   330
         Index           =   1
         Left            =   1935
         TabIndex        =   6
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         ButtonStyle     =   1
         Caption         =   ""
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "projects.frx":869F0
         ColorButton     =   12648447
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         Alignment       =   0
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         RightToLeft     =   -1  'True
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   330
         Index           =   3
         Left            =   885
         TabIndex        =   7
         Top             =   120
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "projects.frx":86D8A
         ColorButton     =   12648447
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin MSComCtl2.DTPicker XPDtbBill 
         Height          =   285
         Left            =   21195
         TabIndex        =   157
         Top             =   4260
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         _Version        =   393216
         Format          =   93519873
         CurrentDate     =   38784
      End
      Begin VB.Frame Frame2 
         Height          =   5415
         Left            =   0
         TabIndex        =   306
         Top             =   3630
         Width           =   18735
         Begin VB.TextBox TxTotalMainDes 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   5520
            TabIndex        =   326
            Top             =   5280
            Visible         =   0   'False
            Width           =   2310
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   1
            Left            =   13440
            TabIndex        =   308
            Top             =   4800
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
            ButtonImage     =   "projects.frx":87124
            DrawFocusRectangle=   0   'False
         End
         Begin ALLButtonS.ALLButton terms_operations 
            Height          =   360
            Index           =   2
            Left            =   15480
            TabIndex        =   311
            Top             =   4800
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   635
            BTYPE           =   3
            TX              =   "«·»‰Êœ «· ›’Ì·Ì…"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":876BE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid FgMainDes 
            Height          =   4260
            Left            =   120
            TabIndex        =   307
            Top             =   480
            Width           =   18360
            _cx             =   32385
            _cy             =   7514
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
            BackColorAlternate=   16777088
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
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"projects.frx":876DA
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   34
            Left            =   6720
            TabIndex        =   329
            Top             =   4920
            Width           =   2430
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   33
            Left            =   960
            TabIndex        =   328
            Top             =   4920
            Width           =   2430
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "«·«Ã„«·Ì «·„‰›–"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   32
            Left            =   3600
            TabIndex        =   327
            Top             =   4920
            Width           =   1830
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "«·«Ã„«·Ì «·›⁄·Ì"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   25
            Left            =   9135
            TabIndex        =   312
            Top             =   4920
            Width           =   1830
         End
         Begin VB.Label Label67 
            Alignment       =   2  'Center
            Caption         =   "«·»‰Êœ «·—∆Ì”Ì…"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   16560
            TabIndex        =   309
            Top             =   120
            Width           =   1815
         End
      End
      Begin C1SizerLibCtl.C1Elastic Frame11 
         Height          =   5445
         Left            =   90
         TabIndex        =   142
         TabStop         =   0   'False
         Top             =   3570
         Width           =   18705
         _cx             =   32994
         _cy             =   9604
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
         BackColor       =   12648447
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "⁄„·Ì«  ﬂ· »‰œ"
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
         Begin C1SizerLibCtl.C1Elastic Frame4 
            Height          =   615
            Left            =   13920
            TabIndex        =   251
            TabStop         =   0   'False
            Top             =   4485
            Width           =   2385
            _cx             =   4207
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
            BackColor       =   12648447
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "œ·«·«  «·«·Ê«‰"
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
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               BackColor       =   &H000000FF&
               Caption         =   "Õ—Ã"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   252
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   252
               Top             =   240
               Width           =   1332
            End
         End
         Begin VB.TextBox txt_opr_total 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   8940
            Locked          =   -1  'True
            TabIndex        =   144
            Top             =   4740
            Width           =   2955
         End
         Begin VB.TextBox TXTNoOFWeek 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   5595
            Locked          =   -1  'True
            TabIndex        =   143
            Top             =   4740
            Width           =   2280
         End
         Begin ALLButtonS.ALLButton terms_operations 
            Height          =   360
            Index           =   1
            Left            =   2865
            TabIndex        =   145
            Top             =   4740
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   635
            BTYPE           =   3
            TX              =   "—ÃÊ⁄ ··»‰Êœ"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   16711680
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":878EB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton opr_items 
            Height          =   315
            Index           =   0
            Left            =   16050
            TabIndex        =   146
            Top             =   2625
            Visible         =   0   'False
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "„Ê«œ "
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   16711680
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":87907
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton employee_details 
            Height          =   315
            Left            =   12240
            TabIndex        =   147
            Top             =   2625
            Visible         =   0   'False
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "»Ì«‰«  «·⁄„«·…"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   16711680
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":87923
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton opr_Expenses 
            Height          =   315
            Index           =   0
            Left            =   10335
            TabIndex        =   148
            Top             =   2625
            Visible         =   0   'False
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "„’«—Ì›"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   16711680
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":8793F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton CMDViewGantt 
            Height          =   360
            Index           =   2
            Left            =   360
            TabIndex        =   149
            Top             =   4740
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   635
            BTYPE           =   3
            TX              =   "⁄—÷ «·Ã«‰ "
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   16711680
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":8795B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton opr_Expenses 
            Height          =   315
            Index           =   2
            Left            =   14115
            TabIndex        =   150
            Top             =   2625
            Visible         =   0   'False
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "«·„⁄œ« "
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   16711680
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":87977
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton CmdProcees 
            Height          =   255
            Left            =   17355
            TabIndex        =   151
            Top             =   4620
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   450
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "projects.frx":87993
            ColorButton     =   12648447
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
            Height          =   3870
            Left            =   120
            TabIndex        =   152
            Top             =   600
            Width           =   18435
            _cx             =   32517
            _cy             =   6826
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
            Cols            =   44
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"projects.frx":87F2D
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   -1  'True
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
         Begin XtremeSuiteControls.CheckBox ChAutoItems 
            Height          =   375
            Left            =   16800
            TabIndex        =   296
            Top             =   5040
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   " Õ„Ì· «·„Ê«œ «·Ì«"
            BackColor       =   12632064
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«”»Ê⁄"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   360
            TabIndex        =   158
            Top             =   705
            Width           =   1095
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«Ã„«·Ì «· ﬂ·›…"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   12000
            TabIndex        =   155
            Top             =   4740
            Width           =   975
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«Ã„«·Ì «·„œ…"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   7980
            TabIndex        =   154
            Top             =   4860
            Width           =   975
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   18195
            TabIndex        =   153
            Top             =   210
            Width           =   255
         End
      End
      Begin C1SizerLibCtl.C1Elastic Fra 
         Height          =   5190
         Index           =   2
         Left            =   120
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   3360
         Width           =   18705
         _cx             =   32994
         _cy             =   9155
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
         Caption         =   "»Ì«‰«  «·œ›⁄« "
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
         Begin VB.TextBox Text12 
            Height          =   735
            Left            =   240
            TabIndex        =   109
            Top             =   375
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VSFlex8Ctl.VSFlexGrid GridSub 
            Height          =   4080
            Left            =   -120
            TabIndex        =   111
            Top             =   735
            Width           =   18765
            _cx             =   33099
            _cy             =   7197
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
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   25
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"projects.frx":8858A
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
         Begin ImpulseButton.ISButton Cmdd 
            Height          =   135
            Left            =   17370
            TabIndex        =   112
            Top             =   4950
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   238
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
            ButtonImage     =   "projects.frx":88963
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   18090
            TabIndex        =   110
            Top             =   375
            Width           =   495
         End
      End
      Begin C1SizerLibCtl.C1Elastic Fra 
         Height          =   5235
         Index           =   3
         Left            =   120
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   3060
         Width           =   18705
         _cx             =   32994
         _cy             =   9234
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
         ForeColor       =   192
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "»Ì«‰«  „Õ«”»Ì…"
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
         Begin C1SizerLibCtl.C1Elastic Fra 
            Height          =   3630
            Index           =   7
            Left            =   720
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   735
            Width           =   17865
            _cx             =   31512
            _cy             =   6403
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
            AccessibleName  =   "&H00E2E9E9&"
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.TextBox txtopening_balance_voucher_id 
               Height          =   525
               Left            =   0
               TabIndex        =   106
               Top             =   0
               Visible         =   0   'False
               Width           =   855
            End
            Begin C1SizerLibCtl.C1Elastic Fra 
               Height          =   1515
               Index           =   0
               Left            =   11160
               TabIndex        =   98
               TabStop         =   0   'False
               Top             =   1995
               Width           =   3105
               _cx             =   5477
               _cy             =   2672
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
               Caption         =   "«·—’Ìœ «·√›  «ÕÏ „” Õ·’« "
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
               Begin VB.TextBox TxtOpenBalance4 
                  Alignment       =   2  'Center
                  Height          =   360
                  Left            =   180
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   645
                  Width           =   1545
               End
               Begin VB.OptionButton OptType4 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   1950
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   375
                  Value           =   -1  'True
                  Width           =   945
               End
               Begin VB.OptionButton OptType4 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ«∆‰"
                  Height          =   255
                  Index           =   1
                  Left            =   1020
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   375
                  Width           =   885
               End
               Begin VB.OptionButton OptType4 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   375
                  Width           =   915
               End
               Begin MSComCtl2.DTPicker Dtp4 
                  Height          =   360
                  Left            =   240
                  TabIndex        =   105
                  Top             =   1110
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   635
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   93519875
                  CurrentDate     =   38718
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   360
                  Index           =   8
                  Left            =   1770
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   675
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ”ÃÌ·"
                  Height          =   345
                  Index           =   9
                  Left            =   1770
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   1080
                  Width           =   1125
               End
            End
            Begin C1SizerLibCtl.C1Elastic Fra 
               Height          =   1515
               Index           =   1
               Left            =   14370
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   1995
               Width           =   3255
               _cx             =   5741
               _cy             =   2672
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
               Caption         =   "«·—’Ìœ «·√›  «ÕÏ  ··«ÃÊ—"
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
               Begin VB.TextBox TxtOpenBalance3 
                  Alignment       =   2  'Center
                  Height          =   330
                  Left            =   270
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   570
                  Width           =   1455
               End
               Begin VB.OptionButton OptType3 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   240
                  Index           =   0
                  Left            =   1935
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   330
                  Value           =   -1  'True
                  Width           =   870
               End
               Begin VB.OptionButton OptType3 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ«∆‰"
                  Height          =   240
                  Index           =   1
                  Left            =   1050
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   330
                  Width           =   840
               End
               Begin VB.OptionButton OptType3 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   240
                  Index           =   2
                  Left            =   105
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   330
                  Width           =   960
               End
               Begin MSComCtl2.DTPicker Dtp3 
                  Height          =   330
                  Left            =   270
                  TabIndex        =   95
                  Top             =   960
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   582
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   93519875
                  CurrentDate     =   38718
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   330
                  Index           =   10
                  Left            =   1770
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   600
                  Width           =   1035
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ”ÃÌ·"
                  Height          =   300
                  Index           =   11
                  Left            =   1770
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   1035
                  Width           =   1035
               End
            End
            Begin C1SizerLibCtl.C1Elastic Fra 
               Height          =   1680
               Index           =   10
               Left            =   7620
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   225
               Width           =   3330
               _cx             =   5874
               _cy             =   2963
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
               Caption         =   "«·—’Ìœ «·√›  «ÕÏ „Ê«œ"
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
               Begin VB.TextBox TxtOpenBalance2 
                  Alignment       =   2  'Center
                  Height          =   435
                  Left            =   420
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   630
                  Width           =   1530
               End
               Begin VB.OptionButton OptType2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   240
                  Index           =   0
                  Left            =   2145
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   390
                  Value           =   -1  'True
                  Width           =   945
               End
               Begin VB.OptionButton OptType2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ«∆‰"
                  Height          =   240
                  Index           =   1
                  Left            =   1260
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   390
                  Width           =   870
               End
               Begin VB.OptionButton OptType2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   240
                  Index           =   2
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   390
                  Width           =   915
               End
               Begin MSComCtl2.DTPicker Dtp2 
                  Height          =   360
                  Left            =   420
                  TabIndex        =   87
                  Top             =   1125
                  Width           =   1545
                  _ExtentX        =   2725
                  _ExtentY        =   635
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   93519875
                  CurrentDate     =   38718
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   420
                  Index           =   18
                  Left            =   1995
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   660
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ”ÃÌ·"
                  Height          =   360
                  Index           =   17
                  Left            =   1995
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   1125
                  Width           =   1125
               End
            End
            Begin C1SizerLibCtl.C1Elastic Fra 
               Height          =   1680
               Index           =   9
               Left            =   11055
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   225
               Width           =   3210
               _cx             =   5662
               _cy             =   2963
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
               Caption         =   "«·—’Ìœ «·√›  «ÕÏ «Ì—«œ« "
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
               Begin VB.OptionButton OptType1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   240
                  Index           =   2
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   390
                  Width           =   915
               End
               Begin VB.OptionButton OptType1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ«∆‰"
                  Height          =   240
                  Index           =   1
                  Left            =   1260
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   390
                  Width           =   870
               End
               Begin VB.OptionButton OptType1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   240
                  Index           =   0
                  Left            =   2055
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   390
                  Value           =   -1  'True
                  Width           =   945
               End
               Begin VB.TextBox TxtOpenBalance1 
                  Alignment       =   2  'Center
                  Height          =   435
                  Left            =   270
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   630
                  Width           =   1530
               End
               Begin MSComCtl2.DTPicker Dtp1 
                  Height          =   360
                  Left            =   300
                  TabIndex        =   79
                  Top             =   1125
                  Width           =   1545
                  _ExtentX        =   2725
                  _ExtentY        =   635
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   93519875
                  CurrentDate     =   38718
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ”ÃÌ·"
                  Height          =   360
                  Index           =   16
                  Left            =   1875
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   1125
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   420
                  Index           =   15
                  Left            =   1875
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   660
                  Width           =   1125
               End
            End
            Begin C1SizerLibCtl.C1Elastic Fra 
               Height          =   1680
               Index           =   8
               Left            =   14370
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   225
               Width           =   3255
               _cx             =   5741
               _cy             =   2963
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
               Caption         =   "«·—’Ìœ «·√›  «ÕÏ „’—Ê›« "
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
               Begin VB.TextBox TxtOpenBalance 
                  Alignment       =   2  'Center
                  Height          =   390
                  Left            =   270
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   555
                  Width           =   1455
               End
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   210
                  Index           =   0
                  Left            =   1935
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   345
                  Value           =   -1  'True
                  Width           =   870
               End
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ«∆‰"
                  Height          =   210
                  Index           =   1
                  Left            =   1050
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   345
                  Width           =   840
               End
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   210
                  Index           =   2
                  Left            =   105
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   345
                  Width           =   960
               End
               Begin MSComCtl2.DTPicker Dtp 
                  Height          =   270
                  Left            =   270
                  TabIndex        =   71
                  Top             =   1005
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   476
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   93519875
                  CurrentDate     =   38718
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   375
                  Index           =   14
                  Left            =   1770
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   585
                  Width           =   1035
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ”ÃÌ·"
                  Height          =   270
                  Index           =   13
                  Left            =   1770
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   1005
                  Width           =   1035
               End
            End
            Begin C1SizerLibCtl.C1Elastic Fra 
               Height          =   1515
               Index           =   4
               Left            =   7620
               TabIndex        =   298
               TabStop         =   0   'False
               Top             =   1995
               Width           =   3330
               _cx             =   5874
               _cy             =   2672
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
               Caption         =   "«·—’Ìœ «·√›  «ÕÏ  Õ  «· ‰›Ì–"
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
               Begin VB.OptionButton OptType5 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   255
                  Index           =   2
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   302
                  Top             =   270
                  Width           =   975
               End
               Begin VB.OptionButton OptType5 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ«∆‰"
                  Height          =   255
                  Index           =   1
                  Left            =   1230
                  RightToLeft     =   -1  'True
                  TabIndex        =   301
                  Top             =   270
                  Width           =   945
               End
               Begin VB.OptionButton OptType5 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   2085
                  RightToLeft     =   -1  'True
                  TabIndex        =   300
                  Top             =   270
                  Value           =   -1  'True
                  Width           =   1020
               End
               Begin VB.TextBox TxtOpenBalance5 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   195
                  RightToLeft     =   -1  'True
                  TabIndex        =   299
                  Top             =   540
                  Width           =   1650
               End
               Begin MSComCtl2.DTPicker Dtp5 
                  Height          =   360
                  Left            =   180
                  TabIndex        =   303
                  Top             =   1005
                  Width           =   1665
                  _ExtentX        =   2937
                  _ExtentY        =   635
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   93519875
                  CurrentDate     =   38718
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ”ÃÌ·"
                  Height          =   330
                  Index           =   24
                  Left            =   1905
                  RightToLeft     =   -1  'True
                  TabIndex        =   305
                  Top             =   975
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   360
                  Index           =   23
                  Left            =   1905
                  RightToLeft     =   -1  'True
                  TabIndex        =   304
                  Top             =   570
                  Width           =   1200
               End
            End
            Begin C1SizerLibCtl.C1Elastic Fra 
               Height          =   1680
               Index           =   5
               Left            =   4200
               TabIndex        =   340
               TabStop         =   0   'False
               Top             =   240
               Width           =   3330
               _cx             =   5874
               _cy             =   2963
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
               Caption         =   "«·—’Ìœ «·√›  «ÕÏ ·Õ”‰ «·«œ«¡"
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
               Begin VB.OptionButton OptType6 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   240
                  Index           =   2
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   344
                  Top             =   390
                  Width           =   915
               End
               Begin VB.OptionButton OptType6 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ«∆‰"
                  Height          =   240
                  Index           =   1
                  Left            =   1260
                  RightToLeft     =   -1  'True
                  TabIndex        =   343
                  Top             =   390
                  Width           =   870
               End
               Begin VB.OptionButton OptType6 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   240
                  Index           =   0
                  Left            =   2145
                  RightToLeft     =   -1  'True
                  TabIndex        =   342
                  Top             =   390
                  Value           =   -1  'True
                  Width           =   945
               End
               Begin VB.TextBox TxtOpenBalance6 
                  Alignment       =   2  'Center
                  Height          =   435
                  Left            =   420
                  RightToLeft     =   -1  'True
                  TabIndex        =   341
                  Top             =   630
                  Width           =   1530
               End
               Begin MSComCtl2.DTPicker Dtp6 
                  Height          =   360
                  Left            =   420
                  TabIndex        =   345
                  Top             =   1125
                  Width           =   1545
                  _ExtentX        =   2725
                  _ExtentY        =   635
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   93519875
                  CurrentDate     =   38718
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ”ÃÌ·"
                  Height          =   360
                  Index           =   44
                  Left            =   1995
                  RightToLeft     =   -1  'True
                  TabIndex        =   347
                  Top             =   1125
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   420
                  Index           =   43
                  Left            =   1995
                  RightToLeft     =   -1  'True
                  TabIndex        =   346
                  Top             =   660
                  Width           =   1125
               End
            End
            Begin C1SizerLibCtl.C1Elastic Fra 
               Height          =   1515
               Index           =   6
               Left            =   4200
               TabIndex        =   348
               TabStop         =   0   'False
               Top             =   1920
               Width           =   3330
               _cx             =   5874
               _cy             =   2672
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
               Caption         =   "«·—’Ìœ «·√›  «ÕÏ ·œ›⁄«  «·„ﬁœ„…"
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
               Begin VB.TextBox TxtOpenBalance8 
                  Alignment       =   2  'Center
                  Height          =   360
                  Left            =   180
                  RightToLeft     =   -1  'True
                  TabIndex        =   355
                  Top             =   600
                  Width           =   1665
               End
               Begin VB.OptionButton OptType8 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   2085
                  RightToLeft     =   -1  'True
                  TabIndex        =   351
                  Top             =   270
                  Value           =   -1  'True
                  Width           =   1020
               End
               Begin VB.OptionButton OptType8 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ«∆‰"
                  Height          =   255
                  Index           =   1
                  Left            =   1110
                  RightToLeft     =   -1  'True
                  TabIndex        =   350
                  Top             =   270
                  Width           =   945
               End
               Begin VB.OptionButton OptType8 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   255
                  Index           =   2
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   349
                  Top             =   270
                  Width           =   975
               End
               Begin MSComCtl2.DTPicker Dtp8 
                  Height          =   360
                  Left            =   180
                  TabIndex        =   354
                  Top             =   1005
                  Width           =   1665
                  _ExtentX        =   2937
                  _ExtentY        =   635
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   93519875
                  CurrentDate     =   38718
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ”ÃÌ·"
                  Height          =   360
                  Index           =   46
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   353
                  Top             =   1080
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   420
                  Index           =   45
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   352
                  Top             =   600
                  Width           =   1125
               End
            End
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   18330
            TabIndex        =   107
            Top             =   375
            Width           =   375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Frame5 
         Height          =   5175
         Left            =   120
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   3885
         Width           =   18705
         _cx             =   32994
         _cy             =   9128
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
         Caption         =   "«·»‰Êœ"
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
         Begin VB.CommandButton Command6 
            BackColor       =   &H000000FF&
            Caption         =   "X"
            Height          =   255
            Left            =   17640
            Style           =   1  'Graphical
            TabIndex        =   310
            Top             =   120
            Width           =   375
         End
         Begin XtremeSuiteControls.CheckBox ChAuto 
            Height          =   375
            Left            =   16920
            TabIndex        =   295
            Top             =   4680
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   " Õ„Ì· «·⁄„·Ì«  «·Ì«"
            BackColor       =   12632064
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.TextBox txt_sub_net 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   8400
            Locked          =   -1  'True
            TabIndex        =   138
            Top             =   4695
            Width           =   1095
         End
         Begin VB.TextBox txt_total_sum 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   345
            Left            =   11505
            Locked          =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   137
            Top             =   4695
            Width           =   1350
         End
         Begin VB.TextBox txt_sub_discount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   9600
            Locked          =   -1  'True
            TabIndex        =   136
            Top             =   4695
            Width           =   1320
         End
         Begin C1SizerLibCtl.C1Elastic Frame17 
            Height          =   930
            Left            =   120
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   450
            Width           =   18465
            _cx             =   32570
            _cy             =   1640
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
            Begin C1SizerLibCtl.C1Elastic Frame19 
               Height          =   930
               Left            =   0
               TabIndex        =   119
               TabStop         =   0   'False
               Top             =   0
               Width           =   13110
               _cx             =   23125
               _cy             =   1640
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
               Begin VB.TextBox Text21 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   11370
                  TabIndex        =   122
                  Top             =   405
                  Width           =   1635
               End
               Begin VB.TextBox Text22 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   9570
                  TabIndex        =   121
                  Top             =   405
                  Width           =   1695
               End
               Begin VB.TextBox Text25 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   4260
                  TabIndex        =   120
                  Top             =   405
                  Width           =   1635
               End
               Begin MSDataListLib.DataCombo DataCombo2 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   123
                  Top             =   405
                  Width           =   3795
                  _ExtentX        =   6694
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DTPicker1 
                  Height          =   300
                  Left            =   7950
                  TabIndex        =   124
                  Top             =   405
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   529
                  _Version        =   393216
                  Format          =   93519873
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker DTPicker3 
                  Height          =   300
                  Left            =   6390
                  TabIndex        =   125
                  Top             =   405
                  Width           =   1320
                  _ExtentX        =   2328
                  _ExtentY        =   529
                  _Version        =   393216
                  Format          =   93519873
                  CurrentDate     =   38784
               End
               Begin VB.Label Label51 
                  Alignment       =   2  'Center
                  Caption         =   "—ﬁ„ «·÷„«‰"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   11595
                  TabIndex        =   131
                  Top             =   120
                  Width           =   1050
               End
               Begin VB.Label Label52 
                  Alignment       =   2  'Center
                  Caption         =   "ﬁÌ„… «·÷„«‰"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   9690
                  TabIndex        =   130
                  Top             =   120
                  Width           =   1065
               End
               Begin VB.Label Label53 
                  Alignment       =   2  'Center
                  Caption         =   " «—ÌŒ »œ«Ì… «·÷„«‰"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   7830
                  TabIndex        =   129
                  Top             =   120
                  Width           =   1515
               End
               Begin VB.Label Label54 
                  Alignment       =   2  'Center
                  Caption         =   "‰Â«Ì… «·÷„«‰"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   6585
                  TabIndex        =   128
                  Top             =   120
                  Width           =   1005
               End
               Begin VB.Label Label56 
                  Alignment       =   2  'Center
                  Caption         =   "»‰ﬂ «·÷„«‰"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   1740
                  TabIndex        =   127
                  Top             =   120
                  Width           =   1530
               End
               Begin VB.Label Label55 
                  Alignment       =   2  'Center
                  Caption         =   "«· „œÌœ"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   4260
                  TabIndex        =   126
                  Top             =   120
                  Width           =   1515
               End
            End
            Begin C1SizerLibCtl.C1Elastic Frame18 
               Height          =   1050
               Left            =   13110
               TabIndex        =   115
               TabStop         =   0   'False
               Top             =   -120
               Width           =   5475
               _cx             =   9657
               _cy             =   1852
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
               Begin VB.OptionButton check4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„ﬁ«Ê· »«ÿ‰"
                  Height          =   345
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   555
                  Width           =   1290
               End
               Begin VB.OptionButton company 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·‘—ﬂ…"
                  Height          =   255
                  Left            =   3060
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   300
                  Width           =   1230
               End
               Begin MSDataListLib.DataCombo DataCombo3 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   118
                  Top             =   555
                  Width           =   2895
                  _ExtentX        =   5106
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Label57 
                  Alignment       =   2  'Center
                  Caption         =   "«·ﬁ«∆„ »«·«⁄„«·"
                  ForeColor       =   &H00C00000&
                  Height          =   315
                  Left            =   4245
                  TabIndex        =   266
                  Top             =   225
                  Width           =   1155
               End
            End
         End
         Begin ALLButtonS.ALLButton terms_operations 
            Height          =   360
            Index           =   0
            Left            =   15570
            TabIndex        =   139
            Top             =   4695
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   635
            BTYPE           =   3
            TX              =   "⁄„·Ì«  «·»‰œ"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16711680
            BCOLO           =   12582912
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":88EFD
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton CmdPand 
            Height          =   360
            Left            =   13650
            TabIndex        =   140
            Top             =   4695
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   635
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
            ButtonImage     =   "projects.frx":88F19
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   4680
            TabIndex        =   271
            Top             =   4770
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ALLButtonS.ALLButton Add 
            Height          =   360
            Left            =   14400
            TabIndex        =   293
            Top             =   4695
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   635
            BTYPE           =   3
            TX              =   "«œ—«Ã ”ÿ—"
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
            BCOL            =   65280
            BCOLO           =   65280
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "projects.frx":894B3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
            Height          =   3285
            Left            =   120
            TabIndex        =   132
            Top             =   1320
            Width           =   18555
            _cx             =   32729
            _cy             =   5794
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
            Cols            =   29
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"projects.frx":894CF
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   -1  'True
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
            Begin VB.PictureBox PicDes 
               BorderStyle     =   0  'None
               Height          =   1635
               Left            =   0
               RightToLeft     =   -1  'True
               ScaleHeight     =   1635
               ScaleWidth      =   10485
               TabIndex        =   133
               Top             =   5280
               Visible         =   0   'False
               Width           =   10485
               Begin VB.TextBox TxtDes 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  Height          =   1125
                  Left            =   30
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   134
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
                  TabIndex        =   135
                  Top             =   0
                  Width           =   4485
               End
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   270
            Index           =   22
            Left            =   7440
            TabIndex        =   272
            Top             =   4770
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   315
            Index           =   21
            Left            =   1380
            TabIndex        =   270
            Top             =   4770
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   315
            Index           =   20
            Left            =   3300
            TabIndex        =   269
            Top             =   4770
            Width           =   1185
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   2520
            TabIndex        =   268
            Top             =   4770
            Width           =   735
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   240
            TabIndex        =   267
            Top             =   4770
            Width           =   1095
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            Caption         =   "«·«Ã„«·Ì"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   12720
            TabIndex        =   141
            Top             =   4695
            Width           =   930
         End
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   $"projects.frx":89930
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
         Height          =   615
         Index           =   5
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   314
         Top             =   0
         Width           =   5040
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   7560
         Picture         =   "projects.frx":899D5
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "     »Ì«‰«  «·„‘«—Ì⁄      "
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
         Height          =   705
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   18825
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   312
      Left            =   21480
      TabIndex        =   221
      Top             =   360
      Visible         =   0   'False
      Width           =   1308
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Format          =   93519873
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DcAccount1 
      Height          =   312
      Left            =   23280
      TabIndex        =   224
      Top             =   3000
      Visible         =   0   'False
      Width           =   1212
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcAccount3 
      Height          =   312
      Left            =   23160
      TabIndex        =   226
      Top             =   3360
      Visible         =   0   'False
      Width           =   1212
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   276
      Index           =   0
      Left            =   20160
      TabIndex        =   232
      Top             =   480
      Width           =   696
      _ExtentX        =   1217
      _ExtentY        =   476
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
      ButtonImage     =   "projects.frx":8D63D
      DrawFocusRectangle=   0   'False
   End
   Begin ALLButtonS.ALLButton opr_Expenses 
      Height          =   372
      Index           =   3
      Left            =   19680
      TabIndex        =   237
      Top             =   3600
      Width           =   2292
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "—ÃÊ⁄ ··⁄„·Ì« "
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   192
      MPTR            =   1
      MICON           =   "projects.frx":8DBD7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid4 
      Height          =   1500
      Left            =   20400
      TabIndex        =   242
      Top             =   4200
      Width           =   9960
      _cx             =   17568
      _cy             =   2646
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
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"projects.frx":8DBF3
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
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«Ã„«·Ì« "
      Height          =   252
      Index           =   7
      Left            =   18720
      RightToLeft     =   -1  'True
      TabIndex        =   238
      Top             =   2520
      Width           =   2532
   End
   Begin VB.Label Label7 
      Caption         =   "Label2"
      Height          =   372
      Left            =   20880
      TabIndex        =   236
      Top             =   3360
      Width           =   852
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   22200
      TabIndex        =   235
      Top             =   3960
      Width           =   372
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   20520
      TabIndex        =   233
      Top             =   0
      Width           =   372
   End
   Begin VB.Label Label25 
      Caption         =   "Label25"
      Height          =   372
      Left            =   21360
      TabIndex        =   231
      Top             =   2640
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   492
      Left            =   21360
      TabIndex        =   230
      Top             =   2160
      Width           =   852
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      Caption         =   "„”·”· «·„‘—Ê⁄"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   23160
      TabIndex        =   229
      Top             =   2640
      Width           =   1212
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·⁄„Ì·"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   22920
      TabIndex        =   228
      Top             =   4080
      Width           =   1332
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·⁄„Ì·"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   22920
      TabIndex        =   227
      Top             =   3720
      Width           =   1332
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   " «ﬁ—» ‰Â«Ì…"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   20880
      TabIndex        =   223
      Top             =   1200
      Visible         =   0   'False
      Width           =   852
   End
End
Attribute VB_Name = "Projects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As String
Public FlgOper As Boolean
Dim last_root As Integer
Dim last_geeral As Integer
Dim last_branch As Integer
Dim mod_flad As String
Dim first_run  As Boolean
Dim Fullcode As String
Dim test1 As Boolean
Dim rs As ADODB.Recordset
Dim RsDev As ADODB.Recordset
Public Monthly As Integer
Dim RsDevsub As ADODB.Recordset
Dim NewGrid As New ClsGrid
Dim RSTransDetails As ADODB.Recordset
Dim RsDetails As ADODB.Recordset
Dim current_terms As String
Public ProjectDes_ID As Integer
Dim current_opr As String
Public LngRow As Double
Public LngCol As Double
Public showAll As Boolean
Dim Pand As Double
Dim Account_Code_dynamic6 As String
Dim Account_Code_dynamic1 As String
Dim Account_Code_dynamic2 As String
Dim Account_Code_dynamic3 As String
Dim Account_Code_dynamic4 As String
Dim Account_Code_dynamic5 As String
Dim Account_Code_dynamic7 As String
Dim Account_Code_dynamic1C As String
Dim Account_Code_dynamic2C As String
Dim Account_Code_dynamic3C As String
Dim Account_Code_dynamic4C As String
Dim Account_Code_dynamic5C As String
Dim Account_Code_dynamic6C As String
Dim Account_Code_dynamic7C As String
Function print_report1(Optional ProjectID As Double = 0, Optional pandid As Double = 0)
   
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.projects_des.oprid, dbo.projects_des.project_no, dbo.projects_des.project_name, dbo.projects_des.[index], dbo.projects_des.des, dbo.projects_des.qty, "
MySQL = MySQL & "                      dbo.projects_des.cost, dbo.projects_des.total, dbo.projects_des.discount, dbo.projects_des.net, dbo.projects_des.project_id, dbo.projects_des.line_no,"
MySQL = MySQL & "                      dbo.projects_des.sub_contractor_id, dbo.projects_des.fullcode, dbo.projects_des.Remark, dbo.terms_operations.total AS Oprtotal, dbo.terms_operations.name,"
MySQL = MySQL & "                      dbo.terms_operations.period, dbo.terms_operations.term_fullcode, dbo.terms_operations.project_id AS Operproject_id, dbo.terms_operations.[count],"
MySQL = MySQL & "                      dbo.terms_operations.salary, dbo.terms_operations.total_items, dbo.terms_operations.total_salary, dbo.terms_operations.total_expenses,"
MySQL = MySQL & "                      dbo.terms_operations.fullcode AS OperFullcodes, dbo.terms_operations.ended, dbo.terms_operations.start_date, dbo.terms_operations.end_date,"
MySQL = MySQL & "                      dbo.terms_operations.StartWeek, dbo.terms_operations.EndWeek, dbo.terms_operations.EarlyEndWeek, dbo.terms_operations.EarlyStartWeek,"
MySQL = MySQL & "                      dbo.terms_operations.Period1, dbo.terms_operations.Critical, dbo.terms_operations.Symbol, dbo.terms_operations.Pre, dbo.terms_operations.EarlyStartDate,"
MySQL = MySQL & "                      dbo.terms_operations.EarlyEndDate, dbo.terms_operations.qty AS Operqty, dbo.terms_operations.periodView, dbo.terms_operations.Actperiod,"
MySQL = MySQL & "                      dbo.terms_operations.unitid, dbo.terms_operations.unitname, dbo.terms_operations.ProjectDes_ID, dbo.terms_operations.expen, dbo.terms_operations.eq,"
MySQL = MySQL & "                      dbo.terms_operations.emps, dbo.terms_operations.matrials, dbo.terms_operations.EquepVal, dbo.terms_operations.hourval, dbo.terms_operations.item_id,"
MySQL = MySQL & "                      dbo.terms_operations.OPRIDD, dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE,"
MySQL = MySQL & "                      dbo.TblProcessDEF.UnitID AS UUnitID, dbo.TblProcessUnites.UnitName AS UUnitName, dbo.TblProcessUnites.UnitNamee, dbo.TblEmpOper.ID AS EmpOID,"
MySQL = MySQL & "                      dbo.TblEmpOper.ProjectID AS EmpOProjectID, dbo.TblEmpOper.Pand AS EmpOPand, dbo.TblEmpOper.Opr AS EmpOOpr, dbo.TblEmpOper.daysalary,"
MySQL = MySQL & "                      dbo.TblEmpOper.[Count] AS EmpOCount, dbo.TblEmpOper.OperCode, dbo.TblEmpOper.EmpID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4,"
MySQL = MySQL & "                      dbo.TblEmployee.Fullcode AS EmpOFullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblEmpOper.JobID, dbo.TblEmpJobsTypes.JobTypeName,"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblExpensiveOper.ID AS ExID, dbo.TblExpensiveOper.ProjectID AS ExProjectID, dbo.TblExpensiveOper.Pand AS ExPand,"
MySQL = MySQL & "                      dbo.TblExpensiveOper.Opr AS ExOpr, dbo.TblExpensiveOper.EsToal, dbo.TblExpensiveOper.[value], dbo.TblExpensiveOper.Des AS ExDes,"
MySQL = MySQL & "                      dbo.TblExpensiveOper.OperCode AS ExOperCode, dbo.TblExpensiveOper.AccountCode, dbo.Expenses_accounts.Account_Name, dbo.terms_operations.id,"
MySQL = MySQL & "                      dbo.TblEquepment.ID AS EqID, dbo.TblEquepment.ProjectID AS EqProjectID, dbo.TblEquepment.Pand AS EqPand, dbo.TblEquepment.Opr AS EqOpr,"
MySQL = MySQL & "                      dbo.TblEquepment.EstHour, dbo.TblEquepment.ActualHour, dbo.TblEquepment.TotalEs, dbo.TblEquepment.[value] AS Eqvalue, dbo.TblEquepment.des AS Eqdes,"
MySQL = MySQL & "                      dbo.TblEquepment.OperCode AS EqOperCode, dbo.TblEquepment.EquepVal AS EqEquepVal, dbo.TblEquepment.ExpensesID, dbo.FixedAssets.code,"
MySQL = MySQL & "                      dbo.FixedAssets.Name AS EqName, dbo.FixedAssets.namee AS EqNameE, dbo.TblMatrials.ID AS MatID, dbo.TblMatrials.Pand AS MatPand,"
MySQL = MySQL & "                      dbo.TblMatrials.[Count] AS MatCount, dbo.TblMatrials.Price, dbo.TblMatrials.Quntapro, dbo.TblMatrials.priceapro, dbo.TblMatrials.ProjectID AS MatProjectID,"
MySQL = MySQL & "                      dbo.TblMatrials.OperCode AS MatOperCode, dbo.TblMatrials.Opr AS MatOpr, dbo.TblMatrials.ItemID, dbo.TblItems.Fullcode AS ItemFullcode, dbo.TblItems.ItemCode,"
MySQL = MySQL & "                      dbo.TblItems.ItemName , dbo.TblItems.ItemNamee"
MySQL = MySQL & " FROM         dbo.TblItems RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblMatrials ON dbo.TblItems.ItemID = dbo.TblMatrials.ItemID RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.terms_operations ON dbo.TblMatrials.Opr = dbo.terms_operations.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEquepment LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets ON dbo.TblEquepment.ExpensesID = dbo.FixedAssets.id ON dbo.terms_operations.id = dbo.TblEquepment.Opr LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.Expenses_accounts RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblExpensiveOper ON dbo.Expenses_accounts.Account_Code = dbo.TblExpensiveOper.AccountCode ON"
MySQL = MySQL & "                      dbo.terms_operations.id = dbo.TblExpensiveOper.Opr LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpOper LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.TblEmpOper.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblEmpOper.EmpID = dbo.TblEmployee.Emp_ID ON dbo.terms_operations.id = dbo.TblEmpOper.Opr LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblProcessUnites LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblProcessDEF ON dbo.TblProcessUnites.UnitID = dbo.TblProcessDEF.UnitID ON"
MySQL = MySQL & "                      dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.projects_des ON dbo.terms_operations.ProjectDes_ID = dbo.projects_des.oprid"
MySQL = MySQL & "  Where (dbo.projects_des.oprid = " & pandid & ") And (dbo.projects_des.project_id = " & ProjectID & ")"



        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessOfPand1.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessOfPand1E.rpt"
            
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
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
Function ChekID(Optional ByRef Name As String) As Boolean
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim sql As String
ChekID = False

If Me.DCPreFix.Text & Me.TxtId.Text <> "" Then
sql = "select Fullcode, Project_name,Project_nameE from  projects where ID<>" & val(txt_project_id.Text) & " and Fullcode='" & Me.DCPreFix.Text & Me.TxtId.Text & " ' "
End If
If sql <> "" Then
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
Name = IIf(IsNull(Rs5("Project_name").value), "", Rs5("Project_name").value)
Else
Name = IIf(IsNull(Rs5("Project_nameE").value), "", Rs5("Project_nameE").value)
End If
ChekID = True
Else
ChekID = False
End If
End If
End Function
Sub ReculcalteDesMain()
  Dim CuntNo As Double
          Dim Pandcode As String
          Dim Total As Double
          Dim count1 As Double
          Dim Count2 As Double
          Dim QtyNo As Double
          Dim Row As Long
          Dim SumValEx As Double
          ReLineGrid
          With FgMainDes
          CuntNo = 0
        SumValEx = 0
        For Row = 1 To .Rows
        If val(.TextMatrix(Row, .ColIndex("ID"))) <> 0 Then
         CuntNo = GetCount(val(.TextMatrix(Row, .ColIndex("ID"))), Total, SumValEx, count1, Count2, QtyNo)
         SetValue val(.TextMatrix(Row, .ColIndex("ID"))), Total, SumValEx, count1, Count2, QtyNo
         End If
        Next Row
        End With
        ReLineGrid
End Sub
Private Sub SaveData()
On Error Resume Next
 Dim LngDevID As Long
            Dim LngOpenID As Long
Dim RsDev1 As ADODB.Recordset
ReculcalteDesMain
    If SystemOptions.UserInterface = EnglishInterface Then

        If DcAccount2.BoundText = "" Then MsgBox "Must Specify Client Name", vbCritical: Exit Sub
        If DcCurrency.BoundText = "" Then
            MsgBox "Must Specify  Currency Name", vbCritical
            DcCurrency.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If DcAccount2.BoundText = "" Then MsgBox "Must Specify client Name", vbCritical: Exit Sub
        If TXTprojectname.Text = "" Then MsgBox "Must Specify project Name", vbCritical: Exit Sub

        If Not IsNumeric(TxtProjectCosts.Text) Then MsgBox "Must Specify project cost", vbCritical: Exit Sub
  
        If Not IsNumeric(txt_total_discount.Text) Then MsgBox "discount must be numeric", vbCritical:  txt_total_discount.Text = 0: Exit Sub
        If Not IsNumeric(dcBranch.BoundText) Then MsgBox "Must select Branch", vbCritical: Exit Sub
        If DataCombo1.BoundText = "" Then MsgBox "Must Specify project Status", vbCritical: Exit Sub
    Else

   '     If DCAccount1.BoundText = "" Then MsgBox "·«»œ „‰  ÕœÌœ «”„ „ﬁ«Ê· «·»«ÿ‰  ", vbCritical: Exit Sub
        If DcAccount2.BoundText = "" Then MsgBox "·«»œ „‰  ÕœÌœ «”„ «·⁄„Ì· «·‰Â«∆Ì", vbCritical: Exit Sub
        If TXTprojectname.Text = "" Then MsgBox "·«»œ „‰  ÕœÌœ «”„   «·„‘—Ê⁄", vbCritical: Exit Sub

        If Not IsNumeric(TxtProjectCosts.Text) Then MsgBox "ÌÃ»  ÕœÌœ ﬁÌ„… «·„‘—Ê⁄", vbCritical: Exit Sub
  
        If Not IsNumeric(txt_total_discount.Text) Then MsgBox "·«»œ „‰  ÕœÌœ «·Œ’„", vbCritical:  txt_total_discount.Text = 0: Exit Sub
        If DcCurrency.BoundText = "" Then
            MsgBox "Õœœ «·⁄„·… «Ê·«", vbCritical
            DcCurrency.SetFocus
           SendKeys "{F4}"
            Exit Sub
        End If

        If Not IsNumeric(dcBranch.BoundText) Then MsgBox "Õœœ «·›—⁄ «Ê·«", vbCritical: Exit Sub
        If DataCombo1.BoundText = "" Then MsgBox "Õœœ Õ«·… «·„‘—Ê⁄", vbCritical: Exit Sub

    End If
  
    'If txtid.text = "" Then
    'txtid.text = get_code
    'End If

    Dim currentcode As String

    If TxtId.Text = "" Then
        currentcode = get_coding(branch_id, "projects", 0, Me.DCPreFix.Text)

        If currentcode = "miniError" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ…  ·Â–« ««ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„"
            Else
                MsgBox "The number fields for this code is to small please change it or call system administrator"
            End If
            Exit Sub
                        
        ElseIf currentcode = "Manual" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "«œŒ· «·ﬂÊœ ÌœÊÌ« ﬂ„« Õœœ  ›Ì  ﬂÊÌœ «·”‰œ« "
            Else
                MsgBox "Please Enter the code manually"
            End If
            Exit Sub
        Else
            TxtId = currentcode
        End If
    End If

    If TxtId.Text = "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            MsgBox "Must enter project code or define coding in your System", vbCritical: Exit Sub
        Else
            MsgBox "·«»œ „‰ ﬂ «»… —ﬁ„ ··„‘—Ê⁄ ·«‰ﬂ ·„  Õœœ  ﬂÊÌœ «·Ì ·…", vbCritical: Exit Sub
        End If
    End If
  
    If Me.OptType(2).value = False Then
        If val(Me.TxtOpenBalance.Text) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ  ··„’—Ê›«   ...!!!"
            Else
                Msg = "You must type the value of the balance of expenses"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

            If TxtOpenBalance.Enabled = True Then
                TxtOpenBalance.SetFocus
            End If

            Exit Sub
        End If
    End If
    
    If Me.OptType1(2).value = False Then
        If val(Me.TxtOpenBalance1.Text) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ ··«ÃÊ— ··«Ì—«œ«   ...!!!"
            Else
                Msg = "You must type the value of the balance of revenue"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

            If TxtOpenBalance1.Enabled = True Then
                TxtOpenBalance1.SetFocus
            End If

            Exit Sub
        End If
    End If
            
    If Me.OptType2(2).value = False Then
        If val(Me.TxtOpenBalance2.Text) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ ··„Ê«œ ...!!!"
            Else
                Msg = "A credit value must be written for the material"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

            If TxtOpenBalance2.Enabled = True Then
                TxtOpenBalance2.SetFocus
            End If

            Exit Sub
        End If
    End If
     
    If Me.OptType3(2).value = False Then
        If val(Me.TxtOpenBalance3.Text) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ ··«ÕÊ— ...!!!"
            Else
                Msg = "The value of the wage balance must be written"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

            If TxtOpenBalance3.Enabled = True Then
                TxtOpenBalance3.SetFocus
            End If

            Exit Sub
        End If
    End If
         If Me.OptType4(2).value = False Then
        If val(Me.TxtOpenBalance4.Text) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ ··„” Œ·’«  ...!!!"
            Else
                Msg = "A balance value must be written for abstracts"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

            If TxtOpenBalance4.Enabled = True Then
                TxtOpenBalance4.SetFocus
            End If

            Exit Sub
        End If
    End If
   If Me.OptType6(2).value = False Then
        If val(Me.TxtOpenBalance6.Text) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ ·Õ”‰ «·«œ«¡ ...!!!"
            Else
                Msg = "A balance value must be written for Good Performance"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

            If TxtOpenBalance6.Enabled = True Then
                TxtOpenBalance6.SetFocus
            End If

            Exit Sub
        End If
    End If
       If Me.OptType8(2).value = False Then
        If val(Me.TxtOpenBalance8.Text) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ  ·œ›⁄… «·„ﬁœ„… ...!!!"
            Else
                Msg = "A balance value must be written for PrePayment"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            If TxtOpenBalance8.Enabled = True Then
                TxtOpenBalance8.SetFocus
            End If

            Exit Sub
        End If
    End If
    
 '   If Me.OptType5(2).value = False Then
 '       If val(Me.TxtOpenBalance5.Text) = 0 Then
 '           Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ  Õ  «· ‰›Ì– ...!!!"
 '           MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'
'            If TxtOpenBalance5.Enabled = True Then
'                TxtOpenBalance5.SetFocus
'            End If
'
'            Exit Sub
'        End If
'    End If
     
 '   total_after_discount = val(TxtProjectCosts) - val(txt_total_discount)
    'terms_operations_Click (1)
    'opr_items_Click (1)
    'opr_Expenses_Click (1)
    'opr_emplyees_name_Click

    If TxtModFlg.Text = "N" Then
'rs.AddNew
       ' XPTxtID.text = CStr(new_id("projects", "id", "", True))
       ' MsgBox XPTxtID.text
        'Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=3"))
        If create_accounts = False Then Exit Sub
      '  rs("expanses_account").value = Account_Code_dynamic1C '  IIf(Trim$(Me.EXPANSES.text) = "", Null, Trim$(Me.EXPANSES.text))
      '  rs("REVENUE_account").value = Account_Code_dynamic2C ' IIf(Trim$(Me.REVENUE.text) = "", Null, Trim$(Me.REVENUE.text))
      '  rs("Material_account").value = Account_Code_dynamic3C '  IIf(Trim$(Me.Material.text) = "", Null, Trim$(Me.Material.text))
      '  rs("Salary_account").value = Account_Code_dynamic4C ' IIf(Trim$(Me.salary.text) = "", Null, Trim$(Me.salary.text))
      '  rs("legal").value = Account_Code_dynamic5C ' IIf(Trim$(Me.legal.text) = "", Null, Trim$(Me.legal.text))
                    
    Else

        If Me.TxtModFlg.Text = "E" Then
        StrSQL = "Delete From TblProjectUser Where ProjectID =" & val(Me.txt_project_id.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From ProjectMainDes Where ProjectID =" & val(Me.txt_project_id.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
           StrSQL = "Delete From ProJectMofrd Where ProjID =" & val(Me.txt_project_id.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From projects_des Where project_id =" & val(Me.txt_project_id.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
                    
           StrSQL = "Delete From Projectssub Where projectid =" & val(Me.txt_project_id.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
                         
        End If
    End If
           rs("id").value = val(Me.txt_project_id.Text)
         rs("UserID").value = IIf(DCboUserName.BoundText <> "", val((DCboUserName.BoundText)), Null)
          rs("EmpId").value = IIf(Me.DCEmP.BoundText = "", Null, (Me.DCEmP.BoundText))
rs("cost_after_discount").value = IIf(Me.total_after_discount.Text = "", Null, val((Me.total_after_discount.Text)))
     rs("EmpId1").value = IIf(Me.DCEmp1.BoundText = "", Null, (Me.DCEmp1.BoundText))
    rs("StartDate").value = DTStartDate.value
     rs("Enddate").value = DTEnddate.value
rs("DpNearEndDate").value = DpNearEndDate.value
''////////
 rs("EmpSalary").value = val(TxtEmpSalary.Text)
 rs("MangSalary").value = val(TxtMangSalary.Text)
 If Option7.value = True Then
  rs("UnderImp").value = 0
 ElseIf Option6.value = True Then
  rs("UnderImp").value = 1
 ElseIf Option8.value = True Then
  rs("UnderImp").value = 2
End If
   If RdTyp(1).value = True Then
   rs("TypeImport").value = 1
   Else
   rs("TypeImport").value = 0
   End If
   rs("Path").value = Me.txtFile.Text
If Option8.value = True Then

Else
      If Option8.value = False Then
        If Me.TxtModFlg.Text = "N" Then
        CretAccount2
        rs("expanses_account").value = Account_Code_dynamic1C '  IIf(Trim$(Me.EXPANSES.text) = "", Null, Trim$(Me.EXPANSES.text))
        rs("REVENUE_account").value = Account_Code_dynamic2C ' IIf(Trim$(Me.REVENUE.text) = "", Null, Trim$(Me.REVENUE.text))
        rs("Material_account").value = Account_Code_dynamic3C '  IIf(Trim$(Me.Material.text) = "", Null, Trim$(Me.Material.text))
        rs("Salary_account").value = Account_Code_dynamic4C ' IIf(Trim$(Me.salary.text) = "", Null, Trim$(Me.salary.text))
        rs("legal").value = Account_Code_dynamic5C ' IIf(Trim$(Me.legal.text) = "", Null, Trim$(Me.legal.text))
        rs("AcountGood").value = Account_Code_dynamic7C
      Else
            If Not IsNull(rs("expanses_account").value) And rs("expanses_account").value <> "" Then
                    
                ModAccounts.EditAccount rs("expanses_account").value, TXTprojectname & " - „’—Ê›«  ", TXTprojectname & "- Expenses", , , , , , , , , , , , , , , , , True
            End If
                         
            If Not IsNull(rs("REVENUE_account").value) And rs("REVENUE_account").value <> "" Then
                    
                ModAccounts.EditAccount rs("REVENUE_account").value, TXTprojectname & " - «Ì—«œ«  ", TXTprojectname & "- Revenue", , , , , , , , , , , , , , , , , True
            End If
                         
            If Not IsNull(rs("Material_account").value) And rs("Material_account").value <> "" Then
                    
                ModAccounts.EditAccount rs("Material_account").value, TXTprojectname & " - „Ê«œ ", TXTprojectname & "- Material", , , , , , , , , , , , , , , , , True
            End If
                         
            If Not IsNull(rs("Salary_account").value) And rs("Salary_account").value <> "" Then
       
                ModAccounts.EditAccount rs("Salary_account").value, TXTprojectname & " - «ÃÊ— ", TXTprojectname & "- Salary", , , , , , , , , , , , , , , , , True
            End If
            
            If Not IsNull(rs("legal").value) And rs("legal").value <> "" Then
       
                ModAccounts.EditAccount rs("legal").value, TXTprojectname & " - „” Œ·’«  ", TXTprojectname & "- bill", , , , , , , , , , , , , , , , , True
            End If
           If Not IsNull(rs("AcountGood").value) And rs("AcountGood").value <> "" Then
       
                ModAccounts.EditAccount rs("AcountGood").value, TXTprojectname & " - Õ”‰ «·«œ«¡ ", TXTprojectname & "- bill", , , , , , , , , , , , , , , , , True
            End If
      End If
      End If
End If

'''//////////
rs("OrderType").value = IIf(val(Me.CBoBasedON.ListIndex) = -1, Null, val(Me.CBoBasedON.ListIndex))
rs("OrderNo").value = IIf(Trim(Me.TXTOrDer_no.Text) = "", Null, Trim(Me.TXTOrDer_no.Text))
rs("TotalMainDes").value = IIf(Trim(Me.lbl(34).Caption) = "", Null, val(Me.lbl(34).Caption))
rs("TotalMainDesExe").value = IIf(Trim(Me.lbl(33).Caption) = "", Null, val(Me.lbl(33).Caption))
''///
  rs("Remarkss").value = IIf(Trim(Me.TxtRemarks.Text) = "", Null, Trim(Me.TxtRemarks.Text))
    rs("End_user_Account").value = IIf(Trim(Me.DcAccount1.Text) = "", Null, Trim(Me.DcAccount1.Text))
    rs("End_user_name").value = IIf(DcAccount2.Text = "", Null, DcAccount2.Text)
    
        rs("End_user_id").value = IIf(DcAccount2.BoundText = "", Null, DcAccount2.BoundText)

        rs("sub_contractor_id").value = IIf(DcAccount4.BoundText = "", Null, DcAccount4.BoundText)

    
    rs("CurrencyID").value = IIf(val(Me.DcCurrency.BoundText) = 0, 1, val(Me.DcCurrency.BoundText))
    
   rs("sub_contractor_Account").value = IIf(DcAccount3.BoundText = "", Null, Trim(DcAccount3.BoundText))
    rs("sub_contractor_name").value = IIf(DcAccount4.Text = "", Null, DcAccount4.Text)
    
    rs("prifix").value = IIf(Trim$(Me.DCPreFix.Text) = "", Null, Trim$(Me.DCPreFix.Text))
    rs("code").value = IIf(Trim$(Me.TxtId.Text) = "", Null, Trim$(Me.TxtId.Text))
    
    rs("Fullcode").value = IIf(Me.DCPreFix.Text & Me.TxtId.Text = "", Null, Me.DCPreFix.Text & Me.TxtId.Text)
    
    rs("Project_name").value = IIf(Trim$(Me.TXTprojectname.Text) = "", Null, Trim$(Me.TXTprojectname.Text))
    rs("Project_namee").value = IIf(Trim$(Me.TXTprojectnamee.Text) = "", Null, Trim$(Me.TXTprojectnamee.Text))
    
    rs("project_cost").value = IIf(val(Me.TxtProjectCosts.Text) = 0, 0, val(Me.TxtProjectCosts.Text))
    rs("general_discount").value = IIf(val(Me.txt_total_discount.Text) = 0, 0, val(Me.txt_total_discount.Text))
    
     rs("DiscountPercentage").value = IIf(val(Me.TxtDiscountPercentage.Text) = 0, 0, val(Me.TxtDiscountPercentage.Text))
     
   ' rs("cost_after_discount").value = rs("project_cost").value - rs("general_discount").value
    rs("net").value = rs("project_cost").value - rs("general_discount").value
     rs("Dept_ID").value = IIf(Trim$(Me.DcbDept.BoundText) = "", Null, Trim$(Me.DcbDept.BoundText))
    rs("branch_no").value = IIf(Trim$(Me.dcBranch.BoundText) = "", Null, Trim$(Me.dcBranch.BoundText))

    rs("Contract_type").value = IIf(Me.DataCombo5.BoundText = "", Null, Me.DataCombo5.BoundText)
    rs("Contract_type_name").value = IIf(Trim$(Me.DataCombo5.Text) = "", Null, Trim$(Me.DataCombo5.Text))
    rs("Project_status").value = IIf(Trim$(Me.DataCombo1.BoundText) = "", Null, val(Me.DataCombo1.BoundText))
  
    '   Rs("departement").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
    '   Rs("project_code").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
 
    rs("branch_no").value = IIf(Not IsNumeric(dcBranch.BoundText), 0, dcBranch.BoundText)
    
   ' rs("End_user_id").value = IIf(Trim$(Me.DCAccount1.text) = "", Null, Trim$(Me.DCAccount1.text))
   ' rs("sub_contractor_id").value = IIf(Trim$(Me.DCAccount3.text) = "", Null, Trim$(Me.DCAccount3.text))
    '   Rs("total").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
    '  Rs("sub_discount_total").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
    '  Rs("net").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
    rs("items_total").value = IIf(Me.XPTxtSum.Text = "", 0, Me.XPTxtSum.Text)
  
  If Ptype(0).value = True Then
  
    rs("Pstate").value = 0
  Else
     rs("Pstate").value = 1
  End If
  
  
    If Me.OptType(2).value = True Then
        rs("OpenBalance").value = 0
        rs("OpenBalanceType").value = Null
    ElseIf Me.OptType(0).value = True Then
        rs("OpenBalance").value = val(Me.TxtOpenBalance.Text)
        rs("OpenBalanceType").value = 0
    ElseIf Me.OptType(1).value = True Then
        rs("OpenBalance").value = val(Me.TxtOpenBalance.Text)
        rs("OpenBalanceType").value = 1
    End If
    
    If Me.OptType1(2).value = True Then
        rs("OpenBalance1").value = 0
        rs("OpenBalanceType1").value = Null
    ElseIf Me.OptType1(0).value = True Then
        rs("OpenBalance1").value = val(Me.TxtOpenBalance1.Text)
        rs("OpenBalanceType1").value = 0
    ElseIf Me.OptType1(1).value = True Then
        rs("OpenBalance1").value = val(Me.TxtOpenBalance1.Text)
        rs("OpenBalanceType1").value = 1
    End If
    
    If Me.OptType2(2).value = True Then
        rs("OpenBalance2").value = 0
        rs("OpenBalanceType2").value = Null
    ElseIf Me.OptType2(0).value = True Then
        rs("OpenBalance2").value = val(Me.TxtOpenBalance2.Text)
        rs("OpenBalanceType2").value = 0
    ElseIf Me.OptType2(1).value = True Then
        rs("OpenBalance2").value = val(Me.TxtOpenBalance2.Text)
        rs("OpenBalanceType2").value = 1
    End If
    
    If Me.OptType3(2).value = True Then
        rs("OpenBalance3").value = 0
        rs("OpenBalanceType3").value = Null
    ElseIf Me.OptType3(0).value = True Then
        rs("OpenBalance3").value = val(Me.TxtOpenBalance3.Text)
        rs("OpenBalanceType3").value = 0
    ElseIf Me.OptType3(1).value = True Then
        rs("OpenBalance3").value = val(Me.TxtOpenBalance3.Text)
        rs("OpenBalanceType3").value = 1
    End If
 
    If Me.OptType4(2).value = True Then
        rs("OpenBalance4").value = 0
        rs("OpenBalanceType4").value = Null
    ElseIf Me.OptType4(0).value = True Then
        rs("OpenBalance4").value = val(Me.TxtOpenBalance4.Text)
        rs("OpenBalanceType4").value = 0
    ElseIf Me.OptType4(1).value = True Then
        rs("OpenBalance4").value = val(Me.TxtOpenBalance4.Text)
        rs("OpenBalanceType4").value = 1
    End If
    
    If Me.OptType5(2).value = True Then
        rs("OpenBalance5").value = 0
        rs("OpenBalanceType5").value = Null
    ElseIf Me.OptType5(0).value = True Then
        rs("OpenBalance5").value = val(Me.TxtOpenBalance5.Text)
        rs("OpenBalanceType5").value = 0
    ElseIf Me.OptType5(1).value = True Then
        rs("OpenBalance5").value = val(Me.TxtOpenBalance5.Text)
        rs("OpenBalanceType5").value = 1
    End If
    If Me.OptType6(2).value = True Then
        rs("OpenBalance6").value = 0
        rs("OpenBalanceType6").value = Null
    ElseIf Me.OptType6(0).value = True Then
        rs("OpenBalance6").value = val(Me.TxtOpenBalance6.Text)
        rs("OpenBalanceType6").value = 0
    ElseIf Me.OptType6(1).value = True Then
        rs("OpenBalance6").value = val(Me.TxtOpenBalance6.Text)
        rs("OpenBalanceType6").value = 1
    End If
    
       If Me.OptType8(2).value = True Then
        rs("OpenBalance8").value = 0
        rs("OpenBalanceType8").value = Null
    ElseIf Me.OptType8(0).value = True Then
        rs("OpenBalance8").value = val(Me.TxtOpenBalance8.Text)
        rs("OpenBalanceType8").value = 0
    ElseIf Me.OptType8(1).value = True Then
        rs("OpenBalance8").value = val(Me.TxtOpenBalance8.Text)
        rs("OpenBalanceType8").value = 1
    End If
    
    ' aladein code add
     rs("JobeDO").value = IIf(Me.Text20(5).Text = "", 0, Me.Text20(5).Text)
     rs("JobeDOPercent").value = IIf(Me.Text20(4).Text = "", 0, Me.Text20(4).Text)
     rs("JobeGet").value = IIf(Me.Text20(3).Text = "", 0, Me.Text20(3).Text)
     rs("JobeRest").value = IIf(Me.Text20(2).Text = "", 0, Me.Text20(2).Text)
     rs("JobeTimeLeft").value = IIf(Me.Text20(1).Text = "", 0, Me.Text20(1).Text)
     rs("JobeTime").value = IIf(Me.Text20(0).Text = "", 0, Me.Text20(0).Text)
     '''''''''''''''''''''''''''''''''''
    If company.value = True Then
    rs.Fields("JobeWork").value = 0
    Else
    rs.Fields("JobeWork").value = 1
    
    End If
    
    If Check4.value = True Then
    rs("JobeContractorID").value = IIf(Me.DataCombo3.BoundText = "", Null, Me.DataCombo3.BoundText)
    Else
     rs("JobeContractorID").value = Null
    End If
    
    rs("Amanhid").value = IIf(Me.AmanhNames.BoundText = "", Null, Me.AmanhNames.BoundText)
    rs("Municipalityid").value = IIf(Me.MunicipalityNames.BoundText = "", Null, Me.MunicipalityNames.BoundText)
    
    ''''''''''''''''''''''''''''''''''
    rs("WarrantyNO").value = IIf(Me.Text21.Text = "", 0, Me.Text21.Text)
    rs("WarrantyValue").value = IIf(Me.Text22.Text = "", 0, Me.Text22.Text)
    rs("WarrDateStart").value = IIf(Me.DTPicker1.value = "", 0, Me.DTPicker1.value)
    rs("WarrDateEnd").value = IIf(Me.DTPicker3.value = "", 0, Me.DTPicker3.value)
    rs("WarrExtension").value = IIf(Me.Text25.Text = "", 0, Me.Text25.Text)
    rs("WarrBank").value = IIf(Me.DataCombo2.BoundText = "", Null, Me.DataCombo2.BoundText)
    rs("ContractNo").value = IIf(Me.TxtContractNo.Text = "", "", Me.TxtContractNo.Text)
    ''''''''' end addy
    
    
    rs("OpenBalanceDate").value = Me.Dtp.value
    Dim Account_Code_dynamic1 As String

    If val(TxtOpenBalance.Text) <> 0 Or val(TxtOpenBalance1.Text) <> 0 Or val(TxtOpenBalance2.Text) <> 0 Or val(TxtOpenBalance3.Text) <> 0 Or val(TxtOpenBalance4.Text) <> 0 Or val(TxtOpenBalance5.Text) <> 0 Or val(TxtOpenBalance6.Text) <> 0 Or val(TxtOpenBalance8.Text) <> 0 Then
        txtopening_balance_voucher_id.Text = get_opening_balance_voucher_id
        Account_Code_dynamic1 = get_account_code_branch(73, my_branch)
                    
        If Account_Code_dynamic1 = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "Branch was not created", vbCritical
            End If
            GoTo ErrTrap
        Else

            If Account_Code_dynamic1 = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "An openning  account for this branch was not selected", vbCritical
                End If
                GoTo ErrTrap
                                 
            End If
        End If

        txtopening_balance_voucher_id.Text = get_opening_balance_voucher_id
        rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.Text)
    Else
        rs("opening_balance_voucher_id").value = Null
    End If
  
    'OPENING Balance Voucher
    Dim StrDes As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrDes = "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.TXTprojectname.Text) & " "
    Else
        StrDes = " Opening Balance For: " & Trim(Me.TXTprojectnamee.Text) & " "
    End If
        
  
    
    If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
           

            LngOpenID = 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
            If Me.OptType(0).value = True Then
        
            '    If ModAccounts.AddNewDev(LngDevID, 1, rs("expanses_account").value, val(Me.TxtOpenBalance.text), 0, StrDes & " - ··„’—Ê›«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
                
            '    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance.text), 1, StrDes & " - ··„’—Ê›«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
                
            ElseIf Me.OptType(1).value = True Then
                 
            '    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance.text), 0, StrDes & " - ··„’—Ê›«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
            '
            '    If ModAccounts.AddNewDev(LngDevID, 2, rs("expanses_account").value, val(Me.TxtOpenBalance.text), 1, StrDes & " - ··„’—Ê›«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
            End If
                 
        End If
    End If

    If Me.OptType1(0).value = True Or Me.OptType1(1).value = True Then
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then

            LngOpenID = 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
            If Me.OptType1(0).value = True Then
        
            '    If ModAccounts.AddNewDev(LngDevID, 1, rs("REVENUE_account").value, val(Me.TxtOpenBalance1.text), 0, StrDes & " - ··«Ì—«œ«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
            '
            '    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance1.text), 1, StrDes & " - ··«Ì—«œ«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
                
            ElseIf Me.OptType1(1).value = True Then
                 
            '    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance1.text), 0, StrDes & " - ··«Ì—«œ«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
            '
            '    If ModAccounts.AddNewDev(LngDevID, 2, rs("REVENUE_account").value, val(Me.TxtOpenBalance1.text), 1, StrDes & " - ··«Ì—«œ«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
            End If
                 
        End If
    End If
 
    If Me.OptType2(0).value = True Or Me.OptType2(1).value = True Then
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then

            LngOpenID = 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
            If Me.OptType2(0).value = True Then
        
            '    If ModAccounts.AddNewDev(LngDevID, 1, rs("Material_account").value, val(Me.TxtOpenBalance2.text), 0, StrDes & " - ··„Ê«œ ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
            '
            '    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance2.text), 1, StrDes & " - ··„Ê«œ ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
                
            ElseIf Me.OptType2(1).value = True Then
                 
            '    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance2.text), 0, StrDes & " - ··„Ê«œ ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
            '
            '    If ModAccounts.AddNewDev(LngDevID, 2, rs("Material_account").value, val(Me.TxtOpenBalance2.text), 1, StrDes & " - ··„Ê«œ ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
            End If
                 
        End If
    End If
  
    If Me.OptType3(0).value = True Or Me.OptType3(1).value = True Then
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then

            LngOpenID = 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
            If Me.OptType3(0).value = True Then
        
            '    If ModAccounts.AddNewDev(LngDevID, 1, rs("Salary_account").value, val(Me.TxtOpenBalance3.text), 0, StrDes & " - ··«ÃÊ— ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
            '
            '    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance3.text), 1, StrDes & " - ··«ÃÊ— ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
                
            ElseIf Me.OptType3(1).value = True Then
                 
            '    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance3.text), 0, StrDes & " - ··«ÃÊ— ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
            '
            '    If ModAccounts.AddNewDev(LngDevID, 2, rs("Salary_account").value, val(Me.TxtOpenBalance3.text), 1, StrDes & " - ··«ÃÊ— ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
            '        GoTo ErrTrap
            '    End If
            End If
                 
        End If
    End If
   
   Dim AccountCode As String
    If Me.OptType4(0).value = True Or Me.OptType4(1).value = True Then
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then

            LngOpenID = 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
       If SystemOptions.Revenueowed = True Then
       AccountCode = rs("legal").value
       Else
       AccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DcAccount2.BoundText))
       End If
       
            If Me.OptType4(0).value = True Then
        
        
                If ModAccounts.AddNewDev(LngDevID, 1, AccountCode, val(Me.TxtOpenBalance4.Text), 0, StrDes & " - ··„” Œ·’«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , val(txt_project_id.Text), , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
                
                If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance4.Text), 1, StrDes & " - ··„” Œ·’«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
                
                ElseIf Me.OptType4(1).value = True Then
                 
                If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance4.Text), 0, StrDes & " - ··„” Œ·’«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
                
                If ModAccounts.AddNewDev(LngDevID, 2, AccountCode, val(Me.TxtOpenBalance4.Text), 1, StrDes & " - ··„” Œ·’«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , val(txt_project_id.Text), , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
        End If
    End If
 ''////////////////////////
    
    'OPENING Balance Voucher


    rs.update
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    ' If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
    ' LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
    ' »‰Êœ «·„‘—Ê⁄
    Set RsDev1 = New ADODB.Recordset
    RsDev1.Open "ProjectMainDes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   Dim i As Integer
    With FgMainDes
        For i = .FixedRows To .Rows - 2
            If .TextMatrix(i, .ColIndex("Name")) <> "" Then

               RsDev1.AddNew
                If .TextMatrix(i, .ColIndex("FullCode")) = "" Then
                    RsDev1("FullCode").value = .TextMatrix(i, .ColIndex("LineNo"))
                Else
                    RsDev1("FullCode").value = .TextMatrix(i, .ColIndex("FullCode"))
                End If
             Pand = val(.TextMatrix(i, .ColIndex("ID")))
       If Me.Checked(0, 0, Pand) = True Then
        Else
       Pand = 1
        maxx 0, 0, Pand
       
        .TextMatrix(i, .ColIndex("ID")) = Pand
       End If
                 RsDev1("Name").value = .TextMatrix(i, .ColIndex("Name"))
                 RsDev1("ProjectID").value = Me.txt_project_id.Text
                 RsDev1("Remarks").value = IIf(.TextMatrix(i, .ColIndex("Remarks")) = "", Null, .TextMatrix(i, .ColIndex("Remarks")))
                 RsDev1("ID").value = IIf(.TextMatrix(i, .ColIndex("ID")) = "", Null, val(.TextMatrix(i, .ColIndex("ID"))))
                 RsDev1("Qty").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("Qty"))), 0, .TextMatrix(i, .ColIndex("Qty")))
                 RsDev1("Price").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("Price"))), 0, .TextMatrix(i, .ColIndex("Price")))
                 RsDev1("Total").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("Total"))), 0, .TextMatrix(i, .ColIndex("Total")))
                 RsDev1("QtyNo").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("QtyNo"))), 0, .TextMatrix(i, .ColIndex("QtyNo")))
                 RsDev1("QtyExe").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("QtyExe"))), 0, .TextMatrix(i, .ColIndex("QtyExe")))
                 RsDev1("PriceExe").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("PriceExe"))), 0, .TextMatrix(i, .ColIndex("PriceExe")))
                 RsDev1("TotalExe").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("TotalExe"))), 0, .TextMatrix(i, .ColIndex("TotalExe")))
                 RsDev1.update
            End If

        Next i
    
    End With
    
    ' »‰Êœ «·„‘—Ê⁄
    Set RsDev1 = New ADODB.Recordset
    RsDev1.Open "projects_des", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    With Fg_Journal

        For i = .FixedRows To .Rows - 2

            '        Dim IntDEV_Type As Integer
            '        Dim SngDEV_Value As Single
            If .TextMatrix(i, .ColIndex("des")) <> "" Then
             If .TextMatrix(i, .ColIndex("by")) = "" Then
               If val(DataCombo3.BoundText) <> 0 And DataCombo3.Text <> "" Then
                .TextMatrix(i, .ColIndex("by")) = DataCombo3.Text
                .TextMatrix(i, .ColIndex("sub_contractor_id")) = val(DataCombo3.BoundText)
              End If
             End If
        Pand = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("oprid"))), 0, val(.TextMatrix(i, .ColIndex("oprid"))))
        
        If Pand = 0 Then
          If Me.Checked(Pand, 0) = True Then
        Else
       Pand = 1
        maxx Pand, 0
        End If
        .TextMatrix(i, .ColIndex("oprid")) = Pand
        End If
        
        
               RsDev1.AddNew
              

                If .TextMatrix(i, .ColIndex("fullcode")) = "" Then
                    RsDev1("fullcode").value = .TextMatrix(i, .ColIndex("LineNo"))
                Else
                    RsDev1("fullcode").value = .TextMatrix(i, .ColIndex("fullcode"))
                End If
                RsDev1("CodeBand").value = .TextMatrix(i, .ColIndex("CodeBand"))
                RsDev1("PanID").value = val(.TextMatrix(i, .ColIndex("PanID")))
                RsDev1("PrMainDesID").value = IIf(val(.TextMatrix(i, .ColIndex("PrMainDesID"))) = 0, Null, val(.TextMatrix(i, .ColIndex("PrMainDesID"))))
                RsDev1("oprid").value = val(.TextMatrix(i, .ColIndex("oprid")))
                RsDev1("QtyNo").value = val(.TextMatrix(i, .ColIndex("QtyNo"))) ' IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("oprid"))), 0, (.TextMatrix(i, .ColIndex("oprid"))))
                RsDev1("project_id").value = Me.txt_project_id.Text
                 RsDev1("des").value = IIf(.TextMatrix(i, .ColIndex("des")) = "", Null, .TextMatrix(i, .ColIndex("des")))
                RsDev1("Remark").value = IIf(.TextMatrix(i, .ColIndex("Remark")) = "", Null, .TextMatrix(i, .ColIndex("Remark")))
                ''//
                RsDev1("PandUnitID").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("qty"))), 0, val(.TextMatrix(i, .ColIndex("PandUnitID"))))
                ''/
                RsDev1("qty").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("qty"))), 0, .TextMatrix(i, .ColIndex("qty")))
                RsDev1("cost").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("cost"))), 0, .TextMatrix(i, .ColIndex("cost")))
                RsDev1("total").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("total"))), 0, .TextMatrix(i, .ColIndex("total")))
                RsDev1("discount").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("discount"))), 0, .TextMatrix(i, .ColIndex("discount")))
                RsDev1("net").value = RsDev1("total").value - RsDev1("discount").value
                RsDev1("line_no").value = .TextMatrix(i, .ColIndex("LineNo"))
                RsDev1("sub_contractor_id").value = val(.TextMatrix(i, .ColIndex("sub_contractor_id")))
                RsDev1("esQty").value = IIf(.TextMatrix(i, .ColIndex("esQty")) = "", Null, .TextMatrix(i, .ColIndex("esQty")))
                RsDev1("QtyExe").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("QtyExe"))), 0, .TextMatrix(i, .ColIndex("QtyExe")))
                RsDev1("PriceExe").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("PriceExe"))), 0, .TextMatrix(i, .ColIndex("PriceExe")))
                RsDev1("TotalExe").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("TotalExe"))), 0, .TextMatrix(i, .ColIndex("TotalExe")))
        
                RsDev1.update
              
            End If

        Next i
    
    End With

    Set RsDevsub = New ADODB.Recordset
    RsDevsub.Open "Projectssub", Cn, adOpenStatic, adLockOptimistic, adCmdTable
terms_operations_Click (1)
'œ›⁄«  «·„‘—Ê⁄
 With GridSub

        For i = .FixedRows To .Rows - 2

            '        Dim IntDEV_Type As Integer
            '        Dim SngDEV_Value As Single
            If .TextMatrix(i, .ColIndex("id")) <> "" Then
                RsDevsub.AddNew
                RsDevsub("projectid").value = Me.txt_project_id
                RsDevsub("subdate").value = IIf(Not IsDate(.TextMatrix(i, .ColIndex("subdate"))), Null, .TextMatrix(i, .ColIndex("subdate")))
              '  RsDevsub("DesTerm").value = IIf(IsNull(.TextMatrix(i, .ColIndex("DesTerm"))), "", .TextMatrix(i, .ColIndex("DesTerm")))
                RsDevsub("ValueTerm").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("ValueTerm"))), 0, .TextMatrix(i, .ColIndex("ValueTerm")))
                RsDevsub("SubValue").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("SubValue"))), 0, .TextMatrix(i, .ColIndex("SubValue")))
                 RsDevsub("rate").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("rate"))), 0, .TextMatrix(i, .ColIndex("rate")))
                RsDevsub("REmarks").value = IIf(IsNull(.TextMatrix(i, .ColIndex("REmarks"))), "", .TextMatrix(i, .ColIndex("REmarks")))
      RsDevsub.update
            End If

        Next i
    
    End With
Dim Rs5 As ADODB.Recordset
    Set Rs5 = New ADODB.Recordset
    Rs5.Open "ProJectMofrd", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 With Me.Grid

        For i = .FixedRows To .Rows - 1
            If val(.TextMatrix(i, .ColIndex("MofrdID"))) <> 0 Then
                Rs5.AddNew
                Rs5("ProjID").value = val(Me.txt_project_id)
                Rs5("MofrdID").value = IIf(IsNull(.TextMatrix(i, .ColIndex("MofrdID"))), 0, val(.TextMatrix(i, .ColIndex("MofrdID"))))
                Rs5("TypeEmp").value = IIf(IsNull(.TextMatrix(i, .ColIndex("TypeEmp"))), 1, val(.TextMatrix(i, .ColIndex("TypeEmp"))))
                Rs5("Valuee").value = IIf(IsNull(.TextMatrix(i, .ColIndex("Valuee"))), 0, val(.TextMatrix(i, .ColIndex("Valuee"))))
                Rs5.update
            End If
        Next i
    End With
    'Retrive
    Command1(1).Enabled = False

    '„Ê«œ «·„‘—Ê⁄
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '       StrSQL = "Delete From Transaction_Details Where project_id =" & Val(Me.txt_project_id.text)
    '       Cn.Execute StrSQL, , adExecuteNoRecords
        
    '   Set RSTransDetails = New ADODB.Recordset
    '   RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '       For RowNum = 1 To FG.Rows - 1
    '       If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
  
    '   RSTransDetails.AddNew
    '   RSTransDetails("Transaction_Details").value = Val(txt_project_id.text)
    '   RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
    '   RSTransDetails("Price").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
    '   RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
    '   RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
    '
    '   RSTransDetails("UnitID").value = _
    '        IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
    '
    '   RSTransDetails.update
    '   End If
    '   Next
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '«·„ÊŸ›Ì‰ «·„”Ã·Ì‰ ›Ì «·„‘—Ê⁄
    'EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
 
    ' Set RsDev = New ADODB.Recordset
    ' Dim Sql As String
    '
    '  With VSFlexGrid1
    '    For I = .FixedRows To .Rows - 2
    '
    '        If .TextMatrix(I, .ColIndex("id")) <> "" Then
    '        Sql = "Select * from TblEmployee where Emp_ID=" & .TextMatrix(I, .ColIndex("id"))
    '        RsDev.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '        RsDev("project_id").value = Me.txt_project_id
    '        RsDev.update
    '        RsDev.Close
    '        End If
    '    Next I
    '
    'End With
    saveOperationDates val(txt_project_id.Text), DTStartDate.value
    SaveUserProject val(txt_project_id.Text)
    'EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
    If Option8.value = True Then
    
    
If Me.TxtModFlg.Text = "N" Then
Account_Code_dynamic6C = ModAccounts.AddNewAccount(Account_Code_dynamic6, TXTprojectname & " - Õ  «· ‰›Ì– ", True, False, TXTprojectnamee & " -Under implementation")
Cn.Execute "Update projects set AccountUnderImp ='" & Account_Code_dynamic6C & "' where id =" & val(txt_project_id.Text) & " "
rs.Resync
End If
If Me.TxtModFlg.Text = "E" Then
     If Not IsNull(rs("AccountUnderImp").value) And rs("AccountUnderImp").value <> "" Then
                ModAccounts.EditAccount rs("AccountUnderImp").value, TXTprojectname & " - Õ  «· ‰›Ì– ", TXTprojectname & "-Under implementation", , , , , , , , , , , , , , , , , True
      End If
 End If
 
    If Me.OptType5(0).value = True Or Me.OptType5(1).value = True Then
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
           

            LngOpenID = 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
            If Me.OptType5(0).value = True Then
        
                If ModAccounts.AddNewDev(LngDevID, 1, rs("AccountUnderImp").value, val(Me.TxtOpenBalance5.Text), 0, StrDes & " - ··„’—Ê›«  ", LngOpenID, , , , Me.Dtp5.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
                
               If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance5.Text), 1, StrDes & " - ··„’—Ê›«  ", LngOpenID, , , , Me.Dtp5.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
                
            ElseIf Me.OptType5(1).value = True Then
                 
               If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance5.Text), 0, StrDes & " - ··„’—Ê›«  ", LngOpenID, , , , Me.Dtp5.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
            
                If ModAccounts.AddNewDev(LngDevID, 2, rs("AccountUnderImp").value, val(Me.TxtOpenBalance5.Text), 1, StrDes & " - ··„’—Ê›«  ", LngOpenID, , , , Me.Dtp5.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
        End If
    End If
    End If
    ''////////////////
       If Me.OptType6(0).value = True Or Me.OptType6(1).value = True Then
       Account_Code_dynamic1 = get_account_code_branch(73, my_branch)
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            LngOpenID = 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
            If Me.OptType6(0).value = True Then
        
                If ModAccounts.AddNewDev(LngDevID, 1, rs("AcountGood").value, val(Me.TxtOpenBalance6.Text), 0, StrDes & " - ·Õ”‰ «·«œ«¡ ", LngOpenID, , , , Me.Dtp6.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
                
               If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance6.Text), 1, StrDes & " - ·Õ”‰ «·«œ«¡ ", LngOpenID, , , , Me.Dtp6.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
                
            ElseIf Me.OptType6(1).value = True Then
                 
               If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance6.Text), 0, StrDes & " - ·Õ”‰ «·«œ«¡ ", LngOpenID, , , , Me.Dtp6.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
            
                If ModAccounts.AddNewDev(LngDevID, 2, rs("AcountGood").value, val(Me.TxtOpenBalance6.Text), 1, StrDes & " - ·Õ”‰ «·«œ«¡ ", LngOpenID, , , , Me.Dtp6.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
        End If
    End If
    ''//////////////////
           If Me.OptType8(0).value = True Or Me.OptType8(1).value = True Then
            Account_Code_dynamic1 = get_account_code_branch(73, my_branch)
            AccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DcAccount2.BoundText), "Account_Code2")
            LngOpenID = 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
            If Me.OptType8(0).value = True Then
        
                If ModAccounts.AddNewDev(LngDevID, 1, AccountCode, val(Me.TxtOpenBalance8.Text), 0, StrDes & " -  ·œ›⁄… „ﬁœ„… ", LngOpenID, , , , Me.Dtp8.value, , , , , , , , , , , , val(txt_project_id.Text), , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
                
               If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance8.Text), 1, StrDes & " -  ·œ›⁄… „ﬁœ„… ", LngOpenID, , , , Me.Dtp8.value, , , , , , , , , , , , val(txt_project_id.Text), , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
                
            ElseIf Me.OptType8(1).value = True Then
                 
               If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance8.Text), 0, StrDes & " -  ·œ›⁄… „ﬁœ„… ", LngOpenID, , , , Me.Dtp8.value, , , , , , , , , , , , val(txt_project_id.Text), , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
            
                If ModAccounts.AddNewDev(LngDevID, 2, AccountCode, val(Me.TxtOpenBalance8.Text), 1, StrDes & " -  ·œ›⁄… „ﬁœ„… ", LngOpenID, , , , Me.Dtp8.value, , , , , , , , , , , , val(txt_project_id.Text), , True, val(txtopening_balance_voucher_id.Text)) = False Then
                    GoTo ErrTrap
                End If
            End If
    End If
    Select Case Me.TxtModFlg.Text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then

                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„‘—Ê⁄" & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " This Project Data Was Saved" & CHR(13)
                Msg = Msg + "Do you want To enter Another Project"
            End If

            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Command1_Click (0)
                Exit Sub
            End If
            
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Amendments have been saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If

    End Select
    
    'If SystemOptions.UserInterface = EnglishInterface Then
    'MsgBox "Saved", vbInformation, ""
    'Else
    'MsgBox " „ Õ›Ÿ «·»Ì«‰« ", vbInformation, ""
    'End If
    TxtModFlg.Text = "R"
    Exit Sub
ErrTrap:

    If SystemOptions.UserInterface = EnglishInterface Then
        MsgBox "Error During saving", vbInformation, ""
    Else
        MsgBox "ÕœÀ Œÿ√ «À‰«¡  Õ›Ÿ «·»Ì«‰« ", vbInformation, ""
    End If

End Sub

Function calcnets()

    With Me.VSFlexGrid1
        txt_employee_count = .Rows - 2
        Me.txt_emp_salary.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
    End With
 
End Function

  Private Sub ReLineGrid2()
    IntCounter = 0
    Dim SUM As Double
    Dim i As Integer
  SUM = 0
    With GridSub

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("subdate")) <> "" Then
            SUM = SUM + val(.TextMatrix(i, .ColIndex("SubValue")))
            IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("id")) = IntCounter
            If SUM <= val(total_after_discount.Text) Then
                
                Else
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "⁄œœ «·œ›⁄«  «ﬂ»— „‰ ﬁÌ„… «·„‘—Ê⁄"
                Else
                    MsgBox "The number of payments is greater than the value of the project"
                End If
                .TextMatrix(i, .ColIndex("SubValue")) = 0
                Exit Sub
                End If
  
            End If

        Next i
   
    End With
End Sub
Private Sub ReLineGrid(Optional current_terms As String = "")
    Dim i As Integer
    Dim IntCounter As Integer
    Dim StartWeek As Double
    Dim EndWeek As Double
    Dim EarlyStartWeek As Double
    Dim EarlyEndWeek As Double
    Dim rs As ADODB.Recordset

    With Fg_Journal
TotalExe = 0
        For i = .FixedRows To .Rows - 1
TotalExe = TotalExe + val(.TextMatrix(i, .ColIndex("TotalExe")))
            If .TextMatrix(i, .ColIndex("des")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                .TextMatrix(i, .ColIndex("discount")) = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("discount"))), 0, .TextMatrix(i, .ColIndex("discount")))
                .TextMatrix(i, .ColIndex("qty")) = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("qty"))), 0, .TextMatrix(i, .ColIndex("qty")))
                .TextMatrix(i, .ColIndex("cost")) = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("cost"))), 0, .TextMatrix(i, .ColIndex("cost")))
                sql = "select sum(total) as total  From terms_operations Where term_fullcode='" & .TextMatrix(i, .ColIndex("fullcode")) & "'"
        
                Set rs = New ADODB.Recordset
                rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If rs.RecordCount > 0 And Not IsNull(rs("total").value) Then
                  '  .TextMatrix(i, .ColIndex("qty")) = ""
                  '  .TextMatrix(i, .ColIndex("cost")) = ""
                  '  .TextMatrix(i, .ColIndex("total")) = rs("total").value
                  '  .TextMatrix(i, .ColIndex("net")) = rs("total").value
         
                Else
                    .TextMatrix(i, .ColIndex("total")) = .TextMatrix(i, .ColIndex("qty")) * .TextMatrix(i, .ColIndex("cost")) * IIf(val(.TextMatrix(i, .ColIndex("QtyNo"))) = 0, 1, val(.TextMatrix(i, .ColIndex("QtyNo"))))
                    .TextMatrix(i, .ColIndex("net")) = .TextMatrix(i, .ColIndex("total")) - .TextMatrix(i, .ColIndex("discount"))
         
                End If
 .TextMatrix(i, .ColIndex("total")) = .TextMatrix(i, .ColIndex("qty")) * .TextMatrix(i, .ColIndex("cost")) * IIf(val(.TextMatrix(i, .ColIndex("QtyNo"))) = 0, 1, val(.TextMatrix(i, .ColIndex("QtyNo"))))
                    .TextMatrix(i, .ColIndex("net")) = .TextMatrix(i, .ColIndex("total")) - .TextMatrix(i, .ColIndex("discount"))
         
    .TextMatrix(i, .ColIndex("TotalExe")) = val(.TextMatrix(i, .ColIndex("QtyExe"))) * val(.TextMatrix(i, .ColIndex("PriceExe")))
    
    .TextMatrix(i, .ColIndex("fullcode")) = IntCounter
               ' If .TextMatrix(i, .ColIndex("CodeBand")) = "" Then
               ' .TextMatrix(i, .ColIndex("CodeBand")) = IntCounter
               ' End If
            End If

        Next i
If .Rows > 1 Then
        Me.txt_total_sum.Text = Round(val(.Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))), Decimal_Places)
        Me.txt_sub_discount.Text = Round(.Aggregate(flexSTSum, .FixedRows, .ColIndex("discount"), .Rows - 1, .ColIndex("discount")), Decimal_Places)
        Me.txt_sub_net.Text = Round(.Aggregate(flexSTSum, .FixedRows, .ColIndex("net"), .Rows - 1, .ColIndex("net")), Decimal_Places)
        If Option8.value = True Then
        OptType5(0).value = True
        
        Me.TxtOpenBalance5.Text = val(lbl(33).Caption)  '.Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalExe"), .Rows - 1, .ColIndex("TotalExe"))
        If val(TxtOpenBalance5.Text) > 0 Then
        TxtOpenBalance5.Enabled = False
        Else
        TxtOpenBalance5.Enabled = True
        End If
        
        End If
      End If
    End With
       
    Label13.Caption = getoprTitle
    IntCounter = 0

    With VSFlexGrid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
  
            End If

        Next i
   
        txt_employee_count = .Rows - 1
        Me.txt_emp_salary.Text = Round(.Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total")), Decimal_Places)

    End With
     
    IntCounter = 0

    With VSFlexGrid2

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                '.TextMatrix(i, .ColIndex("fullcode")) = current_terms & "-" & IntCounter
                .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("EquepVal"))) + val(.TextMatrix(i, .ColIndex("total_expenses"))) + val(.TextMatrix(i, .ColIndex("total_items"))) + val(.TextMatrix(i, .ColIndex("total_salary")))

            End If
        
            If val(.TextMatrix(i, .ColIndex("Period1"))) = 0 Then
                .TextMatrix(i, .ColIndex("Critical")) = 1
            Else
                .TextMatrix(i, .ColIndex("Critical")) = 0
            End If

            get_opr_details .TextMatrix(i, .ColIndex("Pre")), val(.TextMatrix(i, .ColIndex("period"))), val(.TextMatrix(i, .ColIndex("period1"))), StartWeek, EndWeek, EarlyStartWeek, EarlyEndWeek
            .TextMatrix(i, .ColIndex("startweek")) = StartWeek
            .TextMatrix(i, .ColIndex("EndWeek")) = EndWeek
            .TextMatrix(i, .ColIndex("Earlystartweek")) = EarlyStartWeek
            .TextMatrix(i, .ColIndex("EarlyEndWeek")) = EarlyEndWeek

            If val(.TextMatrix(i, .ColIndex("Critical"))) Then
                .Cell(flexcpBackColor, i, 14, i, 14) = vbRed
            Else
                .Cell(flexcpBackColor, i, 14, i, 14) = vbGreen
            End If

        Next i
If .Rows > 1 Then
        Me.txt_opr_total.Text = Round(.Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total")), Decimal_Places)
        Me.TXTNoOFWeek.Text = .Aggregate(flexSTMax, .FixedRows, .ColIndex("EndWeek"), .Rows - 1, .ColIndex("EndWeek"))
        End If
        Dim x As Double
        x = getProjectDuration(val(Me.txt_project_id.Text))
        x = DateDiff("D", DTStartDate.value, DTEnddate.value)
     '   Text4.text = x & "   ÌÊ„ " '& Me.Label13.Caption
         If SystemOptions.UserInterface = ArabicInterface Then
        Text4.Text = x & "   ÌÊ„ " '
       Else
       Text4.Text = x & "   Days "
       End If
       
        If SystemOptions.ProcessPeriodType = 0 Then
        '    DTEnddate.value = DateAdd("d", X, DTStartDate.value)      'day
        ElseIf SystemOptions.ProcessPeriodType = 1 Then
        '    DTEnddate.value = DateAdd("m", X, DTStartDate.value)   'Month
        ElseIf SystemOptions.ProcessPeriodType = 2 Then
        '    DTEnddate.value = DateAdd("yyyy", X, DTStartDate.value)   'Year
        ElseIf SystemOptions.ProcessPeriodType = 3 Then
       '     DTEnddate.value = DateAdd("ww", X, DTStartDate.value)   'week
        End If

    End With

    IntCounter = 0

    With VSFlexGrid3

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("ExpensesID")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
  
            End If

        Next i
   
    End With
    With Grid

        For i = .FixedRows To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("MofrdID"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
  
            End If

        Next i
   
    End With
        IntCounter = 0
Dim SunMAin As Double
Dim SunMAinEx As Double
SunMAin = 0
SunMAinEx = 0
    With Me.FgMainDes
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("Name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                SunMAin = SunMAin + val(.TextMatrix(i, .ColIndex("Total")))
                SunMAinEx = SunMAinEx + val(.TextMatrix(i, .ColIndex("PriceExe")))
            End If
        Next i
   
    End With
    lbl(34).Caption = Round(SunMAin, Decimal_Places)
    lbl(33).Caption = Round(SunMAinEx, Decimal_Places)
End Sub
'REFillOprData
Private Sub REFillOprData(Optional TblProcessDEFID As Double, Optional i As Long = 0)
    'Dim i As Integer
    Dim IntCounter As Integer
    Dim StartWeek As Double
    Dim EndWeek As Double
    Dim EarlyStartWeek As Double
    Dim EarlyEndWeek As Double
    Dim rs As ADODB.Recordset

    With VSFlexGrid2

     '   For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("OPRIDD")) <> "" Then
                sql = "   SELECT      dbo.TblProcessDEF.*, dbo.TblProcessUnites.UnitName, dbo.TblProcessUnites.UnitNamee "
                sql = sql & "  from dbo.TblProcessDEF  "
                sql = sql & "   INNER JOIN"
                sql = sql & " dbo.TblProcessUnites ON dbo.TblProcessDEF.UnitID = dbo.TblProcessUnites.UnitID"
                sql = sql & " Where (TblProcessDEFID = " & TblProcessDEFID & ")"
                                  
                Set rs = New ADODB.Recordset
                rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If rs.RecordCount > 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("unitname")) = IIf(IsNull(rs("unitname").value), "", rs("unitname").value)
                Else
                .TextMatrix(i, .ColIndex("unitname")) = IIf(IsNull(rs("unitnamee").value), "", rs("unitnamee").value)
                End If
                
                    Dim IntervalID As Integer
                    Dim intervaltype As String
                    IntervalID = IIf(IsNull(rs("Intervalid").value), 0, rs("Intervalid").value)
If SystemOptions.UserInterface = ArabicInterface Then
                    If IntervalID = 0 Then
                        intervaltype = "œﬁÌﬁ…"
                    ElseIf IntervalID = 1 Then
                        intervaltype = "”«⁄Â"
                    ElseIf IntervalID = 2 Then
                        intervaltype = "ÌÊ„"
                    ElseIf IntervalID = 3 Then
                        intervaltype = "«”»Ê⁄"
                    ElseIf IntervalID = 4 Then
                        intervaltype = "‘Â—"
                    ElseIf IntervalID = 5 Then
                        intervaltype = "”‰Â"
                    End If
Else
               If IntervalID = 0 Then
                        intervaltype = "Minute"
                    ElseIf IntervalID = 1 Then
                        intervaltype = "Hour"
                    ElseIf IntervalID = 2 Then
                        intervaltype = "Day"
                    ElseIf IntervalID = 3 Then
                        intervaltype = "Week"
                    ElseIf IntervalID = 4 Then
                        intervaltype = "Month"
                    ElseIf IntervalID = 5 Then
                        intervaltype = "Year"
                    End If
End If

                    .TextMatrix(i, .ColIndex("period")) = IIf(IsNull(rs("interval").value), 0, rs("interval").value) * val(.TextMatrix(i, .ColIndex("qty")))
                    .TextMatrix(i, .ColIndex("periodView")) = .TextMatrix(i, .ColIndex("period")) & "  " & intervaltype
                                                         
                End If
                             
            End If

       ' Next i
  
    End With

End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            StrSQL = "Delete From projects Where id=" & val(Me.txt_project_id.Text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
            clear_all Me
                
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "id='" & val(txt_project_id.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If txt_project_id.Text <> "" Then
        StrSQL = "select * From DOUBLE_ENTREY_VOUCHERS where project_id=" & val(txt_project_id.Text)
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

       If Not (RsTemp.EOF Or RsTemp.BOF) Then
             
         If SystemOptions.UserInterface = ArabicInterface Then
             Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·„‘—Ê⁄" & CHR(13)
              Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  „— »ÿ… »Â–« «·„‘—Ê⁄"
          Else
             Msg = "Can't Delete " & CHR(13)
              Msg = Msg + "This project have transactions"
              
          End If
          
            
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
             Exit Sub
         End If
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·„‘—Ê⁄ —ﬁ„ " & CHR(13)
        Msg = Msg + (txt_project_id.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
       Msg = "Delete Project No " & CHR(13)
        Msg = Msg + (txt_project_id.Text) & CHR(13)
        Msg = Msg + " Sure?"

End If

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                Dim StrAccountCode As String
                'StrAccountCode = rs("expanses_account").value
                
              StrAccountCode = IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value)
              If StrAccountCode <> "" Then
                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
               
                Else
                    Exit Sub
                End If
              End If
              If Option8.value = False Then
              If StrAccountCode <> "" Then
              StrAccountCode = IIf(IsNull(rs("REVENUE_account").value), "", rs("REVENUE_account").value)
 
                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
               
                Else
                    Exit Sub
                End If
            End If
              If StrAccountCode <> "" Then
                StrAccountCode = IIf(IsNull(rs("Material_account").value), "", rs("Material_account").value)
                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
               
                Else
                    Exit Sub
                End If
              End If
                StrAccountCode = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
            If StrAccountCode <> "" Then
                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
               
                Else
                    Exit Sub
                End If
            End If
                  StrAccountCode = IIf(IsNull(rs("legal").value), "", rs("legal").value)
             If StrAccountCode <> "" Then
                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
                
                Else
                    Exit Sub
                End If
             End If
           StrAccountCode = IIf(IsNull(rs("AcountGood").value), "", rs("AcountGood").value)
             If StrAccountCode <> "" Then
                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
                
                Else
                    Exit Sub
                End If
             End If
             
             Else
                   StrAccountCode = IIf(IsNull(rs("AccountUnderImp").value), "", rs("AccountUnderImp").value)
              If StrAccountCode <> "" Then
                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
               
                Else
                    Exit Sub
                End If
                End If
              End If
                StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.delete
            
                rs.MoveFirst
                 StrSQL = "Delete From ProjectMainDes Where ProjectID =" & val(Me.txt_project_id.Text)
                 Cn.Execute StrSQL, , adExecuteNoRecords
                  StrSQL = "Delete From TblProjectUser Where ProjectID =" & val(Me.txt_project_id.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            
                 StrSQL = "Delete From ProJectMofrd Where ProjID =" & val(Me.txt_project_id.Text)
                 Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "delete From projects where id =" & val(Me.txt_project_id.Text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                StrSQL = "delete From projects_des where project_id=" & val(Me.txt_project_id.Text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                   StrSQL = "Delete From TblExpensiveOper Where ProjectID=" & val(Me.txt_project_id.Text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblMatrials Where ProjectID=" & val(Me.txt_project_id.Text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblEmpOper Where ProjectID=" & val(Me.txt_project_id.Text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblEquepment Where ProjectID=" & val(Me.txt_project_id.Text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                     StrSQL = "Delete From terms_operations Where project_id=" & val(Me.txt_project_id.Text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
             
                If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ «·Õ–›"
                Else
                MsgBox "Delete Successfully"
                End If
                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                   ListUserSelect.Clear
                    Fg_Journal.Clear flexClearScrollable, flexClearEverything
                    Fg_Journal.Rows = 2
                    Fg_Journal.Enabled = True
                    FgMainDes.Clear flexClearScrollable, flexClearEverything
                    FgMainDes.Rows = 2
                    FgMainDes.Enabled = True
                    
                    txt_total_discount = 0
          
                    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
                    VSFlexGrid1.Rows = 2
                    'VSFlexGrid1.Enabled = True
          
                    VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
                    VSFlexGrid2.Rows = 2
                   ' VSFlexGrid2.Enabled = True
          
                        XPTxtCurrent.Caption = 0
                        XPTxtCount.Caption = 0
                Else
                    XPBtnMove_Click (0)
                End If
            End If
        End If

    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            Msg = "this operation is not available due to lack of records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·Œ“‰… "
            Msg = Msg & CHR(13) & Err.description
        Else
            Msg = "This record can't be deleted for data integration with safe " & CHR(13)
            Msg = Msg & CHR(13) & Err.description
        End If
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If
 
End Sub
Private Sub Add_Click()
    Dim i As Integer
    Dim Msg As String
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "”Ê› Ì „  €ÌÌ— «ﬂÊ«œ «·»‰Êœ Â·  —Ìœ «·„ «»⁄…"
Else
Msg = "Codes will be changed items"
End If
If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
    With Fg_Journal

        If Not .TextMatrix(.Row, .ColIndex("des")) = "" Then

            .AddItem " ", .Row
        End If

    End With
  
  End If
End Sub

Private Sub ALLButton1_Click()
print_report2 , 1
End Sub

Private Sub ALLButton2_Click()
print_report2
End Sub

Private Sub ALLButton3_Click()
      On Error Resume Next
    Dim StrFileName As String
    StrFileName = "C:\" & "\Payrolll.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If
  
      On Error Resume Next
      cd.CancelError = True 'allow escape key/cancel
     cd.filename = "Project"
    cd.ShowSave     'show the dialog screen
    If Err <> 32755 Then    ' User didn't chose Cancel.
   Else
       Exit Sub
    End If
 StrFileName = cd.filename & ".xls"
Me.Fg_Journal.saveGrid StrFileName, flexFileCustomText, True
   
    OpenFile StrFileName
End Sub

Private Sub ALLButton4_Click()
Frame21.Visible = True
End Sub

Private Sub BtnSalary_Click()
Frame20.Visible = True
'C1Elastic2.Visible = True
End Sub

Private Sub RemoveGridRowMofrd()
If Me.TxtModFlg.Text <> "R" Then

    With Me.Grid

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
    End If
End Sub
Private Sub RemoveGridRowMainDes()
If Me.TxtModFlg.Text <> "R" Then
    With Me.FgMainDes
    If SystemOptions.UserInterface = ArabicInterface Then
        StrMSG = "”Ê› Ì „ Õ–› ﬂ· «·»‰Êœ Ê«·⁄„·Ì«  «·„— »ÿÂ »Â–« «·»‰œ Â·  —Ìœ «·Õ–›"
    Else
        StrMSG = "All terms and process connected to this term will be deleted , Are you sure you want to continue"
    End If
        If MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title) = vbYes Then
        If .Row <= 0 Then Exit Sub
        RemoveGridRowMainDes2 val(.TextMatrix(.Row, .ColIndex("ID")))
        .RemoveItem .Row
          End If
    End With

    ReLineGrid
    End If
End Sub
Private Sub RemoveGridRowMainDes2(Optional PrMainDesID As Double)
If PrMainDesID <> 0 Then
Dim i As Long
If Me.TxtModFlg.Text <> "R" Then
    With Me.Fg_Journal
    i = 1
    Do Until (i >= .Rows)
    If val(.TextMatrix(i, .ColIndex("PrMainDesID"))) = PrMainDesID Then
        If i <= 0 Then Exit Sub
        RemoveGridRow22 i
        .RemoveItem i
        
    End If
    i = i + 1
     Loop
    End With
    ReLineGrid
    End If
   End If
End Sub

 

Private Sub ChAuto_Click()
If ChAuto.value = vbChecked Then
terms_operations_Click (0)
Frame11.Visible = False
hideallframe
   
          Frame5.Visible = True
End If
End Sub

Private Sub check4_Click()
If Check4.value = True Then
  
   DataCombo3.Text = ""
   DataCombo3.Enabled = True
  End If
End Sub



Private Sub Cmd_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
If Index = 9 Then
RemoveGridRowMofrd
ElseIf Index = 1 Then
RemoveGridRowMainDes
End If
End If
End Sub

's Private Sub Cmd_Click()
'sa RemoveGridRow
'saEnd Sub

Private Sub Cmdd_Click()
RemoveGridRow
End Sub

Private Sub CmdPand_Click()
RemoveGridRow2
End Sub

Private Sub CmdProcees_Click()
RemoveGridRow1
End Sub

Private Sub CMDViewGantt_Click(Index As Integer)
    terms_operations_Click (1)
    Gantt.show
   Gantt.Init_Chart val(TXTNoOFWeek.Text)
   Gantt.Draw_Data current_terms
End Sub

Private Sub Command1_Click(Index As Integer)
   ' On Error Resume Next
    Dim FirstPeriodDateInthisYear As Date

    Select Case Index

        Case 4
        hideallframe
            Fra(3).Visible = True
          
        Case 0
            'mod_flad = "N"
FlgOper = False
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
        
            TxtModFlg.Text = "N"
            clear_all Me
             Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 2
            RdTyp(0).value = True
            RdTyp_Click (0)
            Grid.Enabled = True
            FgMainDes.Clear flexClearScrollable, flexClearEverything
            FgMainDes.Rows = 2
            FgMainDes.Enabled = True
            ListUserSelect.Clear
            Option7.value = True
            Option6.Enabled = True
            Option7.Enabled = True
            Option8.Enabled = True
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
            Me.Dtp = FirstPeriodDateInthisYear
            Me.Dtp1 = FirstPeriodDateInthisYear
            Me.Dtp2 = FirstPeriodDateInthisYear
            Me.Dtp3 = FirstPeriodDateInthisYear
            Me.Dtp4 = FirstPeriodDateInthisYear
            Me.Dtp5 = FirstPeriodDateInthisYear
            Me.Dtp6 = FirstPeriodDateInthisYear
            Me.Dtp8 = FirstPeriodDateInthisYear
            
            OptType(2).value = True
            OptType1(2).value = True
            OptType2(2).value = True
            OptType3(2).value = True
            OptType4(2).value = True
            OptType5(2).value = True
            OptType6(2).value = True
            OptType8(2).value = True
ChAutoItems.value = vbUnchecked
ChAuto.value = vbUnchecked
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.Rows = 2
            Fg_Journal.Enabled = True
            txt_total_discount = 0 '
          
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 2
            VSFlexGrid1.Enabled = True
          
            VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid2.Rows = 2
           ' VSFlexGrid2.Enabled = True
          
                  GridSub.Clear flexClearScrollable, flexClearEverything
            GridSub.Rows = 2
            GridSub.Enabled = True
            
            Option4.value = True

           ' Me.txt_project_id.text = CStr(new_id("projects", "id", "", True))
            Command1(1).Enabled = True
            Me.dcBranch.BoundText = Current_branch
             Me.DCboUserName.BoundText = user_id
            XPDtbBill.value = Date
            DataCombo1.BoundText = 1
            DcCurrency.BoundText = 1
Ptype(0).value = True
 txt_project_id.Text = CStr(new_id("projects", "id", "", True))
rs.AddNew
rs("id").value = txt_project_id.Text
      
rs.update
CBoBasedON.ListIndex = 0
TxtProjectCosts.Text = 0
 If SystemOptions.ProjectUnderImplemen = True Then
 Option8.value = True
 End If
HidUnder
        Case 1
        If Me.OptType8(0).value = True Or Me.OptType8(1).value = True Then
        AccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DcAccount2.BoundText), "Account_Code2")
            If AccountCode = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» œ›⁄«  „ﬁœ„… ··⁄„Ì·", vbCritical
                Else
                    MsgBox "no selected account "
                End If
                Exit Sub
            End If
        End If
        If val(txt_sub_net.Text) <> val(total_after_discount.Text) Then
                   If SystemOptions.UserInterface = ArabicInterface Then
            '       MsgBox "ÌÃ» «‰  ﬂÊ‰ ﬁÌ„… «·„‘—Ê⁄ „”«ÊÌ… ··’«›Ì"
                   Else
            '       MsgBox "Can not project must be equal to the net value "
                   End If
            '       Exit Sub
        End If
        Dim Name1 As String
           If ChekID(Name1) = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰  ﬂ—«— ﬂÊœ «·„‘—Ê⁄ ·«‰Â „ÊÃÊœ „”»ﬁ« ·„‘—Ê⁄ " & " " & Name1
        Else
        MsgBox "Can Not Repaet Code This Code Already Exists in Project" & " " & Name1
        End If
        Exit Sub
        End If
        my_branch = Me.dcBranch.BoundText
        Me.DCboUserName.BoundText = user_id
            SaveData
TxtId.SetFocus

            'Fg_Journal.Enabled = False
        Case 2 '
      If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
            On Error Resume Next

 
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments Me.DCPreFix.Text & TxtId.Text, "1010201801"

'            If SystemOptions.UserInterface = EnglishInterface Then
'                If DCPreFix.Text & txtid.Text = "" Then MsgBox "Select Project firstly": Exit Sub
'
'            Else
'
'                If DCPreFix.Text & txtid.Text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— „‘—Ê⁄ «Ê·«": Exit Sub
'
'            End If

'            imaged.show

'            If SystemOptions.UserInterface = EnglishInterface Then
'
'                imaged.Label9.Caption = "Attachment For Project "
'                imaged.Caption = "Project Attachment  "
'                imaged.Label6.Caption = "   Project NO"
'                Label5.Caption = "Documents"
'
'            Else
'
'                imaged.Label9.Caption = "„—›ﬁ«    „‘—Ê⁄  —ﬁ„"
'                imaged.Caption = "„—›ﬁ«  „‘—Ê⁄  "
'                imaged.Label6.Caption = "—ﬁ„  «·„‘—Ê⁄"
'
'            End If
'
'            imaged.SUBJECT_NO = DCPreFix.Text & txtid.Text
'            imaged.txtopeation_type = "„—›ﬁ«  „‘—Ê⁄"
'
'            imaged.Adodc1.ConnectionString = Cn.ConnectionString
'            imaged.Adodc1.CommandType = adCmdText
'            imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  „‘—Ê⁄' and subject_no='" & DCPreFix.Text & txtid.Text & "'"
'            imaged.Adodc1.Refresh
'
'            If imaged.Adodc1.Recordset.RecordCount > 0 Then
'
'                imaged.DBPix201.Visible = True
'            Else
'                imaged.DBPix201.Visible = False
'            End If

        Case 3
      
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
            Me.Dtp = FirstPeriodDateInthisYear
            Me.Dtp1 = FirstPeriodDateInthisYear
            Me.Dtp2 = FirstPeriodDateInthisYear
            Me.Dtp3 = FirstPeriodDateInthisYear
            Me.Dtp4 = FirstPeriodDateInthisYear
            Me.Dtp5 = FirstPeriodDateInthisYear
            Me.Dtp6 = FirstPeriodDateInthisYear
            Me.Dtp8 = FirstPeriodDateInthisYear
            
            

            FgMainDes.Rows = FgMainDes.Rows + 1
            FgMainDes.Enabled = True
           Fg_Journal.Rows = Fg_Journal.Rows + 1
            Fg_Journal.Enabled = True
            VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
            VSFlexGrid1.Enabled = True
            VSFlexGrid2.Rows = VSFlexGrid2.Rows + 1
            VSFlexGrid2.Enabled = True
            VSFlexGrid3.Rows = VSFlexGrid3.Rows + 1
            VSFlexGrid3.Enabled = True
            GridSub.Rows = GridSub.Rows + 1
             Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
            GridSub.Enabled = True
            Command1(1).Enabled = True

            'SaveData
        Case 4
 
        Case 5
            FrmProjectSearch.Indx = 0
            Load FrmProjectSearch
            FrmProjectSearch.show vbModal


        Case 6
            Undo

        Case 7

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
        
            Del_Trans
            
    Case 8
hideallframe
   
            Fra(2).Visible = True
          
          
            
    End Select

End Sub
Function hideallframe()
   Frame11.Visible = False
     Frame5.Visible = False
     Frame2.Visible = False
    Fra(2).Visible = False
    Fra(3).Visible = False
    
    
End Function
Function print_report2(Optional Note As String, Optional indexx As Integer = 0)
   Dim Rs2 As ADODB.Recordset
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    sql = "select * from  ProjectMainDes where ProjectID=" & val(txt_project_id.Text) & " "
    Set Rs2 = New ADODB.Recordset
    Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs2.RecordCount > 0 Then
 MySQL = "   SELECT     TOP 100 PERCENT dbo.projects_des.oprid, dbo.projects_des.project_no, dbo.projects_des.project_name, dbo.projects_des.[index], dbo.projects_des.des,"
 MySQL = MySQL & "                     dbo.projects_des.qty, dbo.projects_des.cost, dbo.projects_des.total, dbo.projects_des.discount, dbo.projects_des.net, dbo.projects_des.project_id,"
 MySQL = MySQL & "                     dbo.projects_des.line_no, dbo.projects_des.sub_contractor_id, dbo.projects_des.fullcode, dbo.projects_des.Remark, dbo.terms_operations.total AS Oprtotal,"
 MySQL = MySQL & "                     dbo.terms_operations.name, dbo.terms_operations.period, dbo.terms_operations.term_fullcode, dbo.terms_operations.project_id AS Operproject_id,"
 MySQL = MySQL & "                     dbo.terms_operations.[count], dbo.terms_operations.salary, dbo.terms_operations.total_items, dbo.terms_operations.total_salary,"
 MySQL = MySQL & "                     dbo.terms_operations.total_expenses, dbo.terms_operations.fullcode AS OperFullcodes, dbo.terms_operations.ended, dbo.terms_operations.start_date,"
 MySQL = MySQL & "                     dbo.terms_operations.end_date, dbo.terms_operations.StartWeek, dbo.terms_operations.EndWeek, dbo.terms_operations.EarlyEndWeek,"
 MySQL = MySQL & "                     dbo.terms_operations.EarlyStartWeek, dbo.terms_operations.Period1, dbo.terms_operations.Critical, dbo.terms_operations.Symbol, dbo.terms_operations.Pre,"
 MySQL = MySQL & "                     dbo.terms_operations.EarlyStartDate, dbo.terms_operations.EarlyEndDate, dbo.terms_operations.qty AS Operqty, dbo.terms_operations.periodView,"
 MySQL = MySQL & "                     dbo.terms_operations.Actperiod, dbo.terms_operations.unitid, dbo.terms_operations.unitname, dbo.terms_operations.ProjectDes_ID, dbo.terms_operations.expen,"
 MySQL = MySQL & "                     dbo.terms_operations.eq, dbo.terms_operations.emps, dbo.terms_operations.matrials, dbo.terms_operations.EquepVal, dbo.terms_operations.hourval,"
 MySQL = MySQL & "                     dbo.terms_operations.item_id, dbo.terms_operations.OPRIDD, dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEF.ProcessName,"
 MySQL = MySQL & "                     dbo.TblProcessDEF.ProcessNameE, dbo.TblProcessDEF.UnitID AS UUnitID, dbo.TblProcessUnites.UnitName AS UUnitName, dbo.TblProcessUnites.UnitNamee,"
 MySQL = MySQL & "                     dbo.TblEmpOper.ID AS EmpOID, dbo.TblEmpOper.ProjectID AS EmpOProjectID, dbo.TblEmpOper.Pand AS EmpOPand, dbo.TblEmpOper.Opr AS EmpOOpr,"
 MySQL = MySQL & "                     dbo.TblEmpOper.daysalary, dbo.TblEmpOper.[Count] AS EmpOCount, dbo.TblEmpOper.OperCode, dbo.TblEmpOper.EmpID, TblEmployee_3.Emp_Code,"
 MySQL = MySQL & "                     TblEmployee_3.Emp_Name, TblEmployee_3.Emp_Name1, TblEmployee_3.Emp_Name2, TblEmployee_3.Emp_Name3, TblEmployee_3.Emp_Name4,"
 MySQL = MySQL & "                     TblEmployee_3.Fullcode AS EmpOFullcode, TblEmployee_3.Emp_Namee4, TblEmployee_3.Emp_Namee3, TblEmployee_3.Emp_Namee2,"
 MySQL = MySQL & "                     TblEmployee_3.Emp_Namee1, TblEmployee_3.Emp_Namee, dbo.TblEmpOper.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
 MySQL = MySQL & "                     dbo.TblExpensiveOper.ID AS ExID, dbo.TblExpensiveOper.ProjectID AS ExProjectID, dbo.TblExpensiveOper.Pand AS ExPand, dbo.TblExpensiveOper.Opr AS ExOpr,"
 MySQL = MySQL & "                     dbo.TblExpensiveOper.EsToal, dbo.TblExpensiveOper.[value], dbo.TblExpensiveOper.Des AS ExDes, dbo.TblExpensiveOper.OperCode AS ExOperCode,"
 MySQL = MySQL & "                     dbo.TblExpensiveOper.AccountCode, dbo.Expenses_accounts.Account_Name, dbo.terms_operations.id, dbo.TblEquepment.ID AS EqID,"
 MySQL = MySQL & "                     dbo.TblEquepment.ProjectID AS EqProjectID, dbo.TblEquepment.Pand AS EqPand, dbo.TblEquepment.Opr AS EqOpr, dbo.TblEquepment.EstHour,"
 MySQL = MySQL & "                     dbo.TblEquepment.ActualHour, dbo.TblEquepment.TotalEs, dbo.TblEquepment.[value] AS Eqvalue, dbo.TblEquepment.des AS Eqdes,"
 MySQL = MySQL & "                     dbo.TblEquepment.OperCode AS EqOperCode, dbo.TblEquepment.EquepVal AS EqEquepVal, dbo.TblEquepment.ExpensesID, dbo.FixedAssets.code,"
 MySQL = MySQL & "                     dbo.FixedAssets.Name AS EqName, dbo.FixedAssets.namee AS EqNameE, dbo.TblMatrials.ID AS MatID, dbo.TblMatrials.Pand AS MatPand,"
 MySQL = MySQL & "                     dbo.TblMatrials.[Count] AS MatCount, dbo.TblMatrials.Price, dbo.TblMatrials.Quntapro, dbo.TblMatrials.priceapro, dbo.TblMatrials.ProjectID AS MatProjectID,"
 MySQL = MySQL & "                     dbo.TblMatrials.OperCode AS MatOperCode, dbo.TblMatrials.Opr AS MatOpr, dbo.TblMatrials.ItemID, dbo.TblItems.Fullcode AS ItemFullcode, dbo.TblItems.ItemCode,"
 MySQL = MySQL & "                     dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.projects.id AS projectsID, dbo.projects.End_user_Account, dbo.projects.End_user_name,"
 MySQL = MySQL & "                     dbo.projects.sub_contractor_Account, dbo.projects.sub_contractor_name, dbo.projects.Fullcode AS projectsFullcode, dbo.projects.prifix,"
 MySQL = MySQL & "                     dbo.projects.Code AS projectsCode, dbo.projects.Project_name AS projects_Project_name, dbo.projects.Project_status, dbo.projects.branch_no,"
 MySQL = MySQL & "                     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.projects.Contract_type, dbo.contract_type.name AS contract_type_name,"
 MySQL = MySQL & "                     dbo.projects.CurrencyID, dbo.currency.code AS currency_code, dbo.currency.name AS currency_name, dbo.projects.Project_nameE, dbo.projects.project_cost,"
 MySQL = MySQL & "                     dbo.projects.general_discount, dbo.projects.cost_after_discount, dbo.projects.DiscountPercentage, dbo.projects.StartDate, dbo.projects.EndDate,"
 MySQL = MySQL & "                     dbo.projects.End_user_id, TblCustemers_1.CusName, TblCustemers_1.CusNamee, TblCustemers_1.Fullcode AS Cus_Fullcode,"
 MySQL = MySQL & "                     dbo.projects.sub_contractor_id AS projects_sub_contractor_id, TblCustemers_1.CusName AS Cus_CusName, TblCustemers_1.CusNamee AS Cus_CusNameE,"
 MySQL = MySQL & "                      TblCustemers_1.Fullcode AS Cus_Fullcode1, dbo.projects.EmpId AS projects_empid, TblEmployee_1.Emp_Code AS M_Emp_Code,"
 MySQL = MySQL & "                     TblEmployee_1.Emp_Name AS M_Emp_Name, TblEmployee_1.Emp_Name1 AS M_Emp_Name1, TblEmployee_1.Emp_Name2 AS M_Emp_Name2,"
 MySQL = MySQL & "                     TblEmployee_1.Emp_Name3 AS M_Emp_Name3, TblEmployee_1.Emp_Name4 AS M_Emp_Name4, TblEmployee_1.Fullcode AS M_Fullcode,"
 MySQL = MySQL & "                     TblEmployee_1.Emp_Namee4 AS M_Emp_Namee4, TblEmployee_1.Emp_Namee3 AS M_Emp_Namee3, TblEmployee_1.Emp_Namee2 AS M_Emp_Namee2,"
 MySQL = MySQL & "                     TblEmployee_1.Emp_Namee1 AS M_Emp_Namee1, TblEmployee_1.Emp_Namee AS M_Emp_Namee, dbo.projects.EmpId1, TblEmployee_2.CustNum AS De_DcEmp1,"
 MySQL = MySQL & "                     TblEmployee_2.Emp_Name AS De_Emp_Name, TblEmployee_2.Emp_Name1 AS De_Emp_Name1, TblEmployee_2.Emp_Name2 AS De_Emp_Name2,"
 MySQL = MySQL & "                     TblEmployee_2.Emp_Name3 AS De_Emp_Name3, TblEmployee_2.Emp_Name4 AS De_Emp_Name4, TblEmployee_2.Fullcode AS Del_Fullcode,"
 MySQL = MySQL & "                     TblEmployee_2.Emp_Namee4 AS De_Emp_Namee4, TblEmployee_2.Emp_Namee3 AS De_Emp_Namee3, TblEmployee_2.Emp_Namee2 AS De_Emp_Namee2,"
 MySQL = MySQL & "                     TblEmployee_2.Emp_Namee1 AS De_Emp_Namee1, TblEmployee_2.Emp_Namee AS De_Emp_Namee, dbo.projects.Dept_ID, dbo.TblSection.name AS Sec_name,"
 MySQL = MySQL & "                     dbo.TblSection.namee AS Sec_nameE, dbo.projects.Remarkss, dbo.projects.DpNearEndDate, TblCustemers_1.CusID, TblCustemers_2.CusName AS EndCusName,"
 MySQL = MySQL & "                     TblCustemers_2.CusNamee AS EndCusNameE, TblCustemers_2.Fullcode AS EndFullcode, dbo.currency.nameE, dbo.contract_type.namee AS contract_type_nameE,"
 MySQL = MySQL & "                     dbo.project_status.name AS StatusName, dbo.project_status.namee AS StatusNameE, dbo.Expenses_accounts_eng.Account_NameEng, dbo.projects.TotalMainDes,"
 MySQL = MySQL & "                     dbo.projects_des.PrMainDesID, dbo.ProjectMainDes.ID AS PandMID, dbo.ProjectMainDes.Name AS PandMName, dbo.ProjectMainDes.FullCode AS PandMFullcode,"
 MySQL = MySQL & "                     dbo.ProjectMainDes.Qty AS PandMQty, dbo.ProjectMainDes.Price AS PandMPrice, dbo.ProjectMainDes.Total AS PandMTotal, dbo.ProjectMainDes.Remarks,"
 MySQL = MySQL & "                     dbo.ProjectMainDes.QtyNo , dbo.ProjectMainDes.QtyExe, dbo.ProjectMainDes.PriceExe, dbo.ProjectMainDes.TotalExe, dbo.projects_des.CodeBand"
 MySQL = MySQL & "      FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.projects_des RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.ProjectMainDes ON dbo.projects_des.PrMainDesID = dbo.ProjectMainDes.ID RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.projects ON dbo.ProjectMainDes.ProjectID = dbo.projects.id AND dbo.projects_des.project_id = dbo.projects.id LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.project_status ON dbo.projects.Project_status = dbo.project_status.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblSection ON dbo.projects.Dept_ID = dbo.TblSection.Id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmployee TblEmployee_2 ON dbo.projects.EmpId1 = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmployee TblEmployee_1 ON dbo.projects.EmpId = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblCustemers TblCustemers_1 ON dbo.projects.sub_contractor_id = TblCustemers_1.CusID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblCustemers TblCustemers_2 ON dbo.projects.End_user_id = TblCustemers_2.CusID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.currency ON dbo.projects.CurrencyID = dbo.currency.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.contract_type ON dbo.projects.Contract_type = dbo.contract_type.id ON dbo.TblBranchesData.branch_id = dbo.projects.branch_no LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblProcessUnites LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblProcessDEF ON dbo.TblProcessUnites.UnitID = dbo.TblProcessDEF.UnitID RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.Expenses_accounts RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.Expenses_accounts_eng RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblExpensiveOper ON dbo.Expenses_accounts_eng.Account_Code = dbo.TblExpensiveOper.AccountCode ON"
 MySQL = MySQL & "                     dbo.Expenses_accounts.Account_Code = dbo.TblExpensiveOper.AccountCode RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblItems RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblMatrials ON dbo.TblItems.ItemID = dbo.TblMatrials.ItemID RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.terms_operations ON dbo.TblMatrials.Opr = dbo.terms_operations.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEquepment LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.FixedAssets ON dbo.TblEquepment.ExpensesID = dbo.FixedAssets.id ON dbo.terms_operations.id = dbo.TblEquepment.Opr ON"
 MySQL = MySQL & "                     dbo.TblExpensiveOper.Opr = dbo.terms_operations.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmpOper LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmployee TblEmployee_3 ON dbo.TblEmpOper.EmpID = TblEmployee_3.Emp_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmpJobsTypes ON dbo.TblEmpOper.JobID = dbo.TblEmpJobsTypes.JobTypeID ON dbo.terms_operations.id = dbo.TblEmpOper.Opr ON"
 MySQL = MySQL & "                     dbo.TblProcessDEF.TblProcessDEFID = dbo.terms_operations.OPRIDD ON dbo.projects_des.oprid = dbo.terms_operations.ProjectDes_ID"
    Else
MySQL = " SELECT     TOP 100 PERCENT dbo.projects_des.oprid, dbo.projects_des.project_no, dbo.projects_des.project_name, dbo.projects_des.[index], dbo.projects_des.des, "
MySQL = MySQL & "                      dbo.projects_des.qty, dbo.projects_des.cost, dbo.projects_des.total, dbo.projects_des.discount, dbo.projects_des.net, dbo.projects_des.project_id,"
MySQL = MySQL & "                      dbo.projects_des.line_no, dbo.projects_des.sub_contractor_id, dbo.projects_des.fullcode, dbo.projects_des.Remark, dbo.terms_operations.total AS Oprtotal,"
MySQL = MySQL & "                      dbo.terms_operations.name, dbo.terms_operations.period, dbo.terms_operations.term_fullcode, dbo.terms_operations.project_id AS Operproject_id,"
MySQL = MySQL & "                      dbo.terms_operations.[count], dbo.terms_operations.salary, dbo.terms_operations.total_items, dbo.terms_operations.total_salary,"
MySQL = MySQL & "                      dbo.terms_operations.total_expenses, dbo.terms_operations.fullcode AS OperFullcodes, dbo.terms_operations.ended, dbo.terms_operations.start_date,"
MySQL = MySQL & "                      dbo.terms_operations.end_date, dbo.terms_operations.StartWeek, dbo.terms_operations.EndWeek, dbo.terms_operations.EarlyEndWeek,"
MySQL = MySQL & "                      dbo.terms_operations.EarlyStartWeek, dbo.terms_operations.Period1, dbo.terms_operations.Critical, dbo.terms_operations.Symbol, dbo.terms_operations.Pre,"
MySQL = MySQL & "                      dbo.terms_operations.EarlyStartDate, dbo.terms_operations.EarlyEndDate, dbo.terms_operations.qty AS Operqty, dbo.terms_operations.periodView,"
MySQL = MySQL & "                      dbo.terms_operations.Actperiod, dbo.terms_operations.unitid, dbo.terms_operations.unitname, dbo.terms_operations.ProjectDes_ID, dbo.terms_operations.expen,"
MySQL = MySQL & "                      dbo.terms_operations.eq, dbo.terms_operations.emps, dbo.terms_operations.matrials, dbo.terms_operations.EquepVal, dbo.terms_operations.hourval,"
MySQL = MySQL & "                      dbo.terms_operations.item_id, dbo.terms_operations.OPRIDD, dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEF.ProcessName,"
MySQL = MySQL & "                      dbo.TblProcessDEF.ProcessNameE, dbo.TblProcessDEF.UnitID AS UUnitID, dbo.TblProcessUnites.UnitName AS UUnitName, dbo.TblProcessUnites.UnitNamee,"
MySQL = MySQL & "                      dbo.TblEmpOper.ID AS EmpOID, dbo.TblEmpOper.ProjectID AS EmpOProjectID, dbo.TblEmpOper.Pand AS EmpOPand, dbo.TblEmpOper.Opr AS EmpOOpr,"
MySQL = MySQL & "                      dbo.TblEmpOper.daysalary, dbo.TblEmpOper.[Count] AS EmpOCount, dbo.TblEmpOper.OperCode, dbo.TblEmpOper.EmpID, TblEmployee_3.Emp_Code,"
MySQL = MySQL & "                      TblEmployee_3.Emp_Name, TblEmployee_3.Emp_Name1, TblEmployee_3.Emp_Name2, TblEmployee_3.Emp_Name3, TblEmployee_3.Emp_Name4,"
MySQL = MySQL & "                      TblEmployee_3.Fullcode AS EmpOFullcode, TblEmployee_3.Emp_Namee4, TblEmployee_3.Emp_Namee3, TblEmployee_3.Emp_Namee2,"
MySQL = MySQL & "                      TblEmployee_3.Emp_Namee1, TblEmployee_3.Emp_Namee, dbo.TblEmpOper.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
MySQL = MySQL & "                      dbo.TblExpensiveOper.ID AS ExID, dbo.TblExpensiveOper.ProjectID AS ExProjectID, dbo.TblExpensiveOper.Pand AS ExPand, dbo.TblExpensiveOper.Opr AS ExOpr,"
MySQL = MySQL & "                      dbo.TblExpensiveOper.EsToal, dbo.TblExpensiveOper.[value], dbo.TblExpensiveOper.Des AS ExDes, dbo.TblExpensiveOper.OperCode AS ExOperCode,"
MySQL = MySQL & "                      dbo.TblExpensiveOper.AccountCode, dbo.Expenses_accounts.Account_Name, dbo.terms_operations.id, dbo.TblEquepment.ID AS EqID,"
MySQL = MySQL & "                      dbo.TblEquepment.ProjectID AS EqProjectID, dbo.TblEquepment.Pand AS EqPand, dbo.TblEquepment.Opr AS EqOpr, dbo.TblEquepment.EstHour,"
MySQL = MySQL & "                      dbo.TblEquepment.ActualHour, dbo.TblEquepment.TotalEs, dbo.TblEquepment.[value] AS Eqvalue, dbo.TblEquepment.des AS Eqdes,"
MySQL = MySQL & "                      dbo.TblEquepment.OperCode AS EqOperCode, dbo.TblEquepment.EquepVal AS EqEquepVal, dbo.TblEquepment.ExpensesID, dbo.FixedAssets.code,"
MySQL = MySQL & "                      dbo.FixedAssets.Name AS EqName, dbo.FixedAssets.namee AS EqNameE, dbo.TblMatrials.ID AS MatID, dbo.TblMatrials.Pand AS MatPand,"
MySQL = MySQL & "                      dbo.TblMatrials.[Count] AS MatCount, dbo.TblMatrials.Price, dbo.TblMatrials.Quntapro, dbo.TblMatrials.priceapro, dbo.TblMatrials.ProjectID AS MatProjectID,"
MySQL = MySQL & "                      dbo.TblMatrials.OperCode AS MatOperCode, dbo.TblMatrials.Opr AS MatOpr, dbo.TblMatrials.ItemID, dbo.TblItems.Fullcode AS ItemFullcode, dbo.TblItems.ItemCode,"
MySQL = MySQL & "                      dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.projects.id AS projectsID, dbo.projects.End_user_Account, dbo.projects.End_user_name,"
MySQL = MySQL & "                      dbo.projects.sub_contractor_Account, dbo.projects.sub_contractor_name, dbo.projects.Fullcode AS projectsFullcode, dbo.projects.prifix,"
MySQL = MySQL & "                      dbo.projects.Code AS projectsCode, dbo.projects.Project_name AS projects_Project_name, dbo.projects.Project_status, dbo.projects.branch_no,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.projects.Contract_type, dbo.contract_type.name AS contract_type_name,"
MySQL = MySQL & "                      dbo.projects.CurrencyID, dbo.currency.code AS currency_code, dbo.currency.name AS currency_name, dbo.projects.Project_nameE, dbo.projects.project_cost,"
MySQL = MySQL & "                      dbo.projects.general_discount, dbo.projects.cost_after_discount, dbo.projects.DiscountPercentage, dbo.projects.StartDate, dbo.projects.EndDate,"
MySQL = MySQL & "                      dbo.projects.End_user_id, TblCustemers_1.CusName, TblCustemers_1.CusNamee, TblCustemers_1.Fullcode AS Cus_Fullcode,"
MySQL = MySQL & "                      dbo.projects.sub_contractor_id AS projects_sub_contractor_id, TblCustemers_1.CusName AS Cus_CusName, TblCustemers_1.CusNamee AS Cus_CusNameE,"
MySQL = MySQL & "                      TblCustemers_1.Fullcode AS Cus_Fullcode1, dbo.projects.EmpId AS projects_empid, TblEmployee_1.Emp_Code AS M_Emp_Code,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Name AS M_Emp_Name, TblEmployee_1.Emp_Name1 AS M_Emp_Name1, TblEmployee_1.Emp_Name2 AS M_Emp_Name2,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Name3 AS M_Emp_Name3, TblEmployee_1.Emp_Name4 AS M_Emp_Name4, TblEmployee_1.Fullcode AS M_Fullcode,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Namee4 AS M_Emp_Namee4, TblEmployee_1.Emp_Namee3 AS M_Emp_Namee3, TblEmployee_1.Emp_Namee2 AS M_Emp_Namee2,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Namee1 AS M_Emp_Namee1, TblEmployee_1.Emp_Namee AS M_Emp_Namee, dbo.projects.EmpId1, TblEmployee_2.CustNum AS De_DcEmp1,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Name AS De_Emp_Name, TblEmployee_2.Emp_Name1 AS De_Emp_Name1, TblEmployee_2.Emp_Name2 AS De_Emp_Name2,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Name3 AS De_Emp_Name3, TblEmployee_2.Emp_Name4 AS De_Emp_Name4, TblEmployee_2.Fullcode AS Del_Fullcode,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Namee4 AS De_Emp_Namee4, TblEmployee_2.Emp_Namee3 AS De_Emp_Namee3, TblEmployee_2.Emp_Namee2 AS De_Emp_Namee2,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Namee1 AS De_Emp_Namee1, TblEmployee_2.Emp_Namee AS De_Emp_Namee, dbo.projects.Dept_ID, dbo.TblSection.name AS Sec_name,"
MySQL = MySQL & "                      dbo.TblSection.namee AS Sec_nameE, dbo.projects.Remarkss, dbo.projects.DpNearEndDate, TblCustemers_1.CusID, TblCustemers_2.CusName AS EndCusName,"
MySQL = MySQL & "                      TblCustemers_2.CusNamee AS EndCusNameE, TblCustemers_2.Fullcode AS EndFullcode, dbo.currency.nameE, dbo.contract_type.namee AS contract_type_nameE,"
MySQL = MySQL & "                      dbo.project_status.name AS StatusName, dbo.project_status.namee AS StatusNameE, dbo.Expenses_accounts_eng.Account_NameEng, dbo.projects.TotalMainDes,"
MySQL = MySQL & "                      dbo.projects_des.PrMainDesID, dbo.ProjectMainDes.ID AS PandMID, dbo.ProjectMainDes.Name AS PandMName, dbo.ProjectMainDes.FullCode AS PandMFullcode,"
MySQL = MySQL & "                      dbo.ProjectMainDes.Qty AS PandMQty, dbo.ProjectMainDes.Price AS PandMPrice, dbo.ProjectMainDes.Total AS PandMTotal, dbo.ProjectMainDes.Remarks,"
MySQL = MySQL & "                      dbo.ProjectMainDes.QtyNo , dbo.ProjectMainDes.QtyExe, dbo.ProjectMainDes.PriceExe, dbo.ProjectMainDes.TotalExe, dbo.projects_des.CodeBand"
MySQL = MySQL & " FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.projects LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.ProjectMainDes ON dbo.projects.id = dbo.ProjectMainDes.ProjectID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.projects_des ON dbo.projects.id = dbo.projects_des.project_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.project_status ON dbo.projects.Project_status = dbo.project_status.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblSection ON dbo.projects.Dept_ID = dbo.TblSection.Id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_2 ON dbo.projects.EmpId1 = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.projects.EmpId = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers TblCustemers_1 ON dbo.projects.sub_contractor_id = TblCustemers_1.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers TblCustemers_2 ON dbo.projects.End_user_id = TblCustemers_2.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.currency ON dbo.projects.CurrencyID = dbo.currency.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.contract_type ON dbo.projects.Contract_type = dbo.contract_type.id ON dbo.TblBranchesData.branch_id = dbo.projects.branch_no LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblProcessUnites LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblProcessDEF ON dbo.TblProcessUnites.UnitID = dbo.TblProcessDEF.UnitID RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.Expenses_accounts RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.Expenses_accounts_eng RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblExpensiveOper ON dbo.Expenses_accounts_eng.Account_Code = dbo.TblExpensiveOper.AccountCode ON"
MySQL = MySQL & "                      dbo.Expenses_accounts.Account_Code = dbo.TblExpensiveOper.AccountCode RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblItems RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblMatrials ON dbo.TblItems.ItemID = dbo.TblMatrials.ItemID RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.terms_operations ON dbo.TblMatrials.Opr = dbo.terms_operations.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEquepment LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets ON dbo.TblEquepment.ExpensesID = dbo.FixedAssets.id ON dbo.terms_operations.id = dbo.TblEquepment.Opr ON"
MySQL = MySQL & "                      dbo.TblExpensiveOper.Opr = dbo.terms_operations.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpOper LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_3 ON dbo.TblEmpOper.EmpID = TblEmployee_3.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.TblEmpOper.JobID = dbo.TblEmpJobsTypes.JobTypeID ON dbo.terms_operations.id = dbo.TblEmpOper.Opr ON"
MySQL = MySQL & "                      dbo.TblProcessDEF.TblProcessDEFID = dbo.terms_operations.OPRIDD ON dbo.projects_des.oprid = dbo.terms_operations.ProjectDes_ID"
End If
MySQL = MySQL & "  Where (dbo.projects.ID = " & val(txt_project_id.Text) & ") and  (NOT (dbo.projects_des.oprid IS NULL)) "
MySQL = MySQL & "  order by dbo.projects.ID ,dbo.projects_des.oprid "

If indexx = 1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessofProject1.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessofProject1E.rpt"
            
       End If
  Else
      If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessofProject2.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessofProject2E.rpt"
            
       End If
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
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
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
Function print_report(Optional ProjectID As Double = 0, Optional pandid As Double = 0)
   
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.projects_des.oprid, dbo.projects_des.project_no, dbo.projects_des.project_name, dbo.projects_des.[index], dbo.projects_des.des, dbo.projects_des.qty, "
MySQL = MySQL & "                      dbo.projects_des.cost, dbo.projects_des.total, dbo.projects_des.discount, dbo.projects_des.net, dbo.projects_des.project_id, dbo.projects_des.line_no,"
MySQL = MySQL & "                      dbo.projects_des.sub_contractor_id, dbo.projects_des.fullcode, dbo.projects_des.Remark, dbo.terms_operations.total AS Oprtotal, dbo.terms_operations.name,"
MySQL = MySQL & "                      dbo.terms_operations.id, dbo.terms_operations.period, dbo.terms_operations.term_fullcode, dbo.terms_operations.project_id AS Operproject_id,"
MySQL = MySQL & "                      dbo.terms_operations.[count], dbo.terms_operations.salary, dbo.terms_operations.total_items, dbo.terms_operations.total_salary,"
MySQL = MySQL & "                      dbo.terms_operations.total_expenses, dbo.terms_operations.fullcode AS OperFullcodes, dbo.terms_operations.ended, dbo.terms_operations.start_date,"
MySQL = MySQL & "                      dbo.terms_operations.end_date, dbo.terms_operations.StartWeek, dbo.terms_operations.EndWeek, dbo.terms_operations.EarlyEndWeek,"
MySQL = MySQL & "                      dbo.terms_operations.EarlyStartWeek, dbo.terms_operations.Period1, dbo.terms_operations.Critical, dbo.terms_operations.Symbol, dbo.terms_operations.Pre,"
MySQL = MySQL & "                      dbo.terms_operations.EarlyStartDate, dbo.terms_operations.EarlyEndDate, dbo.terms_operations.qty AS Expr4, dbo.terms_operations.periodView,"
MySQL = MySQL & "                      dbo.terms_operations.Actperiod, dbo.terms_operations.unitid, dbo.terms_operations.unitname, dbo.terms_operations.ProjectDes_ID, dbo.terms_operations.expen,"
MySQL = MySQL & "                      dbo.terms_operations.eq, dbo.terms_operations.emps, dbo.terms_operations.matrials, dbo.terms_operations.EquepVal, dbo.terms_operations.hourval,"
MySQL = MySQL & "                      dbo.terms_operations.item_id, dbo.terms_operations.OPRIDD, dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEF.ProcessName,"
MySQL = MySQL & "                      dbo.TblProcessDEF.ProcessNameE, dbo.TblProcessDEF.UnitID AS UUnitID, dbo.TblProcessUnites.UnitName AS UUnitName, dbo.TblProcessUnites.UnitNamee"
MySQL = MySQL & " FROM         dbo.TblProcessUnites LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblProcessDEF ON dbo.TblProcessUnites.UnitID = dbo.TblProcessDEF.UnitID RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.terms_operations ON dbo.TblProcessDEF.TblProcessDEFID = dbo.terms_operations.OPRIDD RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.projects_des ON dbo.terms_operations.ProjectDes_ID = dbo.projects_des.oprid"
MySQL = MySQL & "  Where (dbo.projects_des.oprid = " & pandid & ") And (dbo.projects_des.project_id = " & ProjectID & ")"



        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessOfPand.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessOfPandE.rpt"
            
       End If
           
       


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
        'GetMsgs 138, vbExclamation
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If
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
Function print_reportPand(Optional project_id As Double = 0, Optional PrMainDesID As Double = 0)
   
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.ProjectMainDes.Name, dbo.ProjectMainDes.FullCode AS MainFullCode, dbo.ProjectMainDes.Qty AS MainQty, dbo.ProjectMainDes.Price, "
MySQL = MySQL & "                       dbo.ProjectMainDes.Total AS MainTotal, dbo.ProjectMainDes.Remarks, dbo.ProjectMainDes.QtyNo AS MainQtyNo, dbo.ProjectMainDes.QtyExe AS MainQtyExe,"
MySQL = MySQL & "                        dbo.ProjectMainDes.PriceExe AS MainPriceExe, dbo.ProjectMainDes.TotalExe AS MainTotalExe, dbo.projects_des.*"
MySQL = MySQL & "   FROM         dbo.projects_des LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.ProjectMainDes ON dbo.projects_des.PrMainDesID = dbo.ProjectMainDes.ID"
MySQL = MySQL & "  Where (Not (dbo.projects_des.oprid Is Null)) and dbo.projects_des.project_id=" & project_id & " and dbo.projects_des.PrMainDesID=" & PrMainDesID & ""
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessOfPandMain.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessOfPandMainE.rpt"
            
       End If
           
    
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
        'GetMsgs 138, vbExclamation
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If
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
Sub CretAccount2()
        Account_Code_dynamic1 = get_account_code_branch(14, my_branch)
        Account_Code_dynamic2 = get_account_code_branch(15, my_branch)
        Account_Code_dynamic3 = get_account_code_branch(27, my_branch)
        Account_Code_dynamic4 = get_account_code_branch(28, my_branch)
        Account_Code_dynamic5 = get_account_code_branch(32, my_branch)
        Account_Code_dynamic7 = get_account_code_branch(152, my_branch)
        Account_Code_dynamic1C = ModAccounts.AddNewAccount(Account_Code_dynamic1, TXTprojectname & " -„’—Ê›«  ", True, False, TXTprojectnamee & " -EXPANSES")
        Account_Code_dynamic2C = ModAccounts.AddNewAccount(Account_Code_dynamic2, TXTprojectname & "-«Ì—«œ«  ", True, False, TXTprojectnamee & " -REVENUE")
        Account_Code_dynamic3C = ModAccounts.AddNewAccount(Account_Code_dynamic3, TXTprojectname & " -„Ê«œ  ", True, False, TXTprojectnamee & " -Material ")
        Account_Code_dynamic4C = ModAccounts.AddNewAccount(Account_Code_dynamic4, TXTprojectname & " -«ÃÊ— ", True, False, TXTprojectnamee & " -salary")
        Account_Code_dynamic5C = ModAccounts.AddNewAccount(Account_Code_dynamic5, TXTprojectname & " -„” Œ·’«  ", True, False, TXTprojectnamee & " -legal")
        If SystemOptions.AllowGoodPerfAccount = True Then
        Account_Code_dynamic7C = ModAccounts.AddNewAccount(Account_Code_dynamic7, TXTprojectname & " -Õ”‰ «·«œ«¡ ", True, False, TXTprojectnamee & " -Good performance")
        End If
End Sub
Function create_accounts() As Boolean
    Dim RsSavRec As ADODB.Recordset
    Dim RsSavRec1 As ADODB.Recordset
    Dim RsSavRec2 As ADODB.Recordset
    Dim RsSavRec3 As ADODB.Recordset
    Dim RsSavRec4 As ADODB.Recordset
    Dim RsSavRec5 As ADODB.Recordset
    Dim RsSavRec6 As ADODB.Recordset
    Dim My_SQL As String

    If 1 = 1 Then
    If Option8.value = True Then
          Account_Code_dynamic6 = get_account_code_branch(142, my_branch)
        
        If Account_Code_dynamic6 = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "Branch was not created", vbCritical
            End If
        
            create_accounts = False
            Exit Function
        Else

            If Account_Code_dynamic6 = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  Õ  «· ‰›Ì–  ··„‘«—Ì⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "no selected account for project under implementation  for this process in this branch"
                End If
        
                create_accounts = False
                Exit Function
            End If
        End If
        create_accounts = True
      Else
      
        Account_Code_dynamic2 = get_account_code_branch(15, my_branch)
        
        If Account_Code_dynamic2 = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "No branch was created", vbCritical
            End If
        
            create_accounts = False
            Exit Function
        Else

            If Account_Code_dynamic2 = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «Ì—«œ«  ··„‘«—Ì⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "no selected account for project revenue in this process"
                End If
                create_accounts = False
                Exit Function
            End If
        End If
        
         
        Account_Code_dynamic1 = get_account_code_branch(14, my_branch)
        
        If Account_Code_dynamic1 = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "No branch was created", vbCritical
            End If
            create_accounts = False
            Exit Function
        Else

            If Account_Code_dynamic1 = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» „’—Ê›«   ··„‘«—Ì⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "No selected account for project expenses in this process", vbCritical
                End If
                create_accounts = False
                Exit Function
            End If
        End If
        
        Account_Code_dynamic2 = get_account_code_branch(15, my_branch)
        
        If Account_Code_dynamic2 = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "No Branch was created", vbCritical
            End If
        
            create_accounts = False
            Exit Function
        Else

            If Account_Code_dynamic2 = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «Ì—«œ«  ··„‘«—Ì⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "No selected account for project revenue in the branch for this process ", vbCritical
                End If
                create_accounts = False
                Exit Function
            End If
        End If
        
        Account_Code_dynamic3 = get_account_code_branch(27, my_branch)
        
        If Account_Code_dynamic3 = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "No Branch was created", vbCritical
            End If
            
            create_accounts = False
            Exit Function
        Else

            If Account_Code_dynamic3 = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» „Ê«œ  ··„‘«—Ì⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "No selected account for materials in the branch for this process", vbCritical
                End If
                
                create_accounts = False
                Exit Function
            End If
        End If
        
        Account_Code_dynamic4 = get_account_code_branch(28, my_branch)
        
        If Account_Code_dynamic4 = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "No Branch was created", vbCritical
            End If
        
            create_accounts = False
            Exit Function
        Else

            If Account_Code_dynamic4 = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «ÃÊ— ··„‘«—Ì⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "No selected account for wages in the branch for this process", vbCritical
                End If
        
                create_accounts = False
                Exit Function
            End If
        End If
        
        Account_Code_dynamic5 = get_account_code_branch(32, my_branch)
        
        If Account_Code_dynamic5 = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "No Branch was created", vbCritical
            End If
            
        
            create_accounts = False
            Exit Function
        Else

            If Account_Code_dynamic5 = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» ‰Ÿ«„Ì ··„‘«—Ì⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "No selected systematic account in the branch for this process ", vbCritical
                End If
        
                create_accounts = False
                Exit Function
            End If
        End If
        If SystemOptions.AllowGoodPerfAccount = True Then
           Account_Code_dynamic5 = get_account_code_branch(152, my_branch)
        
        If Account_Code_dynamic5 = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "No Branch was created", vbCritical
            End If
            
        
            create_accounts = False
            Exit Function
        Else

            If Account_Code_dynamic5 = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» Õ”‰ «·«œ«¡ ··„‘«—Ì⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "No selected systematic account in the branch for this process ", vbCritical
                End If
        
                create_accounts = False
                Exit Function
            End If
        End If
     End If
create_accounts = True
        Exit Function

    End If
       create_accounts = True

        Exit Function
        
    My_SQL = "  select * from project_category WHERE branch_id =" & dcBranch.BoundText & "and category_id=" & DataCombo5.BoundText & " order by id"

    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsSavRec.RecordCount = 0 Then MsgBox "this project type not found in this branch", vbCritical: create_accounts = False: Exit Function

    My_SQL = "  select * from project_category WHERE type='a14' and branch_id =" & dcBranch.BoundText & "and category_id=" & DataCombo5.BoundText & " order by id"
    Set RsSavRec1 = New ADODB.Recordset
    RsSavRec1.CursorLocation = adUseClient
    RsSavRec1.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsSavRec1.RecordCount = 0 Then MsgBox "·„ Ì „  ÕœÌœ Õ”«» „’—Ê›«  ·Â–« «·„‘—Ê⁄", vbCritical: create_accounts = False: Exit Function
 
    My_SQL = "  select * from project_category WHERE type='a15' and branch_id =" & dcBranch.BoundText & "and category_id=" & DataCombo5.BoundText & " order by id"

    Set RsSavRec2 = New ADODB.Recordset
    RsSavRec2.CursorLocation = adUseClient
    RsSavRec2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsSavRec2.RecordCount = 0 Then MsgBox "·„ Ì „  ÕœÌœ Õ”«» «Ì—«œ«  ·Â–« «·„‘—Ê⁄", vbCritical: create_accounts = False:  Exit Function

    My_SQL = "  select * from project_category WHERE type='a27' and branch_id =" & dcBranch.BoundText & "and category_id=" & DataCombo5.BoundText & " order by id"
    Set RsSavRec3 = New ADODB.Recordset
    RsSavRec3.CursorLocation = adUseClient
    RsSavRec3.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsSavRec3.RecordCount = 0 Then MsgBox "·„ Ì „  ÕœÌœ Õ”«» „Ê«œ ·Â–« «·„‘—Ê⁄", vbCritical:  create_accounts = False: Exit Function
 
    My_SQL = "  select * from project_category WHERE type='a28' and branch_id =" & dcBranch.BoundText & "and category_id=" & DataCombo5.BoundText & " order by id"

    Set RsSavRec4 = New ADODB.Recordset
    RsSavRec4.CursorLocation = adUseClient
    RsSavRec4.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsSavRec4.RecordCount = 0 Then MsgBox "·„ Ì „  ÕœÌœ Õ”«» „Ê«œ ·Â–« «·„‘—Ê⁄", vbCritical:  create_accounts = False: Exit Function
 
    My_SQL = "  select * from project_category WHERE type='a32' and branch_id =" & dcBranch.BoundText & "and category_id=" & DataCombo5.BoundText & " order by id"

    Set RsSavRec5 = New ADODB.Recordset
    RsSavRec5.CursorLocation = adUseClient
    RsSavRec5.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsSavRec5.RecordCount = 0 Then MsgBox "·„ Ì „  ÕœÌœ Õ”«» ‰Ÿ«„Ì ·Â–« «·„‘—Ê⁄", vbCritical:  create_accounts = False: Exit Function
 
     
If SystemOptions.AllowGoodPerfAccount = True Then
My_SQL = "  select * from project_category WHERE type='a152' and branch_id =" & dcBranch.BoundText & "and category_id=" & DataCombo5.BoundText & " order by id"
    Set RsSavRec6 = New ADODB.Recordset
    RsSavRec6.CursorLocation = adUseClient
    RsSavRec6.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsSavRec6.RecordCount = 0 Then MsgBox "·„ Ì „  ÕœÌœ Õ”«» Õ”‰ «·«œ«¡", vbCritical:  create_accounts = False: Exit Function
    End If
    
    Account_Code_dynamic1C = ModAccounts.AddNewAccount(RsSavRec1("account_code").value, DCPreFix.Text & Trim$(Me.TxtId.Text) & " -„’—Ê›«  ", True, False, DCPreFix.Text & Trim$(Me.TxtId.Text) & " -Expenses")
    Account_Code_dynamic2C = ModAccounts.AddNewAccount(RsSavRec2("account_code").value, DCPreFix.Text & Trim$(Me.TxtId.Text) & "  -«Ì—«œ«  ", True, False, DCPreFix.Text & Trim$(Me.TxtId.Text) & "  -Revenue")
    Account_Code_dynamic3C = ModAccounts.AddNewAccount(RsSavRec3("account_code").value, DCPreFix.Text & Trim$(Me.TxtId.Text) & "  -„Ê«œ ", True, False, DCPreFix.Text & Trim$(Me.TxtId.Text) & "  -Material")
    Account_Code_dynamic4C = ModAccounts.AddNewAccount(RsSavRec4("account_code").value, DCPreFix.Text & Trim$(Me.TxtId.Text) & "  -«ÃÊ— ", True, False, DCPreFix.Text & Trim$(Me.TxtId.Text) & "  -Salary")
    Account_Code_dynamic5C = ModAccounts.AddNewAccount(RsSavRec5("account_code").value, DCPreFix.Text & Trim$(Me.TxtId.Text) & "  -„” Œ·’«  ", True, False, DCPreFix.Text & Trim$(Me.TxtId.Text) & "  -Legal")
    If SystemOptions.AllowGoodPerfAccount = True Then
    Account_Code_dynamic7C = ModAccounts.AddNewAccount(RsSavRec6("account_code").value, DCPreFix.Text & Trim$(Me.TxtId.Text) & "  -Õ”‰ «·«œ«¡ ", True, False, DCPreFix.Text & Trim$(Me.TxtId.Text) & "  -Good ")
  End If
    ' Me.legal.text = ModAccounts.AddNewAccount(RsSavRec4("account_code").value, DCPreFix.text & Trim$(Me.txtid.text) & "  Legal ", True, False)

    create_accounts = True
End If
End Function
 
Private Sub Command2_Click()
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String

    If Me.dcJobTypeName.BoundText <> "" Then
        If Not IsNumeric(txtCount.Text) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ  ⁄œœ «·«Ì«„   "
            Else
                Msg = " SPecify No of Days  "
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            txtCount.SetFocus
            Exit Sub
        End If
        
        If Not IsNumeric(TxtEmpcount.Text) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ  ⁄œœ «·„ÿ·Ê»Ì‰ „‰ Â–… «·„Â‰…  "
            Else
                Msg = "Specify No oF labors From this Job  "
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            txtCount.SetFocus
            Exit Sub
        End If

        If Option4.value = True Then ' ﬁœÌ— ›ﬁÿ
            StrSQL = "SELECT     ROUND((ISNULL(dbo.TblEmployee.Emp_Salary, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_sakn, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_bus, 0) " & "       + ISNULL(dbo.TblEmployee.Emp_Salary_food, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_others, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_mob, 0) " & "       + ISNULL(dbo.TblEmployee.Emp_Salary_mang, 0)) / 30, 2) AS daysalary, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, " & "       dbo.TblEmployee.JobTypeID , dbo.TblEmployee.project_id, dbo.TblEmpJobsTypes.JobTypeName " & "       FROM         dbo.TblEmployee INNER JOIN" & "       dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID " & "  WHERE      dbo.TblEmployee.JobTypeID =" & val(Me.dcJobTypeName.BoundText)
            ' StrSQL = "SELECT  round((isnull(Emp_Salary,0)+isnull(Emp_Salary_sakn,0) +isnull(Emp_Salary_bus,0)   +isnull(Emp_Salary_food,0)  +isnull(Emp_Salary_others,0)  +isnull(Emp_Salary_mob,0)  +isnull(Emp_Salary_mang,0))/30,2) as daysalary,* from TblEmployee Where  JobTypeID= " & Val(Me.dcJobTypeName.BoundText)
        ElseIf Option5.value = True Then
            ' StrSQL = "SELECT  round((isnull(Emp_Salary,0)+isnull(Emp_Salary_sakn,0) +isnull(Emp_Salary_bus,0)   +isnull(Emp_Salary_food,0)  +isnull(Emp_Salary_others,0)  +isnull(Emp_Salary_mob,0)  +isnull(Emp_Salary_mang,0))/30,2) as daysalary,* from TblEmployee Where  project_id=0 and JobTypeID= " & Val(Me.dcJobTypeName.BoundText) '   Œ’Ì’ ›⁄·Ï
            StrSQL = "SELECT     ROUND((ISNULL(dbo.TblEmployee.Emp_Salary, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_sakn, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_bus, 0) " & "       + ISNULL(dbo.TblEmployee.Emp_Salary_food, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_others, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_mob, 0) " & "       + ISNULL(dbo.TblEmployee.Emp_Salary_mang, 0)) / 30, 2) AS daysalary, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, " & "       dbo.TblEmployee.JobTypeID , dbo.TblEmployee.project_id, dbo.TblEmpJobsTypes.JobTypeName " & "       FROM         dbo.TblEmployee INNER JOIN " & "       dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID " & " WHERE      dbo.TblEmployee.JobTypeID =" & val(Me.dcJobTypeName.BoundText) & " and ( dbo.TblEmployee.project_id =0 OR  dbo.TblEmployee.project_id IS NULL)"
                
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Dim lastrow As Integer
        Dim x As Integer

        If rs.RecordCount > 0 Then
            If Option5.value = True Then
                If rs.RecordCount < val(TxtEmpcount.Text) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "«·⁄œœ «·„ÿ·Ê» „‰ «·⁄„«· €Ì— „ Ê›— Â·  —Ìœ  «· ﬂ„·… »«·⁄œœ «·„ÊÃÊœ" & CHR(13)
                        Msg = Msg & "  ‰⁄„  ﬂ„·…"
                        Msg = Msg & "  ·«  «·€«¡" & CHR(13)
                    Else
                        Msg = "No Of Labors not exist Now,continue with avilable " & CHR(13)
                        Msg = Msg & "  Yes -continue  "
                        Msg = Msg & " No - cancel" & CHR(13)
                    End If

                    x = MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

                    If x = vbNo Then
                        Exit Sub
                    End If
                
                End If
            End If

            rs.MoveFirst
    
            With Me.VSFlexGrid1
                lastrow = .Rows - 1
                .Rows = .Rows + rs.RecordCount

                For i = lastrow To .Rows - 2
                
                    .TextMatrix(i, .ColIndex("LineNo")) = i
                    .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                    .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs("Emp_Code").value), "", rs("Emp_Code").value)
                        
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                        
                    .TextMatrix(i, .ColIndex("jobid")) = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)
                    .TextMatrix(i, .ColIndex("jobname")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                    .TextMatrix(i, .ColIndex("daysalary")) = Round(IIf(IsNull(rs("daysalary").value), 0, rs("daysalary").value), 2)
                    .TextMatrix(i, .ColIndex("Count")) = val(Me.txtCount.Text)
                    .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("daysalary"))) * val(.TextMatrix(i, .ColIndex("Count")))
                    rs.MoveNext
                Next

            End With

            calcnets
        Else

            If Option4.value = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "€Ì— „ Ê›— ⁄„«· »Â–… «·„Â‰…  "
                Else
                    Msg = "No Labors assigned to this job  "
                End If

            ElseIf Option5.value = True Then

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "€Ì— „ Ê›— ⁄„«· »Â–… «·„Â‰… «Ê «‰ ﬂ· «·⁄„«· „Œ’’Ì‰ ·„‘«—Ì⁄ «Ê ⁄„·Ì«  «Œ—Ï  "
                Else
                    Msg = "No Labors assigned to this job Or all Labors Allocated to another Project Process  "
                End If
            End If
                     
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Õœœ «·„Â‰… «Ê·« «·„ÿ·Ê»… «Ê·« "
        Else
            Msg = "Specify Job Firstly "
        End If

        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        dcJobTypeName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

End Sub

Public Function ShowReports()
    On Error Resume Next

    Dim sql As String
 

    Dim Balance As String
    Dim rs  As ADODB.Recordset
 

         
    Dim i As Integer

 
    Dim xApp As New CRAXDRT.Application

    Dim EmpReport As ClsEmployeeReport
    Dim xReport As New CRAXDRT.Report
Dim filename As String

    'Dim rs As ADODB.Recordset
    Dim cCompanyInfo As ClsCompanyInfo
    Set cCompanyInfo = New ClsCompanyInfo
    sql = "SELECT * from projects WHERE     (dbo.projects.id = " & val(txt_project_id.Text) & ")"
    
    sql = " SELECT     200 AS collected, dbo.project_billl.id, dbo.project_billl.ManualNO, dbo.projects.Project_name, dbo.project_billl.project_no, dbo.project_billl.total, "
sql = sql & " dbo.ProjectBillBuy.[Value], dbo.ProjectBillBuy.TxtNoteSerial, dbo.projects.Project_nameE, dbo.projects.Fullcode, dbo.projects.StartDate, dbo.projects.EndDate,"
sql = sql & " dbo.projects.net, dbo.projects.End_user_id, TblCustemers_1.CusName, TblCustemers_1.CusNamee, dbo.projects.WarrantyNO, dbo.projects.WarrantyValue,"
sql = sql & " dbo.projects.WarrDateStart, dbo.projects.WarrDateEnd, dbo.projects.WarrExtension, dbo.projects.WarrBank, dbo.TblMunicipality.name AS amanah,"
sql = sql & " dbo.TblMunicipality.namee AS amanhe, dbo.TblMunicipalityDet.name AS bldya, dbo.TblMunicipalityDet.namee AS bldyae, dbo.BanksData.BankName,"
sql = sql & " dbo.BanksData.BankNamee, dbo.project_status.name AS project_statusA, dbo.project_status.namee AS project_statusE, dbo.contract_type.name AS contract_typeA,"
sql = sql & " dbo.contract_type.namee AS contract_typeE, dbo.project_billl.bill_date, dbo.ProjectBillBuy.RecordDate, TblCustemers_1.CusName AS subcontractname,"
sql = sql & " TblCustemers_1.CusNamee AS subcontractnamee"
sql = sql & "  FROM         dbo.TblCustemers TblCustemers_1 RIGHT OUTER JOIN"
sql = sql & " dbo.projects ON TblCustemers_1.CusID = dbo.projects.JobeContractorID LEFT OUTER JOIN"
sql = sql & " dbo.contract_type ON dbo.projects.Contract_type = dbo.contract_type.id LEFT OUTER JOIN"
sql = sql & " dbo.project_status ON dbo.projects.Project_status = dbo.project_status.id LEFT OUTER JOIN"
sql = sql & " dbo.BanksData ON dbo.projects.WarrBank = dbo.BanksData.BankID LEFT OUTER JOIN"
sql = sql & " dbo.TblMunicipalityDet ON dbo.projects.Municipalityid = dbo.TblMunicipalityDet.ID RIGHT OUTER JOIN"
sql = sql & " dbo.TblCustemers TblCustemers_2 ON dbo.projects.End_user_id = TblCustemers_2.CusID LEFT OUTER JOIN"
sql = sql & " dbo.project_billl LEFT OUTER JOIN"
sql = sql & " dbo.ProjectBillBuy ON dbo.project_billl.id = dbo.ProjectBillBuy.Bill_id ON dbo.projects.id = dbo.project_billl.project_no LEFT OUTER JOIN"
sql = sql & "  dbo.TblMunicipality ON dbo.projects.Amanhid = dbo.TblMunicipality.ID"
 


sql = sql & "  WHERE     (dbo.projects.id = " & val(txt_project_id.Text) & ")"
    
    
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText
       
    If SystemOptions.UserInterface = ArabicInterface Then
    filename = App.path & "\reports\REPORTS NEW\ProjectRPT.rpt "
    
        Set xReport = xApp.OpenReport(filename)
    Else
    filename = App.path & "\reports\REPORTS NEW\ProjectRPTe.rpt"
        Set xReport = xApp.OpenReport(filename)
    End If

    xReport.Database.SetDataSource rs
 
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.TxtPath = filename
    FrmReport.CRViewer.viewReport
 
 '   xReport.reporttitle = cCompanyInfo.ArabCompanyName
     If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(3).AddCurrentValue user_name
       xReport.ParameterFields(12).AddCurrentValue Text4.Text
        xReport.ParameterFields(13).AddCurrentValue "1"
        StrReportTitle = "" '& StrAccountName
 
    Else
 
     '   xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
         
         
     '   StrReportTitle = ""
 
    End If
   
   
    
    
    FrmReport.show
    Screen.MousePointer = vbDefault
    '   xReport.ReportTitle = X
    SendKeys "{RIGHT}"

End Function

Private Sub Command3_Click()
  
      On Error Resume Next
    Dim StrFileName As String
    StrFileName = "C:\" & "\Payrolll.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If
  
      On Error Resume Next
      cd.CancelError = True 'allow escape key/cancel
     cd.filename = "Project"
    cd.ShowSave     'show the dialog screen
    If Err <> 32755 Then    ' User didn't chose Cancel.
   Else
       Exit Sub
    End If
 StrFileName = cd.filename & ".xls"
Me.FgMainDes.saveGrid StrFileName, flexFileCustomText, True
   
    OpenFile StrFileName
End Sub

Private Sub Command5_Click()
Frame20.Visible = False
End Sub

Private Sub Command6_Click()
Frame5.Visible = False
Frame2.Visible = True
End Sub

Private Sub Command7_Click()
Frame21.Visible = False
End Sub

Private Sub company_Click()
If company.value = True Then
 
   DataCombo3.Text = ""
   DataCombo3.Enabled = False
  End If
End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String

        My_SQL = "  select id,name from project_status  "
        fill_combo DataCombo1, My_SQL
    End If

End Sub

Private Sub DataCombo5_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
        My_SQL = "  select id,name from contract_type  "
        fill_combo DataCombo5, My_SQL
    End If

End Sub

Private Sub DcAccount1_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = 13 Then

        On Error Resume Next

        If DcAccount1.Text = "" Then DcAccount2.Text = "": Exit Sub
        DcAccount2.Text = ""
        Dim My_SQL As String

        My_SQL = "select CusName from TblCustemers where CusID='" & DcAccount1.Text & "'"
 
        Set Rec = New ADODB.Recordset
        Rec.CursorLocation = adUseClient

        Rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(Rec.Fields("Account_Name").value) Then
            DcAccount2.Text = Rec.Fields("CusName").value
        Else
            DcAccount2.Text = ""
 
        End If

    End If

 '   If KeyCode = vbKeyF5 Then
      
 '       My_SQL = "  select Account_Code,CusID from TblCustemers  where type=1 "
 '
 '       fill_combo DcAccount1, MySQL

    'End If
        
End Sub

Private Sub DcAccount2_Change()
Dim Fullcode As String
Dim DefaultSalesPersonId As Integer
      Fullcode = ""
        GetCustomersDetail val(DcAccount2.BoundText), DefaultSalesPersonId, Fullcode
        TxtCustCode.Text = Fullcode
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        DCEmp1.BoundText = DefaultSalesPersonId
    End If
End Sub

Private Sub DCAccount2_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 8
            FrmCustemerSearch.show vbModal
            
        End If
End Sub

Private Sub DcAccount3_Click(Area As Integer)
    On Error Resume Next

    If DcAccount3.Text = "" Then Exit Sub
    DcAccount4.Text = ""
    Dim My_SQL As String

    My_SQL = "select CusName from TblCustemers where CusID='" & DcAccount3.Text & "'"
    Dim Rec As ADODB.Recordset
    Set Rec = New ADODB.Recordset
    Rec.CursorLocation = adUseClient

    Rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not IsNull(Rec.Fields("Account_Name").value) Then
        DcAccount4.Text = Rec.Fields("CusName").value
    Else
        DcAccount4.Text = ""
    End If

End Sub

Private Sub DcAccount1_Click(Area As Integer)
    On Error Resume Next

    If DcAccount1.Text = "" Then Exit Sub
    DcAccount2.Text = ""
    Dim My_SQL As String

    My_SQL = "select CusName from TblCustemers where CusID='" & DcAccount1.Text & "'"
    Dim Rec As ADODB.Recordset
    Set Rec = New ADODB.Recordset
    Rec.CursorLocation = adUseClient

    Rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not IsNull(Rec.Fields("Account_Name").value) Then
        DcAccount2.Text = Rec.Fields("CusName").value
    Else
        DcAccount2.Text = ""
    End If

End Sub

Private Sub DcAccount2_Click(Area As Integer)
    On Error Resume Next

    If DcAccount2.Text = "" Then Exit Sub
    DcAccount1.Text = ""
    Dim My_SQL As String

    My_SQL = "select CusID from TblCustemers where CusID =" & DcAccount2.BoundText
    Dim Rec As ADODB.Recordset
    Set Rec = New ADODB.Recordset
    Rec.CursorLocation = adUseClient

    Rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not IsNull(Rec.Fields("CusID").value) Then
        DcAccount1.Text = Rec.Fields("CusID").value
    Else
        DcAccount1.Text = ""
    End If



Dim Fullcode As String
Dim DefaultSalesPersonId As Integer
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then

        
        Fullcode = ""
        GetCustomersDetail val(DcAccount2.BoundText), DefaultSalesPersonId, Fullcode
        TxtCustCode.Text = Fullcode

        DCEmp1.BoundText = DefaultSalesPersonId
    

    
    End If




End Sub

Private Sub DCAccount3_KeyUp(KeyCode As Integer, _
                             Shift As Integer)
    On Error Resume Next

    If KeyCode = 13 Then

        If DcAccount3.Text = "" Then DcAccount4.Text = "": Exit Sub
        DcAccount4.Text = ""
        Dim My_SQL As String

        My_SQL = "select CusName from TblCustemers where CusID='" & DcAccount3.Text & "'"
        Dim Rec As ADODB.Recordset
        Set Rec = New ADODB.Recordset
        Rec.CursorLocation = adUseClient

        Rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(Rec.Fields("Account_Name").value) Then
            DcAccount4.Text = Rec.Fields("CusName").value
        Else
            DcAccount4.Text = ""
             
        End If
 
    End If

    If KeyCode = vbKeyF5 Then
        My_SQL = "  select Account_Code,CusID from TblCustemers  where type=1 "
                
        fill_combo DcAccount3, My_SQL

    End If
        
End Sub

Private Sub DcAccount4_Change()
Dim Fullcode As String
Dim DefaultSalesPersonId As Integer
 
        Fullcode = ""
        GetCustomersDetail val(DcAccount4.BoundText), DefaultSalesPersonId, Fullcode
        TxtCustCode1.Text = Fullcode

        
    

    
  
End Sub

Private Sub DcAccount4_Click(Area As Integer)
    On Error Resume Next

    If DcAccount4.Text = "" Then Exit Sub
    DcAccount3.Text = ""
    Dim My_SQL As String

    My_SQL = "select CusID from TblCustemers where CusID =" & DcAccount4.BoundText
    Dim Rec As ADODB.Recordset
    Set Rec = New ADODB.Recordset
    Rec.CursorLocation = adUseClient

    Rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not IsNull(Rec.Fields("CusID").value) Then
        DcAccount3.Text = Rec.Fields("CusID").value
    Else
        DcAccount3.Text = ""
    End If

Dim Fullcode As String
Dim DefaultSalesPersonId As Integer
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then

        
        Fullcode = ""
        GetCustomersDetail val(DcAccount4.BoundText), DefaultSalesPersonId, Fullcode
        TxtCustCode1.Text = Fullcode

        
    

    
    End If




End Sub

Function gettotal(x As String, filed As String, table As String, filed_search As String) As Double
    Dim My_SQL As String

    My_SQL = "  select Sum(" & filed & ") as total  from " & table & " where " & filed_search & "='" & x & "'"
    Dim Rec As ADODB.Recordset
    Set Rec = New ADODB.Recordset
    Rec.CursorLocation = adUseClient

    Rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    gettotal = IIf(IsNull(Rec.Fields("total").value), 0, Rec.Fields("total").value)

End Function

Private Sub DCAccount4_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
          FrmCustemerSearch.SearchType = 9
             FrmCustemerSearch.show vbModal
          
        End If
End Sub

Private Sub DCboItemsCode_Click(Area As Integer)

    If val(DCboItemsCode.BoundText) = 0 Then Exit Sub
    Text6.Text = get_item_Reserved_qty(val(DCboItemsCode.BoundText))
    Text3.Text = get_item_qty(val(DCboItemsCode.BoundText))
    Text1.Text = get_item_Order_qty(val(DCboItemsCode.BoundText))

End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches dcBranch
    End If

End Sub

Private Sub DCCurrency_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String

        If SystemOptions.UserInterface = ArabicInterface Then
            My_SQL = "  select id,name from currency  order by name  "
        Else
            My_SQL = "  select id,code from currency  order by code  "
        End If

        fill_combo DcCurrency, My_SQL

    End If

End Sub

Private Sub dcEmp_Change()
    If val(DCEmP.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DCEmP.BoundText, EmpCode
    Me.Text10.Text = EmpCode
End Sub

Private Sub DCEmP_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
        
        My_SQL = "  select Emp_ID,Emp_name  from TblEmployee  "
        fill_combo DCEmP, My_SQL
    End If

End Sub

Private Sub DCEmp1_Change()
Dim Fullcode As String
Dim DefaultSalesPersonId As Integer
        TxtCustCode2.Text = GetSalespersonDetail(val(DCEmp1.BoundText))
        
End Sub

Private Sub DCEmp1_Click(Area As Integer)
DCEmp1_Change
End Sub

Private Sub DCPreFix_KeyUp(KeyCode As Integer, _
                           Shift As Integer)
 
    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetPrefix Me.DCPreFix, 0, 0 'val(branch_id)

    End If
        
End Sub

Private Sub DTEnddate_Change()
 Dim x As Double
       
        x = DateDiff("D", DTStartDate.value, DTEnddate.value)
        Monthly = DateDiff("m", DTStartDate.value, DTEnddate.value)
        If SystemOptions.UserInterface = ArabicInterface Then
        Text4.Text = x & "   ÌÊ„ " '
       Else
       Text4.Text = x & "   Days "
       End If
End Sub

Private Sub DTStartDate_Change()
 Dim x As Double
       
        x = DateDiff("D", DTStartDate.value, DTEnddate.value)
        Monthly = DateDiff("m", DTStartDate.value, DTEnddate.value)
              If SystemOptions.UserInterface = ArabicInterface Then
        Text4.Text = x & "   ÌÊ„ " '
       Else
       Text4.Text = x & "   Days "
       End If
End Sub

Private Sub employee_details_Click()
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
    VSFlexGrid1.Enabled = True
          
    If Not VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode")) = "" Then
        Frame10.Visible = True
        Frame10.Enabled = True

        current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode"))

        If SystemOptions.UserInterface = ArabicInterface Then
            Frame10.Caption = "«·⁄„«·Â ›Ì  «·⁄„·Ì… —ﬁ„ : " & current_opr
        Else
            Frame10.Caption = "Labors for Operation NO : " & current_opr
        End If
        
        'Me.txt_emp_salary.text = IIf(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("salary")) = "", 0, VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("salary")))
        'Me.txt_employee_count.text = IIf(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("count")) = "", 0, VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("count")))
        
        Frame10.Visible = True
        StrSQL = "SELECT  * FROM  opr_employee_details Where   opr_fullcode='" & current_opr & "'"

        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            RsDev.MoveFirst
    
            With Me.VSFlexGrid1
                .Rows = .FixedRows + RsDev.RecordCount

                For i = .FixedRows To .Rows - 1
 
                    .TextMatrix(i, .ColIndex("LineNo")) = i
            
                    .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDev("Emp_ID").value), "", RsDev("Emp_ID").value)
            
                    .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(RsDev("Emp_Code").value), "", RsDev("Emp_Code").value)
            
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
            
                    .TextMatrix(i, .ColIndex("jobid")) = IIf(IsNull(RsDev("JobTypeID").value), "", RsDev("JobTypeID").value)
                    .TextMatrix(i, .ColIndex("jobname")) = IIf(IsNull(RsDev("JobTypeName").value), "", RsDev("JobTypeName").value)
                    .TextMatrix(i, .ColIndex("daysalary")) = IIf(IsNull(RsDev("daysalary").value), 0, RsDev("daysalary").value)
                    .TextMatrix(i, .ColIndex("count")) = IIf(IsNull(RsDev("count").value), 0, RsDev("count").value)
        
                    .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(RsDev("total").value), 0, RsDev("total").value)

                    RsDev.MoveNext
                Next i

            End With
    
        End If
    
    End If

    calcnets
    ReLineGrid
End Sub
Sub SetValue(Optional PrMainDesID As Double, Optional Total As Double = 0, Optional SumValEx As Double, Optional Qty As Double, Optional QtyExe As Double, Optional QtyNo As Double, Optional Import As Integer = 0)
Dim i As Integer
With Me.FgMainDes
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("ID"))) = PrMainDesID Then
If Total <> 0 Or Import <> 0 Then
.TextMatrix(i, .ColIndex("QtyNo")) = QtyNo
.TextMatrix(i, .ColIndex("Qty")) = Qty
.TextMatrix(i, .ColIndex("Price")) = Total
.TextMatrix(i, .ColIndex("Total")) = Total
Else
.TextMatrix(i, .ColIndex("QtyNo")) = 0
.TextMatrix(i, .ColIndex("Qty")) = 0
.TextMatrix(i, .ColIndex("Price")) = 0
.TextMatrix(i, .ColIndex("Total")) = 0
End If
If SumValEx <> 0 Then
.TextMatrix(i, .ColIndex("QtyExe")) = QtyExe
.TextMatrix(i, .ColIndex("PriceExe")) = SumValEx
.TextMatrix(i, .ColIndex("TotalExe")) = SumValEx
Else
.TextMatrix(i, .ColIndex("QtyExe")) = 0
.TextMatrix(i, .ColIndex("PriceExe")) = 0
.TextMatrix(i, .ColIndex("TotalExe")) = 0
End If
End If
Next i
End With
ReLineGrid
End Sub
Function GetPandCode(Optional PrMainDesID As Double, Optional Row As Long) As String
Dim i As Integer
With Me.FgMainDes
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("ID"))) = PrMainDesID Then
GetPandCode = .TextMatrix(i, .ColIndex("FullCode"))
End If
Next i
End With
End Function
Function GetCount(Optional PrMainDesID As Double, Optional ByRef SumVal As Double, Optional ByRef SumValEx As Double, Optional ByRef Cout1 As Double, Optional ByRef Cout2 As Double, Optional ByRef QtyNo2 As Double) As Double
Dim i As Integer
Dim Countr As Integer
Dim SumVal2 As Double
Dim SumValEx2 As Double
Dim cont1 As Double
Dim Cont2 As Double

Dim QtyNo As Double
 cont1 = 0
Cont2 = 0
Countr = 0
SumVal2 = 0
SumValEx2 = 0
QtyNo = 0
With Fg_Journal
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("PrMainDesID"))) = PrMainDesID Then
Countr = Countr + 1
SumVal2 = SumVal2 + val(.TextMatrix(i, .ColIndex("net")))
SumValEx2 = SumValEx2 + val(.TextMatrix(i, .ColIndex("PriceExe"))) * val(.TextMatrix(i, .ColIndex("QtyExe")))
cont1 = cont1 + val(.TextMatrix(i, .ColIndex("qty")))
Cont2 = Cont2 + val(.TextMatrix(i, .ColIndex("QtyExe")))
QtyNo = QtyNo + val(.TextMatrix(i, .ColIndex("QtyNo")))
End If
Next i
End With
SumVal = SumVal2
SumValEx = SumValEx2
GetCount = Countr
QtyNo2 = QtyNo
Cout2 = Cont2
Cout1 = cont1
End Function

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                 ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    'Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim SumValEx As Double
    With Fg_Journal
If .TextMatrix(Row, .ColIndex("by")) = "" Then
If val(DataCombo3.BoundText) <> 0 And DataCombo3.Text <> "" Then
.TextMatrix(Row, .ColIndex("by")) = DataCombo3.Text
.TextMatrix(Row, .ColIndex("sub_contractor_id")) = val(DataCombo3.BoundText)
End If
End If
        Select Case .ColKey(Col)
        Case "des"
        If Me.TxtModFlg.Text <> "R" Then
        If .TextMatrix(Row, .ColIndex("des")) <> "" Then
        Pand = 0
         If Me.TxtModFlg.Text = "E" Then
        
        Pand = IIf(Not IsNumeric(.TextMatrix(Row, .ColIndex("oprid"))), 0, val(.TextMatrix(Row, .ColIndex("oprid"))))
        
              End If
          If Me.Checked(Pand, 0) = True Then
        Else
       Pand = 1
        maxx Pand, 0
        End If
        .TextMatrix(Row, .ColIndex("oprid")) = Pand
       End If

 End If
 
  Case "PandName"
  If .TextMatrix(Row, .ColIndex("PandName")) <> "" Then
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("PanID"), False, True)
                .TextMatrix(Row, .ColIndex("PanID")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("des")) = .TextMatrix(Row, .ColIndex("PandName"))
            Fg_Journal_AfterEdit Row, .ColIndex("des")
      End If
    Case "PandUnit"
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("PandUnitID"), False, True)
                .TextMatrix(Row, .ColIndex("PandUnitID")) = StrAccountCode
             
                
            Case "by"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("sub_contractor_id"), False, True)
                .TextMatrix(Row, .ColIndex("sub_contractor_id")) = StrAccountCode
                '.TextMatrix(Row, .ColIndex("id")) = get_Expenses_id(StrAccountCode)
          Case "PrMainDesID", "QtyNo", "qty", "cost", "discount", "QtyExe", "PriceExe"
          Dim CuntNo As Double
          Dim Pandcode As String
          Dim Total As Double
          Dim count1 As Double
          Dim Count2 As Double
          Dim QtyNo As Double
          ReLineGrid
          CuntNo = 0
        SumValEx = 0
         CuntNo = GetCount(val(.TextMatrix(Row, .ColIndex("PrMainDesID"))), Total, SumValEx, count1, Count2, QtyNo)
         Pandcode = GetPandCode(val(.TextMatrix(Row, .ColIndex("PrMainDesID"))), Row)
         .TextMatrix(Row, .ColIndex("CodeBand")) = CuntNo & "-" & Pandcode
         SetValue val(.TextMatrix(Row, .ColIndex("PrMainDesID"))), Total, SumValEx, count1, Count2, QtyNo
        End Select
   
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid

End Sub

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
        Select Case .ColKey(Col)
            
            Case "by"
                Exit Sub
            Case "qty"
            Fg_Journal.ComboList = ""
              Case "QtyExe"
            Fg_Journal.ComboList = ""
              Case "PriceExe"
            Fg_Journal.ComboList = ""
           Case "TotalExe"
           Cancel = True
        End Select

    End With

    Fg_Journal.ComboList = ""

End Sub

Private Sub Fg_Journal_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    With Me.Fg_Journal

        Select Case .ColKey(Col)
        
                Case "PrintttAn"
                  LngRow = Row
 LngCol = Col
print_report1 val(Me.txt_project_id.Text), val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("oprid")))

                      Case "Printtt"
                  LngRow = Row
 LngCol = Col
print_report val(Me.txt_project_id.Text), val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("oprid")))
                 Case "opera"
                  LngRow = Row

 LngCol = Col
             ' ItemProductionDate Row, Col, , 1
             '   Load FrmProcessOfProject
             '   FrmProcessOfProject.show vbModal

                    
                End Select
                End With
End Sub

Private Sub Fg_Journal_Click()

    If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("fullcode")) = "" Then
      ' pand = val(.TextMatrix(Row, .ColIndex("oprid")))
        current_terms = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("fullcode"))
        Pand = val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("oprid")))
    End If

End Sub

Private Sub Fg_Journal_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
With Me.Fg_Journal
Select Case .ColKey(Col)
Case "QtyNo"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("QtyNo")), 0)
Case "qty"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("qty")), 0)
Case "cost"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("cost")), 0)
Case "total"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("total")), 0)
Case "discount"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("discount")), 0)
Case "net"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("net")), 0)
Case "QtyExe"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("QtyExe")), 0)
Case "PriceExe"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("PriceExe")), 0)
Case "TotalExe"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("TotalExe")), 0)
End Select
End With
End Sub

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)
    Dim rs As ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)
        
                                       Case "PrintttAn"
.ColComboList(.ColIndex("PrintttAn")) = "..."
                               Case "Printtt"
.ColComboList(.ColIndex("Printtt")) = "..."

        Case "opera"
         .ColComboList(.ColIndex("opera")) = "..."
             Case "PandName"
                StrSQL = "select * from TblPands "
  Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If rs.RecordCount = 0 Then Exit Sub
              If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = Fg_Journal.BuildComboList(rs, "Name", "ID")
                Else
                 StrComboList = Fg_Journal.BuildComboList(rs, "NameE", "ID")
                 End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                

            Case "PandUnit"
                StrSQL = "select * from TblProcessUnites "
  Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If rs.RecordCount = 0 Then Exit Sub
              If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = Fg_Journal.BuildComboList(rs, "UnitName", "UnitID")
                Else
                 StrComboList = Fg_Journal.BuildComboList(rs, "UnitNamee", "UnitID")
                 End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                 Case "by"
                StrSQL = "select * from TblCustemers where Type=3"
  Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If rs.RecordCount = 0 Then Exit Sub
              
                StrComboList = Fg_Journal.BuildComboList(rs, "CusName", "CusID")
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub
Function GetIDProcess(Optional Name As String, Optional UnitID As Double) As Double
If Name <> "" Then
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select * from TblProcessDEF "
sql = sql & " WHERE     (ProcessName LIKE N'%" & Name & "%') or (ProcessNameE LIKE N'%" & Name & "%')"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetIDProcess = IIf(IsNull(Rs3("TblProcessDEFID").value), 0, Rs3("TblProcessDEFID").value)
Else
GetIDProcess = SaveProcess(Name, UnitID)
End If
End If
End Function
Function GetIDUnit(Optional Name As String) As Double
If Name <> "" Then
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select * from TblProcessUnites "
sql = sql & " WHERE     (UnitName LIKE N'%" & Name & "%') or (UnitNamee LIKE N'%" & Name & "%')"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetIDUnit = IIf(IsNull(Rs3("UnitID").value), 0, Rs3("UnitID").value)
Else
GetIDUnit = SaveUnit(Name)
End If
End If
End Function
Function SaveProcess(Optional Name As String, Optional UnitID As Double) As Double
Dim Rs3 As ADODB.Recordset
Dim ID As Double
Dim StrSQL As String
    Set Rs3 = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblProcessDEF where 1=-1 "
    Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Rs3.AddNew
    ID = CStr(new_id("TblProcessDEF", "TblProcessDEFID", "", True))
    Rs3("TblProcessDEFID") = ID
    Rs3("ProcessName") = Name
    Rs3("ProcessNameE") = Name
    Rs3("UnitID") = UnitID
    Rs3("Interval") = 1
    Rs3("IntervalID") = 0
    Rs3.update
    SaveProcess = ID
End Function
Function SaveUnit(Optional Name As String) As Double
Dim Rs3 As ADODB.Recordset
Dim ID As Double
Dim StrSQL As String
    Set Rs3 = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblProcessUnites where 1=-1 "
    Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Rs3.AddNew
    ID = CStr(new_id("TblProcessUnites", "UnitID", "", True))
    Rs3("UnitID") = ID
    Rs3("UnitName") = Name
    Rs3("UnitNamee") = Name
    Rs3.update
    SaveUnit = ID
End Function
Sub FillColor1(Optional ID As Double)
Dim i As Integer
Dim RowNum As Long
With Me.Fg_Journal
RowNum = 0
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("PrMainDesID"))) = ID Then
   .Cell(flexcpBackColor, i, 1, i, 28) = vbGreen
   If RowNum = 0 Then
   RowNum = i
   End If
   Else
.Cell(flexcpBackColor, i, 1, i, 28) = &H80000005
End If
  Next i
          If RowNum <> 0 Then
                    .Row = RowNum
                    '.Col = Fg.ColIndex("UnitID")
                    .ShowCell RowNum, 1
                    .SetFocus
                    Screen.MousePointer = vbDefault
           End If
     End With
End Sub

Sub FillBand()
Dim i As Integer
With Me.FgMainDes
Fg_Journal.ColComboList(Fg_Journal.ColIndex("PrMainDesID")) = ""
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("ID"))) <> 0 Then
  Fg_Journal.ColComboList(Fg_Journal.ColIndex("PrMainDesID")) = Fg_Journal.ColComboList(Fg_Journal.ColIndex("PrMainDesID")) & "#"
  Fg_Journal.ColComboList(Fg_Journal.ColIndex("PrMainDesID")) = Fg_Journal.ColComboList(Fg_Journal.ColIndex("PrMainDesID")) & "" & .TextMatrix(i, .ColIndex("ID")) & ""
  Fg_Journal.ColComboList(Fg_Journal.ColIndex("PrMainDesID")) = Fg_Journal.ColComboList(Fg_Journal.ColIndex("PrMainDesID")) & ";"
  Fg_Journal.ColComboList(Fg_Journal.ColIndex("PrMainDesID")) = Fg_Journal.ColComboList(Fg_Journal.ColIndex("PrMainDesID")) & "" & .TextMatrix(i, .ColIndex("Name")) & ""
  Fg_Journal.ColComboList(Fg_Journal.ColIndex("PrMainDesID")) = Fg_Journal.ColComboList(Fg_Journal.ColIndex("PrMainDesID")) & "|"
  End If
  Next i
  End With
End Sub
Sub RetriveBandDetials(Optional PrMainDesID As Double)
Dim i As Integer
Dim RsDev As ADODB.Recordset
Dim StrSQL As String
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 2
   If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        StrSQL = " SELECT     dbo.projects_des.fullcode, dbo.projects_des.[index], dbo.projects_des.des, dbo.projects_des.qty, dbo.projects_des.cost, dbo.projects_des.total, "
        StrSQL = StrSQL + "               dbo.projects_des.discount, dbo.projects_des.net, dbo.projects_des.project_id, dbo.projects_des.sub_contractor_id, dbo.TblCustemers.CusName,"
        StrSQL = StrSQL + "               dbo.projects_des.oprid, dbo.projects_des.Remark, dbo.projects_des.esQty, dbo.projects_des.PandUnitID, dbo.TblProcessUnites.UnitName,"
        StrSQL = StrSQL + "               dbo.TblProcessUnites.UnitNamee , dbo.projects_des.QtyNo, dbo.projects_des.CodeBand, dbo.projects_des.PanID, dbo.TblPands.Name, dbo.TblPands.NameE ,"
        StrSQL = StrSQL + "  dbo.projects_des.TotalExe,dbo.projects_des.PriceExe,dbo.projects_des.QtyExe ,dbo.projects_des.PrMainDesID"
        StrSQL = StrSQL + "     FROM         dbo.projects_des LEFT OUTER JOIN"
        StrSQL = StrSQL + "               dbo.TblPands ON dbo.projects_des.PanID = dbo.TblPands.ID LEFT OUTER JOIN"
        StrSQL = StrSQL + "               dbo.TblProcessUnites ON dbo.projects_des.PandUnitID = dbo.TblProcessUnites.UnitID LEFT OUTER JOIN"
        StrSQL = StrSQL + "              dbo.TblCustemers ON dbo.projects_des.sub_contractor_id = dbo.TblCustemers.CusID"
        StrSQL = StrSQL + " Where (project_id =" & val(Me.txt_project_id.Text) & ")"
        If PrMainDesID <> 0 Then
        ' StrSQL = StrSQL + " and dbo.projects_des.PrMainDesID=" & PrMainDesID & ""
         StrSQL = StrSQL + "  ORDER BY dbo.projects_des.SortID,dbo.projects_des.PrMainDesID"
         Else
         StrSQL = StrSQL + "  ORDER BY dbo.projects_des.oprid"
        End If
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RsDev.BOF Or rs.EOF) Then
            RsDev.MoveFirst
    
            With Me.Fg_Journal
                .Rows = .FixedRows + RsDev.RecordCount

                For i = .FixedRows To .Rows - 1
                     .TextMatrix(i, .ColIndex("CodeBand")) = IIf(IsNull(RsDev("CodeBand").value), "", RsDev("CodeBand").value)
                    .TextMatrix(i, .ColIndex("PandUnitID")) = IIf(IsNull(RsDev("PandUnitID").value), "", RsDev("PandUnitID").value)
                    .TextMatrix(i, .ColIndex("PanID")) = IIf(IsNull(RsDev("PanID").value), 0, RsDev("PanID").value)
                    .TextMatrix(i, .ColIndex("PrMainDesID")) = IIf(IsNull(RsDev("PrMainDesID").value), "", RsDev("PrMainDesID").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("PandName")) = IIf(IsNull(RsDev("Name").value), "", RsDev("Name").value)
                     .TextMatrix(i, .ColIndex("PandUnit")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
                     Else
                     .TextMatrix(i, .ColIndex("PandName")) = IIf(IsNull(RsDev("NameE").value), "", RsDev("NameE").value)
                     .TextMatrix(i, .ColIndex("PandUnit")) = IIf(IsNull(RsDev("UnitNamee").value), "", RsDev("UnitNamee").value)
                     End If
                     
                    .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(RsDev("fullcode").value), "", RsDev("fullcode").value)
                     .TextMatrix(i, .ColIndex("oprid")) = IIf(IsNull(RsDev("oprid").value), 0, RsDev("oprid").value)
                     .TextMatrix(i, .ColIndex("QtyNo")) = IIf(IsNull(RsDev("QtyNo").value), 1, RsDev("QtyNo").value)
                    .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("des").value), "", RsDev("des").value)
                    .TextMatrix(i, .ColIndex("Remark")) = IIf(IsNull(RsDev("Remark").value), "", RsDev("Remark").value)
            
                    .TextMatrix(i, .ColIndex("qty")) = IIf(IsNull(RsDev("qty").value), "", RsDev("qty").value)
                    .TextMatrix(i, .ColIndex("esQty")) = IIf(IsNull(RsDev("esQty").value), "", RsDev("esQty").value)
                    .TextMatrix(i, .ColIndex("cost")) = IIf(IsNull(RsDev("cost").value), "", RsDev("cost").value)
           
                    .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(RsDev("total").value), "", RsDev("total").value)
         
                    .TextMatrix(i, .ColIndex("discount")) = IIf(IsNull(RsDev("discount").value), "", RsDev("discount").value)
            
                    .TextMatrix(i, .ColIndex("net")) = IIf(IsNull(RsDev("net").value), "", RsDev("net").value)
                    .TextMatrix(i, .ColIndex("net")) = IIf(IsNull(RsDev("net").value), "", RsDev("net").value)
            
                    .TextMatrix(i, .ColIndex("sub_contractor_id")) = IIf(IsNull(RsDev("sub_contractor_id").value), "", RsDev("sub_contractor_id").value)
            
                    .TextMatrix(i, .ColIndex("by")) = IIf(IsNull(RsDev("CusName").value), "", RsDev("CusName").value)
                    .TextMatrix(i, .ColIndex("QtyExe")) = IIf(IsNull(RsDev("QtyExe").value), "", RsDev("QtyExe").value)
                    .TextMatrix(i, .ColIndex("PriceExe")) = IIf(IsNull(RsDev("PriceExe").value), "", RsDev("PriceExe").value)
                    .TextMatrix(i, .ColIndex("TotalExe")) = IIf(IsNull(RsDev("TotalExe").value), "", RsDev("TotalExe").value)
                    RsDev.MoveNext
                Next i

                Me.txt_total_sum.Text = Round(.Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total")), Decimal_Places)
                Me.txt_sub_discount.Text = Round(.Aggregate(flexSTSum, .FixedRows, .ColIndex("discount"), .Rows - 1, .ColIndex("discount")), Decimal_Places)
                Me.txt_sub_net.Text = Round(.Aggregate(flexSTSum, .FixedRows, .ColIndex("net"), .Rows - 1, .ColIndex("net")), Decimal_Places)
            End With

        End If
        End If
End Sub
Function FillMylist()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
    sql = " SELECT     UserID, UserName"
    sql = sql & "         From dbo.TblUsers"
    sql = sql & " order by UserName"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    ListAllUser.Clear
    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            ListAllUser.AddItem IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
            ListAllUser.ItemData(ListAllUser.NewIndex) = rs("UserID").value
            rs.MoveNext
        Next i

    End If

    rs.Close
End Function
Function FillMylistData(Optional ProjectID As Double)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
    sql = " SELECT     dbo.TblProjectUser.UserID, dbo.TblProjectUser.ProjectID, dbo.TblUsers.UserName"
    sql = sql & "    FROM         dbo.TblProjectUser LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblUsers ON dbo.TblProjectUser.UserID = dbo.TblUsers.UserID"
    sql = sql & "  WHERE     (dbo.TblProjectUser.ProjectID = " & ProjectID & ")"
  
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    ListUserSelect.Clear
    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
                ListUserSelect.AddItem IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
                ListUserSelect.ItemData(ListUserSelect.NewIndex) = rs("UserID").value
            rs.MoveNext
        Next i

    End If

    rs.Close

End Function
Sub SaveUserProject(Optional ProjectID As Double)
Dim i As Integer
Dim sql As String
Dim Rs3 As ADODB.Recordset

If ListUserSelect.ListCount >= 0 Then
sql = "Select * from  TblProjectUser where 1=-1"
Set Rs3 = New ADODB.Recordset
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
For i = 0 To ListUserSelect.ListCount - 1
Rs3.AddNew
Rs3("UserID").value = ListUserSelect.ItemData(i)
Rs3("ProjectID").value = ProjectID
Rs3.update
Next i
End If
End Sub
Private Sub FgMainDes_AfterEdit(ByVal Row As Long, ByVal Col As Long)
   Dim StrAccountCode As String
    Dim Msg As String
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    
    With FgMainDes

        Select Case .ColKey(Col)
        Case "Name"
        If Me.TxtModFlg.Text <> "R" Then
        If .TextMatrix(Row, .ColIndex("Name")) <> "" Then
        Pand = 0
         If Me.TxtModFlg.Text = "E" Then
        
        Pand = IIf(Not IsNumeric(.TextMatrix(Row, .ColIndex("ID"))), 0, val(.TextMatrix(Row, .ColIndex("ID"))))
        
              End If
       If Me.Checked(0, 0, Pand) = True Then
        Else
       Pand = 1
        maxx 0, 0, Pand
        End If
        .TextMatrix(Row, .ColIndex("ID")) = Pand
       End If

 End If
     Case "QtyExe"
               .TextMatrix(Row, .ColIndex("TotalExe")) = val(.TextMatrix(Row, .ColIndex("QtyExe"))) * val(.TextMatrix(Row, .ColIndex("PriceExe")))
       Case "PriceExe"
               .TextMatrix(Row, .ColIndex("TotalExe")) = val(.TextMatrix(Row, .ColIndex("QtyExe"))) * val(.TextMatrix(Row, .ColIndex("PriceExe")))
    Case "QtyNo"
               .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("QtyNo"))) * val(.TextMatrix(Row, .ColIndex("Qty"))) * val(.TextMatrix(Row, .ColIndex("Price")))
    Case "Price"
                .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("QtyNo"))) * val(.TextMatrix(Row, .ColIndex("Qty"))) * val(.TextMatrix(Row, .ColIndex("Price")))
    Case "Qty"
               .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("QtyNo"))) * val(.TextMatrix(Row, .ColIndex("Qty"))) * val(.TextMatrix(Row, .ColIndex("Price")))
        End Select
   
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
    End With

    ReLineGrid
End Sub



Private Sub FgMainDes_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.FgMainDes
Select Case .ColKey(Col)
Case "Remarks"
.ComboList = ""
Case "FullCode"
.ComboList = ""
Case "Name"
.ComboList = ""
Case "Qty"
Cancel = True
Case "Price"
Cancel = True
Case "Total"
Cancel = True
Case "QtyExe"
Cancel = True
Case "PriceExe"
Cancel = True
Case "TotalExe"
Cancel = True
Case "QtyNo"
Cancel = True
End Select
End With
End Sub

Private Sub FgMainDes_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With FgMainDes
Select Case .ColKey(Col)
Case "PrintPand"
print_reportPand val(Me.txt_project_id.Text), val(.TextMatrix(Row, .ColIndex("ID")))
End Select
End With
End Sub

Private Sub FgMainDes_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
With Me.FgMainDes
Select Case .ColKey(Col)
Case "Qty"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("Qty")), 0)
Case "Price"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("Price")), 0)
Case "QtyNo"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("QtyNo")), 0)
Case "QtyExe"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("QtyExe")), 0)
Case "PriceExe"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("PriceExe")), 0)
End Select
End With
End Sub

Private Sub FgMainDes_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FgMainDes
Select Case .ColKey(Col)
Case "PrintPand"
.ColComboList(.ColIndex("PrintPand")) = "..."
End Select
End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MsgBox "hh"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Set SaleReport = Nothing
End Sub



Private Sub Label45_Click()
Dim i As Integer
ListUserSelect.Clear
For i = 0 To ListAllUser.ListCount - 1
ListUserSelect.AddItem ListAllUser.List(i)
ListUserSelect.ItemData(i) = ListAllUser.ItemData(i)
Next i
End Sub

Private Sub Label46_Click()
ListUserSelect.Clear
End Sub

Private Sub Label47_Click()
If ListUserSelect.ListIndex > -1 Then
ListUserSelect.RemoveItem ListUserSelect.ListIndex
End If
End Sub

Private Sub LblSelect_Click()
If ListAllUser.ListIndex = -1 Then Exit Sub
ListUserSelect.AddItem ListAllUser.List(ListAllUser.ListIndex)
ListUserSelect.ItemData(ListUserSelect.NewIndex) = ListAllUser.ItemData(ListAllUser.ListIndex)
End Sub

Private Sub Option6_Click()
HidUnder
End Sub

Private Sub Option7_Click()
HidUnder
End Sub
Sub HidUnder()
If Option8.value = True Then
Fra(0).Visible = False
Fra(1).Visible = False
Fra(8).Visible = False
Fra(9).Visible = False
Fra(10).Visible = False
Else
Fra(0).Visible = True
Fra(1).Visible = True
Fra(8).Visible = True
Fra(9).Visible = True
Fra(10).Visible = True
End If
End Sub

Private Sub OptType6_Click(Index As Integer)
Me.TxtOpenBalance6.Enabled = Not OptType6(2).value
Me.TxtOpenBalance6.Text = IIf(OptType6(2).value = True, 0, Me.TxtOpenBalance6.Text)
End Sub
Private Sub OptType8_Click(Index As Integer)
Me.TxtOpenBalance8.Enabled = Not OptType8(2).value
Me.TxtOpenBalance8.Text = IIf(OptType8(2).value = True, 0, Me.TxtOpenBalance8.Text)
End Sub

Private Sub RdTyp_Click(Index As Integer)
ISButton4.Enabled = False
ISButton3.Enabled = False
If RdTyp(0).value = True Then
Else
ISButton4.Enabled = True
ISButton3.Enabled = True
End If
End Sub
Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
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
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("MofrdID"), False, True)
                .TextMatrix(Row, .ColIndex("MofrdID")) = StrAccountCode
                 End Select

       ' Me.TxtTotal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("value"), .Rows - 1, .ColIndex("value"))

        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
    End With

ErrTrap:
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With Grid
        Select Case .ColKey(Col)
            Case "Valuee"
                .ComboList = ""
        End Select

    End With
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    With Grid
        Select Case .ColKey(Col)

            Case "name"

                StrSQL = " select * from mofrdat "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Grid.BuildComboList(rs, "eq_sys, *mofrad_name", "mofrad_code")
                Else
                    StrComboList = Grid.BuildComboList(rs, "eq_sys, *mofrad_namee", "mofrad_code")
                End If
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With
End Sub

Private Sub GridSub_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim rat As Integer
With GridSub
Select Case .ColKey(Col)
Case "rate"
rat = IIf(Not IsNumeric(.TextMatrix(Row, .ColIndex("rate"))), 0, val(.TextMatrix(Row, .ColIndex("rate"))))
  .TextMatrix(Row, .ColIndex("SubValue")) = rat / 100 * val(total_after_discount.Text)
    
        End Select
            If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
ReLineGrid2
    End With
End Sub


Private Sub RemoveGridRow()
If Me.TxtModFlg.Text <> "R" Then

    With Me.GridSub

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid2
   End If
End Sub

Private Sub RemoveGridRow1()
If Me.TxtModFlg.Text <> "R" Then

    With Me.VSFlexGrid2

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
    End If
End Sub
Private Sub RemoveGridRow22(Optional Rowx As Long)
Dim StrSQL As String
Dim StrMSG As String
Dim pandid As Double
If Me.TxtModFlg.Text <> "R" Then
    With Me.Fg_Journal
        If Rowx <= 0 Then Exit Sub
        pandid = val(Fg_Journal.TextMatrix(Rowx, Fg_Journal.ColIndex("oprid")))
            StrSQL = "Delete From TblExpensiveOper Where Pand=" & pandid & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblMatrials Where Pand=" & pandid & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblEmpOper Where Pand=" & pandid & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblEquepment Where Pand=" & pandid & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                     StrSQL = "Delete From terms_operations Where ProjectDes_ID=" & pandid & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
        .RemoveItem Rowx
      
    End With

    ReLineGrid
    End If
End Sub
Private Sub RemoveGridRow2()
Dim StrSQL As String
Dim StrMSG As String
Dim pandid As Double
If Me.TxtModFlg.Text <> "R" Then

    With Me.Fg_Journal

        If .Row <= 0 Then Exit Sub
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "”Ê› Ì „ Õ–› ﬂ· «·⁄„·Ì«  «·„— »ÿÂ »Â–« «·»‰œ Â·  —Ìœ «·Õ–›"
            Else
                StrMSG = "All process related to this term will be deleted , are you sure you want to continue"
            End If
        If MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title) = vbYes Then
        pandid = val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("oprid")))
            StrSQL = "Delete From TblExpensiveOper Where Pand=" & pandid & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblMatrials Where Pand=" & pandid & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblEmpOper Where Pand=" & pandid & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblEquepment Where Pand=" & pandid & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                     StrSQL = "Delete From terms_operations Where ProjectDes_ID=" & pandid & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
        .RemoveItem .Row
        Else
        Exit Sub
        End If
    End With

    ReLineGrid
    End If
End Sub
Private Sub GridSub_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim LngItemID As Long
    Dim LngStoreID As Long
    Dim rdate As Date
  ' Dim frm As FrmGridAddItemComment
    Dim Frm1 As FrmRegesterDateProject

    'On Error GoTo ErrTrap

    With Me.GridSub

        Select Case .ColKey(Col)

                 Case "subdate"
                  LngRow = Row

 LngCol = Col
             ' ItemProductionDate Row, Col, , 1
                Load FrmRegesterDateProject
                FrmRegesterDateProject.show

                    
                End Select
                End With

End Sub

Private Sub GridSub_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With Me.GridSub

        Select Case .ColKey(Col)

                 Case "subdate"
    
            .ColComboList(.ColIndex("subdate")) = "..."
            Case "rate"
          If .ColComboList(.ColIndex("subdate")) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "ÌÃ» «Œ Ì«—  «—ÌŒ «·œ›⁄Â «Ê·«"
            Else
                MsgBox "payment data must be selected first"
            End If
            Exit Sub
          End If
              Case "SubValue"
          If .ColComboList(.ColIndex("subdate")) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "ÌÃ» «Œ Ì«—  «—ÌŒ «·œ›⁄Â «Ê·«"
            Else
                MsgBox "payment data must be selected first"
            End If
          Exit Sub
          End If
            End Select
            End With
End Sub
Private Sub ImgFavorites_Click()
    AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub
Private Sub ISButton4_Click()
      Grid.Clear flexClearScrollable, flexClearEverything
      Grid.Rows = 1
CD1.ShowOpen
txtFile.Text = CD1.filename
End Sub

Private Sub Label37_Click()
Fra(2).Visible = False
End Sub
Private Sub ISButton3_Click()
Dim CuntNo As Double
Dim Pandcode As String
Dim Total As Double
Dim count1 As Double
Dim Count2 As Double
Dim QtyNo As Double
Dim SumValEx As Double
Dim ProcessID As Double
On Error Resume Next
Dim astrSplit2tems2() As String
If Me.TxtModFlg.Text <> "R" Then
   FgMainDes.Clear flexClearScrollable, flexClearEverything
   FgMainDes.Rows = 1
   Fg_Journal.Clear flexClearScrollable, flexClearEverything
   Fg_Journal.Rows = 2
    Fg_Journal.Enabled = True
    
    If SystemOptions.UserInterface = ArabicInterface Then
        If txtFile.Text = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
    Else
        If txtFile.Text = "" Then MsgBox "Select file first": Exit Sub
    End If
Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Long
Dim NoINdex As Integer
Dim DESINdex As Integer
Dim UnitINdex As Integer
Dim QtyINdex As Integer
Dim PriceINdex As Integer
Dim UnitID As Double
NoINdex = -1
DESINdex = -1
UnitINdex = -1
QtyINdex = -1
PriceINdex = -1
    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")
    ExcelObj.Workbooks.Open txtFile.Text   ' App.Path & "\TrialBalance.xls"
DoEvents
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
 Dim Counter As Integer
 Counter = 0
    With ExcelSheet
    i = 1
    Do Until i > 30

    Select Case .cells(1, i)
    Case "NO"
    NoINdex = i
    Case "DES"
    DESINdex = i
    Case "Unit"
    UnitINdex = i
    Case "Qty"
    QtyINdex = i
    Case "Price"
    PriceINdex = i
    End Select

        i = i + 1
    Loop
    End With
    Dim NO As String
    Dim des As String
    Dim Unit As String
 ''//////////////
 If NoINdex <> -1 Then
     With ExcelSheet
    i = 2
    Do Until .cells(i, NoINdex) & "" = ""
    NO = .cells(i, NoINdex)
    If DESINdex <> -1 Then
    des = .cells(i, DESINdex)
    End If
    If UnitINdex <> -1 Then
    Unit = .cells(i, UnitINdex)
    End If
    If QtyINdex <> -1 Then
    Qty = .cells(i, QtyINdex)
    End If
    If PriceINdex <> -1 Then
    Price = .cells(i, PriceINdex)
    End If
    
  ''//////
 
  If GetNoLetters(NO, "-") = 0 And GetNoLetters(NO, "_") = 0 And GetNoLetters(NO, "\") = 0 And GetNoLetters(NO, "/") = 0 Then
  With FgMainDes
.Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("FullCode")) = NO
        .TextMatrix(.Rows - 1, .ColIndex("Name")) = des
        .TextMatrix(.Rows - 1, .ColIndex("QtyNo")) = Qty
        .TextMatrix(.Rows - 1, .ColIndex("Price")) = Price
        .TextMatrix(.Rows - 1, .ColIndex("PriceExe")) = val(.TextMatrix(.Rows - 1, .ColIndex("Price"))) * val(.TextMatrix(i - 1, .ColIndex("QtyNo")))
        If des <> "" Then
      Pand = IIf(Not IsNumeric(.TextMatrix(.Rows - 1, .ColIndex("ID"))), 0, val(.TextMatrix(.Rows - 1, .ColIndex("ID"))))
       If Me.Checked(0, 0, Pand) = True Then
        Else
       Pand = 1
        maxx 0, 0, Pand
        End If
        .TextMatrix(.Rows - 1, .ColIndex("ID")) = Pand
       End If
 End With
 ElseIf GetNoLetters(NO, "-") >= 1 Or GetNoLetters(NO, "_") >= 1 Or GetNoLetters(NO, "\") >= 2 And GetNoLetters(NO, "/") >= 1 Then
 FillBand
 With Fg_Journal
 Unit = Trim(Unit)
UnitID = GetIDUnit(Unit)
.TextMatrix(.Rows - 1, .ColIndex("PandUnit")) = Unit
.TextMatrix(.Rows - 1, .ColIndex("PandUnitID")) = UnitID
        .TextMatrix(.Rows - 1, .ColIndex("CodeBand")) = NO
        .TextMatrix(.Rows - 1, .ColIndex("des")) = des
        .TextMatrix(.Rows - 1, .ColIndex("PrMainDesID")) = FgMainDes.TextMatrix(FgMainDes.Rows - 1, FgMainDes.ColIndex("ID"))
        .TextMatrix(.Rows - 1, .ColIndex("qty")) = Qty
        .TextMatrix(.Rows - 1, .ColIndex("cost")) = Price
        .TextMatrix(.Rows - 1, .ColIndex("total")) = val(.TextMatrix(.Rows - 1, .ColIndex("cost"))) * val(.TextMatrix(i - 1, .ColIndex("qty")))
Fg_Journal_AfterEdit .Rows - 1, .ColIndex("des")
          ReLineGrid
          CuntNo = 0
        SumValEx = 0
         CuntNo = GetCount(val(.TextMatrix(.Rows - 2, .ColIndex("PrMainDesID"))), Total, SumValEx, count1, Count2, QtyNo)
        SetValue val(.TextMatrix(.Rows - 2, .ColIndex("PrMainDesID"))), Total, SumValEx, count1, Count2, QtyNo, 1
          If Not .TextMatrix(.Rows - 2, .ColIndex("fullcode")) = "" Then
                current_terms = .TextMatrix(.Rows - 2, .ColIndex("fullcode"))
                Pand = val(.TextMatrix(.Rows - 2, .ColIndex("oprid")))
                retrive1 Pand
           End If
          End With
          
   With VSFlexGrid2
    des = Trim(des)
ProcessID = GetIDProcess(des, UnitID)
.TextMatrix(.Rows - 1, .ColIndex("name")) = des
.TextMatrix(.Rows - 1, .ColIndex("OPRIDD")) = ProcessID
.TextMatrix(.Rows - 1, .ColIndex("qty")) = Qty
REFillOprData ProcessID, .Rows - 1
VSFlexGrid2_AfterEdit Rows - 1, val(.ColIndex("qty"))
terms_operations_Click (1)
   End With
 
End If
 If .cells(i, NoINdex) & "" = "" Then Exit Sub
        i = i + 1
    Loop

    End With
     ReLineGrid
  End If
Grid.SetFocus
       ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing
 End If
End Sub
Function GetNoLetters(Optional str As String, Optional Letter As String)
Dim count As Double
count = 0
 For i = 1 To Len(str)
        If mId$(str, i, 1) = Letter Then count = count + 1
    Next
 GetNoLetters = count
End Function
Private Sub Label40_Click()
hideallframe
   
          Frame5.Visible = True
          
          
    
End Sub

Private Sub Label44_Click()
        If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
        ChAuto.value = vbUnchecked
       saveDetails Pand
       End If
       hideallframe
         Frame5.Visible = True
End Sub

Private Sub Label60_Click()
hideallframe
   
          Frame5.Visible = True
          
          
   
End Sub

Private Sub Label63_Click()
C1Elastic2.Visible = False
End Sub

Private Sub Option8_Click()
HidUnder
ReLineGrid
End Sub

Private Sub OptType_Click(Index As Integer)
    Me.TxtOpenBalance.Enabled = Not OptType(2).value
    Me.TxtOpenBalance.Text = IIf(OptType(2).value = True, 0, Me.TxtOpenBalance.Text)
End Sub

Private Sub OptType1_Click(Index As Integer)
    Me.TxtOpenBalance1.Enabled = Not OptType1(2).value
    Me.TxtOpenBalance1.Text = IIf(OptType1(2).value = True, 0, Me.TxtOpenBalance1.Text)
End Sub

Private Sub OptType2_Click(Index As Integer)
    Me.TxtOpenBalance2.Enabled = Not OptType2(2).value
    Me.TxtOpenBalance2.Text = IIf(OptType2(2).value = True, 0, Me.TxtOpenBalance2.Text)
End Sub

Private Sub OptType3_Click(Index As Integer)
    Me.TxtOpenBalance3.Enabled = Not OptType3(2).value
    Me.TxtOpenBalance3.Text = IIf(OptType3(2).value = True, 0, Me.TxtOpenBalance3.Text)
End Sub

Private Sub OptType4_Click(Index As Integer)
    Me.TxtOpenBalance4.Enabled = Not OptType4(2).value
    Me.TxtOpenBalance4.Text = IIf(OptType4(2).value = True, 0, Me.TxtOpenBalance4.Text)
End Sub

Private Sub Form_Load()
'If mdifrmmain.MnuProjects.Visible = False Then
'Frame6.Visible = True
'Else
'Frame6.Visible = False
'End If
 If SystemOptions.AllowGoodPerfAccount = True Then
 Fra(5).Visible = True
 Else
 Fra(5).Visible = False
 End If
Frame21.Visible = False
FillMylist
'equepment
    Dim My_SQL As String
     Me.Left = 0 ' (mdifrmmain.Width - Me.Width) / 2
    Me.Top = 0 '(mdifrmmain.Height - Me.Height) / 2
          If SystemOptions.UserInterface = ArabicInterface Then
                Grid.ColComboList(Grid.ColIndex("TypeEmp")) = "#1; „ÊŸ›|#2; „œÌ—"
      ElseIf SystemOptions.UserInterface = EnglishInterface Then
               Grid.ColComboList(Grid.ColIndex("TypeEmp")) = "#1;Employee |#2;Manager "
      End If
    My_SQL = "  select Account_Code,CusID from TblCustemers  where type=1 "
    My_SQL = My_SQL & " and   BranchId in(" & Current_branchSql & ") "
    My_SQL = My_SQL & " order by CusName "
    fill_combo DcAccount1, My_SQL

    My_SQL = "  select Account_Code,CusID from TblCustemers  where type=1"
    My_SQL = My_SQL & "  and  BranchId in(" & Current_branchSql & ") "
    My_SQL = My_SQL & " order by CusNamee "
Frame20.Visible = False
    fill_combo DcAccount3, My_SQL
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetPrefix Me.DCPreFix, 0, 0 'val(branch_id)
    Dcombos.GetSalesRepData Me.DCEmp1
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetSection Me.DcbDept
    Dcombos.GetPersons Me.DataCombo3
    Dcombos.GetBanks Me.DataCombo2
    Dcombos.GetAmanhNames AmanhNames
    Dcombos.GetMunicipalityNames Me.MunicipalityNames
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select CusID,CusName from TblCustemers  where type=1  "
    Else
        My_SQL = "  select CusID,CusNamee from TblCustemers  where type=1  "
    End If
    My_SQL = My_SQL & "  and  BranchId in(" & Current_branchSql & ") "
    If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = My_SQL & " order by CusName "
    Else
    My_SQL = My_SQL & " order by CusNamee "
    End If

    fill_combo DcAccount2, My_SQL

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select id,code from currency  order by name  "
    Else
        My_SQL = "  select id,code from currency  order by code  "
    End If

    fill_combo DcCurrency, My_SQL
 
    My_SQL = "  select JobTypeID,JobTypeName from TblEmpJobsTypes  order by JobTypeName  "
    fill_combo dcJobTypeName, My_SQL

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select CusID,CusName from TblCustemers  where type=1  "
    Else
        My_SQL = "  select CusID,CusNamee from TblCustemers  where type=1  "
    End If
    My_SQL = My_SQL & "  and  BranchId in(" & Current_branchSql & ") "
    fill_combo DcAccount4, My_SQL
If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "  select id,name from project_status  "
Else
My_SQL = "  select id,namee from project_status  "
End If
    fill_combo DataCombo1, My_SQL
    
If SystemOptions.UserInterface = ArabicInterface Then
My_SQL = "  select id,name from contract_type  "
Else
My_SQL = "  select id,namee from contract_type  "
End If
    fill_combo DataCombo5, My_SQL

    'If SystemOptions.UserInterface = ArabicInterface Then
    'My_SQL = "  select branch_id,branch_name from branches  "
    ' Else
    ' My_SQL = "  select branch_id,branch_namee from branches  "
    ' End If
    '
    'fill_combo Dcbranch, My_SQL

    'Dim Dcombos As ClsDataCombos
    If SystemOptions.UserInterface = ArabicInterface Then
    CBoBasedON.Clear
    CBoBasedON.AddItem "»·« "
    CBoBasedON.AddItem "«„— »Ì⁄"
    Else
    CBoBasedON.Clear
    CBoBasedON.AddItem "NA"
    CBoBasedON.AddItem "Sales Order"
    End If
    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
    End If

    Dcombos.GetBranches dcBranch

    My_SQL = "  select Emp_ID,Emp_name  from TblEmployee  "
    My_SQL = My_SQL & "  where  BranchId in(" & Current_branchSql & ") "
    fill_combo DCEmP, My_SQL

  
    '    Exit Sub
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Set NewGrid.Grid = FG
  '  NewGrid.GridTrans = INVENTORYIN
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.txtTotal = XPTxtSum
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    Set NewGrid.DtpBillDate = Me.XPDtbBill

    Set NewGrid.LblTotalQty = Me.LblTotalQty
    NewGrid.FillGrid
    'FG.WallPaper = BGround.Picture
 
    'SetDtpickerDate XPDtbBill
    'Set Dcombos = New ClsDataCombos
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'LoadSettings
    Set rs = New ADODB.Recordset
    StrSQL = "select * From projects "
     StrSQL = StrSQL & "  where      branch_no in(" & Current_branchSql & ") "
     StrSQL = StrSQL & GetProjectByUser
     
      If SystemOptions.usertype <> UserAdminAll Or val(Current_branch) <> 0 Then
      '  StrSQL = StrSQL & " where   branch_no=" & Current_branch
    End If
    
     StrSQL = StrSQL & " order by id  "
     
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
   ' rs.Open "projects", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"
    'Exit Sub
  
    If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
        ChangeLang
    End If

    ' Me.Width = 10000
  
    If OPEN_NEW_SCREEN = True Then
        Command1_Click (0)
    End If

End Sub

Function ChangeLang()
RdTyp(0).RightToLeft = False
RdTyp(1).RightToLeft = False
RdTyp(0).Caption = "Manual"
RdTyp(1).Caption = "From File"
ISButton4.Caption = "Select Path"
ISButton3.Caption = "Import"
lbl(56).Caption = "Based ON"
Option6.Caption = "Estimation"
Option7.Caption = "Actual"
Option8.Caption = "Under Implem."
Frame8.Caption = "Type"
Cmd(9).Caption = "Delete"
Ptype(0).Caption = "New"
ChAuto.Caption = "Auto"
ALLButton4.Caption = "Users"
lbl(32).Caption = "Total"

Ptype(1).Caption = "Opening"
BtnSalary.Caption = "Data Salary"
Add.Caption = "Insert Row"
Label42.Caption = "Disc.%"
Command1(8).Caption = "  Payments Data"
lbl(20).Caption = "Current Record"
lbl(21).Caption = "NO. Recordes"
    Label43.Caption = "Management"
    With Grid
        .TextMatrix(0, .ColIndex("LineNo")) = "Serial"
        .TextMatrix(0, .ColIndex("name")) = "Name"
        .TextMatrix(0, .ColIndex("Valuee")) = "Value"
        .TextMatrix(0, .ColIndex("TypeEmp")) = "Type"
End With
Label64.Caption = "Data of Salary"
lbl(22).Caption = "By"
Label49(0).Caption = "Executor"
Label49(1).Caption = "Executor%"
Label49(2).Caption = "Collected"
Label49(3).Caption = "Remain"
Label49(4).Caption = "R. Time"
Label49(5).Caption = "Time"
Label58.Caption = "Secretariat"
Label59.Caption = "Municipal"
Label57.Caption = "By"
company.Caption = "company"
Check4.Caption = "Sub-cont"
Label51.Caption = "Gur. No"
Label52.Caption = "Gur. Value"
Label53.Caption = "Gur. Start"
Label54.Caption = "Gur. End"
Label55.Caption = "Gur. Extend"
Label56.Caption = "Gur. Bank"
CmdProcees.Caption = "Delete"
Fra(4).Caption = "Opening Under Implementation"
  OptType5(0).Caption = "Depit"
    OptType5(1).Caption = "Credit"
    OptType5(2).Caption = "Na"
    lbl(23).Caption = "Balance"
    lbl(24).Caption = "Date"
  ''/////////
  Fra(5).Caption = "Opening Good Performance"
    OptType6(0).Caption = "Depit"
    OptType6(1).Caption = "Credit"
    OptType6(2).Caption = "Na"
    lbl(43).Caption = "Balance"
    lbl(44).Caption = "Date"
    
    Fra(6).Caption = "Opening PrePayment"
    OptType8(0).Caption = "Depit"
    OptType8(1).Caption = "Credit"
    OptType8(2).Caption = "Na"
    lbl(45).Caption = "Balance"
    lbl(46).Caption = "Date"
    
    
    Fra(7).Caption = "Accountant Data"
    Fra(8).Caption = "Opening Expenses  Balances "
    OptType(0).Caption = "Depit"
    OptType(1).Caption = "Credit"
    OptType(2).Caption = "Na"
    lbl(14).Caption = "Balance"
    lbl(13).Caption = "Date"

    Fra(9).Caption = "Opening Revenues Balances "
    OptType1(0).Caption = "Depit"
    OptType1(1).Caption = "Credit"
    OptType1(2).Caption = "Na"
    lbl(15).Caption = "Balance"
    lbl(16).Caption = "Date"

    Fra(10).Caption = "Opening  Materials Balances "
    OptType2(0).Caption = "Depit"
    OptType2(1).Caption = "Credit"
    OptType2(2).Caption = "Na"
    lbl(18).Caption = "Balance"
    lbl(17).Caption = "Date"

    Fra(1).Caption = "Opening  Salaries Balances "
    OptType3(0).Caption = "Depit"
    OptType3(1).Caption = "Credit"
    OptType3(2).Caption = "Na"
    lbl(10).Caption = "Balance"
    lbl(11).Caption = "Date"

    Fra(0).Caption = "Opening Invoices Balances  "
    OptType4(0).Caption = "Depit"
    OptType4(1).Caption = "Credit"
    OptType4(2).Caption = "Na"
    lbl(8).Caption = "Balance"
    lbl(9).Caption = "Date"

    Command1(4).Caption = "Opening Balances"

    lbl(40).Caption = " Start D."
    Command3.Caption = "Export"
    ALLButton3.Caption = "Export Details"
    
    Label26.Caption = "Branch"
    temp = XPBtnMove(1).Left
    XPBtnMove(1).Left = XPBtnMove(2).Left
    XPBtnMove(2).Left = temp
    Label36.Caption = "Nearest end"
    Label35.Caption = "Manger"
    temp = XPBtnMove(0).Left
    XPBtnMove(0).Left = XPBtnMove(3).Left
    XPBtnMove(3).Left = temp
  '   SetInterface Me
    lbl(42).Caption = "Users"
    Label16.Caption = "End User ID"
    Label15.Caption = "End User Name"
    Label23.Caption = "Sub-Contractor"
    Label24.Caption = "Contraactor Name"
     lbl(36).Caption = "Name Eng"
    Label6.Caption = "Project Code"
    Label5.Caption = "Staus"
    lbl(35).Caption = "Project Name"
    lbl(37).Caption = "Con. Type"
    lbl(38).Caption = "Project Cost"
    lbl(41).Caption = "Duration"
    Label17.Caption = "End D."
    Label18.Caption = "Earliest end."
    Label22.Caption = "Notes"
    Frame4.Caption = "Color Map"
    Label34.Caption = "Critical"
    'Label21.Caption = "Expanses Account"
    'Label22.Caption = "Revenue Account"
    'Label17.Caption = "Items"
    'Label18.Caption = "Item Description"
    Label19.Caption = "Currency"

    Frame5.Caption = "Terms Data"
    terms_operations(0).Caption = "Terms Operations"
    Frame12.Caption = "Expenses"
    opr_Expenses(1).Caption = "Return To Opr."
    lbl(6).Caption = "Total Expenses"

    With Me.VSFlexGrid3
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("AccountName")) = "Expenses Names"
        .TextMatrix(0, .ColIndex("value")) = "value"
 .TextMatrix(0, .ColIndex("EsToal")) = "Estimated Value"
 
        .TextMatrix(0, .ColIndex("des")) = "des"
 
    End With
CmdPand.Caption = "Delete"
ALLButton1.Caption = "Print"
ALLButton2.Caption = "Detal. Print"

    Frame1.Caption = "Items"

    '   txtid.Alignment = 0
    DataCombo1.RightToLeft = False
    '
   ' CMD_language.Caption = "⁄—»Ì"
    '  Frame4.Visible = True
    Frame3.Visible = True
    '    Frame8.Visible = True
    
    Label9.Caption = "    Projects Data"
    Me.Caption = Label9.Caption
  
    Command1(0).Caption = "new"
    Command1(1).Caption = "save"
    Command1(2).Caption = "Attachments"
    '  SuperLabel2.text = "Search"
    '  Command1(4).Caption = "By ID"
    Command1(5).Caption = "Search"
   
    Label32.Caption = "Discount"
    lbl(39).Caption = "Net Cost"
    Label31.Caption = "Total"
    Command1(3).Caption = "Edit"
    CMDViewGantt(2).Caption = "View Gantt "
    Label12.Caption = "Period"
Label41.Caption = "Employee"
opr_Expenses(2).Caption = "Equipments"
Fra(3).Caption = "Account Info"
Label65.Caption = "Contract No."
    With Me.Fg_Journal
    .TextMatrix(0, .ColIndex("QtyExe")) = "Qty Exe."
    .TextMatrix(0, .ColIndex("PriceExe")) = "Price Exe."
    .TextMatrix(0, .ColIndex("TotalExe")) = "Total Exe."
    
    .TextMatrix(0, .ColIndex("CodeBand")) = "Code"
    .TextMatrix(0, .ColIndex("Printtt")) = "Print"
    .TextMatrix(0, .ColIndex("PrintttAn")) = "Print Detal."
    .TextMatrix(0, .ColIndex("Remark")) = "Remark"
    .TextMatrix(0, .ColIndex("QtyNo")) = "Count"
        .TextMatrix(0, .ColIndex("LineNo")) = "I"
        .TextMatrix(0, .ColIndex("PandUnit")) = "Unit"
        .TextMatrix(0, .ColIndex("PandName")) = "Select Des"
        
        .TextMatrix(0, .ColIndex("fullcode")) = "Term Code"

        .TextMatrix(0, .ColIndex("des")) = "Des"
        .TextMatrix(0, .ColIndex("qty")) = "Qty"
        .TextMatrix(0, .ColIndex("cost")) = "Cost"
        .TextMatrix(0, .ColIndex("total")) = "Total"
        
         .TextMatrix(0, .ColIndex("esQty")) = "Estm. Qty"
        .TextMatrix(0, .ColIndex("EsPrice")) = "Estm.  Cost"
        .TextMatrix(0, .ColIndex("EstTotal")) = "Estm.  Total"
     
        .TextMatrix(0, .ColIndex("discount")) = "Discount"
        .TextMatrix(0, .ColIndex("net")) = "Net"
        .TextMatrix(0, .ColIndex("By")) = "Sub-contarctor"
        .TextMatrix(0, .ColIndex("PrMainDesID")) = "Follow Main Term "
        
    End With
  '  Cmd.Caption = "Delete"
'ma   Frame6.Caption = "Equpiments Data"
lbl(7).Caption = "Totals"
opr_Expenses(3).Caption = "Return To Operations"
Fra(2).Caption = "Payments Data"
lbl(19).Caption = "Estim. Qty"
lbl(12).Caption = "Estim. Price"
  With Me.GridSub
  
 



        .TextMatrix(0, .ColIndex("id")) = "id"
        .TextMatrix(0, .ColIndex("subdate")) = "date"

        .TextMatrix(0, .ColIndex("DesTerm")) = "Finished Term/Process"
        .TextMatrix(0, .ColIndex("rate")) = "Rate"
        .TextMatrix(0, .ColIndex("SubValue")) = "Value"
         
         .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
     
    End With
    
  With Me.VSFlexGrid4
  


        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("FixedAsset")) = "Equipment"

        .TextMatrix(0, .ColIndex("EstHour")) = "Estimated Hour"
        .TextMatrix(0, .ColIndex("ActualHour")) = "Actual Hour"
        .TextMatrix(0, .ColIndex("TotalEs")) = "Total Estimated"
        .TextMatrix(0, .ColIndex("value")) = "Actual Total "
         .TextMatrix(0, .ColIndex("des")) = "Des"
     
    End With
    
    Frame11.Caption = "Terms Operations"

    With Me.VSFlexGrid2
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("Symbol")) = "Symbol"
        .TextMatrix(0, .ColIndex("expensive")) = "Expenses"
        .TextMatrix(0, .ColIndex("startDate")) = "startDate"
        .TextMatrix(0, .ColIndex("EndDate")) = "EndDate"
        .TextMatrix(0, .ColIndex("mat")) = "Materials"
        .TextMatrix(0, .ColIndex("equep")) = "Equip"
        .TextMatrix(0, .ColIndex("employee")) = "Employee"
        .TextMatrix(0, .ColIndex("EquepVal")) = "EquepVal"
        .TextMatrix(0, .ColIndex("Pre")) = "Based On"
        .TextMatrix(0, .ColIndex("Earlystartweek")) = "E. start " & getoprTitle
        .TextMatrix(0, .ColIndex("startweek")) = "start " & getoprTitle
        .TextMatrix(0, .ColIndex("EarlyEndWeek")) = "E. End " & getoprTitle
        .TextMatrix(0, .ColIndex("EndWeek")) = "End " & getoprTitle
        .TextMatrix(0, .ColIndex("Critical")) = "Critical"
        .TextMatrix(0, .ColIndex("fullcode")) = "OPR Code"
        .TextMatrix(0, .ColIndex("name")) = "OPR Name"
        .TextMatrix(0, .ColIndex("period")) = "p " & getoprTitle
        .TextMatrix(0, .ColIndex("period1")) = "Slack " & getoprTitle
        .TextMatrix(0, .ColIndex("total_items")) = "Total Items Cost"
        .TextMatrix(0, .ColIndex("total_salary")) = "Total Salary"
        .TextMatrix(0, .ColIndex("total_expenses")) = "Total Expenses"
        .TextMatrix(0, .ColIndex("total")) = "Total"
        .TextMatrix(0, .ColIndex("qty")) = "Qty"
        .TextMatrix(0, .ColIndex("unitname")) = "Unit Name"
        .TextMatrix(0, .ColIndex("periodView")) = "Default Period"
        .TextMatrix(0, .ColIndex("Actperiod")) = "Actual Period"
    End With

    opr_items(0).Caption = " Items"
    employee_details.Caption = " Labors "
    opr_Expenses(0).Caption = " Expenses"
    terms_operations(1).Caption = "Return To  Terms"
    Label28.Caption = "Total"
    Command1(6).Caption = "Undo"
    Command1(7).Caption = "Delete"
    lbl(4).Visible = False
    lbl(5).Visible = False
  '  Shape1.Visible = False
    Frame10.Caption = "Labors  Data"
    Label27.Caption = "No of Labors "
    Label29.Caption = "Total salaaries"
    opr_emplyees_name.Caption = "Return To Opr."

    Label30.Caption = "ID"
    Label4.Caption = "Count"
    Label11.Caption = "W. Days"
    Label10.Caption = "Start Date"
    Command2.Caption = "Add"

    Label3.Caption = "Select Job Type"
    Option4.Caption = "Estimation"
    Option5.Caption = "Allocation"

    With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("code")) = "Code"
        .TextMatrix(0, .ColIndex("name")) = "Name"
        .TextMatrix(0, .ColIndex("jobname")) = "Job Name"
        .TextMatrix(0, .ColIndex("daysalary")) = "Day Salary"
        .TextMatrix(0, .ColIndex("Count")) = "No.Of.Days"
        .TextMatrix(0, .ColIndex("total")) = "Total"
        .TextMatrix(0, .ColIndex("des")) = "Remark"
    End With

    Frame1.Caption = "OPR Items"
    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = "Status"
    lbl(28).Caption = "Serial"
    lbl(27).Caption = "QTY"
    lbl(26).Caption = "Price"
    lbl(0).Caption = "Avilable"
    lbl(1).Caption = "Reserved"
    lbl(3).Caption = "ON order"
    lbl(2).Caption = "Total"
    opr_items(1).Caption = "Return To Opr."
    
    With FgMainDes
        .TextMatrix(0, .ColIndex("FullCode")) = "Term Code"
        .TextMatrix(0, .ColIndex("Name")) = "Term Name"
        .TextMatrix(0, .ColIndex("QtyNo")) = "Total Count"
        .TextMatrix(0, .ColIndex("Qty")) = "Total Qty"
        .TextMatrix(0, .ColIndex("Price")) = "Total Price"
        .TextMatrix(0, .ColIndex("Total")) = "Actual Net"
        .TextMatrix(0, .ColIndex("QtyExe")) = "Total Implemented Qty"
        .TextMatrix(0, .ColIndex("PriceExe")) = "Total Price for Implemented Qty"
        .TextMatrix(0, .ColIndex("TotalExe")) = "Total openning balance"
        .TextMatrix(0, .ColIndex("Remarks")) = "Notes"
        .TextMatrix(0, .ColIndex("PrintPand")) = "Print"
    End With
    Label67.Caption = "Main Terms"
   lbl(25).Caption = "Total"
    Cmd(1).Caption = "Delete Row"
    terms_operations(2).Caption = "Detailed Terms"
 
End Function

Function SaveAutoVoucher(VType As Integer)
    'Vtype = 0   Œ’Ì’ ›⁄·Ï
    ''Vtype = 3  Œ’Ì’  ﬁœÌ—Ì ›ﬁÿ
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim LngDevID As Long
    Dim voucherid As Integer
    'On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
        
    rs.Open "opr_Employee", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    If Me.TxtModFlg.Text <> "R" Then
 
        Cn.Execute "delete opr_Employee where auto=1 and Project_id=" & val(Me.txt_project_id.Text)
    
        rs.AddNew
        voucherid = CStr(new_id("opr_Employee", "ID", "", True))
        rs("ID").value = voucherid
   
        rs("Start_date").value = XPDtbTrans.value
        rs("Project_id").value = val(txt_project_id)
        rs("opr_type").value = VType
        'Vtype = 0   Œ’Ì’ ›⁄·Ï
        ''Vtype = 3  Œ’Ì’  ﬁœÌ—Ì ›ﬁÿ
        rs("Auto").value = 1
        rs("recorddate").value = Date
        rs("term_Fullcode").value = current_terms
     
        rs("opr_Fullcode").value = current_opr

        rs.update
    
        Set RsDev = New ADODB.Recordset
        
        RsDev.Open "opr_employee_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
        Dim i As Integer

        With Me.VSFlexGrid1

            For i = .FixedRows To .Rows - 1

                If .TextMatrix(i, .ColIndex("id")) <> "" Then
         
                    RsDev.AddNew
                    RsDev("pk_id").value = voucherid
                    RsDev("emp_code").value = .TextMatrix(i, .ColIndex("code"))
                    RsDev("emp_name").value = .TextMatrix(i, .ColIndex("name"))
                    RsDev("JobTypeName").value = .TextMatrix(i, .ColIndex("jobname"))
                    RsDev("JobTypeID").value = .TextMatrix(i, .ColIndex("jobid"))
            
                    RsDev("Emp_id").value = .TextMatrix(i, .ColIndex("id"))
                    RsDev("Start_date").value = XPDtbTrans.value
                    RsDev("Project_id").value = val(Me.txt_project_id)
                    RsDev("opr_type").value = VType
            
                    RsDev("term_Fullcode").value = current_terms
           
                    RsDev("opr_Fullcode").value = current_opr
                    RsDev("daysalary").value = val(.TextMatrix(i, .ColIndex("daysalary")))
                    RsDev("count").value = val(.TextMatrix(i, .ColIndex("count")))
                    RsDev("total").value = val(.TextMatrix(i, .ColIndex("total")))
     
                    If VType = 0 Then
                        save_employee_current_status val(Me.txt_project_id), current_terms, current_opr, val(.TextMatrix(i, .ColIndex("id")))
                    End If

                    RsDev.update
                    
                End If
            
                '
            Next i

        End With
 
    End If

    Exit Function
ErrTrap:
     
End Function

Private Sub opr_emplyees_name_Click()
    calcnets
    VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("total_salary")) = val(txt_emp_salary)
    ReLineGrid
    Frame10.Visible = False
    Set RsDev = New ADODB.Recordset
    Dim sql As String

    '«‰‘«¡ ”‰œ  Œ’Ì’ «·Ì » «—ÌŒ «·„Õœœ
    If Option4.value = True Then
        SaveAutoVoucher (3) ' ﬁœÌ—
    ElseIf Option5.value = True Then
        SaveAutoVoucher (0) '›⁄·Ì
    End If
 
    Exit Sub
 
    If Option5.value = True Then

        With VSFlexGrid1

            For i = .FixedRows To .Rows - 2

                If .TextMatrix(i, .ColIndex("id")) <> "" Then
                    sql = "Select * from TblEmployee where Emp_ID=" & .TextMatrix(i, .ColIndex("id"))
                    RsDev.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
         
                    RsDev("opr_fullcode").value = current_opr
                    RsDev("project_id").value = val(Me.txt_project_id)
                    RsDev("term_id").value = val(current_terms)
                    RsDev("opr_id").value = val(current_opr)
        
                    RsDev.update
                    RsDev.Close
                End If

            Next i
    
        End With

    End If

    'VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("salary")) = IIf(Not IsNumeric(Me.txt_emp_salary.text), 0, Me.txt_emp_salary.text)
    'VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("count")) = IIf(Not IsNumeric(Me.txt_employee_count.text), 0, Me.txt_employee_count.text)
 
    sql = "Select * from terms_operations where fullcode='" & current_opr & "'"
    RsDev.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsDev.RecordCount <= 0 Then Exit Sub
    RsDev("count").value = IIf(Not IsNumeric(Me.txt_employee_count.Text), 0, Me.txt_employee_count.Text)
        
    RsDev("salary").value = IIf(Not IsNumeric(Me.txt_emp_salary.Text), 0, Me.txt_emp_salary.Text)
        
    RsDev.update
    RsDev.Close

End Sub

Private Sub opr_expenses_Click(Index As Integer)

    Select Case Index
Case 2
'ma Frame6.Visible = True

Case 3
'ma Frame6.Visible = False

        Case 0
  
            VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid3.Rows = 2
            VSFlexGrid3.Enabled = True

            If Not VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode")) = "" Then
                Frame12.Visible = True

                current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode"))
                Retrive3 current_opr

                If SystemOptions.UserInterface = ArabicInterface Then
                    Frame1.Caption = "„’«—Ì› «·⁄„·Ì… —ﬁ„ :   " & "  " & current_opr
                Else
                    Frame1.Caption = "Expenses For Operation NO: " & "  " & current_opr
                End If
        
                XPTxtSum.Text = 0
            End If

        Case 1

            VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("total_expenses")) = val(txt_expenses_total)
            ReLineGrid

            StrSQL = "Delete From opr_expenses where opr_fullcode='" & current_opr & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
        
            Set RSTransDetails = New ADODB.Recordset
            RSTransDetails.Open "[opr_expenses]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

            For RowNum = 1 To VSFlexGrid3.Rows - 1

                If VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("ExpensesID")) <> "" Then
  
                    RSTransDetails.AddNew
    
                    RSTransDetails("opr_fullcode").value = current_opr
     
                    RSTransDetails("ExpensesID").value = IIf((VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("ExpensesID")) = ""), Null, val(VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("ExpensesID"))))
                    RSTransDetails("AccountCode").value = IIf((VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("AccountCode")) = ""), Null, VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("AccountCode")))
                    RSTransDetails("AccountName").value = IIf((VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("AccountName")) = ""), Null, VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("AccountName")))
                    RSTransDetails("value").value = IIf((VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("value")) = ""), Null, val(VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("value"))))
                    RSTransDetails("des").value = IIf((VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("des")) = ""), Null, VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("des")))
 
                    RSTransDetails.update
                End If

            Next

            Frame12.Visible = False

    End Select

End Sub

Private Sub opr_items_Click(Index As Integer)
    Dim currentqty As Double

    Select Case Index

        Case 0

            If Not VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode")) = "" Then
                Frame1.Visible = True
                currentqty = val(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("qty")))
                current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("OPRIDD"))
                Retrive2 current_opr, currentqty

                If SystemOptions.UserInterface = ArabicInterface Then
                    Frame1.Caption = "„Ê«œ «·⁄„·Ì… —ﬁ„ :   " & "  " & current_opr
                Else
                    Frame1.Caption = "Items For Operations NO :   " & "  " & current_opr
                End If
        
                ' XPTxtSum.text = 0
                With FG
                    Me.XPTxtSum.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Valu"), .Rows - 1, .ColIndex("Valu"))
                End With

            End If

        Case 1
            VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("total_items")) = val(XPTxtSum)
            ReLineGrid
            StrSQL = "Delete From Transaction_Details where  (payed is null )  and  opr_fullcode='" & current_opr & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
        
            Set RSTransDetails = New ADODB.Recordset
     '       RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  
            For RowNum = 1 To FG.Rows - 1

                If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
  
                    RSTransDetails.AddNew
    
                    RSTransDetails("opr_fullcode").value = current_opr
                    RSTransDetails("Project_id").value = val(txt_project_id.Text)
                    RSTransDetails("term_id").value = val(current_terms)
                    RSTransDetails("opr_id").value = val(current_opr)
    
                    RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                    RSTransDetails("Price").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                    RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                    RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
 
                    RSTransDetails("UnitID").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
 
                    RSTransDetails.update
                End If

            Next

            Frame1.Visible = False

Case 2
'ma Frame6.Enabled = True

Case 3
Frame6.Enabled = False

    End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String

    'On Error GoTo ErrTrap
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

Private Sub Option4_Click()
    XPDtbTrans.Enabled = False
End Sub

Private Sub Option5_Click()
    XPDtbTrans.Enabled = True
End Sub
Sub RetriveAutoProcess(Optional pandid As Double = 0, Optional pandid2 As Double)
Dim sql As String
Dim k As Integer
Dim i As Long
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
If Me.TxtModFlg.Text = "N" Then
    VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid2.Rows = 1
    VSFlexGrid2.Enabled = True
    Else
    VSFlexGrid2.Rows = VSFlexGrid2.Rows - 1
  End If
sql = " SELECT     TblProcessDEFID, ProcessName, ProcessNameE, PandID"
sql = sql & " From dbo.TblProcessDEF"
If Me.TxtModFlg.Text = "N" Then
sql = sql & " Where (pandid = " & pandid & ") "
End If
If Me.TxtModFlg.Text = "E" Then
sql = sql & " Where (pandid = " & pandid & ") and TblProcessDEFID not in( select OPRIDD from terms_operations where  project_id= " & txt_project_id.Text & " and ProjectDes_ID=" & pandid2 & " )"
End If

Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With VSFlexGrid2
k = .Rows
.Rows = .Rows + Rs3.RecordCount
Rs3.MoveFirst
For i = k To .Rows - 1
 .TextMatrix(i, .ColIndex("qty")) = 1
.TextMatrix(i, .ColIndex("OPRIDD")) = IIf(IsNull(Rs3("TblProcessDEFID").value), 0, Rs3("TblProcessDEFID").value)
If val(.TextMatrix(i, .ColIndex("FlgOper"))) <> 1 Then
FillGrid2 val(.TextMatrix(i, .ColIndex("OPRIDD"))), i
End If
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs3("ProcessName").value), "", Rs3("ProcessName").value)
Else
.TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs3("ProcessNameE").value), "", Rs3("ProcessNameE").value)
End If
 REFillOprData val(.TextMatrix(i, .ColIndex("OPRIDD"))), i
 Rs3.MoveNext
Next i
.Rows = .Rows + 1
End With
End If
ReLineGrid
End Sub
Sub FillGrid2(Optional ID As Double = 0, Optional Row As Long)
Dim rs As ADODB.Recordset
Dim sql As String
Dim str As String
Dim i As Integer
Set rs = New ADODB.Recordset
Dim Total As Double

sql = " SELECT     TOP 100 PERCENT dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEFDetails.ItemId, dbo.TblProcessDEFDetails.UnitID, dbo.TblProcessDEFDetails.Price,"
sql = sql & "                      dbo.TblProcessDEFDetails.cost , dbo.TblItems.itemcode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee , dbo.TblItems.Fullcode"
sql = sql & "  FROM         dbo.TblProcessDEF INNER JOIN"
sql = sql & "                       dbo.TblProcessDEFDetails ON dbo.TblProcessDEF.TblProcessDEFID = dbo.TblProcessDEFDetails.TblProcessDEFID INNER JOIN"
sql = sql & "                       dbo.TblItems ON dbo.TblProcessDEFDetails.ItemId = dbo.TblItems.ItemID"
sql = sql & "  Where (dbo.TblProcessDEF.TblProcessDEFID = " & ID & ")"
rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount > 0 Then
rs.MoveFirst
Total = 0
For i = 1 To rs.RecordCount
 str = str & IIf(IsNull(rs("ItemId").value), 0, rs("ItemId").value) & "#"
 str = str & 0 & "#"
 str = str & 0 & "#"
 str = str & 0 & "#"
 str = str & IIf(IsNull(rs("Cost").value), 0, rs("Cost").value) & "#"
 str = str & IIf(IsNull(rs("Price").value), 0, rs("Price").value) & "#"
 str = str & 0 & "#"
  str = str & 0 & "#"
 str = str & Trim("@")
  str = str & CHR(13)
  str = Trim(str)
Total = Total + IIf(IsNull(rs("Cost").value), 0, rs("Cost").value) * IIf(IsNull(rs("Price").value), 0, rs("Price").value)
rs.MoveNext
Next i
End If
With VSFlexGrid2
.TextMatrix(Row, .ColIndex("matrials")) = str
.TextMatrix(Row, .ColIndex("total_items")) = Total
End With

End Sub

Private Sub OptType5_Click(Index As Integer)
Me.TxtOpenBalance5.Enabled = Not OptType5(2).value
    Me.TxtOpenBalance5.Text = IIf(OptType5(2).value = True, 0, Me.TxtOpenBalance5.Text)
End Sub

Private Sub terms_operations_Click(Index As Integer)

    Select Case Index

        Case 0
    hideallframe
            If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("fullcode")) = "" Then
                Frame11.Visible = True
        
                current_terms = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("fullcode"))
          
                Pand = val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("oprid")))
                retrive1 Pand

      If Me.TxtModFlg.Text <> "R" And ChAuto.value = vbChecked Then
                If val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("PanID"))) <> 0 Then
                RetriveAutoProcess val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("PanID"))), Pand
                End If
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                    Frame11.Caption = "⁄„·Ì«  «·»‰œ —ﬁ„ : " & current_terms
                Else
                    Frame11.Caption = "Operations For Term No: " & current_terms
                End If
            End If

        Case 1
        If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
        ChAuto.value = vbUnchecked
       saveDetails Pand
       End If
       hideallframe
   
         Frame5.Visible = True
      Case 2
    If FgMainDes.TextMatrix(FgMainDes.Row, FgMainDes.ColIndex("FullCode")) = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «œŒ«· ﬂÊœ «·»‰œ «·—∆Ì”Ì"
    Else
    MsgBox "Please Enter Code"
    End If
    Exit Sub
    End If
      Frame2.Visible = False
      Frame5.Visible = True
     ' Cn.Execute "update projects_des set SortID=1 where PrMainDesID=" & val(FgMainDes.TextMatrix(FgMainDes.Row, FgMainDes.ColIndex("ID"))) & ""
     ' Cn.Execute "update projects_des set SortID=2 where PrMainDesID <> " & val(FgMainDes.TextMatrix(FgMainDes.Row, FgMainDes.ColIndex("ID"))) & ""
      'RetriveBandDetials val(FgMainDes.TextMatrix(FgMainDes.Row, FgMainDes.ColIndex("ID")))
      FillBand
      FillColor1 val(FgMainDes.TextMatrix(FgMainDes.Row, FgMainDes.ColIndex("ID")))
   
        '    Frame11.Visible = False

    End Select

End Sub
Sub saveDetails(Optional Pand1 As Double = 0)
Dim RsDetails12 As ADODB.Recordset
Dim RsDev2 As ADODB.Recordset
Dim RsDetails1 As ADODB.Recordset
Dim RsDetails11 As ADODB.Recordset
  Dim astrSplit2tems2() As String
Dim astrSplitItems() As String
Dim j As Integer
  Dim st As String
    Dim nElements As Integer
    '  ReLineGrid current_terms

          'ds
        
            ' ⁄„·Ì«  «·»‰Êœ
            '"«·„Ê«œœ"
            'Dim StrSQL As String
                                   Set RsDetails12 = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblExpensiveOper Where (1 = -1)"
   RsDetails12.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
                               Set RsDetails11 = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblEquepment Where (1 = -1)"
   RsDetails11.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
                        Set RsDetails1 = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblEmpOper Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
              Set RsDetails = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblMatrials Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 'sa  If Me.TxtModFlg.text = "E" Then
    StrSQL = "Delete From terms_operations Where ProjectDes_ID =" & Pand & ""
           Cn.Execute StrSQL, , adExecuteNoRecords
   'sa  End If
            Set RsDev2 = New ADODB.Recordset
            RsDev2.Open "terms_operations", Cn, adOpenStatic, adLockOptimistic, adCmdTable
Dim Opr As Double
            Dim i As Integer

            With Me.VSFlexGrid2

                For i = .FixedRows To .Rows - 1

                  
                    If .TextMatrix(i, .ColIndex("name")) <> "" Then

                        RsDev2.AddNew

                        If .TextMatrix(i, .ColIndex("fullcode")) = "" Then
                            RsDev2("fullcode").value = current_terms & "-" & .TextMatrix(i, .ColIndex("LineNo"))
                        Else
                            RsDev2("fullcode").value = .TextMatrix(i, .ColIndex("fullcode"))
                        End If
                    'sa    If Me.TxtModFlg.text = "E" Then
                   
                 StrSQL = "Delete From TblMatrials Where Opr =" & val(.TextMatrix(i, .ColIndex("id"))) & ""
            Cn.Execute StrSQL, , adExecuteNoRecords
             StrSQL = "Delete From TblEmpOper Where Opr =" & val(.TextMatrix(i, .ColIndex("id"))) & ""
            Cn.Execute StrSQL, , adExecuteNoRecords
             StrSQL = "Delete From TblEquepment Where Opr =" & val(.TextMatrix(i, .ColIndex("id"))) & ""
            Cn.Execute StrSQL, , adExecuteNoRecords
              StrSQL = "Delete From TblExpensiveOper Where Opr =" & val(.TextMatrix(i, .ColIndex("id"))) & ""
            Cn.Execute StrSQL, , adExecuteNoRecords
           'sa End If
    '    If Me.TxtModFlg.text = "E" Then
       Opr = val(.TextMatrix(i, .ColIndex("id")))
       If Me.Checked(0, Opr) = True Then
       Else
       Opr = 1
       maxx 0, Opr
       End If
      ' If Me.TxtModFlg.text = "N" Then
     '   OPR = 1
     '  maxx 0, OPR
      ' End If
       
           '   End If
                         RsDev2("project_id").value = val(Me.txt_project_id.Text)
                        RsDev2("term_fullcode").value = current_terms
                        RsDev2("ProjectDes_ID").value = Pand ' ProjectDes_ID
                        RsDev2("id").value = Opr
                       ' RsDev("id").value = .TextMatrix(i, .ColIndex("LineNo"))
                        RsDev2("total").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("total"))), 0, .TextMatrix(i, .ColIndex("total")))
                        RsDev2("name").value = .TextMatrix(i, .ColIndex("name"))
                        RsDev2("period").value = IIf(.TextMatrix(i, .ColIndex("period")) = "", 0, .TextMatrix(i, .ColIndex("period")))
                        RsDev2("count").value = IIf(.TextMatrix(i, .ColIndex("count")) = "", 0, .TextMatrix(i, .ColIndex("count")))
                        RsDev2("salary").value = IIf(.TextMatrix(i, .ColIndex("salary")) = "", 0, .TextMatrix(i, .ColIndex("salary")))
                        RsDev2("total_items").value = IIf(.TextMatrix(i, .ColIndex("total_items")) = "", 0, .TextMatrix(i, .ColIndex("total_items")))
                        RsDev2("total_salary").value = IIf(.TextMatrix(i, .ColIndex("total_salary")) = "", 0, .TextMatrix(i, .ColIndex("total_salary")))
                        RsDev2("total_expenses").value = IIf(.TextMatrix(i, .ColIndex("total_expenses")) = "", 0, .TextMatrix(i, .ColIndex("total_expenses")))
                        RsDev2("EquepVal").value = IIf(.TextMatrix(i, .ColIndex("EquepVal")) = "", 0, .TextMatrix(i, .ColIndex("EquepVal")))
''//
RsDev2("StartDate").value = IIf(.TextMatrix(i, .ColIndex("StartDate")) = "", Null, .TextMatrix(i, .ColIndex("StartDate")))
RsDev2("EndDate").value = IIf(.TextMatrix(i, .ColIndex("EndDate")) = "", Null, .TextMatrix(i, .ColIndex("EndDate")))
                      RsDev2("expen").value = IIf(.TextMatrix(i, .ColIndex("expen")) = "", "", .TextMatrix(i, .ColIndex("expen")))
                       RsDev2("eq").value = IIf(.TextMatrix(i, .ColIndex("eq")) = "", "", .TextMatrix(i, .ColIndex("eq")))
                      RsDev2("emps").value = IIf(.TextMatrix(i, .ColIndex("emps")) = "", "", .TextMatrix(i, .ColIndex("emps")))
                       RsDev2("matrials").value = IIf(.TextMatrix(i, .ColIndex("matrials")) = "", "", .TextMatrix(i, .ColIndex("matrials")))
''//
                        RsDev2("Symbol").value = IIf(.TextMatrix(i, .ColIndex("Symbol")) = "", "", .TextMatrix(i, .ColIndex("Symbol")))
                        RsDev2("Pre").value = IIf(.TextMatrix(i, .ColIndex("Pre")) = "", "", .TextMatrix(i, .ColIndex("Pre")))
                        RsDev2("period1").value = IIf(.TextMatrix(i, .ColIndex("period1")) = "", 0, .TextMatrix(i, .ColIndex("period1")))
                        RsDev2("Earlystartweek").value = IIf(.TextMatrix(i, .ColIndex("Earlystartweek")) = "", 0, .TextMatrix(i, .ColIndex("Earlystartweek")))
                        RsDev2("startweek").value = IIf(.TextMatrix(i, .ColIndex("startweek")) = "", 0, .TextMatrix(i, .ColIndex("startweek")))
                        RsDev2("EarlyEndWeek").value = IIf(.TextMatrix(i, .ColIndex("EarlyEndWeek")) = "", 0, .TextMatrix(i, .ColIndex("EarlyEndWeek")))
                        RsDev2("EndWeek").value = IIf(.TextMatrix(i, .ColIndex("EndWeek")) = "", 0, .TextMatrix(i, .ColIndex("EndWeek")))
                        RsDev2("Critical").value = IIf(.TextMatrix(i, .ColIndex("Critical")) = "", 0, .TextMatrix(i, .ColIndex("Critical")))
                        RsDev2("OPRIDD").value = IIf(.TextMatrix(i, .ColIndex("OPRIDD")) = "", 0, .TextMatrix(i, .ColIndex("OPRIDD")))
                        RsDev2("Actperiod").value = IIf(.TextMatrix(i, .ColIndex("Actperiod")) = "", 0, .TextMatrix(i, .ColIndex("Actperiod")))
                        RsDev2("periodView").value = IIf(.TextMatrix(i, .ColIndex("periodView")) = "", "", .TextMatrix(i, .ColIndex("periodView")))
                        RsDev2("qty").value = IIf(.TextMatrix(i, .ColIndex("qty")) = "", 0, .TextMatrix(i, .ColIndex("qty")))
                        RsDev2("unitname").value = IIf(.TextMatrix(i, .ColIndex("unitname")) = "", 0, .TextMatrix(i, .ColIndex("unitname")))
                        RsDev2("unitid").value = IIf(.TextMatrix(i, .ColIndex("unitid")) = "", 0, .TextMatrix(i, .ColIndex("unitid")))
                        RsDev2.update
                    ''///// «·„Ê«œ
                                   If VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("matrials")) <> "" Then
          st = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("matrials"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         For j = 0 To nElements - 1
          RsDetails.AddNew
                   astrSplit2tems2 = Split(astrSplitItems(j), "#")
        RsDetails("OperCode").value = .TextMatrix(i, .ColIndex("fullcode"))
         RsDetails("ProjectID").value = val(Me.txt_project_id.Text)
         RsDetails("Pand").value = Pand
         RsDetails("Opr").value = IIf(IsNull(RsDev2("id").value), Null, RsDev2("id").value)
         RsDetails("ItemID").value = val(astrSplit2tems2(0))
         RsDetails("Count").value = val(astrSplit2tems2(1))
         RsDetails("Price").value = val(astrSplit2tems2(2))
         'RsDetails("Quntapro").value = val(astrSplit2tems2(3))
         RsDetails("Quntapro").value = val(astrSplit2tems2(4))
         RsDetails("priceapro").value = val(astrSplit2tems2(5))
         RsDetails("catalogID").value = val(astrSplit2tems2(6))
         RsDetails("monthly").value = val(astrSplit2tems2(7))
         RsDetails.update
         Next j
          End If
      ''//////////////
          
          End If
    
                            ''///// «·⁄„«·Â
                                   If VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("emps")) <> "" Then
          st = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("emps"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         For j = 0 To nElements - 1
          RsDetails1.AddNew
         astrSplit2tems2 = Split(astrSplitItems(j), "#")
         RsDetails1("OperCode").value = .TextMatrix(i, .ColIndex("fullcode"))
         RsDetails1("ProjectID").value = val(Me.txt_project_id.Text)
         RsDetails1("Pand").value = Pand
         RsDetails1("Opr").value = IIf(IsNull(RsDev2("id").value), Null, RsDev2("id").value)
         RsDetails1("EmpID").value = val(astrSplit2tems2(0))
         RsDetails1("JobID").value = val(astrSplit2tems2(1))
         RsDetails1("daysalary").value = val(astrSplit2tems2(2))
         RsDetails1("Count").value = val(astrSplit2tems2(3))
                         
         RsDetails1.update
         Next j
          End If
      ''//////////////
                             ''///// «·„⁄œ« 
                                   If VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("eq")) <> "" Then
          st = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("eq"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         For j = 0 To nElements - 1
          RsDetails11.AddNew
         astrSplit2tems2 = Split(astrSplitItems(j), "#")
         RsDetails11("OperCode").value = .TextMatrix(i, .ColIndex("fullcode"))
         RsDetails11("ProjectID").value = val(Me.txt_project_id.Text)
         RsDetails11("Pand").value = Pand
         RsDetails11("Opr").value = IIf(IsNull(RsDev2("id").value), Null, RsDev2("id").value)
         RsDetails11("ExpensesID").value = val(astrSplit2tems2(0))
         RsDetails11("EstHour").value = val(astrSplit2tems2(1))
         RsDetails11("ActualHour").value = val(astrSplit2tems2(2))
         RsDetails11("TotalEs").value = val(astrSplit2tems2(3))
         RsDetails11("value").value = val(astrSplit2tems2(4))
         RsDetails11("des").value = astrSplit2tems2(5)
         RsDetails11("EquepVal").value = val(astrSplit2tems2(6))
              
         RsDetails11.update
       Next j
          End If
      ''//////////////
                                   ''///// «·„’«—Ì›
         If VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("expen")) <> "" Then
          st = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("expen"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         For j = 0 To nElements - 1
          RsDetails12.AddNew
         astrSplit2tems2 = Split(astrSplitItems(j), "#")
         RsDetails12("OperCode").value = .TextMatrix(i, .ColIndex("fullcode"))
         RsDetails12("ProjectID").value = val(Me.txt_project_id.Text)
         RsDetails12("Pand").value = Pand
         RsDetails12("Opr").value = IIf(IsNull(RsDev2("id").value), Null, RsDev2("id").value)
               RsDetails12("AccountCode").value = Replace(Replace(astrSplit2tems2(0), CHR(10), ""), CHR(13), "")
         RsDetails12("EsToal").value = val(astrSplit2tems2(1))
         RsDetails12("value").value = val(astrSplit2tems2(2))
         RsDetails12("Des").value = astrSplit2tems2(3)
        
                         
         RsDetails12.update
         Next j
          End If
      ''//////////////
          
                Next i
    
            End With

End Sub

Function calbetprice()
If Me.TxtModFlg.Text <> "R" Then
Dim discountvalue As Double
Dim netvalue As Double
Dim Projectvalue As Double
Projectvalue = val(TxtProjectCosts.Text)
If val(txt_total_discount) <> 0 Then
discountvalue = val(txt_total_discount)
ElseIf val(TxtDiscountPercentage) <> 0 Then
discountvalue = TxtDiscountPercentage * Projectvalue / 100
End If

total_after_discount.Text = Projectvalue - discountvalue
End If
End Function



Private Sub Text10_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode (Text10.Text), EmpID
        DCEmP.BoundText = EmpID
    End If
End Sub



Private Sub txt_total_discount_KeyUp(KeyCode As Integer, Shift As Integer)
TxtDiscountPercentage.Text = 0
calbetprice
End Sub

Private Sub TxtCustCode2_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode (TxtCustCode2.Text), EmpID
        DCEmp1.BoundText = EmpID
    End If
End Sub

Private Sub TxtDiscountPercentage_KeyUp(KeyCode As Integer, Shift As Integer)
txt_total_discount = 0
calbetprice
End Sub

Private Sub TxtEmpSalary_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtEmpSalary.Text, 0)
End Sub

Private Sub txtid_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        If KeyCode = vbKeyP Then
          print_report2 , 1
        End If
    End If
End Sub

Private Sub txtid_LostFocus()
    'Dim StrSQL As String
    'Dim RsTemp As New ADODB.Recordset
    ' StrSQL = "select * From  projects where fullcode='" & DCPreFix.text & (txtid.text) & "'"
    '            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '            If RsTemp.RecordCount > 0 Then
    '
    '                Msg = "this project code already exist"
    '                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '
    '                Exit Sub
    '            End If
End Sub

Private Sub Retrive3(current_opr As String)
    Dim RsDev As ADODB.Recordset
 
    StrSQL = "SELECT  * from opr_expenses where opr_fullcode='" & current_opr & "'"
  
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid3
   
            .Rows = .FixedRows + RsDev.RecordCount
   
            For i = .FixedRows To .Rows - 1
            
                .TextMatrix(i, .ColIndex("ExpensesID")) = IIf(IsNull(RsDev("ExpensesID").value), "", RsDev("ExpensesID").value)
            
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("AccountCode").value), "", RsDev("AccountCode").value)
            
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("AccountName").value), "", RsDev("AccountName").value)
   
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
          
                .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("des").value), "", RsDev("des").value)
                RsDev.MoveNext
            Next i

            '  If RsDev.RecordCount > 0 Then
            Me.txt_expenses_total.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
            '   End If
        End With

    End If

End Sub

Private Sub Retrive2(current_opr As String, _
                     currentqty As Double)
 
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
 
    'On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Enabled = True
          
    'StrSQL = "select * from Transaction_Details where opr_id=" & current_opr & " and term_id=" & current_terms & " and project_id=" & Me.txt_project_id
 
    StrSQL = " SELECT     TOP 100 PERCENT dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEFDetails.ItemId, dbo.TblProcessDEFDetails.UnitID, dbo.TblProcessDEFDetails.Price, "
    StrSQL = StrSQL & " dbo.TblProcessDEFDetails.Cost, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemCase,"
    StrSQL = StrSQL & " dbo.TblItems.ItemNamee"
    StrSQL = StrSQL & " FROM         dbo.TblProcessDEF INNER JOIN"
    StrSQL = StrSQL & " dbo.TblProcessDEFDetails ON dbo.TblProcessDEF.TblProcessDEFID = dbo.TblProcessDEFDetails.TblProcessDEFID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblUnites ON dbo.TblProcessDEFDetails.UnitID = dbo.TblUnites.UnitID INNER JOIN"
    StrSQL = StrSQL & " dbo.TblItems ON dbo.TblProcessDEFDetails.ItemId = dbo.TblItems.ItemID"
    StrSQL = StrSQL & " WHERE     (dbo.TblProcessDEF.TblProcessDEFID = " & val(current_opr) & ")"

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemId")), "", (RsDetails("ItemId").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemId")), "", Trim(RsDetails("ItemId").value))
            FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Cost")), "", (RsDetails("Cost").value)) * currentqty
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("Valu")) = IIf(IsNull(RsDetails("Cost")), 0, (RsDetails("cost"))) * IIf(IsNull(RsDetails("Price")), 0, (RsDetails("Price").value)) * currentqty
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            RsDetails.MoveNext
        Next Num

    End If

End Sub

Private Sub Retrive2old(current_opr As String)
 
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
 
    'On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Enabled = True
          
    'StrSQL = "select * from Transaction_Details where opr_id=" & current_opr & " and term_id=" & current_terms & " and project_id=" & Me.txt_project_id
 
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where  (payed is null )  and opr_fullcode='" & current_opr & "'"

    'StrSQL = StrSQL + " where opr_id=" & current_opr & " and term_id=" & current_terms & " and project_id=" & Me.txt_project_id

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("Valu")) = IIf(IsNull(RsDetails("Quantity")), 0, (RsDetails("Quantity").value)) * IIf(IsNull(RsDetails("Price")), 0, (RsDetails("Price").value))
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            RsDetails.MoveNext
        Next Num

    End If

End Sub

Private Sub retrive1(Optional pandid As Double = 0)
 
    Dim RsDev3 As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
 
    'On Error GoTo ErrTrap
    VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid2.Rows = 2
  '  VSFlexGrid2.Enabled = True
    txt_opr_total.Text = 0
          
    StrSQL = " select * from terms_operations where ( project_id= " & txt_project_id & " and ProjectDes_ID=" & pandid & " ) ORDER BY id  "
    Set RsDev3 = New ADODB.Recordset
    RsDev3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev3.BOF Or RsDev3.EOF) Then
        RsDev3.MoveFirst
    
        With Me.VSFlexGrid2
            .Rows = .FixedRows + RsDev3.RecordCount

            For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("ProjectDes_ID")) = IIf(IsNull(RsDev3("ProjectDes_ID").value), "", RsDev3("ProjectDes_ID").value)
                .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(RsDev3("fullcode").value), "", RsDev3("fullcode").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDev3("id").value), "", RsDev3("id").value)
             '''/
              .TextMatrix(i, .ColIndex("StartDate")) = IIf(IsNull(RsDev3("StartDate").value), "", RsDev3("StartDate").value)
               .TextMatrix(i, .ColIndex("EndDate")) = IIf(IsNull(RsDev3("EndDate").value), "", RsDev3("EndDate").value)
             .TextMatrix(i, .ColIndex("expen")) = IIf(IsNull(RsDev3("expen").value), "", RsDev3("expen").value)
             .TextMatrix(i, .ColIndex("eq")) = IIf(IsNull(RsDev3("eq").value), "", RsDev3("eq").value)
            .TextMatrix(i, .ColIndex("emps")) = IIf(IsNull(RsDev3("emps").value), "", RsDev3("emps").value)
             .TextMatrix(i, .ColIndex("matrials")) = IIf(IsNull(RsDev3("matrials").value), "", RsDev3("matrials").value)
             
            ''/
            
            
                .TextMatrix(i, .ColIndex("item_id")) = IIf(IsNull(RsDev3("item_id").value), "", RsDev3("item_id").value)
            
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev3("name").value), "", RsDev3("name").value)
            
                .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(RsDev3("id").value), "", RsDev3("id").value)
           
                .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(RsDev3("total").value), "", RsDev3("total").value)
                .TextMatrix(i, .ColIndex("period")) = IIf(IsNull(RsDev3("period").value), "", RsDev3("period").value)
                .TextMatrix(i, .ColIndex("count")) = IIf(IsNull(RsDev3("count").value), "", RsDev3("count").value)
            
                .TextMatrix(i, .ColIndex("salary")) = IIf(IsNull(RsDev3("salary").value), "", RsDev3("salary").value)
 
                .TextMatrix(i, .ColIndex("total_items")) = IIf(IsNull(RsDev3("total_items").value), "", RsDev3("total_items").value)
                .TextMatrix(i, .ColIndex("total_salary")) = IIf(IsNull(RsDev3("total_salary").value), "", RsDev3("total_salary").value)
                .TextMatrix(i, .ColIndex("total_expenses")) = IIf(IsNull(RsDev3("total_expenses").value), "", RsDev3("total_expenses").value)

                .TextMatrix(i, .ColIndex("Symbol")) = IIf(IsNull(RsDev3("Symbol").value), "", RsDev3("Symbol").value)
            
                .TextMatrix(i, .ColIndex("Pre")) = IIf(IsNull(RsDev3("Pre").value), "", RsDev3("Pre").value)
            
                .TextMatrix(i, .ColIndex("period1")) = IIf(IsNull(RsDev3("period1").value), "", RsDev3("period1").value)
            
                .TextMatrix(i, .ColIndex("Earlystartweek")) = IIf(IsNull(RsDev3("Earlystartweek").value), "", RsDev3("Earlystartweek").value)
            
                .TextMatrix(i, .ColIndex("startweek")) = IIf(IsNull(RsDev3("startweek").value), "", RsDev3("startweek").value)
            
                .TextMatrix(i, .ColIndex("EarlyEndWeek")) = IIf(IsNull(RsDev3("EarlyEndWeek").value), "", RsDev3("EarlyEndWeek").value)
            
                .TextMatrix(i, .ColIndex("EndWeek")) = IIf(IsNull(RsDev3("EndWeek").value), "", RsDev3("EndWeek").value)
            
                .TextMatrix(i, .ColIndex("Critical")) = IIf(IsNull(RsDev3("Critical").value), "", RsDev3("Critical").value)
          
                .TextMatrix(i, .ColIndex("OPRIDD")) = IIf(IsNull(RsDev3("OPRIDD").value), "", RsDev3("OPRIDD").value)
            
                .TextMatrix(i, .ColIndex("Actperiod")) = IIf(IsNull(RsDev3("Actperiod").value), "", RsDev3("Actperiod").value)
            
                .TextMatrix(i, .ColIndex("periodView")) = IIf(IsNull(RsDev3("periodView").value), "", RsDev3("periodView").value)
            
                .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(RsDev3("Qty").value), "", RsDev3("Qty").value)
            
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev3("UnitName").value), "", RsDev3("UnitName").value)
            .TextMatrix(i, .ColIndex("EquepVal")) = IIf(IsNull(RsDev3("EquepVal").value), "", RsDev3("EquepVal").value)
                RsDev3.MoveNext
            Next i

            Me.txt_opr_total.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
         .Rows = .Rows + 1
        End With

    End If
          
    ReLineGrid

End Sub
Sub maxx(Optional ByRef Pand As Double = 0, Optional ByRef Opr As Double = 0, Optional ByRef PrMainDesID As Double = 0)
     
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
  Set RsDev = New ADODB.Recordset
  
    If Pand <> 0 Then
   StrSQL = " select max(Pand) as mx from FoxySerial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   Pand = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "FoxySerial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("Pand").value = Pand
RsDev.update
End If

    If Opr <> 0 Then
   StrSQL = " select max(Opr) as mx from FoxySerial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   Opr = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "FoxySerial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("Opr").value = Opr
RsDev.update
End If

    If PrMainDesID <> 0 Then
   StrSQL = " select max(PrMainDesID) as mx from FoxySerial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   PrMainDesID = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
   Set RsDev = New ADODB.Recordset
    RsDev.Open "FoxySerial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("PrMainDesID").value = PrMainDesID
RsDev.update
End If
End Sub
Function Checked(Optional Pand As Double = 0, Optional Opr As Double = 0, Optional PrMainDesID As Double) As Boolean
     Checked = False
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
      Set RsDev = New ADODB.Recordset
    If Pand <> 0 Then
   StrSQL = " select * from FoxySerial where pand=" & Pand & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
    If Opr <> 0 Then
  StrSQL = " select * from FoxySerial where Opr=" & Opr & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
If PrMainDesID <> 0 Then
   StrSQL = " select * from FoxySerial where PrMainDesID=" & PrMainDesID & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If RsDev.RecordCount > 0 Then
   Checked = True
Else
   Checked = False
End If
End If
End Function

Public Sub Retrive(Optional Lngid As Long)
 
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    
        Dim RsDevsub As ADODB.Recordset
    Dim StrSQLsub As String
    
    
    Dim i As Integer
    TxtFillData.Text = "T"
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
            rs.find "id=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    txtopening_balance_voucher_id.Text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)

    If Not (IsNull(rs("OpenBalanceDate").value)) Then
        Me.Dtp.value = rs("OpenBalanceDate").value
        Me.Dtp1.value = rs("OpenBalanceDate").value
        Me.Dtp2.value = rs("OpenBalanceDate").value
        Me.Dtp3.value = rs("OpenBalanceDate").value
        Me.Dtp4.value = rs("OpenBalanceDate").value
        Me.Dtp5.value = rs("OpenBalanceDate").value
        Me.Dtp6.value = rs("OpenBalanceDate").value
        Me.Dtp8.value = rs("OpenBalanceDate").value
     
        ' Me.Dtp.Enabled = True
    Else
        Me.Dtp.value = Date
        Me.Dtp1.value = Date
        Me.Dtp2.value = Date
        Me.Dtp3.value = Date
        Me.Dtp4.value = Date
        Me.Dtp5.value = Date
        Me.Dtp6.value = Date
        Me.Dtp8.value = Date
                    
        '   Me.Dtp.Enabled = False
    End If

    If Not IsNull(rs("OpenBalanceType").value) Then
        Me.TxtOpenBalance.Text = IIf(IsNull(rs("OpenBalance")), "", Trim(rs("OpenBalance")))

        If rs("OpenBalanceType").value = 0 Then
            OptType(0).value = True
            OptType_Click 0
        ElseIf rs("OpenBalanceType").value = 1 Then
            OptType(1).value = True
            OptType_Click 1
        End If
    
    Else
        Me.TxtOpenBalance.Text = 0
        Me.OptType(2).value = True
        OptType_Click 2
    End If

  If IsNull(rs("Pstate").value) Then
Ptype(0).value = True
              
  Else
          If rs("Pstate").value = 0 Then
                
                 Ptype(0).value = True
                Else
                 Ptype(1).value = True
                End If
                
  End If
  If Not IsNull(rs("TypeImport").value) Then
  If (rs("TypeImport").value) = 1 Then
  RdTyp(1).value = True
  Else
  RdTyp(0).value = True
 End If
  Else
  RdTyp(0).value = True
  End If
 Me.txtFile.Text = IIf(IsNull(rs("Path")), "", Trim(rs("Path")))
    If Not IsNull(rs("OpenBalanceType1").value) Then
        Me.TxtOpenBalance1.Text = IIf(IsNull(rs("OpenBalance1")), "", Trim(rs("OpenBalance1")))

        If rs("OpenBalanceType1").value = 0 Then
            OptType1(0).value = True
            OptType1_Click 0
        ElseIf rs("OpenBalanceType1").value = 1 Then
            OptType1(1).value = True
            OptType1_Click 1
        End If
    
    Else
        Me.TxtOpenBalance1.Text = 0
        Me.OptType1(2).value = True
        OptType1_Click 2
    End If

    If Not IsNull(rs("OpenBalanceType2").value) Then
        Me.TxtOpenBalance2.Text = IIf(IsNull(rs("OpenBalance2")), "", Trim(rs("OpenBalance2")))

        If rs("OpenBalanceType2").value = 0 Then
            OptType2(0).value = True
            OptType2_Click 0
        ElseIf rs("OpenBalanceType2").value = 1 Then
            OptType2(1).value = True
            OptType2_Click 1
        End If
    
    Else
        Me.TxtOpenBalance2.Text = 0
        Me.OptType2(2).value = True
        OptType2_Click 2
    End If

    If Not IsNull(rs("OpenBalanceType3").value) Then
        Me.TxtOpenBalance3.Text = IIf(IsNull(rs("OpenBalance3")), "", Trim(rs("OpenBalance3")))

        If rs("OpenBalanceType3").value = 0 Then
            OptType3(0).value = True
            OptType3_Click 0
        ElseIf rs("OpenBalanceType3").value = 1 Then
            OptType3(1).value = True
            OptType3_Click 1
        End If
    
    Else
        Me.TxtOpenBalance3.Text = 0
        Me.OptType3(2).value = True
        OptType3_Click 3
    End If

    If Not IsNull(rs("OpenBalanceType4").value) Then
        Me.TxtOpenBalance4.Text = IIf(IsNull(rs("OpenBalance4")), "", Trim(rs("OpenBalance4")))

        If rs("OpenBalanceType4").value = 0 Then
            OptType4(0).value = True
            OptType4_Click 0
        ElseIf rs("OpenBalanceType4").value = 1 Then
            OptType4(1).value = True
            OptType4_Click 1
        End If
    
    Else
        Me.TxtOpenBalance4.Text = 0
        Me.OptType4(2).value = True
        OptType4_Click 4
    End If
    
        If Not IsNull(rs("OpenBalanceType5").value) Then
        Me.TxtOpenBalance5.Text = IIf(IsNull(rs("OpenBalance5")), "", Trim(rs("OpenBalance5")))

        If rs("OpenBalanceType5").value = 0 Then
            OptType5(0).value = True
            OptType5_Click 0
        ElseIf rs("OpenBalanceType5").value = 1 Then
            OptType5(1).value = True
            OptType5_Click 1
        End If
    
    Else
        Me.TxtOpenBalance5.Text = 0
        Me.OptType5(2).value = True
        OptType5_Click 2
    End If
    If Not IsNull(rs("OpenBalanceType6").value) Then
        Me.TxtOpenBalance6.Text = IIf(IsNull(rs("OpenBalance6")), "", Trim(rs("OpenBalance6")))
        If rs("OpenBalanceType6").value = 0 Then
            OptType6(0).value = True
            OptType6_Click 0
        ElseIf rs("OpenBalanceType6").value = 1 Then
            OptType6(1).value = True
            OptType6_Click 1
        End If
    
    Else
        Me.TxtOpenBalance6.Text = 0
        Me.OptType6(2).value = True
        OptType6_Click 2
    End If
    '''//////
        If Not IsNull(rs("OpenBalanceType8").value) Then
        Me.TxtOpenBalance8.Text = IIf(IsNull(rs("OpenBalance8")), "", Trim(rs("OpenBalance8")))
        If rs("OpenBalanceType8").value = 0 Then
            OptType8(0).value = True
            OptType8_Click 0
        ElseIf rs("OpenBalanceType8").value = 1 Then
            OptType8(1).value = True
            OptType8_Click 1
        End If
    
    Else
        Me.TxtOpenBalance8.Text = 0
        Me.OptType8(2).value = True
        OptType8_Click 2
    End If
    
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Enabled = True
    XPTxtSum.Text = 0
          
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 2
    Fg_Journal.Enabled = True
    FgMainDes.Clear flexClearScrollable, flexClearEverything
    FgMainDes.Rows = 2
    FgMainDes.Enabled = True
    
    txt_total_sum.Text = 0
          
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
    VSFlexGrid1.Enabled = True
         
            GridSub.Clear flexClearScrollable, flexClearEverything
    GridSub.Rows = 2
  '''///////
  Me.TxtEmpSalary.Text = IIf(IsNull(rs("EmpSalary").value), 0, rs("EmpSalary").value)
  Me.TxtMangSalary.Text = IIf(IsNull(rs("MangSalary").value), 0, rs("MangSalary").value)
  Me.lbl(34).Caption = IIf(IsNull(rs("TotalMainDes").value), 0, rs("TotalMainDes").value)
  Me.lbl(33).Caption = IIf(IsNull(rs("TotalMainDesExe").value), 0, rs("TotalMainDesExe").value)
  
  Me.CBoBasedON.ListIndex = IIf(IsNull(rs("OrderType").value), -1, rs("OrderType").value)
  Me.TXTOrDer_no.Text = IIf(IsNull(rs("OrderNo").value), "", rs("OrderNo").value)
  ''////
Me.DCEmp1.BoundText = IIf(IsNull(rs("EmpId1").value), "", rs("EmpId1").value)
Me.DCEmP.BoundText = IIf(IsNull(rs("EmpId").value), "", rs("EmpId").value)

    txt_project_id.Text = IIf(IsNull(rs("id").value), 0, val(rs("id").value))
    DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    DcAccount1.BoundText = IIf(IsNull(rs("End_user_Account").value), "", rs("End_user_Account").value)
    DcCurrency.BoundText = IIf(IsNull(rs("CurrencyID").value), 1, rs("CurrencyID").value)

    DTStartDate.value = IIf(IsNull(rs("StartDate").value), Date, rs("StartDate").value)
        DTEnddate.value = IIf(IsNull(rs("Enddate").value), Date, rs("Enddate").value)
    DpNearEndDate.value = IIf(IsNull(rs("DpNearEndDate").value), Date, rs("DpNearEndDate").value)
     TxtRemarks.Text = IIf(IsNull(rs("Remarkss").value), "", rs("Remarkss").value)

    Me.DcAccount2.BoundText = IIf(IsNull(rs("End_user_id").value), "", rs("End_user_id").value)

    DcAccount3.BoundText = IIf(IsNull(rs("sub_contractor_Account").value), "", rs("sub_contractor_Account").value)
    DcAccount4.BoundText = IIf(IsNull(rs("sub_contractor_id").value), "", rs("sub_contractor_id").value)

    DCPreFix.Text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)

    TxtId.Text = IIf(IsNull(rs("code").value), "", rs("code").value)

    Me.TXTprojectname.Text = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
    Me.TXTprojectnamee.Text = IIf(IsNull(rs("Project_namee").value), "", rs("Project_namee").value)

    Me.TxtProjectCosts.Text = IIf(IsNull(rs("project_cost").value), 0, rs("project_cost").value)

    Me.txt_total_discount.Text = IIf(IsNull(rs("general_discount").value), 0, rs("general_discount").value)
    Me.TxtDiscountPercentage.Text = IIf(IsNull(rs("DiscountPercentage").value), 0, rs("DiscountPercentage").value)
    
    Me.total_after_discount.Text = IIf(IsNull(rs("cost_after_discount").value), 0, rs("cost_after_discount").value)

    Me.dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    Me.DcbDept.BoundText = IIf(IsNull(rs("Dept_ID").value), "", rs("Dept_ID").value)
If IsNull(rs("Project_status").value) Then
    Me.DataCombo1.BoundText = ""
Else
    Me.DataCombo1.BoundText = IIf(IsNull(rs("Project_status").value), "", val(rs("Project_status").value))
End If

    Me.DataCombo5.BoundText = IIf(IsNull(rs("Contract_type").value), "", rs("Contract_type").value)
    Me.txt_total_sum.Text = Round(IIf(IsNull(rs("total").value), 0, rs("total").value), Decimal_Places)

    Me.txt_sub_discount.Text = IIf(IsNull(rs("sub_discount_total").value), "", rs("sub_discount_total").value)

    Me.txt_sub_net.Text = IIf(IsNull(rs("net").value), "", rs("net").value)

    'Me.EXPANSES.text = IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value)
    'Me.REVENUE.text = IIf(IsNull(rs("REVENUE_account").value), "", rs("REVENUE_account").value)

    'Me.Material.text = IIf(IsNull(rs("Material_account").value), "", rs("Material_account").value)
    'Me.legal.text = IIf(IsNull(rs("legal").value), "", rs("legal").value)

    'Me.salary.text = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)

    'Me.txtProject_account.text = IIf(IsNull(rs("Project_account").value), "", rs("Project_account").value)

    Me.XPTxtSum.Text = IIf(IsNull(rs("items_total").value), 0, rs("items_total").value)
    
    ' aladein addy
    Me.Text20(5).Text = IIf(IsNull(rs("JobeDO").value), 0, rs("JobeDO").value)
    Me.Text20(4).Text = IIf(IsNull(rs("JobeDOPercent").value), 0, rs("JobeDOPercent").value)
    Me.Text20(3).Text = IIf(IsNull(rs("JobeGet").value), 0, rs("JobeGet").value)
    Me.Text20(2).Text = IIf(IsNull(rs("JobeRest").value), 0, rs("JobeRest").value)
    Me.Text20(1).Text = IIf(IsNull(rs("JobeTimeLeft").value), 0, rs("JobeTimeLeft").value)
    Me.Text20(0).Text = IIf(IsNull(rs("JobeTime").value), 0, rs("JobeTime").value)
    ''''''''''''''''''''''''''''''
    If Not IsNull(rs.Fields("JobeWork").value) Then
    
    
     If rs.Fields("JobeWork").value = 0 Then
     company.value = True
     Else
     Check4.value = True
     End If
   Else
   company.value = True
   
   End If
   If Not IsNull(rs("UnderImp").value) Then
   If rs("UnderImp").value = 0 Then
   Option7.value = True
   ElseIf rs("UnderImp").value = 1 Then
   Option6.value = True
   ElseIf rs("UnderImp").value = 2 Then
   Option8.value = True
   End If
   Else
   Option7.value = True
   End If
    Me.DataCombo3.BoundText = IIf(IsNull(rs("JobeContractorID").value), 0, rs("JobeContractorID").value)
    Me.AmanhNames.BoundText = IIf(IsNull(rs("Amanhid").value), 0, rs("Amanhid").value)
    Me.MunicipalityNames.BoundText = IIf(IsNull(rs("Municipalityid").value), 0, rs("Municipalityid").value)
    
     ''''''''''''''''''''''''''''''
     Me.TxtContractNo.Text = IIf(IsNull(rs("ContractNo").value), "", rs("ContractNo").value)
    Me.Text21.Text = IIf(IsNull(rs("WarrantyNO").value), 0, rs("WarrantyNO").value)
    Me.Text22.Text = IIf(IsNull(rs("WarrantyValue").value), 0, rs("WarrantyValue").value)
    Me.DTPicker1.value = IIf(IsNull(rs("WarrDateStart").value), 0, rs("WarrDateStart").value)
    Me.DTPicker3.value = IIf(IsNull(rs("WarrDateEnd").value), 0, rs("WarrDateEnd").value)
    Me.Text25.Text = IIf(IsNull(rs("WarrExtension").value), 0, rs("WarrExtension").value)
    Me.DataCombo2.BoundText = IIf(IsNull(rs("WarrBank").value), 0, rs("WarrBank").value)
       ''''   «·»‰Êœ «·—∆Ì”Ì…
           'œ›⁄«  «·„‘—Ê⁄
        '»‰Êœ «·„‘—Ê⁄
    '-----------------------------------------------------------------------------
 
        StrSQL = "SELECT    * from ProjectMainDes "
        StrSQL = StrSQL + " Where (ProjectID =" & Me.txt_project_id.Text & ")"
        Set RsDevsub = New ADODB.Recordset
        RsDevsub.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RsDevsub.BOF Or RsDevsub.EOF) Then
            RsDevsub.MoveFirst
            With Me.FgMainDes
                .Rows = .FixedRows + RsDevsub.RecordCount
                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, .ColIndex("LineNo")) = i
                    .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(RsDevsub("ID").value), 0, RsDevsub("ID").value)
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsDevsub("Name").value), "", RsDevsub("Name").value)
                    .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(RsDevsub("FullCode").value), "", RsDevsub("FullCode").value)
                    .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDevsub("Remarks").value), "", RsDevsub("Remarks").value)
                    .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(RsDevsub("Qty").value), 0, RsDevsub("Qty").value)
                    .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDevsub("Price").value), 0, RsDevsub("Price").value)
                    .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(RsDevsub("Total").value), 0, RsDevsub("Total").value)
                    .TextMatrix(i, .ColIndex("QtyNo")) = IIf(IsNull(RsDevsub("QtyNo").value), 0, RsDevsub("QtyNo").value)
                    .TextMatrix(i, .ColIndex("QtyExe")) = IIf(IsNull(RsDevsub("QtyExe").value), 0, RsDevsub("QtyExe").value)
                    .TextMatrix(i, .ColIndex("PriceExe")) = IIf(IsNull(RsDevsub("PriceExe").value), 0, RsDevsub("PriceExe").value)
                    .TextMatrix(i, .ColIndex("TotalExe")) = IIf(IsNull(RsDevsub("TotalExe").value), 0, RsDevsub("TotalExe").value)
                    
                    RsDevsub.MoveNext
                Next i
            End With

        End If
    '»‰Êœ «·„‘—Ê⁄
    FillBand
    '-----------------------------------------------------------------------------
 RetriveBandDetials
    'œ›⁄«  «·„‘—Ê⁄
        '»‰Êœ «·„‘—Ê⁄
    '-----------------------------------------------------------------------------
 
        StrSQL = "SELECT    * from Projectssub "
 
        StrSQL = StrSQL + " Where (projectid =" & Me.txt_project_id.Text & ")"
        Set RsDevsub = New ADODB.Recordset
        RsDevsub.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDevsub.BOF Or RsDevsub.EOF) Then
            RsDevsub.MoveFirst
    
            With Me.GridSub
                .Rows = .FixedRows + RsDevsub.RecordCount

                For i = .FixedRows To .Rows - 1
    
                    .TextMatrix(i, .ColIndex("id")) = i
    
                    .TextMatrix(i, .ColIndex("subdate")) = IIf(Not IsDate(RsDevsub("subdate").value), "", RsDevsub("subdate").value)
            
                   ' .TextMatrix(i, .ColIndex("DesTerm")) = IIf(IsNull(RsDevsub("DesTerm").value), "", RsDevsub("DesTerm").value)
            .TextMatrix(i, .ColIndex("rate")) = IIf(IsNull(RsDevsub("rate").value), "", RsDevsub("rate").value)
                    .TextMatrix(i, .ColIndex("ValueTerm")) = IIf(IsNull(RsDevsub("ValueTerm").value), "", RsDevsub("ValueTerm").value)
           
                    .TextMatrix(i, .ColIndex("SubValue")) = IIf(IsNull(RsDevsub("SubValue").value), "", RsDevsub("SubValue").value)
        
                    .TextMatrix(i, .ColIndex("REmarks")) = IIf(IsNull(RsDevsub("REmarks").value), "", RsDevsub("REmarks").value)
            
            
                    RsDevsub.MoveNext
                Next i

              
            End With

        End If
   ''///////////////
       Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 2
            Grid.Enabled = True
        Dim Rs5 As ADODB.Recordset
        StrSQL = " SELECT     dbo.ProJectMofrd.ID, dbo.ProJectMofrd.ProjID, dbo.ProJectMofrd.TypeEmp, dbo.ProJectMofrd.Valuee, dbo.ProJectMofrd.MofrdID, dbo.mofrdat.mofrad_namee, "
        StrSQL = StrSQL & "               dbo.mofrdat.mofrad_name"
        StrSQL = StrSQL & "    FROM         dbo.ProJectMofrd LEFT OUTER JOIN"
        StrSQL = StrSQL & "               dbo.mofrdat ON dbo.ProJectMofrd.MofrdID = dbo.mofrdat.mofrad_code"
        StrSQL = StrSQL & "     Where (dbo.ProJectMofrd.ProjID = " & val(txt_project_id.Text) & ")"
        Set Rs5 = New ADODB.Recordset
        Rs5.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (Rs5.BOF Or Rs5.EOF) Then
            Rs5.MoveFirst
            With Me.Grid
                .Rows = .FixedRows + Rs5.RecordCount
                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, .ColIndex("LineNo")) = i
                    .TextMatrix(i, .ColIndex("MofrdID")) = IIf(IsNull(Rs5("MofrdID").value), 0, Rs5("MofrdID").value)
                    .TextMatrix(i, .ColIndex("TypeEmp")) = IIf(IsNull(Rs5("TypeEmp").value), "", Rs5("TypeEmp").value)
                    .TextMatrix(i, .ColIndex("Valuee")) = IIf(IsNull(Rs5("Valuee").value), 0, Rs5("Valuee").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs5("mofrad_name").value), "", Rs5("mofrad_name").value)
                    Else
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs5("mofrad_namee").value), "", Rs5("mofrad_namee").value)
                    End If
                    
                    Rs5.MoveNext
                Next i
            End With
        End If
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        '„Ê«œ «·„‘—Ê⁄
 
        'StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & _
        '"ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
        'StrSQL = StrSQL + " where Project_id=" & Val(txt_project_id.text)

        'Set RsDetails = New ADODB.Recordset
        'RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        'If Not (RsDetails.EOF Or RsDetails.BOF) Then
        '    FG.Rows = RsDetails.RecordCount + 1
        '    For Num = 1 To RsDetails.RecordCount
        '        FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
        '        FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
        '        FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
        '        FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
        '        FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
        '        FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
        '        FG.TextMatrix(Num, FG.ColIndex("Valu")) = IIf(IsNull(RsDetails("Quantity")), 0, (RsDetails("Quantity").value)) * IIf(IsNull(RsDetails("Price")), 0, (RsDetails("Price").value))
        '        FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
        '        FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        '        RsDetails.MoveNext
        '    Next Num
        'End If
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
        'EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
        'StrSQL = "SELECT  * FROM  TblEmployee Where (project_id =" & Me.txt_project_id.text & ")"
        '    Set RsDev = New ADODB.Recordset
        '    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        '    If Not (RsDev.BOF Or Rs.EOF) Then
        '      RsDev.MoveFirst
        '
        '    With Me.VSFlexGrid1
        '    .Rows = .FixedRows + RsDev.RecordCount
        '    For I = .FixedRows To .Rows - 1
        '        .TextMatrix(I, .ColIndex("id")) = IIf(IsNull(RsDev("Emp_ID").value), _
        '            "", RsDev("Emp_ID").value)
        '
        '        .TextMatrix(I, .ColIndex("code")) = IIf(IsNull(RsDev("Emp_Code").value), _
        '            "", RsDev("Emp_Code").value)
        '
        '                .TextMatrix(I, .ColIndex("name")) = IIf(IsNull(RsDev("Emp_Name").value), _
        '            "", RsDev("Emp_Name").value)
        '
        '        .TextMatrix(I, .ColIndex("LineNO")) = I
        '        RsDev.MoveNext
        '    Next I
        '    End With
    
        '    End If
        'EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
   ' End If
   
    '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), _
    '  .Rows - 1, .ColIndex("CreditValue"))
    '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), _
    '  .Rows - 1, .ColIndex("DebitValue"))

    '-----------------------------------------------------------------------------
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    FillMylistData val(txt_project_id.Text)
    ReLineGrid
'    txtid.SetFocus
    
    Exit Sub
ErrTrap:

End Sub

Private Sub TxtMangSalary_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtMangSalary.Text, 0)
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap
       Option6.Enabled = False
            Option7.Enabled = False
            Option8.Enabled = False
            FgMainDes.Enabled = True
            VSFlexGrid2.Enabled = True
            GridSub.Enabled = True
            Grid.Enabled = True
            Fg_Journal.Enabled = True
    Select Case Me.TxtModFlg.Text
        Case "R"
      '  FgMainDes.Enabled = False
        VSFlexGrid2.Enabled = True
        Grid.Enabled = True
        Fg_Journal.Enabled = True
GridSub.Enabled = False
            Fra(7).Enabled = False

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„‘—Ê⁄«  "
            Else
                Me.Caption = "Projects"
            End If
         
            VSFlexGrid3.Enabled = False
            Ele(2).Enabled = False
            FG.Enabled = False
            Frame3.Enabled = False
           ' VSFlexGrid2.Editable = flexEDNone
        
          '  Fg_Journal.Editable = flexEDNone
            VSFlexGrid3.Editable = flexEDNone
        
            Me.Command1(0).Enabled = True 'ÃœÌœ
            Me.Command1(3).Enabled = True ' ⁄œÌ·
            Me.Command1(1).Enabled = False 'Õ›Ÿ
            Me.Command1(7).Enabled = True 'Õ–›
            Me.Command1(6).Enabled = False ' —«Ã⁄
            Me.Command1(5).Enabled = True '»ÕÀ
            Me.Command1(2).Enabled = True '„—›ﬁ« 
            Command3.Enabled = True ' ﬁ—Ì—
            ALLButton3.Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
 
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Command1(7).Enabled = False
                Me.Command1(3).Enabled = False
            
            End If
        
        Case "N"
              Option6.Enabled = True
            Option7.Enabled = True
            Option8.Enabled = True
        GridSub.Enabled = True
            Fra(7).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«·„‘—Ê⁄«  (ÃœÌœ)"
            Else
                Me.Caption = " Projects(New Record)"
            End If
        
            VSFlexGrid3.Enabled = True
            Ele(2).Enabled = True
            FG.Enabled = True
            Frame3.Enabled = True
          '  VSFlexGrid2.Editable = flexEDKbdMouse
        
            Fg_Journal.Editable = flexEDKbdMouse
            VSFlexGrid3.Editable = flexEDKbdMouse
        
            Me.Command1(0).Enabled = False 'ÃœÌœ
            Me.Command1(3).Enabled = False ' ⁄œÌ·
            Me.Command1(1).Enabled = True 'Õ›Ÿ
            Me.Command1(7).Enabled = False 'Õ–›
            Me.Command1(6).Enabled = True ' —«Ã⁄
            Me.Command1(5).Enabled = False '»ÕÀ
            Me.Command1(2).Enabled = True '„—›ﬁ« 
            Command3.Enabled = False ' ﬁ—Ì—
           ALLButton3.Enabled = False
        Case "E"
        GridSub.Enabled = True
            Fra(7).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«·„‘—Ê⁄« (  ⁄œÌ· )"
            Else
                Me.Caption = "Projects (Edit Current Record)"
            End If

            VSFlexGrid3.Enabled = True
            Ele(2).Enabled = True
            FG.Enabled = True
            Frame3.Enabled = True
         '   VSFlexGrid2.Editable = flexEDKbdMouse
        
            Fg_Journal.Editable = flexEDKbdMouse
            VSFlexGrid3.Editable = flexEDKbdMouse
        
            Me.Command1(0).Enabled = False 'ÃœÌœ
            Me.Command1(3).Enabled = False ' ⁄œÌ·
            Me.Command1(1).Enabled = True 'Õ›Ÿ
            Me.Command1(7).Enabled = False 'Õ–›
            Me.Command1(6).Enabled = True ' —«Ã⁄
            Me.Command1(5).Enabled = False '»ÕÀ
            Me.Command1(2).Enabled = True '„—›ﬁ« 
            Command3.Enabled = False ' ﬁ—Ì—
             ALLButton3.Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
  
    End Select

    Exit Sub
ErrTrap:

End Sub
 
Private Sub TXTOrDer_no_Change()
    Dim Transaction_Type As Integer
    If val(CBoBasedON.ListIndex) = 1 Then
        Transaction_Type = 6
    Else
      Transaction_Type = 0
         Exit Sub
    End If

    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder Me.TXTOrDer_no, Transaction_Type
    End If
End Sub
Public Sub RetriveOrder(Optional order_no As String = "", _
                        Optional Transaction_Type As Integer = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 2
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Refresh

     If Transaction_Type = 6 Then
        StrSQL = "Select * from transactions where  Transaction_Type=" & Transaction_Type & " and order_no='" & order_no & "'"
    End If


    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Sub
    Else
        DcAccount2.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    '    Me.DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").value), 1, rs("Currency_id").value)
    '    Me.DCboStoreName.BoundText = IIf(IsNull(rs("storeid").value), "", rs("storeid").value)
    '    Me.Dcbranch.BoundText = IIf(IsNull(rs("Branchid").value), "", rs("Branchid").value)

    'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
 
    
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

StrSQL = StrSQL + "order by dbo.Transaction_Details.id"


    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""
With Fg_Journal
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Fg_Journal.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            .TextMatrix(Num, .ColIndex("fullcode")) = IIf(IsNull(RsDetails("Fullcode")), "", (RsDetails("Fullcode").value))
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(Num, .ColIndex("des")) = IIf(IsNull(RsDetails("ItemName")), "", Trim(RsDetails("ItemName").value))
            Else
            .TextMatrix(Num, .ColIndex("des")) = IIf(IsNull(RsDetails("ItemNamee")), "", Trim(RsDetails("ItemNamee").value))
            End If
            .TextMatrix(Num, .ColIndex("qty")) = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value))
            .TextMatrix(Num, .ColIndex("PandUnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            .TextMatrix(Num, .ColIndex("PandUnit")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            .TextMatrix(Num, .ColIndex("cost")) = IIf(IsNull(RsDetails("ShowPrice")), 0, (RsDetails("ShowPrice").value))
            .TextMatrix(Num, .ColIndex("total")) = val(.TextMatrix(Num, .ColIndex("cost"))) * val(.TextMatrix(Num, .ColIndex("qty"))) * IIf(val(.TextMatrix(Num, .ColIndex("qty"))) = 0, 1, val(.TextMatrix(Num, .ColIndex("qty"))))
            
            Fg_Journal_AfterEdit Num, .ColIndex("des")
            RsDetails.MoveNext

        Next Num

    End If
End With
    Screen.MousePointer = vbDefault

    Exit Sub
  End If
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub TXTOrDer_no_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
If KeyCode = vbKeyF3 Then
 If val(CBoBasedON.ListIndex) = 1 Then
Dim transactionName As String
Dim transactiontype As Integer
   transactiontype = 6
                      If SystemOptions.UserInterface = ArabicInterface Then
                          transactionName = "»ÕÀ ⁄‰ «Ê«„— «·»Ì⁄"
                        Else
                        transactionName = "Search  Sales Order"
                        End If
                 Order_no_search.show
        Order_no_search.RetrunType = 19
       Order_no_search.Label1(2).Caption = transactionName
       Order_no_search.lblSpecificsearch = transactiontype
End If
End If
End If
End Sub

Private Sub TxtProjectCosts_Change()
calbetprice
End Sub







Private Sub TXTprojectname_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub TXTprojectnamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid1

        Select Case .ColKey(Col)
 
            Case "name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
             
                StrSQL = "SELECT  * from TblEmployee Where Emp_id=" & val(StrAccountCode)
                Set rs = Nothing
            
                If StrAccountCode <> "" Then
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        .TextMatrix(Row, .ColIndex("code")) = IIf(IsNull(rs("Emp_Code").value), "", rs("Emp_Code").value)
                    End If
                End If
            
                '.TextMatrix(Row, .ColIndex("id")) = get_Expenses_id(StrAccountCode)
        
            Case "code"
                  
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT  * from TblEmployee Where Emp_Code=" & .TextMatrix(Row, Col)
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
          
                    .TextMatrix(Row, .ColIndex("id")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                    
                    .TextMatrix(Row, .ColIndex("name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                Else
                    .TextMatrix(Row, .ColIndex("id")) = ""
                    .TextMatrix(Row, .ColIndex("name")) = ""
                End If

        End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
        txt_employee_count = .Rows - 2
        Me.txt_emp_salary.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
   
    End With

    ReLineGrid
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid1

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
        Select Case .ColKey(Col)
            
            Case "name"
                Exit Sub
        End Select

    End With

    VSFlexGrid1.ComboList = ""
End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid1

        Select Case .ColKey(Col)

            Case "name"
                StrSQL = "select * from TblEmployee"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "Emp_Name", "Emp_ID")
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Public Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)

    With VSFlexGrid2
  
        On Error Resume Next
        Dim StrAccountCode As String
        Dim Msg As String
        Dim rs As New ADODB.Recordset
        Dim StrSQL As String
        Dim ClsAcc As New ClsAccounts
        Dim LngRow As Long
        Dim code  As Double

        With VSFlexGrid2

            Select Case .ColKey(Col)

                Case "name"
                    code = .ComboData
                    .TextMatrix(Row, .ColIndex("OPRIDD")) = code
                    .TextMatrix(Row, .ColIndex("name")) = .ComboItem
                    .TextMatrix(Row, .ColIndex("qty")) = 1
                    REFillOprData code, Row

                Case "qty"
                    code = val(.TextMatrix(Row, .ColIndex("OPRIDD")))
                    REFillOprData code, Row
            End Select
      
            If Row = .Rows - 1 Then
                .Rows = .Rows + 1
            End If
  
        End With

        ReLineGrid
 
    End With

    If Me.TxtModFlg <> "E" Then Exit Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    'Grid.TextMatrix(Row, Grid.ColIndex("Code"))
    'Grid.TextMatrix(Row, Grid.ColIndex("Name"))
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////

End Sub

Private Sub VSFlexGrid2_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid2
Select Case .ColKey(Col)
Case "LineNo"
Cancel = True
Case "qty"
.ComboList = ""
Case "periodView"
.ComboList = ""
Case "Actperiod"
.ComboList = ""
Case "Pre"
.ComboList = ""
Case "Period1"
.ComboList = ""
Case "StartDate"
.ComboList = ""
Case "EndDate"
.ComboList = ""
Case "EarlyStartWeek"
.ComboList = ""
Case "StartWeek"
.ComboList = ""
Case "EarlyEndWeek"
.ComboList = ""
Case "EndWeek"
.ComboList = ""
Case "Critical"
.ComboList = ""
Case "mat"
.ComboList = ""
Case "equep"
.ComboList = ""
Case "employee"
.ComboList = ""
Case "expensive"
.ComboList = ""
Case "total_items"
Cancel = True
Case "total_salary"
Cancel = True
Case "EquepVal"
Cancel = True
Case "total_expenses"
Cancel = True
Case "total"
Cancel = True
End Select
    End With

End Sub

Private Sub VSFlexGrid2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    With Me.VSFlexGrid2

        Select Case .ColKey(Col)
                Case "StartDate"
                                  LngRow = Row

 LngCol = Col
Load FrmDateOpProject
FrmDateOpProject.Index = 0
FrmDateOpProject.show vbModal

        Case "EndDate"
                                  LngRow = Row

 LngCol = Col
Load FrmDateOpProject
FrmDateOpProject.Index = 0
FrmDateOpProject.show vbModal

        

        Case "expensive"
                                  LngRow = Row

 LngCol = Col
Load FrmExchangeOper
FrmExchangeOper.show vbModal

        Case "equep"
                                  LngRow = Row

 LngCol = Col
 Load FrmEquepment
  FrmEquepment.show vbModal

        Case "employee"
                                  LngRow = Row

 LngCol = Col
Load FrmEmpOper
FrmEmpOper.show vbModal
                 Case "mat"
                                           LngRow = Row

 LngCol = Col

             ' ItemProductionDate Row, Col, , 1
             Monthly = DateDiff("m", DTStartDate.value, DTEnddate.value)
                Load FrmMatrialsOp
                FrmMatrialsOp.show vbModal

                    
                End Select
                End With
End Sub

Private Sub VSFlexGrid2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
With Me.VSFlexGrid2
Select Case .ColKey(Col)
Case "qty"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("qty")), 0)
Case "periodView"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("periodView")), 0)
Case "Actperiod"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("Actperiod")), 0)
Case "Period1"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("Period1")), 0)
Case "Period1"
KeyAscii = KeyAscii_Num(KeyAscii, .TextMatrix(Row, .ColIndex("Period1")), 0)
End Select
End With
End Sub

Private Sub VSFlexGrid2_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim LngItemID As Integer
    Dim MyStrList As String

    With Me.VSFlexGrid2

        Select Case .ColKey(Col)
                               Case "StartDate"
.ColComboList(.ColIndex("StartDate")) = "..."
                       Case "EndDate"
.ColComboList(.ColIndex("EndDate")) = "..."

                       Case "expensive"
.ColComboList(.ColIndex("expensive")) = "..."
        
               Case "equep"
.ColComboList(.ColIndex("equep")) = "..."

        Case "employee"
.ColComboList(.ColIndex("employee")) = "..."

Case "mat"
.ColComboList(.ColIndex("mat")) = "..."

            Case "name"
            
If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = " SELECT     ProcessName, TblProcessDEFID"
Else
                          StrSQL = " SELECT     ProcessNamee, TblProcessDEFID"
      
End If

                StrSQL = StrSQL + " from dbo.TblProcessDEF"
                
                StrSQL = StrSQL + " ORDER BY TblProcessDEFID"
                 
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
               If SystemOptions.UserInterface = ArabicInterface Then

                    MyStrList = .BuildComboList(rs, "ProcessName", "TblProcessDEFID")
                Else
                MyStrList = .BuildComboList(rs, "ProcessNameE", "TblProcessDEFID")
                End If
                    '                    Grid.ColComboList = MyStrList
                    VSFlexGrid2.ColComboList(.ColIndex("name")) = "|" & MyStrList
                Else
                    Cancel = True
                End If
            
        End Select

    End With

End Sub

Function get_opr_details(Strbasedon As String, Period As Double, period1 As Double, Optional ByRef StartWeek As Double, Optional ByRef lastweek As Double, Optional ByRef EarlyStartWeek As Double, Optional ByRef Earlylastweek As Double)
    On Error Resume Next
    Dim astrSplitItems() As String
    Dim lastend As Double

    If Strbasedon = "" Then
        StartWeek = 0
        lastweek = Period
        EarlyStartWeek = 0
        Earlylastweek = Period
    Else

        astrSplitItems = Split(Strbasedon, ",")
        lastend = 0

        For i = 0 To 20

            If lastend < getlastend(astrSplitItems(i)) Then
                lastend = getlastend(astrSplitItems(i))
            End If

        Next i
  
        StartWeek = lastend + period1
        EarlyStartWeek = lastend

        lastweek = StartWeek + Period

        Earlylastweek = lastweek - period1

    End If

End Function

Function getlastend(str As String) As Double
    Dim i As Integer

    With Me.VSFlexGrid2

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("Symbol")) = str Then
                getlastend = val(.TextMatrix(i, .ColIndex("EndWeek")))

            End If

        Next i

    End With

End Function

Private Sub VSFlexGrid2_Click()

    If Not VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode")) = "" Then
      
        current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode"))
    End If

End Sub
Private Sub VSFlexGrid3_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)

    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid3

        Select Case .ColKey(Col)
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
  
            Case "value"
                Dim sgl As String
  
        End Select

        Me.txt_expenses_total.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
   
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub


Private Sub VSFlexGrid3_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid3

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel h= True
            '  End If
        End If

        Select Case .ColKey(Col)

            Case "value"
                .ComboList = ""

            Case "des"
                .ComboList = ""
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub VSFlexGrid3_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid3

        Select Case .ColKey(Col)

            Case "AccountName"
                StrSQL = "select * from Expenses_accounts"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub VSFlexGrid4_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid4
  Select Case .ColKey(Col)
Case "FixedAsset"
Dim rs As ADODB.Recordset
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Namee  "
                Else
                                        StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Name  "
                    
                End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Name", "id")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Namee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
.Rows = .Rows + 1
End Select
End With

End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error Resume Next
    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
    End If
terms_operations_Click 1
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
Frame5.Visible = True
    Retrive
    Exit Sub
ErrTrap:
End Sub



