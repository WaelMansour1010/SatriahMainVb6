VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form End_oF_service 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ”«» „ﬂ«›√… ‰Â«Ì… «·Œœ„…"
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14730
   Icon            =   "End_oF_service.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11910
   ScaleWidth      =   14730
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
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   14880
      TabIndex        =   51
      Top             =   6240
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   315
         Left            =   120
         TabIndex        =   52
         Top             =   1680
         Width           =   1215
      End
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   255
      Left            =   14880
      TabIndex        =   50
      Top             =   1800
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«·”»»"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "End_oF_service.frx":6852
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   " ÕœÌœ «·ﬂ·"
      Height          =   195
      Left            =   17040
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmd_CALC_NET 
      Caption         =   "Õ”«» «·’«›Ì"
      Height          =   315
      Left            =   15360
      TabIndex        =   48
      Top             =   5280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      Height          =   1815
      Left            =   14880
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox txt_salry_total 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   45
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox mang1 
         Height          =   195
         Left            =   3000
         TabIndex        =   43
         Top             =   1080
         Width           =   135
      End
      Begin VB.TextBox Txtsalary 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   42
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtsaknm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   34
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtbusm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   33
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtfoodm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtanotherm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   31
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2040
         TabIndex        =   30
         Top             =   4080
         Width           =   615
      End
      Begin VB.Frame Frame5 
         Caption         =   "ÿ—Ìﬁ… «·Õ”«»"
         Height          =   3255
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   4800
         Visible         =   0   'False
         Width           =   4935
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "End_oF_service.frx":686E
            Left            =   2160
            List            =   "End_oF_service.frx":6875
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            Caption         =   "‘Â—Ì"
            Height          =   195
            Index           =   0
            Left            =   2400
            TabIndex        =   24
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "”‰ÊÌ"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   1200
            Width           =   4695
         End
         Begin VB.Frame Frame6 
            Height          =   495
            Left            =   240
            TabIndex        =   15
            Top             =   600
            Width           =   1935
            Begin VB.CommandButton Command1 
               Caption         =   "="
               Height          =   315
               Index           =   4
               Left            =   120
               TabIndex        =   20
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "+"
               Height          =   315
               Index           =   3
               Left            =   1560
               TabIndex        =   19
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "-"
               Height          =   315
               Index           =   2
               Left            =   1200
               TabIndex        =   18
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "*"
               Height          =   315
               Index           =   1
               Left            =   840
               TabIndex        =   17
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "/"
               Height          =   315
               Index           =   5
               Left            =   480
               TabIndex        =   16
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               Caption         =   "«Õ„«·Ì «·„ﬁ«„"
               ForeColor       =   &H000000FF&
               Height          =   15
               Left            =   0
               TabIndex        =   21
               Top             =   2520
               Width           =   1935
            End
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   1560
            Width           =   4695
         End
         Begin VB.CommandButton Command6 
            Caption         =   "„Ê«›ﬁ"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   285
            Index           =   41
            Left            =   3720
            TabIndex        =   29
            Top             =   240
            Width           =   915
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·Õ”«»"
            Height          =   285
            Index           =   42
            Left            =   3480
            TabIndex        =   28
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰ ÌÃ…"
            Height          =   285
            Index           =   43
            Left            =   3600
            TabIndex        =   27
            Top             =   2040
            Width           =   1155
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   285
            Index           =   44
            Left            =   1680
            TabIndex        =   26
            Top             =   2040
            Width           =   1155
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   4440
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   4800
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   5280
         Width           =   615
      End
      Begin VB.TextBox TXTMOBM 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TXTMANGM 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox sal1 
         Height          =   195
         Left            =   6240
         TabIndex        =   6
         Top             =   240
         Width           =   135
      End
      Begin VB.CheckBox sakn1 
         Height          =   195
         Left            =   6240
         TabIndex        =   5
         Top             =   600
         Width           =   135
      End
      Begin VB.CheckBox bus1 
         Height          =   195
         Left            =   6240
         TabIndex        =   4
         Top             =   1080
         Width           =   135
      End
      Begin VB.CheckBox another1 
         Height          =   195
         Left            =   6240
         TabIndex        =   3
         Top             =   1440
         Width           =   135
      End
      Begin VB.CheckBox food1 
         Height          =   195
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   135
      End
      Begin VB.CheckBox mob1 
         Height          =   195
         Left            =   3000
         TabIndex        =   1
         Top             =   600
         Width           =   135
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«Ã„«·Ì"
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   46
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·—« » «·«”«”Ì"
         Height          =   285
         Index           =   0
         Left            =   4920
         TabIndex        =   41
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "»œ· «·”ﬂ‰"
         Height          =   285
         Index           =   34
         Left            =   5040
         TabIndex        =   40
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "»œ· „Ê«’·« "
         Height          =   285
         Index           =   35
         Left            =   5160
         TabIndex        =   39
         Top             =   990
         Width           =   915
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "»œ· ÿ⁄«„"
         Height          =   285
         Index           =   37
         Left            =   1920
         TabIndex        =   38
         Top             =   240
         Width           =   915
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "»œ·«  «Œ—Ï"
         Height          =   285
         Index           =   38
         Left            =   5160
         TabIndex        =   37
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "»œ· ÃÊ«·"
         Height          =   285
         Index           =   36
         Left            =   1920
         TabIndex        =   36
         Top             =   600
         Width           =   915
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "»œ· ≈‘—«›"
         Height          =   285
         Index           =   45
         Left            =   1920
         TabIndex        =   35
         Top             =   1080
         Width           =   915
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Height          =   1545
      Left            =   120
      TabIndex        =   53
      Top             =   13680
      Visible         =   0   'False
      Width           =   8235
      _cx             =   14526
      _cy             =   2725
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"End_oF_service.frx":687C
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
   Begin C1SizerLibCtl.C1Elastic C1Elastic3 
      Height          =   11910
      Left            =   0
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   0
      Width           =   14730
      _cx             =   25982
      _cy             =   21008
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic100 
         Height          =   1575
         Left            =   0
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   10215
         Width           =   14610
         _cx             =   25770
         _cy             =   2778
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   360
            Index           =   0
            Left            =   12870
            TabIndex        =   57
            Top             =   870
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   635
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
            ColorToggledText=   -2147483631
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   360
            Index           =   1
            Left            =   11190
            TabIndex        =   58
            Top             =   870
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   635
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
            Height          =   360
            Index           =   2
            Left            =   9300
            TabIndex        =   59
            Top             =   870
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   635
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
            Height          =   360
            Index           =   3
            Left            =   7050
            TabIndex        =   60
            Top             =   870
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   635
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
            Height          =   360
            Index           =   4
            Left            =   5985
            TabIndex        =   61
            Top             =   870
            Width           =   1035
            _ExtentX        =   1826
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
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   360
            Index           =   5
            Left            =   5190
            TabIndex        =   62
            Top             =   870
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   635
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
            Height          =   360
            Index           =   6
            Left            =   45
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   870
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   635
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   360
            Index           =   7
            Left            =   3915
            TabIndex        =   64
            Top             =   870
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   635
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄…"
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
            Height          =   360
            Left            =   1035
            TabIndex        =   65
            Top             =   870
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   635
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   735
            Index           =   2
            Left            =   8895
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   0
            Width           =   11325
            _cx             =   19976
            _cy             =   1296
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
            Begin VB.CommandButton Command11 
               Caption         =   "≈⁄«œ… ≈‰‘«¡"
               Height          =   195
               Left            =   4590
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   0
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.CommandButton Command7 
               Caption         =   "Õ–› ﬁÌœ «·«” Õﬁ«ﬁ"
               Height          =   375
               Left            =   7275
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   240
               Width           =   1860
            End
            Begin VB.CommandButton Command5 
               Caption         =   "≈‰‘«¡ ﬁÌœ «·«” Õﬁ«ﬁ"
               Height          =   375
               Left            =   9135
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   210
               Width           =   1740
            End
            Begin VB.CommandButton Command8 
               Caption         =   "ﬂ‘› Õ”«»"
               Height          =   375
               Left            =   225
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   240
               Width           =   1500
            End
            Begin VB.CommandButton Command9 
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   375
               Left            =   2190
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   240
               Width           =   1110
            End
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   270
               Width           =   2835
            End
            Begin VB.TextBox TxtNoteID 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   -90
               Visible         =   0   'False
               Width           =   690
            End
            Begin XtremeSuiteControls.CheckBox chkGE 
               Height          =   255
               Left            =   9135
               TabIndex        =   74
               Top             =   0
               Width           =   1980
               _Version        =   786432
               _ExtentX        =   3492
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«·ﬁÌœ ⁄·Ì  «—ÌŒ «·Õ—ﬂ…"
               ForeColor       =   8388608
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   180
               Index           =   35
               Left            =   5940
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   330
               Width           =   930
            End
         End
         Begin ImpulseButton.ISButton Accredit 
            Height          =   360
            Left            =   1905
            TabIndex        =   76
            Top             =   870
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   635
            ButtonPositionImage=   1
            Caption         =   "«—”«· ··«⁄ „«œ"
            BackColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   -2147483635
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            Height          =   225
            Left            =   690
            TabIndex        =   80
            Top             =   240
            Width           =   1530
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            Height          =   225
            Left            =   4710
            TabIndex        =   79
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   6
            Left            =   6570
            TabIndex        =   78
            Top             =   240
            Width           =   2070
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   7
            Left            =   2535
            TabIndex        =   77
            Top             =   240
            Width           =   1800
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   10125
         Left            =   0
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   0
         Width           =   14730
         _cx             =   25982
         _cy             =   17859
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   645
            Index           =   0
            Left            =   0
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   0
            Width           =   14670
            _cx             =   25876
            _cy             =   1138
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   21.75
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
            Caption         =   "    Õ”«» „ﬂ«›√… ‰Â«Ì… «·Œœ„…  "
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
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Text            =   "TxtModFlg"
               Top             =   120
               Visible         =   0   'False
               Width           =   1095
            End
            Begin ImpulseButton.ISButton XPBtnMove 
               Height          =   375
               Index           =   0
               Left            =   1215
               TabIndex        =   84
               Top             =   90
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
               ButtonImage     =   "End_oF_service.frx":69AB
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
               Left            =   150
               TabIndex        =   85
               Top             =   90
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
               ButtonImage     =   "End_oF_service.frx":6D45
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
               Left            =   1740
               TabIndex        =   86
               Top             =   90
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
               ButtonImage     =   "End_oF_service.frx":70DF
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
               Left            =   675
               TabIndex        =   87
               Top             =   90
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
               ButtonImage     =   "End_oF_service.frx":7479
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
            Begin Dynamic_Byte.NourHijriCal Txt_DateEndLincH 
               Height          =   255
               Left            =   0
               TabIndex        =   88
               Top             =   0
               Visible         =   0   'False
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
            End
            Begin VB.Image ImgFavorites 
               Height          =   390
               Left            =   2280
               Picture         =   "End_oF_service.frx":7813
               Stretch         =   -1  'True
               Top             =   0
               Width           =   525
            End
         End
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   9585
            Left            =   0
            TabIndex        =   89
            Top             =   480
            Width           =   14715
            _cx             =   25956
            _cy             =   16907
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
            Caption         =   "«·»Ì«‰«  «·√”«”Ì…|Õ«·… «·«⁄ „«œ"
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   9165
               Index           =   1
               Left            =   45
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   45
               Width           =   14625
               _cx             =   25797
               _cy             =   16166
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
               Begin VB.TextBox TxtReqNo 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   240
                  MaxLength       =   10
                  TabIndex        =   136
                  Top             =   705
                  Width           =   1815
               End
               Begin VB.TextBox TxtTotalDis 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   465
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   135
                  Top             =   7695
                  Width           =   3045
               End
               Begin VB.TextBox TxtAddOther 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   465
                  MaxLength       =   10
                  TabIndex        =   134
                  Top             =   5580
                  Width           =   1950
               End
               Begin VB.TextBox TxtJbsatust 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Height          =   285
                  Left            =   4605
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   133
                  Top             =   2010
                  Visible         =   0   'False
                  Width           =   615
               End
               Begin VB.CommandButton Comman1122 
                  Caption         =   "⁄—÷ «·„”Ì— «·Õ«·Ì"
                  Height          =   300
                  Left            =   6375
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
                  Top             =   5700
                  Visible         =   0   'False
                  Width           =   1680
               End
               Begin VB.TextBox txtDayval 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   330
                  Left            =   0
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   131
                  Top             =   6270
                  Visible         =   0   'False
                  Width           =   480
               End
               Begin VB.TextBox txtSal 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   465
                  MaxLength       =   10
                  TabIndex        =   130
                  Top             =   4155
                  Width           =   1830
               End
               Begin VB.TextBox txtCount 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   2625
                  MaxLength       =   20
                  TabIndex        =   129
                  Top             =   4155
                  Width           =   645
               End
               Begin VB.TextBox txtTicketValue 
                  Alignment       =   2  'Center
                  Height          =   300
                  Left            =   2910
                  MaxLength       =   10
                  TabIndex        =   128
                  Top             =   4515
                  Width           =   960
               End
               Begin VB.TextBox txtCustom 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   465
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   127
                  Top             =   5220
                  Width           =   3045
               End
               Begin VB.TextBox Text6 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FF80FF&
                  Height          =   285
                  Left            =   7440
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   3555
                  Width           =   600
               End
               Begin VB.TextBox Text5 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FF80FF&
                  Height          =   285
                  Left            =   6885
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   3555
                  Width           =   570
               End
               Begin VB.TextBox Text4 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0FF&
                  Height          =   285
                  Left            =   9480
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   3555
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.TextBox Text3 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0FF&
                  Height          =   285
                  Left            =   8895
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   3555
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.TextBox txtreasons 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   240
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   122
                  Top             =   2130
                  Width           =   4500
               End
               Begin VB.TextBox TxtVlueVaction 
                  Alignment       =   2  'Center
                  Height          =   300
                  Left            =   465
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   121
                  Top             =   6645
                  Width           =   1005
               End
               Begin VB.TextBox TxtVSa 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0FF&
                  Height          =   285
                  Left            =   8895
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
                  Top             =   3555
                  Width           =   1830
               End
               Begin VB.TextBox TXTLastTotal 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FFFF&
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
                  ForeColor       =   &H00C00000&
                  Height          =   360
                  Left            =   990
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   119
                  Top             =   8655
                  Width           =   1305
               End
               Begin VB.TextBox TxtCash 
                  Alignment       =   2  'Center
                  Height          =   300
                  Left            =   2505
                  MaxLength       =   10
                  TabIndex        =   118
                  Top             =   6990
                  Width           =   1005
               End
               Begin VB.TextBox TXTAdvanceTotal 
                  Alignment       =   2  'Center
                  Height          =   300
                  Left            =   2505
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   117
                  Top             =   6645
                  Width           =   1005
               End
               Begin VB.TextBox text1 
                  Alignment       =   1  'Right Justify
                  Height          =   1425
                  Left            =   6135
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   116
                  Top             =   4260
                  Width           =   8205
               End
               Begin VB.TextBox txtnet 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   465
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   115
                  Top             =   5925
                  Width           =   1950
               End
               Begin VB.TextBox txtnum 
                  Alignment       =   2  'Center
                  Height          =   300
                  Left            =   990
                  MaxLength       =   10
                  TabIndex        =   114
                  Top             =   8310
                  Width           =   1305
               End
               Begin VB.Frame Frame1 
                  Height          =   480
                  Left            =   3495
                  TabIndex        =   107
                  Top             =   8190
                  Width           =   1950
                  Begin VB.CommandButton Command1 
                     Caption         =   "/"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   9
                     Left            =   120
                     TabIndex        =   112
                     Top             =   120
                     Width           =   375
                  End
                  Begin VB.CommandButton Command1 
                     Caption         =   "*"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   8
                     Left            =   480
                     TabIndex        =   111
                     Top             =   120
                     Width           =   375
                  End
                  Begin VB.CommandButton Command1 
                     Caption         =   "-"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   7
                     Left            =   840
                     TabIndex        =   110
                     Top             =   120
                     Width           =   375
                  End
                  Begin VB.CommandButton Command1 
                     Caption         =   "+"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   6
                     Left            =   1200
                     TabIndex        =   109
                     Top             =   120
                     Width           =   375
                  End
                  Begin VB.CommandButton Command1 
                     Caption         =   "="
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   0
                     Left            =   1440
                     TabIndex        =   108
                     Top             =   120
                     Width           =   495
                  End
                  Begin VB.Label Label5 
                     Alignment       =   2  'Center
                     Caption         =   "«Õ„«·Ì «·„ﬁ«„"
                     ForeColor       =   &H000000FF&
                     Height          =   15
                     Left            =   0
                     TabIndex        =   113
                     Top             =   2520
                     Width           =   1935
                  End
               End
               Begin VB.TextBox TXTid 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   11940
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   705
                  Width           =   1545
               End
               Begin VB.TextBox txtyear 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Height          =   285
                  Left            =   11190
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   3555
                  Width           =   600
               End
               Begin VB.TextBox txtmonth 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Height          =   285
                  Left            =   11775
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   3555
                  Width           =   645
               End
               Begin VB.TextBox txtday 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Height          =   285
                  Left            =   12405
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   3555
                  Width           =   600
               End
               Begin VB.TextBox txtEmpCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   11940
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   1065
                  Width           =   1545
               End
               Begin VB.TextBox Text2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FF80FF&
                  Height          =   285
                  Left            =   8025
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   3555
                  Width           =   660
               End
               Begin VB.TextBox txttotal 
                  Alignment       =   2  'Center
                  Height          =   300
                  Left            =   3750
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   100
                  Top             =   3210
                  Width           =   2055
               End
               Begin VB.TextBox TxtRate 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   300
                  Left            =   2505
                  MaxLength       =   20
                  TabIndex        =   99
                  Top             =   3210
                  Width           =   765
               End
               Begin VB.TextBox TxtNetEnd 
                  Alignment       =   2  'Center
                  Height          =   300
                  Left            =   345
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   98
                  Top             =   3210
                  Width           =   2070
               End
               Begin VB.TextBox TxtEndService 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   2505
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   97
                  Top             =   3555
                  Width           =   1260
               End
               Begin VB.TextBox TxtCusTiket 
                  Alignment       =   2  'Center
                  Height          =   270
                  Left            =   2910
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   96
                  Top             =   4860
                  Width           =   960
               End
               Begin VB.TextBox TxtDiffTekit 
                  Alignment       =   2  'Center
                  Height          =   270
                  Left            =   465
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   95
                  Top             =   4860
                  Width           =   1005
               End
               Begin VB.TextBox TxtDiffEnd 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   345
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   94
                  Top             =   3555
                  Width           =   885
               End
               Begin VB.TextBox TxtDiscounts 
                  Alignment       =   2  'Center
                  Height          =   300
                  Left            =   465
                  MaxLength       =   10
                  TabIndex        =   93
                  Top             =   6990
                  Width           =   1005
               End
               Begin VB.TextBox TxtTicktConract 
                  Alignment       =   2  'Center
                  Height          =   300
                  Left            =   465
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   92
                  Top             =   4515
                  Width           =   1005
               End
               Begin VB.TextBox TxtDisSalary 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   465
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   91
                  Top             =   7335
                  Width           =   3045
               End
               Begin XtremeSuiteControls.CheckBox AlowAssest 
                  Height          =   255
                  Left            =   3030
                  TabIndex        =   137
                  Top             =   8070
                  Visible         =   0   'False
                  Width           =   2655
                  _Version        =   786432
                  _ExtentX        =   4683
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "Ì „ Œ’„ «·⁄Âœ „‰ ≈Ã„«·Ì «·„” Õﬁ"
                  ForeColor       =   8388608
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboEmp 
                  Height          =   315
                  Left            =   6135
                  TabIndex        =   138
                  Top             =   1065
                  Width           =   5655
                  _ExtentX        =   9975
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "7"
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker txtdate 
                  Height          =   285
                  Left            =   9375
                  TabIndex        =   139
                  Top             =   705
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  _Version        =   393216
                  Format          =   102039553
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo dctype 
                  Height          =   315
                  Left            =   6135
                  TabIndex        =   140
                  Top             =   705
                  Width           =   1800
                  _ExtentX        =   3175
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "7"
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo dcby 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   141
                  Top             =   1425
                  Width           =   4500
                  _ExtentX        =   7938
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  ListField       =   "7"
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VSFlex8UCtl.VSFlexGrid Fg 
                  Height          =   1410
                  Left            =   6135
                  TabIndex        =   142
                  Top             =   1545
                  Width           =   8385
                  _cx             =   14790
                  _cy             =   2487
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
                  AllowUserResizing=   1
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"End_oF_service.frx":B47B
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
               Begin MSDataListLib.DataCombo DataCombo1 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   143
                  Top             =   1785
                  Width           =   4500
                  _ExtentX        =   7938
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  ListField       =   "7"
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                  Height          =   1335
                  Left            =   6135
                  TabIndex        =   144
                  TabStop         =   0   'False
                  Top             =   7695
                  Width           =   8445
                  _cx             =   14896
                  _cy             =   2355
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
                  Begin VSFlex8Ctl.VSFlexGrid Grid 
                     Height          =   675
                     Left            =   225
                     TabIndex        =   145
                     Top             =   315
                     Width           =   15795
                     _cx             =   27861
                     _cy             =   1191
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
                     Cols            =   65
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   0
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"End_oF_service.frx":B5AA
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
                  Begin C1SizerLibCtl.C1Elastic Ele 
                     Height          =   1050
                     Index           =   3
                     Left            =   17130
                     TabIndex        =   146
                     TabStop         =   0   'False
                     Top             =   1560
                     Width           =   3945
                     _cx             =   6959
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
                     BackColor       =   14871017
                     ForeColor       =   -2147483630
                     FloodColor      =   6553600
                     ForeColorDisabled=   -2147483631
                     Caption         =   "≈Œ Ì«— «· «—ÌŒ"
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
                     Begin VB.ComboBox CboYear 
                        Height          =   315
                        Index           =   0
                        Left            =   90
                        RightToLeft     =   -1  'True
                        Style           =   2  'Dropdown List
                        TabIndex        =   148
                        Top             =   240
                        Width           =   1755
                     End
                     Begin VB.ComboBox CmbMonth 
                        Height          =   315
                        Index           =   0
                        Left            =   90
                        RightToLeft     =   -1  'True
                        Style           =   2  'Dropdown List
                        TabIndex        =   147
                        Top             =   540
                        Width           =   1755
                     End
                     Begin ImpulseButton.ISButton CmdOk 
                        Height          =   315
                        Left            =   90
                        TabIndex        =   149
                        Top             =   855
                        Width           =   1755
                        _ExtentX        =   3096
                        _ExtentY        =   556
                        ButtonStyle     =   1
                        ButtonPositionImage=   1
                        Caption         =   "⁄—÷  "
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
                        ButtonImage     =   "End_oF_service.frx":BDA1
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "”‰…"
                        Height          =   15
                        Index           =   52
                        Left            =   90
                        RightToLeft     =   -1  'True
                        TabIndex        =   151
                        Top             =   1815
                        Width           =   1755
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "‘Â—"
                        Height          =   15
                        Index           =   53
                        Left            =   90
                        RightToLeft     =   -1  'True
                        TabIndex        =   150
                        Top             =   1860
                        Width           =   1755
                     End
                  End
                  Begin MSDataListLib.DataCombo Dcemp 
                     Height          =   315
                     Left            =   1665
                     TabIndex        =   152
                     Top             =   1560
                     Width           =   6240
                     _ExtentX        =   11007
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
                     Caption         =   "„ÊŸ› „Õœœ"
                     DataField       =   "Õœœ"
                     Height          =   315
                     Index           =   54
                     Left            =   7620
                     RightToLeft     =   -1  'True
                     TabIndex        =   155
                     Top             =   1575
                     Width           =   1995
                  End
                  Begin VB.Label Label6 
                     Alignment       =   1  'Right Justify
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
                     Height          =   225
                     Index           =   1
                     Left            =   25260
                     RightToLeft     =   -1  'True
                     TabIndex        =   154
                     Top             =   0
                     Width           =   1860
                  End
                  Begin VB.Label XPLbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "—« » «·‘Â— «·Õ«·”"
                     ForeColor       =   &H00FF0000&
                     Height          =   255
                     Index           =   39
                     Left            =   13440
                     TabIndex        =   153
                     Top             =   0
                     Width           =   2505
                  End
               End
               Begin MSDataListLib.DataCombo Dcbranch 
                  Bindings        =   "End_oF_service.frx":C13B
                  Height          =   315
                  Left            =   240
                  TabIndex        =   156
                  Top             =   1065
                  Width           =   4500
                  _ExtentX        =   7938
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
               Begin ImpulseButton.ISButton ISButton2 
                  Height          =   435
                  Left            =   3495
                  TabIndex        =   157
                  ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                  Top             =   8655
                  Width           =   1950
                  _ExtentX        =   3440
                  _ExtentY        =   767
                  Caption         =   "≈Õ ”«» «· ’›Ì…"
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
                  ButtonImage     =   "End_oF_service.frx":C150
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  LowerToggledContent=   0   'False
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   1665
                  Index           =   18
                  Left            =   6255
                  TabIndex        =   158
                  TabStop         =   0   'False
                  Top             =   6045
                  Width           =   8295
                  _cx             =   14631
                  _cy             =   2937
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
                  Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
                     Height          =   1200
                     Left            =   0
                     TabIndex        =   159
                     Top             =   0
                     Width           =   15900
                     _cx             =   28046
                     _cy             =   2117
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
                     AllowUserResizing=   1
                     SelectionMode   =   0
                     GridLines       =   1
                     GridLinesFixed  =   2
                     GridLineWidth   =   1
                     Rows            =   50
                     Cols            =   8
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"End_oF_service.frx":129B2
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
               Begin MSComCtl2.DTPicker date1 
                  Height          =   345
                  Left            =   3375
                  TabIndex        =   160
                  Top             =   2475
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   609
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   102039553
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker date2 
                  Height          =   345
                  Left            =   240
                  TabIndex        =   161
                  Top             =   2475
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   609
                  _Version        =   393216
                  Format          =   102039553
                  CurrentDate     =   38784
               End
               Begin XtremeSuiteControls.CheckBox ChEndServ 
                  Height          =   255
                  Left            =   5430
                  TabIndex        =   162
                  Top             =   3555
                  Visible         =   0   'False
                  Width           =   255
                  _Version        =   786432
                  _ExtentX        =   450
                  _ExtentY        =   450
                  _StockProps     =   79
                  ForeColor       =   8388608
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox ChSalar 
                  Height          =   255
                  Left            =   5430
                  TabIndex        =   163
                  Top             =   4155
                  Visible         =   0   'False
                  Width           =   255
                  _Version        =   786432
                  _ExtentX        =   450
                  _ExtentY        =   450
                  _StockProps     =   79
                  ForeColor       =   8388608
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox ChValTekt 
                  Height          =   240
                  Left            =   5430
                  TabIndex        =   164
                  Top             =   4515
                  Visible         =   0   'False
                  Width           =   255
                  _Version        =   786432
                  _ExtentX        =   450
                  _ExtentY        =   423
                  _StockProps     =   79
                  ForeColor       =   8388608
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox ChCustom 
                  Height          =   255
                  Left            =   5430
                  TabIndex        =   165
                  Top             =   5220
                  Visible         =   0   'False
                  Width           =   255
                  _Version        =   786432
                  _ExtentX        =   450
                  _ExtentY        =   450
                  _StockProps     =   79
                  ForeColor       =   8388608
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox ChCusTiket 
                  Height          =   255
                  Left            =   5430
                  TabIndex        =   166
                  Top             =   4860
                  Visible         =   0   'False
                  Width           =   255
                  _Version        =   786432
                  _ExtentX        =   450
                  _ExtentY        =   450
                  _StockProps     =   79
                  ForeColor       =   8388608
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox ChAddOther 
                  Height          =   240
                  Left            =   5430
                  TabIndex        =   167
                  Top             =   5580
                  Visible         =   0   'False
                  Width           =   255
                  _Version        =   786432
                  _ExtentX        =   450
                  _ExtentY        =   423
                  _StockProps     =   79
                  ForeColor       =   8388608
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox ChAdvanceTotal 
                  Height          =   240
                  Left            =   5430
                  TabIndex        =   168
                  Top             =   6645
                  Visible         =   0   'False
                  Width           =   255
                  _Version        =   786432
                  _ExtentX        =   450
                  _ExtentY        =   423
                  _StockProps     =   79
                  ForeColor       =   8388608
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox ChCash 
                  Height          =   240
                  Left            =   5430
                  TabIndex        =   169
                  Top             =   6990
                  Visible         =   0   'False
                  Width           =   255
                  _Version        =   786432
                  _ExtentX        =   450
                  _ExtentY        =   423
                  _StockProps     =   79
                  ForeColor       =   8388608
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "»‰«¡ ⁄·Ï ÿ·» «‰Â«¡ «·Œœ„… —ﬁ„"
                  Height          =   405
                  Index           =   22
                  Left            =   1920
                  TabIndex        =   224
                  Top             =   705
                  Width           =   2490
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ÌÊ„"
                  Height          =   405
                  Index           =   21
                  Left            =   2280
                  TabIndex        =   223
                  Top             =   4155
                  Width           =   300
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·Œ’Ê„« "
                  Height          =   300
                  Index           =   20
                  Left            =   2910
                  TabIndex        =   222
                  Top             =   7695
                  Width           =   2355
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈÷«›«  «Œ—Ï"
                  Height          =   285
                  Index           =   11
                  Left            =   3375
                  TabIndex        =   221
                  Top             =   5580
                  Width           =   1890
               End
               Begin VB.Label lblbr 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·›—⁄"
                  Height          =   300
                  Left            =   4845
                  RightToLeft     =   -1  'True
                  TabIndex        =   220
                  Top             =   1065
                  Width           =   1095
               End
               Begin VB.Shape Shape4 
                  BorderColor     =   &H000080FF&
                  Height          =   960
                  Left            =   6135
                  Top             =   2970
                  Width           =   8445
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—« » «·‘Â— «·Õ«·Ï ≈÷«›«  ·⁄œœ"
                  Height          =   285
                  Index           =   19
                  Left            =   3255
                  TabIndex        =   219
                  Top             =   4155
                  Width           =   2250
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… „Œ’’ «·«Ã«“…"
                  Height          =   285
                  Index           =   18
                  Left            =   3855
                  TabIndex        =   218
                  Top             =   5220
                  Width           =   1410
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· –«ﬂ—  ÿ»ﬁ« ··⁄ﬁœ"
                  Height          =   270
                  Index           =   17
                  Left            =   1455
                  TabIndex        =   217
                  Top             =   4515
                  Width           =   1410
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·ÊŸÌ›…"
                  Height          =   270
                  Index           =   5
                  Left            =   4695
                  TabIndex        =   216
                  Top             =   1785
                  Width           =   1245
               End
               Begin VB.Shape Shape3 
                  Height          =   6255
                  Left            =   120
                  Top             =   2895
                  Width           =   5805
               End
               Begin VB.Shape Shape2 
                  BorderColor     =   &H000000FF&
                  Height          =   1680
                  Left            =   345
                  Top             =   6525
                  Width           =   5580
               End
               Begin VB.Shape Shape1 
                  BorderColor     =   &H00008000&
                  Height          =   2235
                  Left            =   345
                  Top             =   4035
                  Width           =   5580
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " -  „Œ’Ê„« „‰Â"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   14.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   300
                  Index           =   16
                  Left            =   3495
                  TabIndex        =   215
                  Top             =   6270
                  Width           =   1890
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " + „÷«›« ≈·ÌÂ   "
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   14.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00008000&
                  Height          =   285
                  Index           =   15
                  Left            =   3975
                  TabIndex        =   214
                  Top             =   3795
                  Width           =   1410
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·⁄Âœ «·⁄Ì‰Ì…"
                  ForeColor       =   &H00FF0000&
                  Height          =   270
                  Index           =   14
                  Left            =   13110
                  TabIndex        =   213
                  Top             =   5700
                  Width           =   1290
               End
               Begin VB.Label Label15 
                  Alignment       =   2  'Center
                  Caption         =   "‘Â—"
                  Height          =   255
                  Left            =   7320
                  RightToLeft     =   -1  'True
                  TabIndex        =   212
                  Top             =   3315
                  Width           =   720
               End
               Begin VB.Label Label14 
                  Alignment       =   2  'Center
                  Caption         =   "”‰…"
                  Height          =   255
                  Left            =   6765
                  RightToLeft     =   -1  'True
                  TabIndex        =   211
                  Top             =   3315
                  Width           =   690
               End
               Begin VB.Label Label13 
                  Alignment       =   2  'Center
                  Caption         =   "‘Â—"
                  Height          =   255
                  Left            =   9480
                  RightToLeft     =   -1  'True
                  TabIndex        =   210
                  Top             =   3315
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.Label Label12 
                  Alignment       =   2  'Center
                  Caption         =   "”‰…"
                  Height          =   255
                  Left            =   8895
                  RightToLeft     =   -1  'True
                  TabIndex        =   209
                  Top             =   3315
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Label Label9 
                  Alignment       =   2  'Center
                  Caption         =   "⁄œœ «Ì«„ «·€Ì«» "
                  Height          =   240
                  Left            =   6885
                  RightToLeft     =   -1  'True
                  TabIndex        =   208
                  Top             =   3090
                  Width           =   1680
               End
               Begin VB.Label Label8 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Õœœ «·„›—œ« "
                  ForeColor       =   &H00FF0000&
                  Height          =   270
                  Left            =   12645
                  RightToLeft     =   -1  'True
                  TabIndex        =   207
                  Top             =   1305
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„·«ÕŸ« "
                  Height          =   255
                  Index           =   4
                  Left            =   4695
                  TabIndex        =   206
                  Top             =   2130
                  Width           =   1245
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·«Ã«“«  »œÊ‰ —« »"
                  Height          =   270
                  Index           =   13
                  Left            =   3495
                  TabIndex        =   205
                  Top             =   6990
                  Width           =   1890
               End
               Begin VB.Label Label10 
                  Alignment       =   2  'Center
                  Caption         =   "ÌÊ„"
                  Height          =   255
                  Left            =   9375
                  RightToLeft     =   -1  'True
                  TabIndex        =   204
                  Top             =   3315
                  Width           =   750
               End
               Begin VB.Label Label7 
                  Alignment       =   2  'Center
                  Caption         =   "⁄œœ «Ì«„ « «·«Ã«“«  »œÊ‰ —« »"
                  Height          =   240
                  Left            =   8670
                  RightToLeft     =   -1  'True
                  TabIndex        =   203
                  Top             =   3090
                  Width           =   2055
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "’«›Ì «·„ﬂ«›√…"
                  ForeColor       =   &H00C00000&
                  Height          =   300
                  Index           =   12
                  Left            =   2040
                  TabIndex        =   202
                  Top             =   8775
                  Width           =   1410
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·”·›"
                  Height          =   270
                  Index           =   10
                  Left            =   4335
                  TabIndex        =   201
                  Top             =   6645
                  Width           =   1050
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "‘—Õ «·«” Õﬁ«ﬁ"
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   9
                  Left            =   13110
                  TabIndex        =   200
                  Top             =   4035
                  Width           =   1290
               End
               Begin VB.Label Label6 
                  Alignment       =   2  'Center
                  Caption         =   "„œ… «·Œœ„…"
                  Height          =   240
                  Index           =   0
                  Left            =   11310
                  RightToLeft     =   -1  'True
                  TabIndex        =   199
                  Top             =   3090
                  Width           =   1575
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·ﬁ«∆„ »«·«‰Â«¡"
                  Height          =   270
                  Index           =   3
                  Left            =   4695
                  TabIndex        =   198
                  Top             =   1425
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "‰Ê⁄ ‰Â«Ì… «·Œœ„…"
                  Height          =   285
                  Index           =   2
                  Left            =   7920
                  TabIndex        =   197
                  Top             =   705
                  Width           =   1290
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· «—ÌŒ"
                  Height          =   285
                  Index           =   1
                  Left            =   10815
                  TabIndex        =   196
                  Top             =   705
                  Width           =   675
               End
               Begin VB.Label OPR 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  ForeColor       =   &H000000FF&
                  Height          =   270
                  Left            =   2160
                  TabIndex        =   195
                  Top             =   6885
                  Width           =   540
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„”·”·"
                  Height          =   240
                  Index           =   3
                  Left            =   12990
                  TabIndex        =   194
                  Top             =   705
                  Width           =   1290
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·«÷«›« "
                  Height          =   285
                  Index           =   7
                  Left            =   3975
                  TabIndex        =   193
                  Top             =   5925
                  Width           =   1290
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„…"
                  Height          =   300
                  Index           =   6
                  Left            =   2400
                  TabIndex        =   192
                  Top             =   8310
                  Width           =   585
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì ‰Â«Ì… «·Œœ„… «·„” Õﬁ"
                  Height          =   285
                  Index           =   5
                  Left            =   3615
                  TabIndex        =   191
                  Top             =   2970
                  Width           =   2130
               End
               Begin VB.Label Label4 
                  Alignment       =   2  'Center
                  Caption         =   "”‰…"
                  Height          =   255
                  Left            =   11070
                  RightToLeft     =   -1  'True
                  TabIndex        =   190
                  Top             =   3315
                  Width           =   720
               End
               Begin VB.Label Label3 
                  Alignment       =   2  'Center
                  Caption         =   "‘Â—"
                  Height          =   255
                  Left            =   11775
                  RightToLeft     =   -1  'True
                  TabIndex        =   189
                  Top             =   3315
                  Width           =   765
               End
               Begin VB.Label Label2 
                  Alignment       =   2  'Center
                  Caption         =   "ÌÊ„"
                  Height          =   255
                  Left            =   12525
                  RightToLeft     =   -1  'True
                  TabIndex        =   188
                  Top             =   3315
                  Width           =   480
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  Caption         =   "«·„œ…"
                  Height          =   255
                  Index           =   0
                  Left            =   705
                  RightToLeft     =   -1  'True
                  TabIndex        =   187
                  Top             =   1065
                  Width           =   765
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   " «·„ÊŸ›"
                  Height          =   225
                  Index           =   1
                  Left            =   12990
                  TabIndex        =   186
                  Top             =   1065
                  Width           =   1290
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  Caption         =   "ÌÊ„"
                  Height          =   255
                  Left            =   7920
                  RightToLeft     =   -1  'True
                  TabIndex        =   185
                  Top             =   3315
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«Ã„«·Ì"
                  Height          =   270
                  Index           =   10
                  Left            =   13110
                  TabIndex        =   184
                  Top             =   3090
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  Caption         =   "0"
                  Height          =   165
                  Index           =   11
                  Left            =   13230
                  TabIndex        =   183
                  Top             =   3555
                  Width           =   1260
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   " «—ÌŒ »œ¡ «·⁄„·"
                  Height          =   300
                  Index           =   9
                  Left            =   4725
                  TabIndex        =   182
                  Top             =   2475
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   " «—ÌŒ ‰Â«Ì… «·⁄„·"
                  ForeColor       =   &H00FF0000&
                  Height          =   300
                  Index           =   0
                  Left            =   1920
                  TabIndex        =   181
                  Top             =   2475
                  Width           =   1290
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„⁄œ·"
                  Height          =   285
                  Index           =   23
                  Left            =   2505
                  TabIndex        =   180
                  Top             =   2970
                  Width           =   705
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "x"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Index           =   24
                  Left            =   3255
                  TabIndex        =   179
                  Top             =   3090
                  Width           =   315
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "’«›Ì ‰Â«Ì… «·Œœ„… ÿ»ﬁ« ·· ’›Ì…"
                  Height          =   285
                  Index           =   25
                  Left            =   0
                  TabIndex        =   178
                  Top             =   2970
                  Width           =   2460
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„Œ’’ ‰Â«Ì… «·Œœ„…"
                  Height          =   285
                  Index           =   26
                  Left            =   3150
                  TabIndex        =   177
                  Top             =   3555
                  Width           =   2115
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… „Œ’’  –«ﬂ—"
                  Height          =   270
                  Index           =   27
                  Left            =   3855
                  TabIndex        =   176
                  Top             =   4860
                  Width           =   1410
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "›—ﬁ «· –«ﬂ—"
                  Height          =   405
                  Index           =   28
                  Left            =   1800
                  TabIndex        =   175
                  Top             =   4860
                  Width           =   900
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "  «Ã„«·Ì «·⁄Âœ"
                  Height          =   270
                  Index           =   29
                  Left            =   990
                  TabIndex        =   174
                  Top             =   6645
                  Width           =   1470
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "›—ﬁ ‰Â«Ì… «·Œœ„…"
                  Height          =   405
                  Index           =   30
                  Left            =   1095
                  TabIndex        =   173
                  Top             =   3555
                  Width           =   1365
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "  Œ’Ê„«  «Œ—Ï"
                  Height          =   270
                  Index           =   31
                  Left            =   990
                  TabIndex        =   172
                  Top             =   6990
                  Width           =   1470
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «· –«ﬂ— «·›⁄·Ì…"
                  Height          =   270
                  Index           =   32
                  Left            =   3855
                  TabIndex        =   171
                  Top             =   4515
                  Width           =   1410
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—« » «·‘Â— «·Õ«·Ì Œ’Ê„« "
                  Height          =   285
                  Index           =   33
                  Left            =   3615
                  TabIndex        =   170
                  Top             =   7335
                  Width           =   1890
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic4 
               Height          =   9165
               Left            =   15360
               TabIndex        =   225
               TabStop         =   0   'False
               Top             =   45
               Width           =   14625
               _cx             =   25797
               _cy             =   16166
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
               Begin VSFlex8UCtl.VSFlexGrid GRID2 
                  Height          =   6915
                  Left            =   0
                  TabIndex        =   226
                  Tag             =   "1"
                  Top             =   330
                  Width           =   28035
                  _cx             =   49451
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
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"End_oF_service.frx":12AF0
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
               Begin VB.Label Label110 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                  Height          =   255
                  Left            =   16815
                  RightToLeft     =   -1  'True
                  TabIndex        =   227
                  Top             =   7320
                  Width           =   6555
               End
            End
         End
      End
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "......"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   10800
      Width           =   5295
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      Caption         =   "„›—œ«  «·—« »"
      Height          =   285
      Index           =   8
      Left            =   14760
      TabIndex        =   47
      Top             =   2640
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   285
      Index           =   2
      Left            =   14160
      TabIndex        =   44
      Top             =   1395
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "End_oF_service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim month_salary As Double
Dim day_salary As Double
Dim net_value As Double
Dim net_value1 As Double
Dim FixedOrChanged(40) As Integer
Dim AddOrDiscount(40) As Integer
Dim ViewComp(40) As Boolean
Dim Account_code(40) As String
Dim Account_code1(40) As String
Dim ZmamAccount(40) As String
Dim AdvPaymentdAccount(40) As String
Dim componentname(40) As String
Dim Advance As Double

Private Sub GetAdvanceValues(IntMonth As Integer, _
                             IntYear As Integer)
    Dim Rs8 As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim LngFindRow As Long
    On Error GoTo hErr
    StrSQL = "Select Emp_ID,Sum(TotalAdvance)as CCC From ( SELECT QryAllEmpAdvance.Emp_ID,QryA" & "llEmpAdvance.TotalAdvance FROM   dbo.QryAllEmpAdvance(" & IntMonth & "," & IntYear & ") QryAllEmpAdvance )" & "Xtable Group By Emp_ID"
    Set Rs8 = New ADODB.Recordset
    Rs8.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs8.BOF Or Rs8.EOF Then
        Exit Sub
    End If

    With Me.Grid
        Rs8.MoveFirst
        .Cell(flexcpText, .FixedRows, .ColIndex("TotalAdvance"), .Rows - 1, .ColIndex("TotalAdvance")) = 0

        For i = 1 To Rs8.RecordCount
            LngFindRow = .FindRow(Rs8("Emp_ID").value, .FixedRows, .ColIndex("Emp_ID"), False, True)

            If LngFindRow <> -1 Then
                If Not (IsNull(Rs8("CCC").value)) Then
                    .TextMatrix(LngFindRow, .ColIndex("TotalAdvance")) = Round(Rs8("CCC").value, 0)
                End If
            End If

            Rs8.MoveNext
        Next i

    End With

hErr:
    'Stop
End Sub
Private Sub CalculateNets()
    Dim i As Integer
    Dim SngHourPrice As Single
    Dim SngOverTimePrice As Single

    Dim NetTotal As Single
    Dim SngTemp As Single
    Dim TotalAddtion As Double
    Dim TotalDiscount As Double
    Dim ColumnName As String
    Dim SngTotal As Double
    Dim j As Integer
    'On Error GoTo ErrTrap
    On Error Resume Next

    With Me.Grid

        If .FixedRows = .Rows Then Exit Sub

        For i = .FixedRows To .Rows - 1
            '     SngHourPrice = Val(.TextMatrix(i, .ColIndex("Emp_Salary"))) / Val(.TextMatrix(i, .ColIndex("DefWorkHours")))
            '     If .TextMatrix(i, .ColIndex("OverTime")) <> "" Then
            '         SngTemp = ConvertHoursToMints(.TextMatrix(i, .ColIndex("OverTime")))
            '         SngTemp = SngTemp * (1 / 60)
            '         SngOverTimePrice = SngTemp * SngHourPrice
            '         .TextMatrix(i, .ColIndex("OverTimePrice")) = SngOverTimePrice
            '         If SngOverTimePrice < 0 Then
            '             .Cell(flexcpForeColor, i, .ColIndex("OverTimePrice")) = vbRed
            '         End If
            '     End If

            TotalAddtion = 0
            TotalDiscount = 0

            For j = 1 To 40
                ColumnName = "Comp" & j

                If AddOrDiscount(j) = 0 Then
                    TotalAddtion = TotalAddtion + val(.TextMatrix(i, .ColIndex(ColumnName)))
                Else
                    TotalDiscount = TotalDiscount + val(.TextMatrix(i, .ColIndex(ColumnName)))
                End If

            Next j
        
            .TextMatrix(i, .ColIndex("total1")) = val(.TextMatrix(i, .ColIndex("Mokafea"))) + TotalAddtion
            .TextMatrix(i, .ColIndex("total2")) = val(.TextMatrix(i, .ColIndex("TotalAdvance"))) + val(.TextMatrix(i, .ColIndex("TotalDiscount"))) + TotalDiscount
            .TextMatrix(i, .ColIndex("EmpTotalNet")) = val(.TextMatrix(i, .ColIndex("total1"))) - val(.TextMatrix(i, .ColIndex("total2")))

            If i Mod 2 = 0 Then
                .Cell(flexcpBackColor, i, 1, i, 41) = &HE0E0E0
     
            End If
        
        Next i
    
    End With

    Exit Sub
ErrTrap:
    'Resume
End Sub

Public Sub FillGridWithData()
    Dim i As Integer
    Dim j As Integer
Dim countFlag As Integer
    Dim Rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim LstDay As Date
    Dim FrstDay As Date
    Dim StrTxt As String
    Dim My_SQL As String
    Dim StrWhere As String
    Dim StrGrp As String
    Dim IntMonth As Integer
    Dim IntYear As Integer
    Dim Msg As String
    Dim ColumnName As String
    Dim TotalAddtion As Double
    Dim TotalDiscount As Double

 
    Set Rs1 = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

'    If Me.CmbMonth.ListIndex = -1 Then Exit Sub
'    If Me.CboYear.ListIndex = -1 Then Exit Sub
countFlag = 1
 

    IntYear = year(date2.value)
    IntMonth = month(date2.value)

 
        Dim ID As String
 
    My_SQL = " Select  lastHolidaydate,BignDateWork,  fullcode,groupid,  BranchId,Emp_ID,Emp_Code,Emp_Name,DepartmentID,project_id ,cost_center_id,IsNUll(Emp_Salary,0)as Emp_Salary,IsNUll(Emp_Salary_sakn,0)as Emp_Salary_sakn,IsNUll(Emp_Salary_bus,0)as Emp_Salary_bus,IsNUll(Emp_Salary_food,0)as Emp_Salary_food,IsNUll(Emp_Salary_others,0)as Emp_Salary_others,IsNUll(Emp_Salary_mob,0)as Emp_Salary_mob,IsNUll(Emp_Salary_mang,0)as Emp_Salary_mang,  IsNUll( TotalDiscount,0)as TotalDiscount,IsNUll(TotalMokafea, 0) As TotalMokafea,(IsNUll(Emp_Salary,0)+IsNUll( TotalMokafea,0))-(IsNUll(TotalDiscount,0)) as EmpTotalNet ,JobTypeName, JobTypeNamee,branch_name,branch_namee,projectFullcode,Project_name,Project_nameE" & CHR(13)
  My_SQL = My_SQL + "  From (" & CHR(13)

  My_SQL = My_SQL + "  SELECT     TOP 100 PERCENT dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.BranchId, dbo.TblEmployee.project_id, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.cost_center_id, SUM(QryAllDiscountWithMkafea.TotalDiscount) AS TotalDiscount, SUM(QryAllDiscountWithMkafea.Mokafea) AS TotalMokafea," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee," & CHR(13)
  My_SQL = My_SQL + "                       dbo.projects.Fullcode AS projectFullcode, dbo.projects.Project_name, dbo.projects.Project_nameE" & CHR(13)
  My_SQL = My_SQL + " FROM         dbo.TblEmpJobsTypes INNER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID LEFT OUTER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.projects ON dbo.TblEmployee.project_id = dbo.projects.id LEFT OUTER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.QryAllDiscountWithMkafea(" & IntMonth & ", " & IntYear & ") QryAllDiscountWithMkafea ON dbo.TblEmployee.Emp_ID = QryAllDiscountWithMkafea.Emp_ID" & CHR(13)

 
        My_SQL = My_SQL + " and dbo.TblEmployee.BignDateWork<" & SQLDate(date2.value, True)
                If Me.DcboEmp.Text <> "" Then
            My_SQL = My_SQL + " Where  dbo.TblEmployee.Emp_id=" & val(DcboEmp.BoundText) ' & "'"
        End If

 'DcboEmpName
 My_SQL = My_SQL + "  GROUP BY dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID, dbo.TblEmployee.BranchId, " & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus," & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang," & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.cost_center_id, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.project_id, dbo.TblEmpJobsTypes.JobTypeName," & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.projects.Fullcode, dbo.projects.Project_name," & CHR(13)
My_SQL = My_SQL + "                      dbo.Projects.Project_nameE" & CHR(13)
My_SQL = My_SQL + " ORDER BY dbo.TblEmployee.Fullcode" & CHR(13)

My_SQL = My_SQL + "  )XTable"


    Rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If Rs1.RecordCount > 0 Then
            .Rows = Rs1.RecordCount + 1
            Rs1.MoveFirst
Dim CountDays As Double
 
Dim MonthDayNo  As Double

MonthDayNo = DaysInMonth(date2.value)

            For i = 1 To .Rows - 1
         countFlag = 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
            .TextMatrix(i, .ColIndex("BignDateWork")) = IIf(IsNull(Rs1.Fields("BignDateWork").value), "", Rs1.Fields("BignDateWork").value)
            .TextMatrix(i, .ColIndex("lastHolidaydate")) = IIf(IsNull(Rs1.Fields("lastHolidaydate").value), "", Rs1.Fields("lastHolidaydate").value)

           
           CountDays = Day(date2.value)
           
            .TextMatrix(i, .ColIndex("CountDays")) = CountDays
            
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(Rs1.Fields("DepartmentID").value), "", Rs1.Fields("DepartmentID").value)
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(Rs1.Fields("BranchId").value), 1, Rs1.Fields("BranchId").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(Rs1.Fields("project_id").value), "", Rs1.Fields("project_id").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(Rs1.Fields("Emp_ID").value), "", Rs1.Fields("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs1.Fields("fullcode").value), "", Rs1.Fields("fullcode").value)
                .TextMatrix(i, .ColIndex("cost_center_id")) = IIf(IsNull(Rs1.Fields("cost_center_id").value), "", Rs1.Fields("cost_center_id").value)
     
                
                      If SystemOptions.UserInterface = ArabicInterface Then
           .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs1.Fields("JobTypeName").value), "", Rs1.Fields("JobTypeName").value)
           .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1.Fields("Emp_Name").value), "", Rs1.Fields("Emp_Name").value)
           Else
           .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs1.Fields("JobTypeNamee").value), "", Rs1.Fields("JobTypeNamee").value)
           .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1.Fields("Emp_Namee").value), "", Rs1.Fields("Emp_Namee").value)
           End If
                TotalAddtion = 0
                TotalDiscount = 0

                For j = 1 To 40
                    ColumnName = "Comp" & j

                    If ViewComp(j) = True Then
                        If FixedOrChanged(j) = 0 Then
                            .TextMatrix(i, .ColIndex(ColumnName)) = GetEmployeeSalaryAccordingToComponent(val(.TextMatrix(i, .ColIndex("Emp_ID"))), CStr(j), , date2.value)
                                           
                                           If countFlag = 1 Then
                                           
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / MonthDayNo * CountDays, 2)
                                           End If
                                           
                        Else
                            .TextMatrix(i, .ColIndex(ColumnName)) = GetEmployeeChangedSalary(val(.TextMatrix(i, .ColIndex("Emp_ID"))), j, val(CboYear(0).Text), CmbMonth(0).ListIndex + 1)
                                                     
                        End If
                    End If
    
                Next j
    
                 '         .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs1.Fields("TotalDiscount").value), "", Round(rs1.Fields("TotalDiscount").value, Decimal_Places))
             
                '.TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs1.Fields("TotalMokafea").value), "", Round(rs1.Fields("TotalMokafea").value, Decimal_Places))
              
                Rs1.MoveNext
            
            Next

            Rs1.Close
        End If

        'GetAdvanceValues IntMonth, IntYear
        ' GetWorkHours
        CalculateNets
        .Rows = .Rows + 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        Else
            .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "Total"
        End If

        .IsSubtotal(.Rows - 1) = True
        Dim SngTotal As Single
 
        For j = 1 To 40
            ColumnName = "Comp" & j
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex(ColumnName), .Rows - 1, .ColIndex(ColumnName))
            .TextMatrix(.Rows - 1, .ColIndex(ColumnName)) = SngTotal
     
        Next j
      
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .Rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.Rows - 1, .ColIndex("Mokafea")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .Rows - 1, .ColIndex("TotalAdvance"))
        .TextMatrix(.Rows - 1, .ColIndex("TotalAdvance")) = SngTotal
         
          'TXTAdvanceTotal.Text = Advance - SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .Rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.Rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .Rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.Rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .Rows, .ColIndex("total1"))
        .TextMatrix(.Rows - 1, .ColIndex("total1")) = SngTotal
         Txtsalary = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .Rows, .ColIndex("total2"))
        .TextMatrix(.Rows - 1, .ColIndex("total2")) = SngTotal
        TxtDisSalary.Text = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .Rows, .ColIndex("EmpTotalNet"))
        .TextMatrix(.Rows - 1, .ColIndex("EmpTotalNet")) = SngTotal


        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 10
        .Cell(flexcpFontName, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
    End With
 

'rs1.Close
Set Rs1 = Nothing

'    Coloring
ErrTrap:
End Sub
Private Sub YearMonth()
    Dim i As Integer
    Dim IntDefIndex As Integer
    'CmbMonth.Clear
    For i = 1 To 12
        CmbMonth(0).AddItem MonthName(i)
    Next
    CmbMonth(0).ListIndex = month(Date) - 1
    ''''''''''
    CboYear(0).Clear
    For i = 2000 To 2050
        CboYear(0).AddItem i
        If i = year(Date) Then
            IntDefIndex = CboYear(0).NewIndex
        End If
    Next
    CboYear(0).ListIndex = IntDefIndex
End Sub
Function getTitlesName() As Boolean
Grid.ColHidden(Grid.ColIndex("TotalAdvance")) = False
getTitlesName = True
    Dim sql As String
    Dim Rs1 As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
    Dim ColumnName As String
    Dim i As Integer
    sql = "select * from mofrad order by id  "
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
 
        For i = 1 To Rs1.RecordCount
            FixedOrChanged(i) = IIf(IsNull(Rs1("FixedOrChanged").value), 0, Rs1("FixedOrChanged").value)
            AddOrDiscount(i) = IIf(IsNull(Rs1("AddOrDiscount").value), 0, Rs1("AddOrDiscount").value)
            ViewComp(i) = IIf(IsNull(Rs1("ViewComp").value), False, Rs1("ViewComp").value)
            Account_code(i) = IIf(IsNull(Rs1("Account_Code").value), "", Rs1("Account_Code").value)
             Account_code1(i) = IIf(IsNull(Rs1("Account_Code1").value), "", Rs1("Account_Code1").value)
             
            
      '      If Account_Code(i) = "" Then
      ''      MsgBox " ·„ Ì „ —»ÿ «·Õ”«» «·Œ«’ » " & ViewComp(i), vbCritical
       '     getTitlesName = False
       '     Exit Function
       '     End If
            
            
            ZmamAccount(i) = IIf(IsNull(Rs1("ZmamAccount").value), 0, Rs1("ZmamAccount").value)
            AdvPaymentdAccount(i) = IIf(IsNull(Rs1("AdvPaymentdAccount").value), 0, Rs1("AdvPaymentdAccount").value)
            
            

            
            
              'AdvPaymentdAccount
            If SystemOptions.UserInterface = ArabicInterface Then
                componentname(i) = IIf(IsNull(Rs1("name").value), "", Rs1("name").value)
            Else
                componentname(i) = IIf(IsNull(Rs1("namee").value), "", Rs1("namee").value)
            End If
             
             
         '   If ViewComp(i) = True And Account_Code(i) = "" And (ZmamAccount(i) <> "True" And AdvPaymentdAccount(i) <> "True") Then
         '   MsgBox " ·„ Ì „ —»ÿ «·Õ”«» «·Œ«’ » " & componentname(i), vbCritical
         '   getTitlesName = False
          
           ' Unload Me
         '     Exit Function
         '   End If
              
              
            With Me.Grid
             
                ColumnName = "Comp" & i

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(0, .ColIndex(ColumnName)) = IIf(IsNull(Rs1("name").value), "", Rs1("name").value)
                Else
                    .TextMatrix(0, .ColIndex(ColumnName)) = IIf(IsNull(Rs1("namee").value), "", Rs1("namee").value)
                End If
                     
                If ViewComp(i) = True Then
                    .ColHidden(.ColIndex(ColumnName)) = False
                Else
                    .ColHidden(.ColIndex(ColumnName)) = True
                End If
                     
            End With
             
 
             
            Rs1.MoveNext
             
        Next i
  
    End If
 
    Rs1.Close
End Function
Private Sub ShowComponent()
'If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    On Error Resume Next
If Me.DcboEmp.BoundText = "" Then Exit Sub
'firstrun = False
     If getTitlesName = True Then
   
   End If
    DoEvents
     
    FillGridWithData
 
    Dim i As Integer
        With Grid
For i = 1 To 40

                If val((.TextMatrix(.Rows - 1, .ColIndex("Comp" & i & "")))) = 0 Then
                  .ColHidden(.ColIndex("Comp" & i)) = True
                End If


                If val((.TextMatrix(.Rows - 1, .ColIndex("sgn")))) = 0 Then
                  .ColHidden(.ColIndex("sgn")) = True
                End If
               If val((.TextMatrix(.Rows - 1, .ColIndex("TotalAdvance")))) = 0 Then
                  .ColHidden(.ColIndex("TotalAdvance")) = True
                End If
                
                          If val((.TextMatrix(.Rows - 1, .ColIndex("TotalDiscount")))) = 0 Then
                  .ColHidden(.ColIndex("TotalDiscount")) = True
                 End If
                
                          If val((.TextMatrix(.Rows - 1, .ColIndex("Mokafea")))) = 0 Then
                  .ColHidden(.ColIndex("Mokafea")) = True
                End If
Next i
End With
'End If
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
  MySQL = " SELECT     dbo.End_of_serviceDetails.ID, dbo.End_of_serviceDetails.IDEndS, dbo.End_of_serviceDetails.MValue, dbo.End_of_serviceDetails.Selected, "
  MySQL = MySQL & "                     dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.mofrdat.mofrad_code, dbo.End_of_service.opr_date, dbo.End_of_service.sal, dbo.End_of_service.sakn,"
  MySQL = MySQL & "                     dbo.End_of_service.bus, dbo.End_of_service.another, dbo.End_of_service.food, dbo.End_of_service.mob, dbo.End_of_service.mang,"
  MySQL = MySQL & "                     dbo.End_of_service.total_salary, dbo.End_of_service.start_date, dbo.End_of_service.[end _date], dbo.End_of_service.daycount, dbo.End_of_service.monthcount,"
  MySQL = MySQL & "                     dbo.End_of_service.yearcount, dbo.End_of_service.total, dbo.End_of_service.opr, dbo.End_of_service.num, dbo.End_of_service.net, dbo.End_of_service.sal1,"
  MySQL = MySQL & "                     dbo.End_of_service.sakn1, dbo.End_of_service.bus1, dbo.End_of_service.another1, dbo.End_of_service.food1, dbo.End_of_service.mob1,"
  MySQL = MySQL & "                     dbo.End_of_service.mang1, dbo.End_of_service.record_date, dbo.End_of_service.Type, dbo.jopstatus.name, dbo.jopstatus.namee, dbo.jopstatus.resignationInt,"
  MySQL = MySQL & "                     dbo.jopstatus.Vacation, dbo.End_of_service.Reaons, dbo.End_of_service.Des, dbo.End_of_service.TotalAdvance, dbo.End_of_service.TotalCash,"
  MySQL = MySQL & "                     dbo.End_of_service.LastTotal, dbo.End_of_service.VwithoutSa, dbo.End_of_service.TxtVlueVaction, dbo.TblUsers.UserName, dbo.End_of_service.EmpID,"
  MySQL = MySQL & "                     dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
  MySQL = MySQL & "                     dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2,"
  MySQL = MySQL & "                     dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Fullcode, dbo.jopstatus.id AS Expr1, dbo.End_of_service.id AS idM,"
  MySQL = MySQL & "                     dbo.End_of_service.NoDayeSa, dbo.End_of_service.Custom, dbo.End_of_service.Ticket, dbo.End_of_service.LastMonth, dbo.End_of_serviceDetails.IDMofrd,"
  MySQL = MySQL & "                     dbo.End_of_serviceDetails.TypeM, dbo.End_of_service.TotalDis, dbo.End_of_service.AddOther, dbo.End_of_service.AlowAssest, dbo.End_of_service.ChCash,"
  MySQL = MySQL & "                     dbo.End_of_service.ChAdvanceTotal, dbo.End_of_service.ChAddOther, dbo.End_of_service.ChCusTiket, dbo.End_of_service.ChCustom,"
  MySQL = MySQL & "                     dbo.End_of_service.ChValTekt, dbo.End_of_service.ChSalar, dbo.End_of_service.ChEndServ, dbo.End_of_service.CusTiket, dbo.End_of_service.EndService,"
  MySQL = MySQL & "                     dbo.End_of_service.NetEnd, dbo.End_of_service.Rate, dbo.End_of_service.ReqNo, dbo.End_of_service.Discounts, dbo.End_of_service.DiffEnd,"
  MySQL = MySQL & "                     dbo.End_oF_service.DiffTekit ,dbo.End_of_service.TicktConract"
  MySQL = MySQL & "    FROM         dbo.jopstatus RIGHT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblUsers RIGHT OUTER JOIN"
  MySQL = MySQL & "                     dbo.mofrdat RIGHT OUTER JOIN"
  MySQL = MySQL & "                     dbo.End_of_serviceDetails ON dbo.mofrdat.mofrad_code = dbo.End_of_serviceDetails.IDMofrd RIGHT OUTER JOIN"
  MySQL = MySQL & "                     dbo.End_of_service INNER JOIN"
  MySQL = MySQL & "                     dbo.TblEmployee ON dbo.End_of_service.EmpID = dbo.TblEmployee.Emp_ID ON dbo.End_of_serviceDetails.IDEndS = dbo.End_of_service.id ON"
  MySQL = MySQL & "                     dbo.TblUsers.UserID = dbo.End_of_service.UserID ON dbo.jopstatus.id = dbo.End_of_service.Type"
    MySQL = MySQL & "   Where  (dbo.End_oF_service.id = " & val(Me.TXTid.Text) & ") and (dbo.End_of_serviceDetails.TypeM  IS NULL)"
    
  'MySQL = MySQL & "   Where (dbo.End_of_serviceDetails.Selected=1) And (dbo.End_oF_service.id = " & val(Me.TXTid.Text) & ") and (dbo.End_of_serviceDetails.TypeM  IS NULL)"
 'MySQL = MySQL & "                           Where (dbo.TblInjury.id = " & val(XPTxtID.text) & ")"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEndService.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEndService.rpt"
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
        Msg = "No Data"
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
       
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
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
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
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

Function print_report22(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.End_of_serviceDetails.ID, dbo.End_of_serviceDetails.IDEndS, dbo.End_of_serviceDetails.MValue, dbo.End_of_serviceDetails.Selected,"
 MySQL = MySQL & "                       dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.mofrdat.mofrad_code, dbo.End_of_serviceDetails.IDMofrd, dbo.End_of_service.opr_date,"
  MySQL = MySQL & "                      dbo.End_of_service.sal, dbo.End_of_service.sakn, dbo.End_of_service.bus, dbo.End_of_service.another, dbo.End_of_service.food, dbo.End_of_service.mob,"
 MySQL = MySQL & "                       dbo.End_of_service.mang, dbo.End_of_service.total_salary, dbo.End_of_service.start_date, dbo.End_of_service.[end _date], dbo.End_of_service.daycount,"
 MySQL = MySQL & "                       dbo.End_of_service.monthcount, dbo.End_of_service.yearcount, dbo.End_of_service.total, dbo.End_of_service.opr, dbo.End_of_service.num, dbo.End_of_service.net,"
 MySQL = MySQL & "                       dbo.End_of_service.sal1, dbo.End_of_service.sakn1, dbo.End_of_service.bus1, dbo.End_of_service.another1, dbo.End_of_service.food1, dbo.End_of_service.mob1,"
 MySQL = MySQL & "                       dbo.End_of_service.mang1, dbo.End_of_service.record_date, dbo.End_of_service.Type, dbo.jopstatus.name, dbo.jopstatus.namee, dbo.jopstatus.resignationInt,"
MySQL = MySQL & "                        dbo.jopstatus.Vacation, dbo.End_of_service.Reaons, dbo.End_of_service.Des, dbo.End_of_service.TotalAdvance, dbo.End_of_service.TotalCash,"
 MySQL = MySQL & "                       dbo.End_of_service.LastTotal, dbo.End_of_service.VwithoutSa, dbo.End_of_service.TxtVlueVaction, dbo.TblUsers.UserName, dbo.End_of_service.EmpID,"
 MySQL = MySQL & "                       dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
 MySQL = MySQL & "                       dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2,"
 MySQL = MySQL & "                       dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Fullcode, dbo.jopstatus.id, dbo.End_of_service.id AS idM ,dbo.End_of_serviceDetails.TypeM "
MySQL = MySQL & "   FROM         dbo.mofrdat RIGHT OUTER JOIN"
MySQL = MySQL & "                        dbo.jopstatus RIGHT OUTER JOIN"
 MySQL = MySQL & "                       dbo.End_of_service INNER JOIN"
 MySQL = MySQL & "                       dbo.TblEmployee ON dbo.End_of_service.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                       dbo.End_of_serviceDetails ON dbo.End_of_service.id = dbo.End_of_serviceDetails.IDEndS LEFT OUTER JOIN"
  MySQL = MySQL & "                      dbo.TblUsers ON dbo.End_of_service.UserID = dbo.TblUsers.UserID ON dbo.jopstatus.id = dbo.End_of_service.Type ON"
MySQL = MySQL & "                        dbo.mofrdat.mofrad_code = dbo.End_of_serviceDetails.IDMofrd"
MySQL = MySQL & "   Where (dbo.End_of_serviceDetails.Selected = 1) And (dbo.End_oF_service.id = " & val(Me.TXTid.Text) & ") and (dbo.End_of_serviceDetails.TypeM  IS NULL) "
 'MySQL = MySQL & "                           Where (dbo.TblInjury.id = " & val(XPTxtID.text) & ")"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEndService.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEndService.rpt"
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
        Msg = "No Data"
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

    xReport.ParameterFields(3).AddCurrentValue user_name
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
'   xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
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
Private Sub ALLButton1_Click()
    Frame2.Visible = True
End Sub
Public Sub CreatLog_File_for_error(str As String)
    Dim StrLogFileName As String
    Dim IntFreeFile As Integer
    Dim ss As String

    StrLogFileName = App.path & "\employee_account_error.txt"

    If Dir(StrLogFileName) <> "" Then
        Kill StrLogFileName
    End If

    ss = "»Ì«‰ »«”„«¡ «·„ÊŸ›Ì‰ «·–Ì‰ ·œÌÂ„ „‘«ﬂ·  "
    ss = ss & vbCrLf & "Byte Informations Systems "
    ss = ss & vbCrLf & "BYTE "
    ss = ss & vbCrLf & "Create Date:- " & Now
    ss = ss & vbCrLf & str & vbCrLf
    IntFreeFile = FreeFile

    Open StrLogFileName For Output As #IntFreeFile
    Print #IntFreeFile, ss
    Close #IntFreeFile
End Sub

Private Sub AlowAssest_Click()
Calcul
End Sub

Private Sub another1_Click()
    'calc_total

End Sub

Private Sub bus1_Click()
    'calc_total

End Sub

Private Sub ChAddOther_Click()
Calcul
End Sub

Private Sub ChAdvanceTotal_Click()
Calcul
End Sub

Private Sub ChCash_Click()
Calcul
End Sub

Private Sub ChCusTiket_Click()
Calcul
End Sub

Private Sub ChCustom_Click()
Calcul
End Sub

Private Sub Check1_Click()

    If Check1.value = vbChecked Then
        sal1.value = vbChecked
        sakn1.value = vbChecked
        food1.value = vbChecked
        another1.value = vbChecked
        mang1.value = vbChecked
        mob1.value = vbChecked
        bus1.value = vbChecked

    Else
        sal1.value = Unchecked
        sakn1.value = Unchecked
        food1.value = Unchecked
        another1.value = Unchecked
        mang1.value = Unchecked
        mob1.value = Unchecked
        bus1.value = Unchecked

    End If

End Sub
Sub chendservice(Optional ByRef Y As Integer, Optional ByRef X As Integer)
    Dim StrSQL As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    StrSQL = "SELECT     id,name,resignationInt"
StrSQL = StrSQL & " From dbo.jopstatus"
StrSQL = StrSQL & " WHERE     (resignationInt = 1 OR resignationInt = 2) and id= " & Y & ""
Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If Rs3.RecordCount > 0 Then
 X = val(Rs3("resignationInt").value) - 1
   Else
   X = 3
   End If
   
End Sub

Private Sub ChEndServ_Click()
Calcul
End Sub

Private Sub ChSalar_Click()
Calcul
End Sub

Private Sub ChValTekt_Click()
Calcul
End Sub
Sub StanderCalCulte()
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        ReLineGrid
        'calc_total
        
        Dim X As Integer
    
        chendservice val(dctype.BoundText), X
   
        txt_salry_total = Round(val(lbl(11).Caption), 2)
        month_salary = Round(lbl(11).Caption, 2)
        day_salary = Round(val(lbl(11).Caption) / 30, 2)
        TxtCash.Text = Round((val(Me.TxtVSa.Text) * day_salary), 2)
        txtCount.Text = Day(date2.value)
        'If X = 0 Then
        TxtRate.Text = 1
  
        If SystemOptions.EndServiceMore5Year = True Then
            cal_intervalOptional
        Else
            cal_interval
        End If
    
    'Else
    
        If X = 1 Then
            If val(txtyear.Text) < 2 Then
                TxtRate.Text = 0
                If SystemOptions.UserInterface = ArabicInterface Then
                    text1.Text = "«ﬁ· «Ê Ì”«ÊÌ 2 ”‰Ê«  ·Ì” ·Â «Ì „” Õﬁ« "
                Else
                    text1.Text = "Less than 2 Years No dues"
                End If
            ElseIf val(txtyear.Text) = 2 And (val(txtmonth.Text) = 0 And val(txtday.Text) = 0) Then
                TxtRate.Text = 0
                If SystemOptions.UserInterface = ArabicInterface Then
                    text1.Text = "«ﬁ· «Ê Ì”«ÊÌ 2 ”‰Ê«  ·Ì” ·Â «Ì „” Õﬁ« "
                Else
                    text1.Text = "Less than 2 Years No dues"
                End If
            ElseIf val(txtyear.Text) = 2 And (val(txtmonth.Text) > 0 Or val(txtday.Text) > 0) Then
                TxtRate.Text = 1 / 3 ' Round(1 / 3, 2)
                If SystemOptions.UserInterface = ArabicInterface Then
                    'Text1.text = Text1.text & vbNewLine
                    text1.Text = text1.Text & "»„« «‰Â« «” ﬁ«·… Ì „ «Õ ”«» À·À1/3 «·„ﬂ«›∆…"
                Else
                    text1.Text = text1.Text & "calculated by one-third 1/3"
                End If
            ElseIf val(txtyear.Text) > 2 And val(txtyear.Text) < 5 Then
                TxtRate.Text = 1 / 3 ' Round(1 / 3, 2)
                If SystemOptions.UserInterface = ArabicInterface Then
                    'Text1.text = Text1.text & vbNewLine
                    text1.Text = text1.Text & "»„« «‰Â« «” ﬁ«·… Ì „ «Õ ”«» À·À1/3 «·„ﬂ«›∆…"
                Else
                    text1.Text = text1.Text & "calculated by one-third 1/3"
                End If
            ElseIf val(txtyear.Text) = 5 And (val(txtmonth.Text) = 0 And val(txtday.Text) = 0) Then
                TxtRate.Text = 1 / 3 ' Round(1 / 3, 2)
                If SystemOptions.UserInterface = ArabicInterface Then
                    'Text1.text = Text1.text & vbNewLine
                    text1.Text = text1.Text & "»„« «‰Â« «” ﬁ«·… Ì „ «Õ ”«» À·À1/3 «·„ﬂ«›∆…"
                Else
                    text1.Text = text1.Text & "calculated by one-third 1/3"
                End If
            ElseIf val(txtyear.Text) = 5 And (val(txtmonth.Text) > 0 Or val(txtday.Text) > 0) Then
                TxtRate.Text = 2 / 3 ' Round(2 / 3, 2)
                If SystemOptions.UserInterface = ArabicInterface Then
                    'Text1.text = Text1.text & vbNewLine
                    text1.Text = text1.Text & "»„« «‰Â« «” ﬁ«·… Ì „ «Õ ”«» À·ÀÌ2/3 «·„ﬂ«›∆…"
                Else
                    text1.Text = text1.Text & "calculated by two-third 2/3"
                End If
            ElseIf val(txtyear.Text) > 5 And val(txtyear.Text) < 10 Then
                TxtRate.Text = 2 / 3 ' Round(2 / 3, 2)
                If SystemOptions.UserInterface = ArabicInterface Then
                    'Text1.text = Text1.text & vbNewLine
                    text1.Text = text1.Text & "»„« «‰Â« «” ﬁ«·… Ì „ «Õ ”«» À·ÀÌ2/3 «·„ﬂ«›∆…"
                Else
                    text1.Text = text1.Text & "calculated by two-third 2/3"
                End If
            ElseIf val(txtyear.Text) = 10 And (val(txtmonth.Text) = 0 And val(txtday.Text) = 0) Then
                TxtRate.Text = 2 / 3
                If SystemOptions.UserInterface = ArabicInterface Then
                    'Text1.text = Text1.text & vbNewLine
                    text1.Text = text1.Text & "»„« «‰Â« «” ﬁ«·… Ì „ «Õ ”«» À·ÀÌ2/3 «·„ﬂ«›∆…"
                Else
                    text1.Text = text1.Text & "calculated by two-third 2/3"
                End If
            ElseIf val(txtyear.Text) = 10 And (val(txtmonth.Text) > 0 Or val(txtday.Text) > 0) Then
                TxtRate.Text = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                    'Text1.text = Text1.text & vbNewLine
                    text1.Text = text1.Text & "»„« «‰Â« «” ﬁ«·… Ì „ «Õ ”«» ﬂ«„·  «·„ﬂ«›∆…"
                Else
                    text1.Text = text1.Text & "calculated All"
                End If
            ElseIf val(txtyear.Text) > 10 Then
                TxtRate.Text = 1
                If SystemOptions.UserInterface = ArabicInterface Then
                    'Text1.text = Text1.text & vbNewLine
                    text1.Text = text1.Text & "»„« «‰Â« «” ﬁ«·… Ì „ «Õ ”«» ﬂ«„·  «·„ﬂ«›∆…"
                Else
                    text1.Text = text1.Text & "calculated All"
                End If
                TxtRate.Text = Round(val(TxtRate.Text), 2)
                'cal_intervaldeparture
            End If
        End If
        Calcul
    End If
End Sub

Private Sub cmd_CALC_NET_Click()
StanderCalCulte
End Sub

Private Sub Cmd_Click(Index As Integer)

    On Error GoTo ErrTrap

    Select Case Index
        Case 0
            clear_all Me
            dctype.BoundText = 0
            Grid.Clear flexClearScrollable, flexClearEverything
            Me.dcby.BoundText = user_id
            EmptyTxet
            Me.TXTid.Text = CStr(new_id("End_of_service", "id", "", True))
            TxtModFlg.Text = "N"
            Me.Dcbranch.BoundText = Current_branch
            Txt_DateEndLincH.value = ToHijriDate(Date)
            txtnet.Text = 0
            TxtVlueVaction.Text = 0
            txtCount.Text = 0
            txtSal.Text = 0
            GRID2.Clear flexClearScrollable, flexClearEverything
            GRID2.Rows = 1
        Case 1
            If chkGE.value = xtpChecked Then '«·ﬁÌœ ⁄·Ì «·Õ—ﬂ…
                If ChekClodePeriod(txtdate.value) = True Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                    Else
                        MsgBox "Please Change Date Becouse This is Period is Closed"
                    End If
                    Exit Sub
                End If
            Else
                If ChekClodePeriod(date2.value) = True Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                    Else
                        MsgBox "Please Change Date Becouse This is Period is Closed"
                    End If
                    Exit Sub
                End If
            End If
            
            If TxtNoteSerial.Text <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· .Ì—ÃÏ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ"
                Else
                    MsgBox "Can Not edit .Delete JE"
                End If
                Exit Sub
            End If
            
            If ScreenAproved(val(TXTid.Text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
                Exit Sub
            End If
            
            If ChekPayment() = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰  ⁄œÌ· Â–Â «·⁄„·Ì… „— »ÿÂ »«·”œ«œ"
                Else
                    MsgBox "Can Not Update This is process linked payment"
                End If
                Exit Sub
            Else
                TxtModFlg.Text = "E"
                Me.dcby.BoundText = user_id
                Me.Dcbranch.BoundText = Current_branch
                DcboEmp_Click (0)
            End If
        Case 2
            Dim Account_Code_dynamic As String
            Account_Code_dynamic = get_account_code_branch(139, my_branch)

            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else
                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«÷«›«  «·«Œ—Ì  ", vbCritical
                    GoTo ErrTrap
                End If
            End If
            Account_Code_dynamic = get_account_code_branch(140, my_branch)

            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else
                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··Œ’Ê„«  «·«Œ—Ì    ", vbCritical
                    GoTo ErrTrap
                End If
            End If
            
            Account_Code_dynamic = get_account_code_branch(141, my_branch)

            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else
                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«Ã«“«  »œÊ‰ —« »  ", vbCritical
                    GoTo ErrTrap
                End If
            End If
            
            
                        If SystemOptions.ProvisionsByManagement Then
                    mTempAccNo = GetDepAccByEmp(val(DcboEmp.BoundText), 2)
                    Account_Code_dynamic = IIf(Trim(mTempAccNo) = "NO account", get_account_code_branch(94, my_branch), Trim(mTempAccNo))

               Else
                   Account_Code_dynamic = get_account_code_branch(94, my_branch)
               End If
               
               
               
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            ElseIf Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»       „’—Ê› ‰Â«ÌÂ «·Œœ„Â  ", vbCritical
                    GoTo ErrTrap
                End If
 
    
            If chkGE.value = xtpChecked Then '«·ﬁÌœ ⁄·Ì «·Õ—ﬂ…
                If ChekClodePeriod(txtdate.value) = True Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                    Else
                        MsgBox "Please Change Date Becouse This is Period is Closed"
                    End If
                    Exit Sub
                End If
            Else
                If ChekClodePeriod(date2.value) = True Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                    Else
                        MsgBox "Please Change Date Becouse This is Period is Closed"
                    End If
                    Exit Sub
                End If
            End If
           
            If AlowAssest.value = vbUnchecked Then
                If val(TxtVlueVaction.Text) <> 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ  ”·Ì„ «·⁄Âœ «·‰ﬁœÌ… «Ê·«"
                    Else
                        MsgBox "Can Not Save Please Drive Assest"
                    End If
                    TxtVlueVaction.SetFocus
                    Exit Sub
                End If
            End If
            ISButton2_Click
            SaveData
        Case 3
            Undo
        Case 4
            If chkGE.value = xtpChecked Then '«·ﬁÌœ ⁄·Ì «·Õ—ﬂ…
                If ChekClodePeriod(txtdate.value) = True Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                    Else
                        MsgBox "Please Change Date Becouse This is Period is Closed"
                    End If
                    Exit Sub
                End If
            Else
                If ChekClodePeriod(date2.value) = True Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                    Else
                        MsgBox "Please Change Date Becouse This is Period is Closed"
                    End If
                    Exit Sub
                End If
            End If
 
            If TxtNoteSerial.Text <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «·Õ–› .Ì—ÃÏ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ"
                Else
                    MsgBox "Can Not delete .Delete JE"
                End If
                Exit Sub
            End If
      
            Del_ProfData

        Case 5
            Load FrmEnserviceSearch
            FrmEnserviceSearch.Index = 0
            FrmEnserviceSearch.Show
        Case 6
            Unload Me
        Case 7
            print_report val(Me.TXTid.Text)
    End Select
    Exit Sub
ErrTrap:
End Sub

'Sub retcountHoliday(Optional EmpID As Integer, Optional ByRef contt As Integer)
' Dim Sql As String
'    Dim mofrad_name As String
'    Dim valuee As Double
'    Dim Rs1 As New ADODB.Recordset
'    Dim Balance As Double
' Dim cont As Integer
' cont = 0
'    Dim I As Integer
''StrSQL = "SELECT count(SpecificHolidyaType1) as cont    * from dbo.TblEmpHolidaysDetails Where (1 = -1) and (SpecificHolidyaType1=1)"
'       Sql = " select * from TblEmpHolidaysDetails "
'Sql = Sql & "  Where ( dbo.TblEmpHolidaysDetails.Emp_id = " & val(EmpID) & ") and (SpecificHolidyaType1=1) "
'Rs1.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'      If Rs1.RecordCount > 0 Then
'      For I = 1 To Rs1.RecordCount
'   If Not IsNull(Rs1("todate").value) Then
'                 strdated = DateDiff("d", Rs1("fromdate").value, Rs1("todate").value)
'
'              cont = cont + strdated
'               End If
'               Rs1.MoveNext
'          Next I
'
'  contt = cont
'End If
    
'End Sub
Public Sub Retrive(Optional Lngid As Long = 0)
    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
         XPTxtCurrent.Caption = 0
         XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        If Lngid <> 0 Then
             rs.Find "id=" & Lngid, , adSearchForward, adBookmarkFirst
             If rs.EOF Or rs.BOF Then
                 Exit Sub
           End If
         End If
    End If
Grid.Clear flexClearScrollable, flexClearEverything
VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
If Not IsNull(rs("AlowAssest").value) Then
    If rs("AlowAssest").value = 1 Then
    AlowAssest.value = vbChecked
    Else
    AlowAssest.value = vbUnchecked
    End If
 Else
 AlowAssest.value = vbUnchecked
 End If
  Me.TxtDiscounts.Text = IIf(IsNull(rs("Discounts").value), 0, (rs("Discounts").value))
    Me.TxtDiffEnd.Text = IIf(IsNull(rs("DiffEnd").value), 0, (rs("DiffEnd").value))
    Me.TXTid.Text = IIf(IsNull(rs("id").value), 0, (rs("id").value))
    Me.TxtReqNo.Text = IIf(IsNull(rs("ReqNo").value), 0, (rs("ReqNo").value))
    Me.TxtNoteID.Text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)
    Me.TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    Me.TxtJbsatust.Text = IIf(IsNull(rs("jopstatusid").value), 0, rs("jopstatusid").value)
    TxtAddOther.Text = IIf(IsNull(rs("AddOther").value), 0, rs("AddOther").value)
    TxtTotalDis.Text = IIf(IsNull(rs("TotalDis").value), 0, rs("TotalDis").value)
    Me.txtEmpCode.Text = IIf(IsNull(rs("emp_code").value), 0, rs("emp_code").value)
    Me.DcboEmp.BoundText = (IIf(IsNull(rs("EmpID").value), 0, (rs("EmpID").value)))
    Me.Txtsalary.Text = IIf(IsNull(rs("sal").value), 0, Trim(rs("sal").value))
    Me.txtsaknm.Text = IIf(IsNull(rs("sakn").value), 0, Trim(rs("sakn").value))
    Me.txtbusm.Text = IIf(IsNull(rs("bus").value), 0, Trim(rs("bus").value))
    Me.TxtDiffTekit.Text = IIf(IsNull(rs("DiffTekit").value), 0, Trim(rs("DiffTekit").value))
    Me.txtanotherm.Text = IIf(IsNull(rs("another").value), 0, Trim(rs("another").value))
    Me.txtfoodm.Text = IIf(IsNull(rs("food").value), 0, Trim(rs("food").value))
    TxtVSa.Text = IIf(IsNull(rs("VwithoutSa").value), 0, Trim(rs("VwithoutSa").value))
    Me.TXTMOBM.Text = IIf(IsNull(rs("mob").value), 0, Trim(rs("mob").value))
    TxtTicktConract.Text = IIf(IsNull(rs("TicktConract").value), 0, Trim(rs("TicktConract").value))
    Me.TxtDisSalary.Text = IIf(IsNull(rs("DisSalary").value), 0, Trim(rs("DisSalary").value))
    Me.TXTMANGM.Text = IIf(IsNull(rs("mang").value), 0, Trim(rs("mang").value))

    Me.txt_salry_total.Text = IIf(IsNull(rs("total_salary").value), 0, Trim(rs("total_salary").value))
 
    Me.txtdate.value = IIf(Not IsDate(rs("record_date").value), Date, rs("record_date").value)
    dctype.BoundText = IIf(rs("Type").value = 0, "", Trim(rs("Type").value))
    dcby.BoundText = IIf(rs("by_employee").value = 0, "", Trim(rs("by_employee").value))
    Me.txtreasons.Text = IIf(IsNull(rs("Reaons").value), "", Trim(rs("Reaons").value))
    Me.text1.Text = IIf(IsNull(rs("Des").value), "", Trim(rs("Des").value))

    txtSal.Text = IIf(IsNull(rs("LastMonth").value), 0, rs("LastMonth").value)
    txtTicketValue.Text = IIf(IsNull(rs("Ticket").value), 0, rs("Ticket").value)
    txtCustom.Text = IIf(IsNull(rs("Custom").value), "0", rs("Custom").value)
    txtCount.Text = IIf(IsNull(rs("NoDayeSa").value), Null, rs("NoDayeSa").value)

    

       Me.TxtVlueVaction.Text = IIf(IsNull(rs("TxtVlueVaction").value), 0, Trim(rs("TxtVlueVaction").value))
    If rs("sal1").value = True Then
        sal1.value = Checked
    Else
        sal1.value = Unchecked

    End If
    
    If rs("sakn1").value = True Then
        sakn1.value = Checked
    Else
        sakn1.value = Unchecked

    End If

    If rs("bus1").value = True Then
        bus1.value = Checked
    Else
        bus1.value = Unchecked

    End If
    
    If rs("another1").value = True Then
        another1.value = Checked
    Else
        another1.value = Unchecked

    End If
    
    If rs("food1").value = True Then
        food1.value = Checked
    Else
        food1.value = Unchecked

    End If
    
    If rs("mob1").value = True Then
        mob1.value = Checked
    Else
        mob1.value = Unchecked

    End If
    
    If rs("mang1").value = True Then
        mang1.value = Checked
    Else
        mang1.value = Unchecked
    End If
    
    Me.date1.value = IIf(Not IsDate(rs("start_date").value), Date, rs("start_date").value)
    Me.date2.value = IIf(Not IsDate(rs("end _date").value), Date, rs("end _date").value)
 
    Me.txtday.Text = IIf(IsNull(rs("daycount").value), 0, Trim(rs("daycount").value))
    Me.txtmonth.Text = IIf(IsNull(rs("monthcount").value), 0, Trim(rs("monthcount").value))
    Me.txtyear.Text = IIf(IsNull(rs("yearcount").value), 0, Trim(rs("yearcount").value))
 
    Me.txttotal.Text = IIf(IsNull(rs("total").value), 0, Trim(rs("total").value))
    Me.OPR.Caption = IIf(IsNull(rs("opr").value), 0, Trim(rs("opr").value))
    Me.txtnum.Text = IIf(IsNull(rs("num").value), 0, Trim(rs("num").value))
    Me.txtnet.Text = IIf(IsNull(rs("net").value), 0, Trim(rs("net").value))
   
    Me.TXTAdvanceTotal.Text = IIf(IsNull(rs("TotalAdvance").value), 0, (rs("TotalAdvance").value))
    Me.TxtCash.Text = IIf(IsNull(rs("TotalCash").value), 0, (rs("TotalCash").value))
    Me.TXTLastTotal.Text = IIf(IsNull(rs("LastTotal").value), 0, Trim(rs("LastTotal").value))
    Me.Dcbranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    ''//////////
    Me.TxtRate.Text = IIf(IsNull(rs("Rate").value), 0, Trim(rs("Rate").value))
    Me.TxtNetEnd.Text = IIf(IsNull(rs("NetEnd").value), 0, Trim(rs("NetEnd").value))
    Me.TxtEndService.Text = IIf(IsNull(rs("EndService").value), 0, Trim(rs("EndService").value))
    Me.TxtCusTiket.Text = IIf(IsNull(rs("CusTiket").value), 0, Trim(rs("CusTiket").value))
     If rs("ChEndServ").value = True Then
        ChEndServ.value = Checked
    Else
        ChEndServ.value = Unchecked
    End If
    
    
      If rs("chkGE").value = True Then
        chkGE.value = Checked
    Else
        chkGE.value = Unchecked
    End If
    
    
    If rs("ChSalar").value = True Then
        ChSalar.value = Checked
    Else
        ChSalar.value = Unchecked
    End If
    If rs("ChValTekt").value = True Then
        ChValTekt.value = Checked
    Else
        ChValTekt.value = Unchecked
    End If
    If rs("ChCustom").value = True Then
        ChCustom.value = Checked
    Else
        ChCustom.value = Unchecked
    End If
    If rs("ChCusTiket").value = True Then
        ChCusTiket.value = Checked
    Else
        ChCusTiket.value = Unchecked
    End If
    If rs("ChAddOther").value = True Then
        ChAddOther.value = Checked
    Else
        ChAddOther.value = Unchecked
    End If
    If rs("ChAdvanceTotal").value = True Then
        ChAdvanceTotal.value = Checked
    Else
        ChAdvanceTotal.value = Unchecked
    End If
    If rs("ChCash").value = True Then
        ChCash.value = Checked
    Else
        ChCash.value = Unchecked
    End If
    
'''''
 Me.dcby.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
Set RsDetails1 = New ADODB.Recordset
StrSQL = " SELECT     dbo.End_of_serviceDetails.ID, dbo.End_of_serviceDetails.IDEndS, dbo.End_of_serviceDetails.MValue, dbo.End_of_serviceDetails.Selected, "
StrSQL = StrSQL & "                      dbo.mofrdat.mofrad_name , dbo.mofrdat.mofrad_namee, dbo.mofrdat.mofrad_code, dbo.End_of_serviceDetails.IDMofrd"
StrSQL = StrSQL & " FROM         dbo.End_of_serviceDetails LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.mofrdat ON dbo.End_of_serviceDetails.IDMofrd = dbo.mofrdat.mofrad_code"

StrSQL = StrSQL & "  Where (dbo.End_of_serviceDetails.IDEndS =" & val(Me.TXTid.Text) & ") and(dbo.End_of_serviceDetails.Selected=1) and (dbo.End_of_serviceDetails.TypeM  IS NULL) "
    
    RsDetails1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
     If Not (RsDetails1.BOF Or RsDetails1.EOF) Then
       With Me.Fg
      '  RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails1.RecordCount

        For i = .FixedRows To .Rows - 1
    
            .TextMatrix(i, .ColIndex("Serial")) = i
            .TextMatrix(i, .ColIndex("mofrdID")) = (IIf(IsNull(RsDetails1("mofrad_code").value), 0, RsDetails1("mofrad_code").value)) 'RsDetails1("Value").value
             .TextMatrix(i, .ColIndex("selected")) = -1 'l(IIf(IsNull(RsDetails1("Selected").value), 0, RsDetails1("Selected").value)) 'RsDetails1("Mainte").value
            .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDetails1("mofrad_namee").value), 0, RsDetails1("mofrad_namee").value) 'RsDetails1("count").value
             .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDetails1("mofrad_name").value), 0, RsDetails1("mofrad_name").value)
           .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDetails1("MValue").value), "", RsDetails1("MValue").value) ' RsDetails1("comp").value
          
            RsDetails1.MoveNext
         
        Next i
End With
    End If
   
    ReLineGrid
RsDetails1.Close
''''''''
' Set RsDetails1 = Nothing '''///////////
 Dim RsDev As ADODB.Recordset
 Set RsDev = New ADODB.Recordset
StrSQL = " SELECT     TOP 100 PERCENT dbo.End_of_serviceDetails.IDEndS, dbo.End_of_serviceDetails.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee, dbo.End_of_serviceDetails.TypeM, dbo.End_of_serviceDetails.DeliverDate, dbo.End_of_serviceDetails.ReciveDate,"
StrSQL = StrSQL & "                      dbo.End_of_serviceDetails.IDMofrd , dbo.TblAssestes.AsName, dbo.TblAssestes.AsCode, dbo.TblAssestes.AsestName"
StrSQL = StrSQL & " FROM         dbo.End_of_serviceDetails LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAssestes ON dbo.End_of_serviceDetails.IDMofrd = dbo.TblAssestes.AsID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.End_of_serviceDetails.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & " Where (dbo.End_of_serviceDetails.IDEndS = " & val(TXTid.Text) & ") And (dbo.End_of_serviceDetails.TypeM = 1)"

    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
     If Not (RsDev.BOF Or RsDev.EOF) Then
       With Me.VSFlexGrid2
      '  RsDetails.MoveFirst
        .Rows = .FixedRows + RsDev.RecordCount

        For i = .FixedRows To .Rows - 1
                 .TextMatrix(i, .ColIndex("MofrdID")) = IIf(IsNull(RsDev("IDMofrd").value), "", RsDev("IDMofrd").value)
                .TextMatrix(i, .ColIndex("AsCode")) = IIf(IsNull(RsDev("AsCode").value), "", RsDev("AsCode").value)
                .TextMatrix(i, .ColIndex("DeliverDate")) = IIf(IsNull(RsDev("DeliverDate").value), "", RsDev("DeliverDate").value)
                .TextMatrix(i, .ColIndex("ReciveDate")) = IIf(IsNull(RsDev("ReciveDate").value), "", RsDev("ReciveDate").value)
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(RsDev("EmpID").value), "", RsDev("EmpID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDev("AsName").value), "", RsDev("AsName").value)
                .TextMatrix(i, .ColIndex("Emp_NameTo")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
                Else
                 .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDev("AsName").value), "", RsDev("AsName").value)
                .TextMatrix(i, .ColIndex("Emp_NameTo")) = IIf(IsNull(RsDev("Emp_Namee").value), "", RsDev("Emp_Namee").value)
                End If
            RsDev.MoveNext
         
        Next i
End With
    End If
   
    ReLineGrid
    fillapprovData
RsDev.Close
 Set RsDev = Nothing
 
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    ShowComponent
ReLineGrid
    Exit Sub
ErrTrap:
End Sub
Sub RtriverAsse(Optional EmpID As Integer = 0)
Dim sql As String
Dim i As Integer
Dim RsDev As ADODB.Recordset
sql = " SELECT     TOP 100 PERCENT dbo.TblAssestes.AsID, dbo.TblAssestes.AsName, dbo.TblAssestes.AsCode, TblEmployee_2.Emp_Name, TblEmployee_2.Fullcode,"
sql = sql & "                      TblEmployee_2.Emp_Namee, dbo.TblEmpAsest.ToEmId, TblEmployee_1.Emp_Name AS Emp_NameTo, TblEmployee_1.Fullcode AS FullcodeTo,"
sql = sql & "                       TblEmployee_1.Emp_Namee AS Emp_NameToE, dbo.TblEmpAsest.DeliverDate, dbo.TblEmpAsest.PostedDate, dbo.TblEmpAsestDetails.Qunt,"
sql = sql & "                       dbo.TblEmpAsestDetails.DIFF , dbo.TblEmpAsestDetails.FlagAs, dbo.TblEmpAsest.TypeAsset, dbo.TblEmpAsest.EmpAsestID"
sql = sql & "  FROM         dbo.TblEmpAsest LEFT OUTER JOIN"
sql = sql & "                       dbo.TblEmpAsestDetails ON dbo.TblEmpAsest.EmpAsID = dbo.TblEmpAsestDetails.IDAseset LEFT OUTER JOIN"
sql = sql & "                       dbo.TblAssestes ON dbo.TblEmpAsestDetails.AsID = dbo.TblAssestes.AsID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblEmployee TblEmployee_1 ON dbo.TblEmpAsest.ToEmId = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblEmployee TblEmployee_2 ON dbo.TblEmpAsest.EmpAsestID = TblEmployee_2.Emp_ID"
sql = sql & "  Where (dbo.TblEmpAsestDetails.FlagAs Is Null) And (dbo.TblEmpAsest.EmpAsestID =" & EmpID & ")"
Set RsDev = New ADODB.Recordset
       RsDev.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
 VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
           VSFlexGrid2.Rows = 1
    If (RsDev.RecordCount > 0) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid2
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("MofrdID")) = IIf(IsNull(RsDev("AsID").value), "", RsDev("AsID").value)
            
                .TextMatrix(i, .ColIndex("AsCode")) = IIf(IsNull(RsDev("AsCode").value), "", RsDev("AsCode").value)
                .TextMatrix(i, .ColIndex("DeliverDate")) = IIf(IsNull(RsDev("DeliverDate").value), "", RsDev("DeliverDate").value)
                .TextMatrix(i, .ColIndex("ReciveDate")) = IIf(IsNull(RsDev("PostedDate").value), "", RsDev("PostedDate").value)
            
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(RsDev("ToEmId").value), "", RsDev("ToEmId").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDev("AsName").value), "", RsDev("AsName").value)
                .TextMatrix(i, .ColIndex("Emp_NameTo")) = IIf(IsNull(RsDev("Emp_NameTo").value), "", RsDev("Emp_NameTo").value)
                Else
                 .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDev("AsName").value), "", RsDev("AsName").value)
                .TextMatrix(i, .ColIndex("Emp_NameTo")) = IIf(IsNull(RsDev("Emp_NameToE").value), "", RsDev("Emp_NameToE").value)
                End If
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
    RsDev.Close
End Sub
Private Sub SaveData()
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
  ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then

        If Me.DcboEmp.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» «Œ Ì«— «”„ «·„ÊŸ› "
            Else
            Msg = "Please Select Employee"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboEmp.SetFocus
             SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
 With Me.VSFlexGrid2

        For i = 1 To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("MofrdID"))) <> 0 Then
            If .TextMatrix(i, .ColIndex("DeliverDate")) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ Ì—ÃÏ  ”·Ì„ «·⁄Âœ «Ê·«"
            Else
            MsgBox "Can Not Save Please Drive Assest"
            End If
        Exit Sub
        End If
        End If
        Next i
        End With
        
        If Me.dcby.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» «Œ Ì«— «”„ «·ﬁ«∆„ »«·⁄„·Ì… "
            Else
            Msg = "Please Select Employee"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcby.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
 
        If Me.dctype.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» «Œ Ì«— ‰Ê⁄ ⁄„·Ì… «·«‰Â«¡ "
            Else
            Msg = "Please Select Type of End Service"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dctype.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
 
        If Me.txtnum.Text <> "" Then
            If Not (IsNumeric(Me.txtnum.Text)) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "«·ﬁÌ„… ·«»œ «‰  ﬂÊ‰ —ﬁ„"
                Else
                Msg = "Only Value"
                End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.txtnum.SetFocus
                SelectText Me.txtnum
                Exit Sub
            End If
        End If
 If BeginTrans = False Then
        Cn.BeginTrans
        BeginTrans = True
  End If
        
  If Me.TxtModFlg.Text = "E" Then
  StrSQL = "Delete From End_of_serviceDetails Where IDEndS=" & val(TXTid.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
  StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
                 Cn.Execute StrSQL, , adExecuteNoRecords
  End If
        If TxtModFlg.Text = "N" Then
            '  Dim RsTemp As New ADODB.Recordset
            StrSQL = "select * From End_of_service where emp_code='" & (Me.txtEmpCode.Text) & " '"
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If RsTemp.RecordCount > 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = " „ Õ”«» «·„ﬂ«›√… ·Â–« «·„ÊŸ› „‰ ﬁ»·" & CHR(13)
                Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & CHR(13)
                Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                Else
                Msg = "This account has been rewarded by the employee before"
                End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If
            
            rs.AddNew
            rs("id").value = val(Me.TXTid.Text)
           ElseIf Me.TxtModFlg.Text = "E" Then
            StrSQL = "Delete From End_of_serviceDetails Where IDEndS=" & val(Me.TXTid.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If

        rs("emp_code").value = IIf(txtEmpCode.Text = "", 0, Trim(txtEmpCode.Text))
        rs("EmpID").value = val(IIf(DcboEmp.BoundText = "", 0, (DcboEmp.BoundText)))
    
        rs("sal").value = IIf(Txtsalary.Text = "", 0, Trim(Txtsalary.Text))
    If AlowAssest.value = vbChecked Then
    rs("AlowAssest").value = 1
    End If
        rs("DiffEnd").value = val(TxtDiffEnd.Text)
        rs("TxtVlueVaction").value = IIf(TxtVlueVaction.Text = "", 0, Trim(TxtVlueVaction.Text))
        rs("sakn").value = IIf(txtsaknm.Text = "", 0, Trim(txtsaknm.Text))
        rs("DiffTekit").value = IIf(TxtDiffTekit.Text = "", 0, val(TxtDiffTekit.Text))
        rs("VwithoutSa").value = IIf(TxtVSa.Text = "", 0, Trim(TxtVSa.Text))
        rs("bus").value = IIf(txtbusm.Text = "", 0, Trim(txtbusm.Text))
        rs("another").value = IIf(txtanotherm.Text = "", 0, Trim(txtanotherm.Text))
        rs("food").value = IIf(txtfoodm.Text = "", 0, Trim(txtfoodm.Text))
        rs("AddOther").value = IIf(TxtAddOther.Text = "", 0, val(TxtAddOther.Text))
        rs("TotalDis").value = IIf(TxtTotalDis.Text = "", 0, val(TxtTotalDis.Text))
        rs("mob").value = IIf(TXTMOBM.Text = "", 0, Trim(TXTMOBM.Text))
        rs("mang").value = IIf(TXTMANGM.Text = "", 0, Trim(TXTMANGM.Text))
        rs("total_salary").value = IIf(txt_salry_total.Text = "", 0, Trim(txt_salry_total.Text))
        rs("Discounts").value = IIf(Not IsNumeric(TxtDiscounts.Text), 0, val(TxtDiscounts.Text))
        rs("LastMonth").value = IIf(Not IsNumeric(txtSal.Text), 0, val(txtSal.Text))
        rs("Ticket").value = IIf(Not IsNumeric(txtTicketValue.Text), 0, val(txtTicketValue.Text))
        rs("Custom").value = IIf(Not IsNumeric(txtCustom.Text), 0, val(txtCustom.Text))
        rs("NoDayeSa").value = IIf(Not IsNumeric(txtCount.Text), 0, val(txtCount.Text))
        rs("TicktConract").value = IIf(TxtTicktConract.Text = "", 0, val(TxtTicktConract.Text))
        rs("jopstatusid").value = IIf(Not IsNumeric(TxtJbsatust.Text), 0, val(TxtJbsatust.Text))
        rs("DisSalary").value = val(TxtDisSalary.Text)
        If sal1.value = Checked Then
            rs("sal1").value = 1
      
        Else
            rs("sal1").value = 0

        End If
    
        If sakn1.value = Checked Then
       
            rs("sakn1").value = 1
        Else
     
            rs("sakn1").value = 0

        End If

        If bus1.value = Checked Then
            rs("bus1").value = 1
       
        Else
     
            rs("bus1").value = 0

        End If
    
        If another1.value = Checked Then
            rs("another1").value = 1
        Else
            rs("another1").value = 0

        End If
    
        If food1.value = Checked Then
            rs("food1").value = 1
        Else
            rs("food1").value = 0

        End If
    
        If mob1.value = Checked Then
            rs("mob1").value = 1
        Else
            rs("mob1").value = 0

        End If
    
        If mang1.value = Checked Then
            rs("mang1").value = 1
        Else
            rs("mang1").value = 0

        End If
    
        rs("start_date").value = Me.date1.value
        rs("end _date").value = Me.date2.value
   
        rs("record_date").value = Me.txtdate.value
        rs("Type").value = IIf(dctype.BoundText = "", 0, val(dctype.BoundText))
        rs("by_employee").value = IIf(dcby.BoundText = "", 0, val(dcby.BoundText))
        rs("Reaons").value = IIf(txtreasons.Text = "", "", Trim(txtreasons.Text))
        rs("Des").value = IIf(text1.Text = "", "", Trim(text1.Text))

        rs("daycount").value = IIf(txtday.Text = "", 0, Trim(txtday.Text))
        rs("monthcount").value = IIf(txtmonth.Text = "", 0, Trim(txtmonth.Text))
        rs("yearcount").value = IIf(txtyear.Text = "", 0, Trim(txtyear.Text))
       
        rs("total").value = IIf(Not IsNumeric(txttotal.Text), 0, val(txttotal.Text))

        rs("opr").value = IIf(Me.OPR.Caption = "", "+", Trim(Me.OPR.Caption))

        rs("num").value = IIf(txtnum.Text = "", 0, Trim(txtnum.Text))
        rs("ReqNo").value = IIf(Not IsNumeric(TxtReqNo.Text), 0, val(TxtReqNo.Text))
        rs("net").value = IIf(Not IsNumeric(txtnet.Text), 0, val(txtnet.Text))
        rs("TotalAdvance").value = IIf(Not IsNumeric(Me.TXTAdvanceTotal.Text), 0, val(TXTAdvanceTotal.Text))
        rs("TotalCash").value = IIf(Not IsNumeric(Me.TxtCash.Text), 0, val(TxtCash.Text))
        rs("LastTotal").value = IIf(Not IsNumeric(Me.TXTLastTotal.Text), 0, val(TXTLastTotal.Text))
        rs("UserID").value = IIf(Me.dcby.BoundText = "", Null, Me.dcby.BoundText)
        rs("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
        rs("rate").value = IIf(Not IsNumeric(Me.TxtRate.Text), 0, val(TxtRate.Text))
        rs("NetEnd").value = IIf(Not IsNumeric(Me.TxtNetEnd.Text), 0, val(TxtNetEnd.Text))
        rs("EndService").value = IIf(Not IsNumeric(Me.TxtEndService.Text), 0, val(TxtEndService.Text))
        rs("CusTiket").value = IIf(Not IsNumeric(Me.TxtCusTiket.Text), 0, val(TxtCusTiket.Text))
        If ChEndServ.value = Checked Then
            rs("ChEndServ").value = 1
        Else
            rs("ChEndServ").value = 0
        End If
        
        
        If chkGE.value = Checked Then
            rs("chkGE").value = 1
        Else
            rs("chkGE").value = 0
        End If
        
         
         
        If ChSalar.value = Checked Then
            rs("ChSalar").value = 1
        Else
            rs("ChSalar").value = 0
        End If
        If ChValTekt.value = Checked Then
            rs("ChValTekt").value = 1
        Else
            rs("ChValTekt").value = 0
        End If
           If ChCustom.value = Checked Then
            rs("ChCustom").value = 1
        Else
            rs("ChCustom").value = 0
        End If
           If ChCusTiket.value = Checked Then
            rs("ChCusTiket").value = 1
        Else
            rs("ChCusTiket").value = 0
        End If
         If ChAddOther.value = Checked Then
            rs("ChAddOther").value = 1
        Else
            rs("ChAddOther").value = 0
        End If
          If ChAdvanceTotal.value = Checked Then
            rs("ChAdvanceTotal").value = 1
        Else
            rs("ChAdvanceTotal").value = 0
        End If
        If ChCash.value = Checked Then
            rs("ChCash").value = 1
        Else
            rs("ChCash").value = 0
        End If
        
        rs.update
        
    StrSQL = "update TblEmployee Set   jopstatusid=" & val(dctype.BoundText) & " ,workstate=0   , endWork=" & SQLDate(date2, True) & " where Emp_ID=" & val(DcboEmp.BoundText)
'  StrSQL = StrSQL & " , endWork=" & SQLDate(date2, True)
    
    Cn.Execute StrSQL
    If val(TxtReqNo.Text) <> 0 Then
     StrSQL = "update TBLRegisterHoliday Set   FlagPayed=1   where ID=" & val(TxtReqNo.Text)
    Cn.Execute StrSQL
    End If
    
        If Change_filed_value(Me.DcboEmp.BoundText, "Emp_ID", "jopstatusid", "TblEmployee", Me.dctype.BoundText) Then
        End If
    
        If Change_filed_value(Me.DcboEmp.BoundText, "Emp_ID", "workstate", "TblEmployee", 0) Then
        End If
    
        If Change_filed_value(Me.DcboEmp.BoundText, "Emp_ID", "Notsstkala", "TblEmployee", Me.txtreasons.Text) Then
        End If
    ''''''
          Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from dbo.End_of_serviceDetails Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      ' RsDetails1.Open "TblCardAuthorizationReformDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
If Fg.Rows > 1 Then
       For i = Me.Fg.FixedRows To Fg.Rows - 1
       If val(Fg.TextMatrix(i, Fg.ColIndex("mofrdID"))) <> 0 Then
              If Fg.TextMatrix(i, Fg.ColIndex("mofrd")) = "" Then
              If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» «Œ Ì«—  «”„ «·„›—œ!! "
            Else
            Msg = "Please Select Component"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
           ' Me.TxtCliientName.SetFocus
           ' SendKeys "{F4}"
           Cn.CommitTrans
            Exit Sub
        End If
       
           RsDetails1.AddNew
          RsDetails1("IDEndS").value = val(TXTid.Text)
        RsDetails1("IDMofrd").value = val(Fg.TextMatrix(i, Fg.ColIndex("mofrdID")))
        If Fg.Cell(flexcpChecked, i, Fg.ColIndex("selected")) = flexChecked Then
       RsDetails1("Selected").value = -1
       Else
         RsDetails1("Selected").value = 0
         End If
           RsDetails1("MValue").value = val(Fg.TextMatrix(i, Fg.ColIndex("value")))
         ' RsDetails1("selected").value = val(XPTxtID.text)
   
         RsDetails1.update

       End If
           Next i
        End If
    ''''''
      RsDetails1.Close
        Set RsDetails1 = Nothing
    ''/////////////
    Dim RsDev As ADODB.Recordset
             Set RsDev = New ADODB.Recordset
         StrSQL = "SELECT     *  from dbo.End_of_serviceDetails Where (1 = -1)"
         RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      With Me.VSFlexGrid2

        For i = 1 To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("MofrdID"))) <> 0 Then
                RsDev.AddNew
                RsDev("IDEndS").value = val(Me.TXTid.Text)
                RsDev("IDMofrd").value = val(.TextMatrix(i, .ColIndex("MofrdID")))
                RsDev("EmpID").value = val(.TextMatrix(i, .ColIndex("EmpID")))
                RsDev("DeliverDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("DeliverDate"))), .TextMatrix(i, .ColIndex("DeliverDate")), Null)
                RsDev("ReciveDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("ReciveDate"))), .TextMatrix(i, .ColIndex("ReciveDate")), Null)
                RsDev("TypeM").value = 1
                RsDev.update
                    
            End If
            
            '
        Next i

    End With
        UpdateAdvance val(Me.DcboEmp.BoundText)
        Cn.CommitTrans
        BeginTrans = False
        
     
           XPTxtCurrent.Caption = rs.AbsolutePosition
           XPTxtCount.Caption = rs.RecordCount
     
        Select Case Me.TxtModFlg.Text

            Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„ÊŸ› " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
Else
Msg = "This record already saved" & CHR(13)
Msg = Msg & "you need to eneter another record "
End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                  MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
                End If
        End Select

        TxtModFlg.Text = "R"
    End If

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
         Msg = "You can not save data " & CHR(13)
        Msg = Msg + "It has been enter  incorrect data " & CHR(13)
        Msg = Msg + "Make sure of the validity of the data and try again"
         End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If BeginTrans = True Then
        Cn.RollbackTrans
        BeginTrans = False
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Sorry...error during Saving"
    End If
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
            rs.Find "id='" & val(Me.TXTid.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub CmdExit_Click()
    Frame2.Visible = False
End Sub
Private Sub Comman1122_Click()
ShowComponent
ReLineGrid
End Sub

Private Sub Command1_Click(Index As Integer)
    OPR.Caption = Command1(Index).Caption
End Sub
 Sub GetEmployeeSalaryAccordingToComponentEndservice(Emp_id As Integer)
                                                    
    Dim sql As String
    Dim mofrad_name As String
    Dim valuee As Double
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim Mofradd As String
    Dim i As Integer
    Mofradd = ""
  
  '  sql = "SELECT     dbo.EmpSalaryComponent.[Value],dbo.mofrdat.mofrad_name,dbo.mofrdat.mofrad_type "
  '  sql = sql & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
  '  sql = sql & " dbo.mofrdat ON dbo.EmpSalaryComponent.AccountCode = dbo.mofrdat.mofrad_code"
  '  sql = sql & " WHERE   (dbo.EmpSalaryComponent.emp_ID = " & Emp_id & ")"
 sql = " SELECT     dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.EmpSalaryComponent.[Value], dbo.mofrdat.mofrad_type , dbo.EmpSalaryComponent.AccountCode"
sql = sql & "  FROM         dbo.mofrad INNER JOIN"
sql = sql & "                       dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type INNER JOIN"
sql = sql & "                       dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode"
sql = sql & "  Where (dbo.EmpSalaryComponent.Emp_id = " & Emp_id & ") And (dbo.mofrad.Aloc2 = 1)"
      rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
  With Me.Fg
  .Rows = 1
 End With
    If rs.RecordCount > 0 Then
  With Me.Fg
  .Rows = rs.RecordCount + 1
      For i = 1 To rs.RecordCount
       .TextMatrix(i, .ColIndex("Serial")) = i
      .TextMatrix(i, .ColIndex("mofrdID")) = IIf(IsNull(rs("AccountCode").value), 0, rs("AccountCode").value)
       .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(rs("mofrad_name").value), "", rs("mofrad_name").value)
 .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(rs("value").value), 0, rs("value").value)
 .TextMatrix(i, .ColIndex("selected")) = -1
 rs.MoveNext
      Next i
 End With
     End If
    rs.Close
    ReLineGrid
End Sub






Private Sub Command11_Click()
Dim Msg As String
 If chkGE.value = xtpChecked Then '«·ﬁÌœ ⁄·Ì «·Õ—ﬂ…
          If ChekClodePeriod(txtdate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
 Else
         If ChekClodePeriod(date2.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
 
 
 End If
  Me.TxtModFlg.Text = "E"
   ISButton2_Click
     SaveData
 createVoucher
      If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " „ «‰‘«¡ «·ﬁÌœ"
    Else
        Msg = "Create Successfully   "
    End If
    MsgBox Msg
End Sub

Private Sub Command5_Click()
 If chkGE.value = xtpChecked Then '«·ﬁÌœ ⁄·Ì «·Õ—ﬂ…
          If ChekClodePeriod(txtdate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
 Else
         If ChekClodePeriod(date2.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
 
 
 End If
 
 

If TxtNoteSerial.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Ì—ÃÏ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ «Ê·« "
    Else
        Msg = "Please Delete GL  "
    End If
    MsgBox Msg
   Exit Sub
  End If

 createVoucher
     If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " „ «‰‘«¡ «·ﬁÌœ"
    Else
        Msg = "Create Successfully   "
    End If
    MsgBox Msg
End Sub

Private Sub Command7_Click()
Dim Msg As String
Dim StrSQL As String
Dim X As Integer
If chkGE.value = xtpChecked Then '«·ﬁÌœ ⁄·Ì «·Õ—ﬂ…
          If ChekClodePeriod(txtdate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
 Else
         If ChekClodePeriod(date2.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
 
 
 End If
  If ChekPayment() = True Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "·«Ì„ﬂ‰ Õ–› Â–« «·ﬁÌœ „— »ÿÂ »«·”œ«œ"
 Else
 MsgBox "Can Not Dlete  This is voucher linked payment"
 End If
 Exit Sub
 End If
              
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " √ﬂÌœ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ  "
    Else
        Msg = "Confirm Delete  "
    End If
        X = MsgBox(Msg, vbCritical + vbYesNo)

      If X = vbYes Then

        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "update   End_oF_service set NoteSerial=null,NoteID=null Where ID=" & val(Me.TXTid.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        rs.Resync
        Retrive (val(Me.TXTid.Text))
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " „  Õ–› ﬁÌœ «·«” Õﬁ«ﬁ  "
        Else
            Msg = " This voucher deleted  "
        End If
        MsgBox Msg
       End If
        
End Sub

Private Sub Command8_Click()
Dim StrTempAccountCode As String
            Dim FirstPeriod As Date
            getFirstPeriodDateInthisYear FirstPeriod
                   'StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
 
       StrTempAccountCode = get_EMPLOYEE_Account(val(DcboEmp.BoundText), "Account_Code1")    '«·«ÃÊ— «·„” Õﬁ…
            '    StrAccountCode = Employee_account
         
         
            ShowReport StrTempAccountCode, DcboEmp.Text, FirstPeriod, Date
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub

Private Sub date2_Change()
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
txtCount.Text = Day(date2.value)
End If
End Sub


Private Sub DcboEmp_Change()
DcboEmp_Click (0)
If val(DcboEmp.BoundText) > 0 Then
RtriverAsse val(DcboEmp.BoundText)
 With Me.VSFlexGrid2

        For i = 1 To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("MofrdID"))) <> 0 Then
            If .TextMatrix(i, .ColIndex("DeliverDate")) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «Õ ”«» «· ’›Ì… Ì—ÃÏ  ”·Ì„ «·⁄Âœ «Ê·«"
            Else
            MsgBox "Can Not Calculate Please Drive Assest"
            End If
            ISButton2.Enabled = False
        Exit Sub
        End If
        End If
        Next i
        End With
  End If
End Sub

Private Sub DcboEmp_Click(Area As Integer)
    
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
 ISButton2.Enabled = True
If Me.dctype.Text = "" Or val(Me.dctype.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " ÌÃ» «Œ Ì«— ‰Ê⁄ ‰Â«Ì… «·Œœ„…"
Else
MsgBox "Select Tpe of End Work"
End If
 dctype.SetFocus
Exit Sub
End If

End If

End Sub

Function calc_total()
If Me.TxtModFlg.Text <> "R" Then
   ' Me.TXTAdvanceTotal.text = getEmployeeAdvance(val(Me.DcboEmp.BoundText))
    Dim Total As Double

    If sal1.value = vbChecked Then
        Total = Total + val(Txtsalary.Text)
    End If
    
    If Me.sakn1.value = vbChecked Then
        Total = Total + val(txtsaknm.Text)
    End If
    
    If Me.bus1.value = vbChecked Then
        Total = Total + val(txtbusm.Text)
    End If
    
    If Me.another1.value = vbChecked Then
        Total = Total + val(txtanotherm.Text)
    End If
    
    If Me.food1.value = vbChecked Then
        Total = Total + val(txtfoodm.Text)
    End If
    
    If Me.mob1.value = vbChecked Then
        Total = Total + val(TXTMOBM.Text)
    End If
    
    If Me.mang1.value = vbChecked Then
        Total = Total + val(TXTMANGM.Text)
    End If
    
    txt_salry_total = lbl(11).Caption
    month_salary = lbl(11).Caption
    day_salary = val(lbl(11).Caption) / 30
    cal_interval
  End If
End Function

Function get_employee_information(ID As Integer, Optional ByRef AccountCode As String, Optional ByRef Account_code2 As String, Optional ByRef Account_Code5 As String, Optional ByRef Account_code4 As String)
If Me.TxtModFlg.Text <> "R" Then
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String

    sql = "Select * from TblEmployee where Emp_ID=" & ID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rs3.RecordCount > 0 Then
        Txtsalary.Text = IIf(IsNull(Rs3("Emp_Salary").value), 0, Rs3("Emp_Salary").value)
        txtsaknm.Text = IIf(IsNull(Rs3("Emp_Salary_sakn").value), 0, Rs3("Emp_Salary_sakn").value)
         AccountCode = IIf(IsNull(Rs3("Account_Code").value), "", Rs3("Account_Code").value)
         Account_Code5 = IIf(IsNull(Rs3("Account_Code5").value), "", Rs3("Account_Code5").value)
         Account_code2 = IIf(IsNull(Rs3("Account_Code2").value), "", Rs3("Account_Code2").value)
         Account_code4 = IIf(IsNull(Rs3("Account_code4").value), "", Rs3("Account_code4").value)
        txtbusm.Text = IIf(IsNull(Rs3("Emp_Salary_bus").value), 0, Rs3("Emp_Salary_bus").value)
    
        txtanotherm.Text = IIf(IsNull(Rs3("Emp_Salary_others").value), 0, Rs3("Emp_Salary_others").value)
        txtfoodm.Text = IIf(IsNull(Rs3("Emp_Salary_food").value), 0, Rs3("Emp_Salary_food").value)
      
        TXTMOBM.Text = IIf(IsNull(Rs3("Emp_Salary_mob").value), 0, Rs3("Emp_Salary_mob").value)
        TXTMANGM.Text = IIf(IsNull(Rs3("Emp_Salary_mang").value), 0, Rs3("Emp_Salary_mang").value)
        Me.date1.value = IIf(Not IsDate(Rs3("BignDateWork").value), Date, Rs3("BignDateWork").value)
        If Me.TxtModFlg.Text = "N" Then
       TxtJbsatust.Text = IIf(IsNull(Rs3("jopstatusid").value), 0, Rs3("jopstatusid").value)
       End If
    Else
    Account_code2 = ""
    Account_Code5 = ""
        txt_salry_total = 0
    End If

    Rs3.Close
End If
End Function



Private Sub DcboEmp_LostFocus()
If val(DcboEmp.BoundText) > 0 Then
RtriverAsse val(DcboEmp.BoundText)
 With Me.VSFlexGrid2

        For i = 1 To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("MofrdID"))) <> 0 Then
            If .TextMatrix(i, .ColIndex("DeliverDate")) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «Õ ”«» «· ’›Ì… Ì—ÃÏ  ”·Ì„ «·⁄Âœ «Ê·«"
            Else
            MsgBox "Can Not Calculate Please Drive Assest"
            End If
            ISButton2.Enabled = False
        Exit Sub
        End If
        End If
        Next i
        End With
  End If
End Sub

Private Sub Dcbranch_Change()
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.Text = ""
End If
End Sub

Private Sub Dcbranch_Click(Area As Integer)
Dcbranch_Change
End Sub

Private Sub dctype_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
'DcboEmp_Click (0)
EmptyTxet
End If
End Sub

Private Sub fg_Click()
ReLineGrid
End Sub

Sub RetriveRquestEnd(Optional ByRef ID As Double)
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
sql = "Select * from TBLRegisterHoliday where id =" & ID & " and FlagPayed  is null"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
dctype.BoundText = IIf(IsNull(Rs3("jopstatusid").value), 0, Rs3("jopstatusid").value)
DcboEmp.BoundText = IIf(IsNull(Rs3("EmpID").value), 0, Rs3("EmpID").value)
date2.value = IIf(IsNull(Rs3("EndWork").value), Date, Rs3("EndWork").value)
txtreasons.Text = IIf(IsNull(Rs3("Notsstkala").value), "", Rs3("Notsstkala").value)
Else
txtreasons.Text = ""
dctype.BoundText = 0
DcboEmp.BoundText = 0
End If
Rs3.Close
End Sub
Private Sub Del_ProfData()

    Dim Msg As String
    Dim StrSQL As String
    
    On Error GoTo ErrTrap
    
    If ChekPayment() = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ Õ–› Â–Â «·⁄„·Ì… „— »ÿÂ »«·”œ«œ"
        Else
            MsgBox "Can Not delete This is process linked payment"
        End If
        Exit Sub
    Else
        If Me.TXTid.Text <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
                Msg = Msg + (Me.TXTid.Text) & CHR(13)
                Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
            Else
                Msg = "Confirm Delete"
            End If

            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
                If val(TxtJbsatust.Text) = 1 Or val(TxtJbsatust.Text) = 0 Then
                    StrSQL = "update TblEmployee Set   jopstatusid=1 ,workstate=1  where Emp_ID=" & val(DcboEmp.BoundText)
                    Cn.Execute StrSQL
                Else
                    StrSQL = "update TblEmployee Set   jopstatusid=" & val(TxtJbsatust.Text) & " ,workstate=0  where Emp_ID=" & val(DcboEmp.BoundText)
                    Cn.Execute StrSQL
                End If
                
                If val(TxtReqNo.Text) <> 0 Then
                    StrSQL = "update TBLRegisterHoliday Set   FlagPayed=null   where ID=" & val(TxtReqNo.Text)
                    Cn.Execute StrSQL
                End If
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Update TblEmpAdvanceRequestDetails Set Payed=null    Where EndSrvID=" & val(TXTid.Text) & "  "
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                Deletepost Me.Name, "End_of_service", "id", 0, val(Dcbranch.BoundText), val(TXTid.Text), TXTid
                
                If Not rs.RecordCount < 1 Then
                    rs.delete
                    rs.MoveFirst
                    StrSQL = "Delete From End_of_serviceDetails Where IDEndS=" & val(TXTid.Text) & ""
                    Cn.Execute StrSQL, , adExecuteNoRecords
                    
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
            
            GRID2.Clear flexClearScrollable, flexClearEverything
            GRID2.Rows = 1
            
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
            Else
                Msg = "This process is not available because there is no record"
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtModFlg_Change
            Exit Sub
        End If
        TxtModFlg_Change
    End If
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–… «·⁄„·Ì… "
        Else
            Msg = "You can not delete this record" & CHR(13) & "There are all linked data"
        End If
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If
End Sub
Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter  As Integer
   lbl(11).Caption = 0
    IntCounter = 0

    With Fg

        For i = .FixedRows To .Rows - 1

       
 If Fg.Cell(flexcpChecked, i, Fg.ColIndex("selected")) = flexChecked Then
                
                lbl(11).Caption = val(lbl(11).Caption) + val(Fg.TextMatrix(i, Fg.ColIndex("value")))
        
            End If
        Next i
 
    End With
     IntCounter = 0

If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
txtSal.Text = 0
    With Grid

        For i = .FixedRows To .Rows - 1
       If .TextMatrix(i, .ColIndex("Emp_Code")) <> "" Then
    txtSal.Text = val(txtSal.Text) + val(.TextMatrix(i, .ColIndex("total1")))
        
            End If
        Next i
 txtSal.Text = Round(val(txtSal.Text), 2)
    End With
    
End If

End Sub
Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    Dim My_SQL As String

    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetUsers Me.dcby
    Dcombos.GetJobEndService dctype
    
    'My_SQL = "  select  id,name  from jopstatus where id>1 "
    
    If SystemOptions.AllowDynamicEdit = True Then
         Command11.Visible = True
    End If

    'fill_combo dctype, My_SQL
    
    YearMonth
    
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    
    'Resize_Form Me
    'AddTip
    
    Set rs = New ADODB.Recordset
    rs.Open "[End_of_service]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2
    'Me.TxtModFlg.Text = "R"

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:

End Sub
Function ChekPayment() As Boolean
Dim sql As String
ChekPayment = False
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = "Select NoteID from  Notes where TxtEndService=" & val(TXTid.Text) & " and CashingType=10  "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChekPayment = True
Else
ChekPayment = False
End If
End Function

Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "«À»«  ‰Â«Ì… Œœ„… „ÊŸ› »”‰œ —ﬁ„" & TXTid.Text & " ··„ÊŸ› " & DcboEmp.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
tablename = "End_of_service"
Filedname = "ID"
NoteSerial1 = val(TXTid)
Notevalue = 0

 notytype = 9050
Notevalue = val(txtSal.Text)
 

 BranchID = val(Dcbranch.BoundText)
 If chkGE.value = xtpChecked Then
 NoteDate = (txtdate.value)
 Else
 
NoteDate = (date2.value)
End If
If Notevalue <> 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TxtNoteID.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial
                                     Else
                                                 If TxtNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                            CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                                 TxtNoteID.Text = NoteID
                                                                TxtNoteSerial.Text = NoteSerial
                                                   Else
                                                                 sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   sql = sql & " where NoteID=" & val(TxtNoteID.Text)
                                                                   Cn.Execute sql
                                                               
                                                 End If
                                       
                                End If
'If val(txtSal.Text) >= 0 Then
CREATE_VOUCHER_GE val(TxtNoteID.Text), BranchID, user_id, NoteDate
'Else
'CREATE_VOUCHER_GE2 val(TxtNoteID.Text), BranchID, user_id, NoteDate
'End If
 rs.Resync adAffectCurrent
     End If

End Function
Function check_employee_accounts() As Boolean
    Dim Employee_account As String
    Dim error_string As String
    error_string = ""
    check_employee_accounts = True
    Dim i As Integer

    With Grid

        For i = .FixedRows To .Rows - 2
                   If val(.TextMatrix(i, .ColIndex("BranchId"))) = 0 Then
                   error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡    ÕœÌœ «·›—⁄ «· «»⁄ ·Â"
        
                check_employee_accounts = False
                   End If
                   
                   
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code")

            If Employee_account = "" Or (Employee_account) = Null Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«» –„ …"
        
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«» –„ … ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
       
                check_employee_accounts = False
            End If
            
            
  Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
                    If Employee_account = "" Or (Employee_account) = Null Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«» «·«ÃÊ— «·„” Õﬁ…"
        
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«» «·«ÃÊ— «·„” Õﬁ… ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
       
                check_employee_accounts = False
            End If
            
            
            
  Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code3")
                    If Employee_account = "" Or (Employee_account) = Null Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«»   «·„œ›Ê⁄«  «·„ﬁœ„…"
        
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«»    «·„œ›Ê⁄«  «·„ﬁœ„… ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
       
                check_employee_accounts = False
            End If
            
            
            '     If Val(.TextMatrix(i, .ColIndex("Emp_Salary"))) = 0 Then
            '     error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & " ·„ Ì „  ÕœÌœ —« » «”«”Ì ·Â  " & vbCrLf
            '
            '    check_employee_accounts = False
            '
            '     End If
            If error_string <> "" Then
            CreatLog_File_for_error (error_string)
       End If
        Next i

    End With

    Dim X As Integer
    Dim StrLogFileName As String

    If error_string <> "" Then
        X = MsgBox("Â·  —Ìœ › Õ «·„·› ··„—«Ã⁄Â", vbCritical + vbYesNo, "ÌÊÃœ Œÿ√ ›Ì Õ”«»«  «·„ÊŸ›Ì‰  Ì„ﬂ‰ „—«Ã⁄ … ›Ì „·› «·«Œÿ«¡")

        If X = vbYes Then
            StrLogFileName = App.path & "\employee_account_error.txt"
            ShellExecute 0&, vbNullString, StrLogFileName, vbNullString, vbNullString, vbNormalFocus
        End If
    End If

End Function
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
Dim BasicSalaryAccount As String
Dim StrSQL As String
Dim Benfts As Double
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
    Dim depit_side As String
    Dim credit_side As String
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Dim j As Integer
    Dim ColumnName As String
    Dim SalaryAccount As String
    Dim BonusAccount As String
    Dim DiscountAccount As String
    
        Msg = "«À»«  ‰Â«Ì… «·Œœ„… „ÊŸ› »”‰œ —ﬁ„" & TXTid.Text & " ··„ÊŸ› " & DcboEmp.Text
    If check_employee_accounts = False Then
        Exit Function
    End If
    Dim mTempAccNo As String
    
        
BasicSalaryAccount = ""
 notes_id = general_noteid
                  
    For j = 1 To 40
        ColumnName = "Comp" & j

        If ViewComp(j) = True Then
                                  
            If CheckAccountToJE(Account_code(j)) = False Then
                Account_code(j) = SalaryAccount
            End If
        End If
    
    Next j
     my_branch = val(Dcbranch.BoundText)
 
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
    '«·ÿ—› «·„œÌ‰ «·«÷«ﬁ« 
     
    Dim CValue As Double
    Dim Branch As Integer
    Dim ProjectID As Integer
    
    BranchID = 1
    
    With Grid
BranchID = .TextMatrix(1, .ColIndex("BranchId"))
End With
    With Grid

        For j = 1 To 40
            ColumnName = "Comp" & j
            If ViewComp(j) = True And AddOrDiscount(j) = 0 Then '«·ŸÂÊ— Ê«÷«›… Ê·Ì” –„„ Ê·Ì” „ﬁœ„
                       If BasicSalaryAccount = "" Then
                                                                        BasicSalaryAccount = Account_code(j)
                                                 End If
                                                 
                If ZmamAccount(j) <> True And AdvPaymentdAccount(j) <> True Then
                                   
                                
                        CValue = GetComponentValuePerBranch(BranchID, ColumnName)
                        CValue = Round(CValue, 2)
                               If val(txtCount.Text) = 0 Then
                               CValue = 0
                        
                                                 
                               End If
                               
                        If CValue > 0 Then
                                            
                
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 0, Msg & " —« » «·‘Â— «·Õ«·Ì »⁄œœ  " & txtCount & " ÌÊ„  " & .TextMatrix(0, .ColIndex(ColumnName)), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If

                            line_no = line_no + 1
                        End If

 '                       rsBranch.MoveNext
'                    Next branch
                             
                End If
                             
            End If
    
        Next j
       
                                      

        '''///////////≈÷«›«  «Œ—Ï
               my_branch = BranchID
                total_value = val(TxtAddOther.Text)
        total_value = Round(total_value, 2)
         If total_value > 0 Then
          credit_side = get_account_code_branch(139, my_branch)
         depit_side = get_EMPLOYEE_Account(DcboEmp.BoundText, "Account_Code4")
         
                    If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, Abs(Round(total_value, 2)), 0, Msg + " Õ”«» «·«÷«›«   ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    line_no = line_no + 1
                 '    If ModAccounts.AddNewDev(LngDevID, line_no, depit_side, Abs(Round(total_value, 2)), 1, Msg + " Õ”«» ‰Â«Ì… «·Œœ„… «÷«›«   ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
                 '       GoTo ErrTrap
                 '   End If
                 '   line_no = line_no + 1
            End If
           '''/////////////////////////

        For i = .FixedRows To .Rows - 2
    Benfts = val(.TextMatrix(i, .ColIndex("total1"))) + val(TxtAddOther.Text)
    Benfts = Round(Benfts, 2)
            If Benfts > 0 And val(val(txtCount.Text)) <> 0 Then        '«·«ÃÊ— «·„” Õﬁ… œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Benfts, 1, Msg & " —« » «·‘Â— «·Õ«·Ì »⁄œœ  " & txtCount & " ÌÊ„  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
     
     
              '   If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) < 0 And val(val(txtCount.Text)) <> 0 Then         '«·«ÃÊ— «·„” Õﬁ… „œÌ‰
              '  Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
              '  StrAccountCode = Employee_account
        '
        '        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Round(Abs(.TextMatrix(i, .ColIndex("EmpTotalNet"))), 2), 0, Msg & " —« » «·‘Â— «·Õ«·Ì »⁄œœ  " & txtCount & " ÌÊ„  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
        ''            GoTo ErrTrap
         '       End If

         '       line_no = line_no + 1
           ' End If
            
 '   For j = 1 To 40 ' Œ’Ê„« 
'
'            ColumnName = "Comp" & j
'
'            If ViewComp(j) = True And AddOrDiscount(j) = -1 Then
'                If ZmamAccount(j) <> True And AdvPaymentdAccount(j) <> True Then
'
'depit_side = get_account_code_branch(139, my_branch)
'         credit_side = get_EMPLOYEE_Account(DcboEmp.BoundText, "Account_Code")
'
'                        CValue = GetComponentValuePerBranch(BranchID, ColumnName)
'                        CValue = Round(CValue, 2)
'                               If val(txtCount.Text) = 0 Then
'                               CValue = 0
'                               End If
'
''                        If CValue > 0 Then
 '                  '      SystemOptions.ProjectEmployeeGV = True
 'If SystemOptions.ProjectDiscountPolicy = 1 Then
 '
 '                           If ModAccounts.AddNewDev(LngDevID, line_no, Account_code1(j), CValue, 1, Msg & "   Œ’Ê„«  " & " —« » «·‘Â— «·Õ«·Ì »⁄œœ  " & txtCount & " ÌÊ„  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
 '                               GoTo ErrTrap
 '                           End If
 '
 '                           Else
 '
 '                                    If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code(j), CValue, 1, Msg & " —« » «·‘Â— «·Õ«·Ì »⁄œœ  " & txtCount & " ÌÊ„  " & "   Œ’Ê„«  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
 '                               GoTo ErrTrap
 '                           End If
 '
 '
 '
'End If
'                            line_no = line_no + 1
'                        End If
                                    
 '                       rsBranch.MoveNext
 '                   Next branch
 '
'                End If
'            End If
'
'        Next j
            
            
'            For j = 1 To 40
'                ColumnName = "Comp" & j
'
'                If ViewComp(j) = True And ZmamAccount(j) = True Then
'
'                    Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
'                    StrAccountCode = Employee_account
''
 '                   If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 And val(val(txtCount.Text)) <> 0 Then
 '                       If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Round(.TextMatrix(i, .ColIndex(ColumnName)), 2), 1, Msg & " –„„ " & " —« » «·‘Â— «·Õ«·Ì »⁄œœ  " & txtCount & " ÌÊ„  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
 '                           GoTo ErrTrap
 '                       End If
'
'                        line_no = line_no + 1
'                    End If
'
'                End If
'
'            Next j
            '''///////////Œ’Ê„«  «Œ—Ï
            Dim discount As Double
            discount = 0
              StrAccountCode = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code")
         If val(TxtDiscounts.Text) > 0 And val(val(txtCount.Text)) <> 0 Then
                  If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Round(val(TxtDiscounts.Text), 2), 0, Msg & "”œ«œ ”·› " & " —« » «·‘Â— «·Õ«·Ì    ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                   GoTo ErrTrap
                End If
               line_no = line_no + 1
           End If
        If val(TxtDisSalary.Text) > 0 And val(val(txtCount.Text)) <> 0 Then
                  If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Round(val(TxtDisSalary.Text), 2), 0, Msg & "”œ«œ ”·› " & " —« » «·‘Â— «·Õ«·Ì    ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                   GoTo ErrTrap
                End If
               line_no = line_no + 1
        End If
                If val(TxtCash.Text) > 0 And val(val(txtCount.Text)) <> 0 Then
                  If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Round(val(TxtCash.Text), 2), 0, Msg & "”œ«œ ”·› " & " —« » «·‘Â— «·Õ«·Ì    ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                   GoTo ErrTrap
                End If
               line_no = line_no + 1
        End If
           discount = val(TxtDiscounts.Text) + val(TxtDisSalary.Text) + val(TxtCash.Text)
           discount = Round(discount, 2)
           credit_side = get_account_code_branch(139, my_branch)
                   If discount > 0 And val(val(txtCount.Text)) <> 0 Then
                  If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, discount, 1, Msg & "”œ«œ ”·› " & " —« » «·‘Â— «·Õ«·Ì    ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                   GoTo ErrTrap
                End If
               line_no = line_no + 1
        End If
         '   If val(.TextMatrix(i, .ColIndex("TotalAdvance"))) > 0 And val(val(txtCount.Text)) <> 0 Then        '«·”·› œ«∆‰
         '       Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
         '       StrAccountCode = Employee_account
        '
        '        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Round(.TextMatrix(i, .ColIndex("TotalAdvance")), 2), 1, Msg & "”œ«œ ”·› " & " —« » «·‘Â— «·Õ«·Ì    ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
        '            GoTo ErrTrap
        '        End If
'
'                line_no = line_no + 1
'            End If
'*********************************************************«Œ—Ì «÷«›…
CValue = 0 'val(TxtSalEntitOther)
If CValue > 0 Then
CValue = Round(CValue, 2)
                       If ModAccounts.AddNewDev(LngDevID, line_no, BasicSalaryAccount, CValue, 0, Msg & "   «Œ—Ì «÷«›… ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If


                   line_no = line_no + 1


                 Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, CValue, 1, Msg & "   «Œ—Ì «÷«›…  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
End If


'*************************************************************************************************************************
'*********************************************************«Œ—Ì Œ’„
CValue = 0
If CValue > 0 Then


                 Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        CValue = Round(CValue, 2)
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, CValue, 0, Msg & "   «Œ—Ì Œ’„   ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
                
                                       If ModAccounts.AddNewDev(LngDevID, line_no, BasicSalaryAccount, CValue, 1, Msg & "   «Œ—Ì Œ’„ ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If


                   line_no = line_no + 1


End If


'****************************************************************************”·› ﬂ«ﬂ·…
CValue = 0 'val(TxtAdvance)
             If CValue > 0 Then  '«·”·› œ«∆‰
             CValue = Round(CValue, 2)
                          Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, CValue, 0, Msg & "   ”œ«œ ”·›     ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
                
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, CValue, 1, Msg & "”œ«œ ”·› ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
            
 '****************************************************************************”·› ﬂ«ﬂ·…
'*******************************„œ›Ê⁄«  „ﬁ
            For j = 1 To 40
                ColumnName = "Comp" & j

                If ViewComp(j) = True And AdvPaymentdAccount(j) = True Then
                     
                    Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code3") 'œ›⁄«  „ﬁœ„…
                    StrAccountCode = Employee_account
                                 If AddOrDiscount(j) = 0 Then
                                                    If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                                                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Round(.TextMatrix(i, .ColIndex(ColumnName)), 2), 0, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                                            GoTo ErrTrap
                                                        End If
                                
                                                        line_no = line_no + 1
                                                    End If
                        
                        Else
                        
                                                If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                                                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Round(.TextMatrix(i, .ColIndex(ColumnName))), 1, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                                            GoTo ErrTrap
                                                        End If
                                
                                                        line_no = line_no + 1
                                                    End If

                        
                        
                        End If
                        
                 
                End If

            Next j
                 

            
'*******************************„œ›Ê⁄«  „ﬁ
 
        Next i

    End With
           my_branch = BranchID
                total_value = val(TxtEndService.Text) - val(TxtNetEnd.Text)
                total_value = Round(total_value, 2)
                depit_side = get_EMPLOYEE_Account(DcboEmp.BoundText, "Account_Code4")
         
         
         If SystemOptions.ProvisionsByManagement Then
                    mTempAccNo = GetDepAccByEmp(val(DcboEmp.BoundText), 3)
                    credit_side = IIf(Trim(mTempAccNo) = "NO account", get_account_code_branch(56, my_branch), Trim(mTempAccNo))


        Else
            credit_side = get_account_code_branch(56, my_branch)
        End If
         
         If total_value < 0 Then
         
                    If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, Abs(Round(total_value, 2)), 0, Msg + "›—Êﬁ«  ‰Â«Ì… «·Œœ„…  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    line_no = line_no + 1
                     If ModAccounts.AddNewDev(LngDevID, line_no, depit_side, Abs(Round(total_value, 2)), 1, Msg + " ›—Êﬁ«  ‰Â«Ì… «·Œœ„…  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    line_no = line_no + 1
             ElseIf total_value > 0 Then
           If ModAccounts.AddNewDev(LngDevID, line_no, depit_side, Abs(Round(total_value, 2)), 0, Msg + " ›—Êﬁ«  ‰Â«Ì… «·Œœ„…  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    line_no = line_no + 1
                     If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, Abs(Round(total_value, 2)), 1, Msg + " ›—Êﬁ«  ‰Â«Ì… «·Œœ„…  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    line_no = line_no + 1
                    
                    
                End If
    ''////////// –«ﬂ—
       my_branch = BranchID
                total_value = val(TxtCusTiket.Text) - val(txtTicketValue.Text)
                total_value = Round(total_value, 2)
                depit_side = get_EMPLOYEE_Account(DcboEmp.BoundText, "Account_Code5")
                
            If SystemOptions.ProvisionsByManagement Then
                    mTempAccNo = GetDepAccByEmp(val(DcboEmp.BoundText), 2)
                    credit_side = IIf(Trim(mTempAccNo) = "NO account", get_account_code_branch(94, my_branch), Trim(mTempAccNo))

               Else
                   credit_side = get_account_code_branch(94, my_branch)
               End If
                                
         
         If total_value < 0 Then
         
                    If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, Abs(Round(total_value, 2)), 0, Msg + "  –«ﬂ—  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    line_no = line_no + 1
                     If ModAccounts.AddNewDev(LngDevID, line_no, depit_side, Abs(Round(total_value, 2)), 1, Msg + "  –«ﬂ—  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    line_no = line_no + 1
             ElseIf total_value > 0 Then
           If ModAccounts.AddNewDev(LngDevID, line_no, depit_side, Abs(Round(total_value, 2)), 0, Msg + "  –«ﬂ—  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    line_no = line_no + 1
                     If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, Abs(Round(total_value, 2)), 1, Msg + "  –«ﬂ—  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    line_no = line_no + 1
                    
                    
                End If
  ''////////////
      ''//////////«÷«›« 

            '''''''''//////////////Œ’Ê„«  «Œ—Ï
     '           my_branch = BranchID
     '           total_value = val(TxtDiscounts.Text)
     '            total_value = Round(total_value, 2)
     '           Befents = Round(val(txtSal.Text), 2)
     ''   depit_side = get_account_code_branch(140, my_branch)
      '   If total_value > 0 Then
      '   CValue = total_value
      '                If Befents >= CValue And CValue <> 0 Then
      '                  TepValue = CValue
      '                  Befents = Befents - CValue
      '                  CValue = 0
      '                  Else
      '                  TepValue = Befents
      '
      '                  CValue = CValue - Befents
      '                  Befents = 0
      '                  End If
      '
      '                 credit_side = get_EMPLOYEE_Account(DcboEmp.BoundText, "Account_Code1")
      '
      '              If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, Abs(Round(TepValue, 2)), 0, Msg + " ‰Â«Ì… «·Œœ„… Œ’Ê„«   ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
      '                  GoTo ErrTrap
      '              End If
      '              line_no = line_no + 1
      '               If ModAccounts.AddNewDev(LngDevID, line_no, depit_side, Abs(Round(TepValue, 2)), 1, Msg + " Õ”«» Œ’Ê„«  ‰Â«Ì… «·Œœ„… ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
      '                  GoTo ErrTrap
      '              End If
      '              line_no = line_no + 1
      '
      '            total_value = val(TxtDiffEnd.Text)
      '           credit_side = get_EMPLOYEE_Account(DcboEmp.BoundText, "Account_Code4")
      '           If CValue > 0 And total_value > 0 Then
      '        If total_value >= CValue Then
      '
      '                  TepValue = CValue
      '                  total_value = 0
      '                  CValue = 0
      '                  Else
      '                  TepValue = total_value
      '                  CValue = CValue - total_value
      '                  End If
      '
      '                If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, Abs(Round(TepValue, 2)), 0, Msg + " ‰Â«Ì… «·Œœ„… Œ’Ê„«   ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
      '                  GoTo ErrTrap
      '              End If
      '              line_no = line_no + 1
      '               If ModAccounts.AddNewDev(LngDevID, line_no, depit_side, Abs(Round(TepValue, 2)), 1, Msg + " Õ”«» Œ’Ê„«  ‰Â«Ì… «·Œœ„… ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
      '                  GoTo ErrTrap
      '              End If
      '              line_no = line_no + 1
      '              End If
      '             If CValue >= 0 Then
      '             credit_side = get_EMPLOYEE_Account(DcboEmp.BoundText, "Account_Code")
      '            If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, Abs(Round(CValue, 2)), 0, Msg + " ‰Â«Ì… «·Œœ„… Œ’Ê„«   ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
      '                  GoTo ErrTrap
      '              End If
      '              line_no = line_no + 1
      '               If ModAccounts.AddNewDev(LngDevID, line_no, depit_side, Abs(Round(CValue, 2)), 1, Msg + " Õ”«» Œ’Ê„«  ‰Â«Ì… «·Œœ„… ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
      '                  GoTo ErrTrap
      '              End If
      '              line_no = line_no + 1
      '            End If
      '      End If
      '''//////////Œ’Ê„« 
    '   my_branch = BranchID
    '    total_value = val(TxtDiscounts.Text)
    '    total_value = Round(total_value, 2)
    '     If total_value > 0 Then
    '      depit_side = get_account_code_branch(140, my_branch)
    '     credit_side = get_EMPLOYEE_Account(DcboEmp.BoundText, "Account_Code1")
    '
    '                If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, Abs(Round(total_value, 2)), 0, Msg + " ‰Â«Ì… «·Œœ„… Œ’Ê„«   ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
    '                    GoTo ErrTrap
    '                End If
    '                line_no = line_no + 1
    '                 If ModAccounts.AddNewDev(LngDevID, line_no, depit_side, Abs(Round(total_value, 2)), 1, Msg + " Õ”«» Œ’Ê„«  ‰Â«Ì… «·Œœ„… ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
    '                    GoTo ErrTrap
    '                End If
            '        line_no = line_no + 1
    '        End If
              ''//////////«Ã«“«  »œÊ‰ —« »
   '    my_branch = BranchID
   '             total_value = val(TxtCash.Text)
   '              total_value = Round(total_value, 2)
   '      If total_value > 0 Then
   '       depit_side = get_account_code_branch(141, my_branch)
   '      credit_side = get_EMPLOYEE_Account(DcboEmp.BoundText, "Account_Code4")
   '
   '                 If ModAccounts.AddNewDev(LngDevID, line_no, credit_side, Abs(Round(total_value, 2)), 0, Msg + " Õ”«» ‰Â«Ì… Œœ„… «Ã«“«  »œÊ‰ —« »  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
   '                     GoTo ErrTrap
   '                 End If
   '                 line_no = line_no + 1
   '                  If ModAccounts.AddNewDev(LngDevID, line_no, depit_side, Abs(Round(total_value, 2)), 1, Msg + " Õ”«» «Ã«“«  »œÊ‰ —« »  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , val(DcboEmp.BoundText)) = False Then
   '                     GoTo ErrTrap
   '                 End If
   '                 line_no = line_no + 1
   '         End If
    

  If SystemOptions.ProjectEmployeeGV = True Then
'rs.Close
    Dim sql As String
    
    Dim Balance As Double
Dim mofradAccount As String
Dim mofradAccount1 As String
Dim Emp_id As Double
Dim Salary_account As String
 Dim Project_name As String
 Dim mofradname As String
  Dim AddOrDiscount1 As Integer
        sql = "SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.mofrad.Account_Code AS mofradAccount,  dbo.mofrad.Account_Code1 AS mofradAccount1, dbo.TblChangedComponentRegisterDetails.projectid,"
sql = sql & " dbo.Projects.Salary_account , dbo.Projects.Project_name, dbo.MOFRAD.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount"
sql = sql & " FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & "                       dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.ZmamAccount = 0) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(" & SQLDate(NoteDate, True) & " )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR(" & SQLDate(NoteDate, True) & "))"
sql = sql & " GROUP BY dbo.mofrad.Account_Code,dbo.mofrad.Account_Code1, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount"
 
    
  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText 'stop

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     mofradAccount = IIf(IsNull(rs("mofradAccount").value), "", rs("mofradAccount").value)
     mofradAccount1 = IIf(IsNull(rs("mofradAccount1").value), "", rs("mofradAccount1").value)
     
    'mofradAccount1
     
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), 0, rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
     
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & "", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    '
                     '            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , Notedate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
                     '       GoTo ErrTrap
                     '   End If
        
        
    
                             If SystemOptions.ProjectDiscountPolicy = 1 Then
                             
                                        If mofradAccount1 <> "" Then
                                        Salary_account = mofradAccount1
                                        End If
                            
                             
                             End If
                             
                                line_no = line_no + 1
                                                             If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
                        
                        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
        
            line_no = line_no + 1
        
        
         
                        
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    
 
'«·„‘«—Ì⁄ Ê·ﬂ‰ –„„
 Dim empAccount_Codezmam As String
 Dim emp_Name As String
            sql = " SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account,"
sql = sql & " dbo.projects.Project_name, dbo.mofrad.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code,"
sql = sql & " dbo.TblEmployee.emp_name , dbo.TblEmployee.Account_Code"
sql = sql & "  FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & " dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.ZmamAccount = 1) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(  " & SQLDate(NoteDate, True) & " )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR( " & SQLDate(NoteDate, True) & " ))"
sql = sql & " GROUP BY dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
sql = sql & " dbo.TblEmployee.Account_Code"
 
 
    
  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText '0000000

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     empAccount_Codezmam = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     emp_Name = IIf(IsNull(rs("emp_name").value), "", rs("emp_name").value)
     ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                                 If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If



    rs.Close
  

    
   ' Õ„Ì· «·„’—Ê›«  ⁄·Ï «·„‘«—Ì⁄
    
       sql = "SELECT      SUM(ROUND(dbo.EmpSalaryComponent.[Value] * dbo.opr_employee_details.[interval] / 30, 2)) AS Total, dbo.mofrad.Account_Code, "
sql = sql & " dbo.mofrad.AddOrDiscount, dbo.EmpSalaryComponent.EntIncresDataM, dbo.projects.Salary_account, 2006 + dbo.opr_Employee.Years AS [year],"
sql = sql & " dbo.opr_Employee.Months, SUM(dbo.opr_employee_details.[interval]) AS Intervals, dbo.opr_employee_details.ProjectID, dbo.mofrdat.mofrad_name,"
sql = sql & " dbo.Projects.Project_name , dbo.TblEmployee.BranchId"
sql = sql & " FROM         dbo.opr_employee_details INNER JOIN"
sql = sql & " dbo.projects ON dbo.opr_employee_details.ProjectID = dbo.projects.id INNER JOIN"
sql = sql & " dbo.opr_Employee ON dbo.opr_employee_details.pk_id = dbo.opr_Employee.id INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.opr_employee_details.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.EmpSalaryComponent ON dbo.opr_employee_details.Emp_id = dbo.EmpSalaryComponent.emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad INNER JOIN"
sql = sql & " dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type ON dbo.EmpSalaryComponent.AccountCode = dbo.mofrdat.mofrad_code"
sql = sql & " GROUP BY dbo.mofrad.Account_Code, dbo.EmpSalaryComponent.EntIncresDataM, dbo.projects.Salary_account, 2006 + dbo.opr_Employee.Years, dbo.opr_Employee.Months,"
sql = sql & " dbo.MOFRAD.AddOrDiscount , dbo.opr_employee_details.ProjectID, dbo.mofrdat.mofrad_name, dbo.Projects.Project_name, dbo.TblEmployee.BranchId"
sql = sql & " HAVING      (dbo.EmpSalaryComponent.EntIncresDataM IS NULL  OR"
sql = sql & "  dbo.EmpSalaryComponent.EntIncresDataM >= " & SQLDate(NoteDate, True) & " )"

sql = sql & "   AND (dbo.opr_Employee.Months = " & CmbMonth(0).ListIndex & ") AND (2006 + dbo.opr_Employee.Years = " & CboYear(0).Text & ")"


sql = sql & " ORDER BY dbo.opr_employee_details.ProjectID"

 
    
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     mofradAccount = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("mofrad_name").value), "", rs("mofrad_name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
             ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                                 If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    
    
    
    
    
    
    
    
'«·„‘«—Ì⁄ Ê·ﬂ‰ œ›⁄«  „ﬁœ„…
 'Dim empAccount_Codezmam As String
 'Dim emp_name As String
            sql = " SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account,"
sql = sql & " dbo.projects.Project_name, dbo.mofrad.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code,"
sql = sql & " dbo.TblEmployee.emp_name , dbo.TblEmployee.Account_Code3"
sql = sql & "  FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & " dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.AdvPaymentdAccount = 1) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(   " & SQLDate(NoteDate, True) & "  )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR( " & SQLDate(NoteDate, True) & "  ))"
sql = sql & " GROUP BY dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
sql = sql & " dbo.TblEmployee.Account_Code3"
 
 
    
  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     empAccount_Codezmam = IIf(IsNull(rs("Account_Code3").value), "", rs("Account_Code3").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     emp_Name = IIf(IsNull(rs("emp_name").value), "", rs("emp_name").value)
     ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
     
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                                 If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    

End If


'«· √„Ì‰« 


'    rs.Close
    
       
       sql = " "

'
 

' project gv

'    Create_dev2 = True
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
'    Create_dev2 = False
  
'********************************************************************
  End Function

  
Function GetComponentValuePerBranch(BramchId As Integer, componentname As String) As Double
    Dim SUM As Double
    SUM = 0
    Dim i As Integer

    With Grid

        For i = .FixedRows To .Rows - 2
    
            If val(.TextMatrix(i, .ColIndex(componentname))) > 0 And val(.TextMatrix(i, .ColIndex("BranchId"))) = BramchId Then
                SUM = SUM + val(.TextMatrix(i, .ColIndex(componentname)))
            End If

        Next i

    End With

    GetComponentValuePerBranch = SUM
End Function
Sub EmptyTxet()
DcboEmp.BoundText = 0
ChEndServ.value = vbUnchecked
ChSalar.value = vbUnchecked
ChValTekt.value = vbUnchecked
ChCustom.value = vbUnchecked
ChCusTiket.value = vbUnchecked
ChAddOther.value = vbUnchecked
ChAdvanceTotal.value = vbUnchecked
ChCash.value = vbUnchecked
AlowAssest.value = vbUnchecked
 Grid.Clear flexClearScrollable, flexClearEverything
 Fg.Clear flexClearScrollable, flexClearEverything
 VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
 text1.Text = ""
 txttotal.Text = 0
 TXTLastTotal.Text = 0
 txtnet.Text = 0
 TxtTotalDis.Text = 0
End Sub


Private Sub ChangeLang()
    XPLbl(28).Caption = "Difference"
    XPLbl(30).Caption = "Difference"
    XPLbl(29).Caption = "Total Assets"
    XPLbl(21).Caption = "Day"
    XPLbl(31).Caption = "Others Discount"
    XPLbl(17).Caption = "Tickets from Cont."
    XPLbl(33).Caption = "Salary"
    XPLbl(39).Caption = "Month Salary"
    chkGE.Caption = "GE By trans Date"
    
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    
    lblbr.Caption = "Branch"
    XPLbl(27).Caption = "Custom Tickets"
    XPLbl(26).Caption = "Custom End Service"
    Label6(0).Caption = "Duration"
    lbl(6).Caption = "Current Record"
    lbl(7).Caption = "No. Recordes"
    AlowAssest.RightToLeft = False
    Label1(35).Caption = "No.GL"
    AlowAssest.Caption = "Discount  Value of Covenant"
    Command8.Caption = "Acc.Statement"
    XPLbl(20).Caption = "Total Discount"
    Command5.Caption = "Create JE"
    XPLbl(25).Caption = "Net"
    XPLbl(23).Caption = "Average"
    XPLbl(22).Caption = "Order No"
    'Me.Label1.Caption = "Duration"
    'Frame9.Caption = "Accounting"
    Command9.Caption = "Print GL"
    Command7.Caption = "Delete GL"
    Comman1122.Caption = "Show"
    XPLbl(14).Caption = "Covenant"
    Me.Label7.Caption = "Leave without pay"
    Label10.Caption = "Day"
    Label13.Caption = "Month"
    Label12.Caption = "Year"
    Label11.Caption = "Day"
    ISButton2.Caption = "Calculate"
    Label15.Caption = "Month"
    Label14.Caption = "Year"
    XPLbl(14).Caption = "Select"
    Label8.Caption = "Select"
    XPLbl(15).Caption = "Added"
    XPLbl(16).Caption = "Discount"
    ALLButton1.Caption = "Reason"
    lbl(5).Caption = "Job"
    XPLbl(19).Caption = "Curr Salary"
    XPLbl(32).Caption = "Ticket Value"
    XPLbl(18).Caption = "Vacation Value"
    CmdExit.Caption = "Hide"
    XPLbl(0).Caption = "Basic"
    XPLbl(34).Caption = "Housing"
    XPLbl(35).Caption = "Transportation"
    XPLbl(38).Caption = "Others"
    XPLbl(37).Caption = "Food"
    XPLbl(13).Caption = "Total vacations without pay"
    XPLbl(34).Caption = "Housing"
    XPLbl(36).Caption = "Mobile"
    XPLbl(45).Caption = "Supervision"
    XPLbl(4).Caption = "Total"
    lbl(10).Caption = "Total"
    lbl(9).Caption = "Start Date"
    lbl(0).Caption = "End Date"
    'Label6.Caption = " Interval"
    Label2.Caption = " Days"
    Label3.Caption = " Months"
    Label4.Caption = " Years"
    XPLbl(9).Caption = "Des'"
    XPLbl(5).Caption = "Total"
    XPLbl(6).Caption = "value"
    XPLbl(7).Caption = "ReSult"
    XPLbl(10).Caption = "Advance"
    XPLbl(11).Caption = "Add Other"
    XPLbl(12).Caption = "Net"
    cmd_CALC_NET.Caption = "Calc Nets"
    
    Me.Caption = "End Of Service"
    Ele(0).Caption = Me.Caption
    
    XPLbl(3).Caption = "OPR#"
    XPLbl(1).Caption = "Emp Code"
    XPLbl(2).Caption = "Emp Name"
    lbl(1).Caption = "Date"
    lbl(2).Caption = "Type"
    lbl(3).Caption = "By"
    lbl(2).Caption = "Type"
    XPLbl(8).Caption = "Salaries"
    Check1.Caption = "Select All"
    lbl(4).Caption = "Remarks"
    Label9.Caption = "No.Days absence"
    Me.Cmd(0).Caption = "&New"
    Me.Cmd(1).Caption = "&Edit"
    Me.Cmd(2).Caption = "&Save"
    Me.Cmd(3).Caption = "&Undo"
    Me.Cmd(4).Caption = "&Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "E&xit"
    Me.Cmd(7).Caption = "&Print"
    Me.CmdHelp.Caption = "&Help"

    With Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("selected")) = "Select"
        .TextMatrix(0, .ColIndex("mofrd")) = "MofrdName"
        .TextMatrix(0, .ColIndex("value")) = "Value"
    End With

    With VSFlexGrid1
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("selected")) = "Select"
        .TextMatrix(0, .ColIndex("mofrd")) = "MofrdName"
        .TextMatrix(0, .ColIndex("value")) = "Value"
    End With

    With Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("Emp_Code")) = "Code"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
    End With
       
    With Me.VSFlexGrid2
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("AsCode")) = "No"
        .TextMatrix(0, .ColIndex("mofrd")) = "Name"
        .TextMatrix(0, .ColIndex("ReciveDate")) = "ReciveDate"
        .TextMatrix(0, .ColIndex("DeliverDate")) = "DeliverDate"
        .TextMatrix(0, .ColIndex("Emp_NameTo")) = "Recipient Name"""
    End With
    
    With GRID2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level"
        .TextMatrix(0, .ColIndex("EmpName")) = "Employee"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approve Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Notes"
    End With
End Sub
Function cal_intervalOptional()
If Me.TxtModFlg.Text <> "R" Then
    Dim astrSplitItems() As String
    Dim Result As String
    Dim total_year_m As Double
    Dim total_month_m As Double
    Dim total_day_m As Double
    Dim Total As Double
    Dim interval1 As Double
    Dim interval2 As Double

    Dim diff_year As Integer
    Result = ExactAge(date1, date2)

    If Result = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " «—ÌŒ »œ«Ì… «·⁄„· ÂÊ ‰›”  «—ÌŒ «·⁄„· «Ê «ﬂ»— „‰Â ﬁ„ » €ÌÌ— ﬁÌ„ «· «—ÌŒ ", vbCritical
        Else
            MsgBox "Date of start of work is the same date as the end of the work ", vbCritical
        End If
    End If

    astrSplitItems = Split(Result, "-")
    txtyear = astrSplitItems(0)
    txtmonth = astrSplitItems(1)
    txtday = astrSplitItems(2)
    text1.Text = ""

    If txtyear < 5 Then
        total_year_m = (0.5 * month_salary) * txtyear
        total_month_m = ((0.5 * month_salary) / 12) * txtmonth
        total_day_m = ((0.5 * month_salary) / 365) * txtday
        Total = total_year_m + total_month_m + total_day_m

        If SystemOptions.UserInterface = ArabicInterface Then

            text1.Text = "⁄œœ «·”‰Ê«  «ﬁ· „‰ 5 Ì „ «Õ ”«» ‰’› ‘Â— ⁄‰ ﬂ·  ”‰…" & vbNewLine
            text1.Text = text1.Text + "ﬁÌ„… ‰’› «·‘Â—  ”«ÊÌ  " & (0.5 * month_salary) & vbNewLine
   
            text1.Text = text1.Text + "«Ã„«·Ì ”‰Ê«  «·Œœ„…  " & txtyear & "  ”‰…  " & "*" & Round((0.5 * month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
            text1.Text = text1.Text + "«Ã„«·Ì ‘ÂÊ— «·Œœ„…  " & txtmonth & "  ‘Â—  " & "*" & Round(((0.5 * month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
            text1.Text = text1.Text + "«Ã„«·Ì «Ì«„ «·Œœ„…  " & txtday & "  ÌÊ„  " & "*" & Round(((0.5 * month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
        Else
  
            text1.Text = " Less than 5 years is calculated half a month for each year " & vbNewLine
            text1.Text = text1.Text + " Value equal to a half months  " & (0.5 * month_salary) & vbNewLine
   
            text1.Text = text1.Text + "  Years   " & txtyear & "  Years  " & "*" & Round((0.5 * month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
            text1.Text = text1.Text + " Months  " & txtmonth & "  Months  " & "*" & Round(((0.5 * month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
            text1.Text = text1.Text + " Days  " & txtday & "  Days  " & "*" & Round(((0.5 * month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
        End If
    
    ElseIf txtyear >= 5 Then
    
        If txtyear = 5 Then
                 
            total_year_m = (0.5 * month_salary) * 5
            total_month_m = (month_salary / 12) * txtmonth
            total_day_m = ((month_salary) / 365) * txtday

            If SystemOptions.UserInterface = ArabicInterface Then

                text1.Text = text1.Text + "⁄œœ «·”‰Ê«   5    ”‰Ê«    Ì „ «Õ ”«» ‰’› ‘Â— ⁄‰ ﬂ·Â ”‰… Ê‘Â— ⁄‰ «·«‘Â— «·“Ì«œÂ ⁄‰ «·Œ„” ”‰Ê« " & vbNewLine
                text1.Text = text1.Text + "ﬁÌ„…  «·‘Â—  ”«ÊÌ  " & (month_salary) & vbNewLine
                text1.Text = text1.Text + "ﬁÌ„… ‰’› «·‘Â—  ”«ÊÌ  " & (0.5 * month_salary) & vbNewLine
                    
                text1.Text = text1.Text + "«Ã„«·Ì ”‰Ê«  «·Œœ„…  " & txtyear & "  ”‰…  " & "*" & Round((0.5 * month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì ‘ÂÊ— «·Œœ„…  " & txtmonth & "  ‘Â—  " & "*" & Round(((month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì «Ì«„ «·Œœ„…  " & txtday & "  ÌÊ„  " & "*" & Round(((month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
            Else
                    
                text1.Text = text1.Text + " The number of years greater than 5 a month is calculated for each year" & vbNewLine
                text1.Text = text1.Text + "  Month Value =   " & (month_salary) & vbNewLine
                text1.Text = text1.Text + "Value equal to a half months =  " & (0.5 * month_salary) & vbNewLine
                    
                text1.Text = text1.Text + " Years  " & txtyear & "  Years  " & "*" & Round((0.5 * month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
                text1.Text = text1.Text + " Months  " & txtmonth & "  Years  " & "*" & Round(((month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
                text1.Text = text1.Text + " Days " & txtday & "  Days  " & "*" & Round(((month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
                     
            End If

            Total = total_year_m + total_month_m + total_day_m
        ElseIf txtyear.Text > 5 Then
            diff_year = txtyear - 5
            interval1 = (0.5 * month_salary) * 5
            interval2 = (month_salary) * diff_year
            total_year_m = interval1 + interval2
            total_month_m = (month_salary / 12) * txtmonth
            total_day_m = ((month_salary) / 365) * txtday
            Total = total_year_m + total_month_m + total_day_m

            If SystemOptions.UserInterface = ArabicInterface Then
                 
                text1.Text = text1.Text + "⁄œœ «·”‰Ê«   «ﬂ»— „‰ 5   ”‰Ê«    " & vbNewLine
                text1.Text = text1.Text + "ﬁÌ„…  «·‘Â—  ”«ÊÌ  " & (month_salary) & vbNewLine
                text1.Text = text1.Text + "ﬁÌ„… ‰’› «·‘Â—  ”«ÊÌ  " & (0.5 * month_salary) & vbNewLine
                text1.Text = text1.Text + "Õ”«» «Ê· 5 ”‰Ê«  " & vbNewLine
                    
                text1.Text = text1.Text + "«Ã„«·Ì «Ê· 5 ”‰Ê«   = 5  ”‰…  " & "*" & Round((0.5 * month_salary), 2) & "=" & Round(interval1, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì »«ﬁÌ «·„œÂ" & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì ”‰Ê«  «·Œœ„… «·„ »ﬁÌ…  " & diff_year & "  ”‰…  " & "*" & Round((month_salary), 2) & "=" & Round(interval2, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì ‘ÂÊ— «·Œœ„… «·„ »ﬁÌ… " & txtmonth & "  ‘Â—  " & "*" & Round(((month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì «Ì«„ «·Œœ„… «·„ »ﬁÌ… " & txtday & "  ÌÊ„  " & "*" & Round(((month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
            Else
                 
                text1.Text = text1.Text + " The number of years greater than 5  " & vbNewLine
                text1.Text = text1.Text + " Month Value=  " & (month_salary) & vbNewLine
                text1.Text = text1.Text + "Value equal to a half months   " & (0.5 * month_salary) & vbNewLine
                text1.Text = text1.Text + " First 5 Years  " & vbNewLine
                    
                text1.Text = text1.Text + "First 5 Years Total= 5 Years  " & "*" & Round((0.5 * month_salary), 2) & "=" & Round(interval1, 2) & vbNewLine
                text1.Text = text1.Text + "Total rest of the term " & vbNewLine
                text1.Text = text1.Text + " Total Years of service remaining  " & diff_year & "  Months  " & "*" & Round((month_salary), 2) & "=" & Round(interval2, 2) & vbNewLine
                text1.Text = text1.Text + " Total Months of service remaining " & txtmonth & "  Months  " & "*" & Round(((month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
                text1.Text = text1.Text + " Total Days of service remaining  " & txtday & "  Days  " & "*" & Round(((month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
            End If
                
        End If

    
    End If
 
    '  month_salary = total
    '  day_salary total / 30
    If Not IsNumeric(txtnum.Text) Then txtnum.Text = 0
    If Not IsNumeric(TxtNetEnd.Text) Then TxtNetEnd.Text = 0
    Me.txttotal.Text = Round(Total, 2)
    
End If
End Function
Function cal_interval()
If Me.TxtModFlg.Text <> "R" Then
    Dim astrSplitItems() As String
    Dim Result As String
    Dim total_year_m As Double
    Dim total_month_m As Double
    Dim total_day_m As Double
    Dim Total As Double
    Dim interval1 As Double
    Dim interval2 As Double

    Dim diff_year As Integer
    Result = ExactAge(date1, date2)

    If Result = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " «—ÌŒ »œ«Ì… «·⁄„· ÂÊ ‰›”  «—ÌŒ «·⁄„· «Ê «ﬂ»— „‰Â ﬁ„ » €ÌÌ— ﬁÌ„ «· «—ÌŒ ", vbCritical
        Else
            MsgBox "Date of start of work is the same date as the end of the work ", vbCritical
        End If
    End If

    astrSplitItems = Split(Result, "-")
    txtyear = astrSplitItems(0)
    txtmonth = astrSplitItems(1)
    txtday = astrSplitItems(2)
    text1.Text = ""

    If txtyear < 5 Then
        total_year_m = (0.5 * month_salary) * txtyear
        total_month_m = ((0.5 * month_salary) / 12) * txtmonth
        total_day_m = ((0.5 * month_salary) / 365) * txtday
        Total = total_year_m + total_month_m + total_day_m

        If SystemOptions.UserInterface = ArabicInterface Then

            text1.Text = "⁄œœ «·”‰Ê«  «ﬁ· „‰ 5 Ì „ «Õ ”«» ‰’› ‘Â— ⁄‰ ﬂ·  ”‰…" & vbNewLine
            text1.Text = text1.Text + "ﬁÌ„… ‰’› «·‘Â—  ”«ÊÌ  " & (0.5 * month_salary) & vbNewLine
   
            text1.Text = text1.Text + "«Ã„«·Ì ”‰Ê«  «·Œœ„…  " & txtyear & "  ”‰…  " & "*" & Round((0.5 * month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
            text1.Text = text1.Text + "«Ã„«·Ì ‘ÂÊ— «·Œœ„…  " & txtmonth & "  ‘Â—  " & "*" & Round(((0.5 * month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
            text1.Text = text1.Text + "«Ã„«·Ì «Ì«„ «·Œœ„…  " & txtday & "  ÌÊ„  " & "*" & Round(((0.5 * month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
        Else
  
            text1.Text = " Less than 5 years is calculated half a month for each year " & vbNewLine
            text1.Text = text1.Text + " Value equal to a half months  " & (0.5 * month_salary) & vbNewLine
   
            text1.Text = text1.Text + "  Years   " & txtyear & "  Years  " & "*" & Round((0.5 * month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
            text1.Text = text1.Text + " Months  " & txtmonth & "  Months  " & "*" & Round(((0.5 * month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
            text1.Text = text1.Text + " Days  " & txtday & "  Days  " & "*" & Round(((0.5 * month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
        End If
    
    ElseIf txtyear >= 5 And txtyear < 10 Then
    
        If txtyear = 5 Then
                 
            total_year_m = (0.5 * month_salary) * 5
            total_month_m = (month_salary / 12) * txtmonth
            total_day_m = ((month_salary) / 365) * txtday

            If SystemOptions.UserInterface = ArabicInterface Then

                text1.Text = text1.Text + "⁄œœ «·”‰Ê«   5    ”‰Ê«    Ì „ «Õ ”«» ‰’› ‘Â— ⁄‰ ﬂ·Â ”‰… Ê‘Â— ⁄‰ «·«‘Â— «·“Ì«œÂ ⁄‰ «·Œ„” ”‰Ê« " & vbNewLine
                text1.Text = text1.Text + "ﬁÌ„…  «·‘Â—  ”«ÊÌ  " & (month_salary) & vbNewLine
                text1.Text = text1.Text + "ﬁÌ„… ‰’› «·‘Â—  ”«ÊÌ  " & (0.5 * month_salary) & vbNewLine
                    
                text1.Text = text1.Text + "«Ã„«·Ì ”‰Ê«  «·Œœ„…  " & txtyear & "  ”‰…  " & "*" & Round((0.5 * month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì ‘ÂÊ— «·Œœ„…  " & txtmonth & "  ‘Â—  " & "*" & Round(((month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì «Ì«„ «·Œœ„…  " & txtday & "  ÌÊ„  " & "*" & Round(((month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
            Else
                    
                text1.Text = text1.Text + " The number of years greater than 5 a month is calculated for each year" & vbNewLine
                text1.Text = text1.Text + "  Month Value =   " & (month_salary) & vbNewLine
                text1.Text = text1.Text + "Value equal to a half months =  " & (0.5 * month_salary) & vbNewLine
                    
                text1.Text = text1.Text + " Years  " & txtyear & "  Years  " & "*" & Round((0.5 * month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
                text1.Text = text1.Text + " Months  " & txtmonth & "  Years  " & "*" & Round(((month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
                text1.Text = text1.Text + " Days " & txtday & "  Days  " & "*" & Round(((month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
                     
            End If

            Total = total_year_m + total_month_m + total_day_m
        ElseIf txtyear.Text > 5 Then
            diff_year = txtyear - 5
            interval1 = (0.5 * month_salary) * 5
            interval2 = (month_salary) * diff_year
            total_year_m = interval1 + interval2
            total_month_m = (month_salary / 12) * txtmonth
            total_day_m = ((month_salary) / 365) * txtday
            Total = total_year_m + total_month_m + total_day_m

            If SystemOptions.UserInterface = ArabicInterface Then
                 
                text1.Text = text1.Text + "⁄œœ «·”‰Ê«   «ﬂ»— „‰ 5   ”‰Ê«  Ê«ﬁ· „‰ 10 ”‰Ê«   " & vbNewLine
                text1.Text = text1.Text + "ﬁÌ„…  «·‘Â—  ”«ÊÌ  " & (month_salary) & vbNewLine
                text1.Text = text1.Text + "ﬁÌ„… ‰’› «·‘Â—  ”«ÊÌ  " & (0.5 * month_salary) & vbNewLine
                text1.Text = text1.Text + "Õ”«» «Ê· 5 ”‰Ê«  " & vbNewLine
                    
                text1.Text = text1.Text + "«Ã„«·Ì «Ê· 5 ”‰Ê«   = 5  ”‰…  " & "*" & Round((0.5 * month_salary), 2) & "=" & Round(interval1, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì »«ﬁÌ «·„œÂ" & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì ”‰Ê«  «·Œœ„… «·„ »ﬁÌ…  " & diff_year & "  ”‰…  " & "*" & Round((month_salary), 2) & "=" & Round(interval2, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì ‘ÂÊ— «·Œœ„… «·„ »ﬁÌ… " & txtmonth & "  ‘Â—  " & "*" & Round(((month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì «Ì«„ «·Œœ„… «·„ »ﬁÌ… " & txtday & "  ÌÊ„  " & "*" & Round(((month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
            Else
                 
                text1.Text = text1.Text + " The number of years greater than 5 and less than 10 " & vbNewLine
                text1.Text = text1.Text + " Month Value=  " & (month_salary) & vbNewLine
                text1.Text = text1.Text + "Value equal to a half months   " & (0.5 * month_salary) & vbNewLine
                text1.Text = text1.Text + " First 5 Years  " & vbNewLine
                    
                text1.Text = text1.Text + "First 5 Years Total= 5 Years  " & "*" & Round((0.5 * month_salary), 2) & "=" & Round(interval1, 2) & vbNewLine
                text1.Text = text1.Text + "Total rest of the term " & vbNewLine
                text1.Text = text1.Text + " Total Years of service remaining  " & diff_year & "  Months  " & "*" & Round((month_salary), 2) & "=" & Round(interval2, 2) & vbNewLine
                text1.Text = text1.Text + " Total Months of service remaining " & txtmonth & "  Months  " & "*" & Round(((month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
                text1.Text = text1.Text + " Total Days of service remaining  " & txtday & "  Days  " & "*" & Round(((month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
            End If
                
        End If
    
    ElseIf txtyear >= 10 Then
        total_year_m = (month_salary) * txtyear
        total_month_m = (month_salary / 12) * txtmonth
        total_day_m = ((month_salary) / 365) * txtday
        Total = total_year_m + total_month_m + total_day_m
   
        If SystemOptions.UserInterface = ArabicInterface Then
                 
            text1.Text = "⁄œœ «·”‰Ê«  «ﬂ»— „‰ 10 Ì „ «Õ ”«»   ‘Â— ⁄‰ ﬂ·  ”‰…" & vbNewLine
            text1.Text = text1.Text + "ﬁÌ„…   «·‘Â—  ”«ÊÌ  " & (month_salary) & vbNewLine
   
            text1.Text = text1.Text + "«Ã„«·Ì ”‰Ê«  «·Œœ„…  " & txtyear & "  ”‰…  " & "*" & Round((month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
            text1.Text = text1.Text + "«Ã„«·Ì ‘ÂÊ— «·Œœ„…  " & txtmonth & "  ‘Â—  " & "*" & Round(((month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
            text1.Text = text1.Text + "«Ã„«·Ì «Ì«„ «·Œœ„…  " & txtday & "  ÌÊ„  " & "*" & Round(((month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
   
        Else
            text1.Text = " The number of years greater than 10 a month is calculated for each year " & vbNewLine
            text1.Text = text1.Text + " Month Vakue=  " & (month_salary) & vbNewLine
   
            text1.Text = text1.Text + " Years Total  " & txtyear & "  Years  " & "*" & Round((month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
            text1.Text = text1.Text + "Months Total  " & txtmonth & "  Months  " & "*" & Round(((month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
            text1.Text = text1.Text + " Dyas Total " & txtday & "  Dyas  " & "*" & Round(((month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
    
        End If
    
    End If
 
    '  month_salary = total
    '  day_salary total / 30
    If Not IsNumeric(txtnum.Text) Then txtnum.Text = 0
    If Not IsNumeric(TxtNetEnd.Text) Then TxtNetEnd.Text = 0
    Me.txttotal.Text = Round(Total, 2)
    
End If
End Function
Sub Calcul()
If Me.TxtModFlg.Text <> "R" Then
  If val(txtTicketValue.Text) <> 0 Then
  TxtDiffTekit.Text = Round(val(txtTicketValue.Text) - val(TxtCusTiket.Text), 2)
  Else
  TxtDiffTekit.Text = 0
  End If
 'If val(TxtNetEnd.Text) <> 0 Then
  TxtDiffEnd.Text = Round(val(TxtNetEnd.Text) - val(TxtEndService.Text), 2)
 ' Else
'  TxtDiffEnd.Text = 0
 ' End If
  
TxtNetEnd.Text = Round(val(txttotal.Text) * val(TxtRate.Text), 2)
txtnet.Text = 0
'If ChSalar.value = vbChecked Then
txtnet.Text = val(txtnet.Text) + val(txtSal.Text)
'End If
'If ChValTekt.value = vbChecked Then
txtnet.Text = val(txtnet.Text) + val(txtTicketValue.Text)
'End If
'If ChCustom.value = vbChecked Then
txtnet.Text = val(txtnet.Text) + val(txtCustom.Text)
'txtnet.Text = val(txtnet.Text) + val(txtCustom.Text) + val(TxtDiffTekit.Text)
'End If
'If ChCusTiket.value = vbChecked Then
'''''''''''''''''''txtnet.text = val(txtnet.text) + val(TxtCusTiket.text)
'End If
'If ChAddOther.value = vbChecked Then
txtnet.Text = val(txtnet.Text) + val(TxtAddOther.Text)
'End If
txtnet.Text = Round(val(txtnet.Text), 2)
TxtTotalDis.Text = 0
'If ChAdvanceTotal.value = vbChecked Then
TxtTotalDis.Text = val(TxtTotalDis.Text) + val(TXTAdvanceTotal.Text)
'End If
'If ChCash.value = vbChecked Then
TxtTotalDis.Text = val(TxtTotalDis.Text) + val(TxtCash.Text)
TxtTotalDis.Text = val(TxtTotalDis.Text) + val(TxtDiscounts.Text)
'End If
'If AlowAssest.value = vbChecked Then
TxtTotalDis.Text = val(TxtTotalDis.Text) + val(TxtVlueVaction.Text) + val(TxtDisSalary.Text)
TxtTotalDis.Text = Round(val(TxtTotalDis.Text), 2)
'End If
TXTLastTotal.Text = val(txtnet.Text) + val(TxtNetEnd.Text) - val(TxtTotalDis.Text)
'If ChEndServ.value = Checked Then
''''''''''''''TXTLastTotal.text = val(TXTLastTotal.text) + val(TxtEndService.text)
'End If
    Select Case OPR.Caption
        Case "+"
            TXTLastTotal.Text = val(TXTLastTotal.Text) + val(txtnum.Text)
        Case "-"
            TXTLastTotal.Text = val(TXTLastTotal.Text) - val(txtnum.Text)
        Case "*"
            TXTLastTotal.Text = val(txtnum.Text) * val(txttotal.Text) + val(TXTLastTotal.Text) - val(TxtNetEnd.Text)

        Case "/"

            If txtnum.Text = 0 Then
                '    MsgBox "·« Ì„ﬂ‰ «·ﬁ”„… ⁄·Ï ’›—", vbCritical
               TXTLastTotal.Text = val(TXTLastTotal.Text) + val(txtnum.Text)
            Else
                TXTLastTotal.Text = (val(txttotal.Text) / val(txtnum.Text)) + val(TXTLastTotal.Text) - val(TxtNetEnd.Text)
            End If

    End Select

    'txttotal.text = Format(total, SystemOptions.SysDefCurrencyForamt)
    TXTLastTotal.Text = Round(val(TXTLastTotal.Text), 2)
    
    End If
   
End Sub
Function cal_intervaldeparture(Optional X As Integer)
If Me.TxtModFlg.Text <> "R" Then
    Dim astrSplitItems() As String
    Dim Result As String
    Dim total_year_m As Double
    Dim total_month_m As Double
    Dim total_day_m As Double
    Dim Total As Double
    Dim interval1 As Double
    Dim interval2 As Double

    Dim diff_year As Integer
    Result = ExactAge(date1, date2)

    If Result = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " «—ÌŒ »œ«Ì… «·⁄„· ÂÊ ‰›”  «—ÌŒ «·⁄„· «Ê «ﬂ»— „‰Â ﬁ„ » €ÌÌ— ﬁÌ„ «· «—ÌŒ ", vbCritical
        Else
            MsgBox "Date of start of work is the same date as the end of the work ", vbCritical
        End If
    End If

    astrSplitItems = Split(Result, "-")
    txtyear = astrSplitItems(0)
    txtmonth = astrSplitItems(1)
    txtday = astrSplitItems(2)
    text1.Text = ""

    If txtyear < 5 And txtyear >= 2 And (txtmonth > 0 Or txtday > 0) Then
        total_year_m = ((1 / 3) * month_salary) * txtyear
        total_month_m = (((1 / 3) * month_salary) / 12) * txtmonth
        total_day_m = (((1 / 3) * month_salary) / 365) * txtday
        Total = total_year_m + total_month_m + total_day_m

        If SystemOptions.UserInterface = ArabicInterface Then

            text1.Text = "⁄œœ «·”‰Ê«  «ﬁ· „‰ 5 Ì „ «Õ ”«» À·À ‘Â— ⁄‰ ﬂ·  ”‰…" & vbNewLine
            text1.Text = text1.Text + "ﬁÌ„… À·À «·‘Â—  ”«ÊÌ  " & ((1 / 3) * month_salary) & vbNewLine
   
            text1.Text = text1.Text + "«Ã„«·Ì ”‰Ê«  «·Œœ„…  " & txtyear & "  ”‰…  " & "*" & Round(((1 / 3) * month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
            text1.Text = text1.Text + "«Ã„«·Ì ‘ÂÊ— «·Œœ„…  " & txtmonth & "  ‘Â—  " & "*" & Round((((1 / 3) * month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
            text1.Text = text1.Text + "«Ã„«·Ì «Ì«„ «·Œœ„…  " & txtday & "  ÌÊ„  " & "*" & Round((((1 / 3) * month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
        Else
  
            text1.Text = " Less than 5 years is calculated One-third a month for each year " & vbNewLine
            text1.Text = text1.Text + " Value equal to a One-third months  " & ((1 / 3) * month_salary) & vbNewLine
   
            text1.Text = text1.Text + "  Years   " & txtyear & "  Years  " & "*" & Round(((1 / 3) * month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
            text1.Text = text1.Text + " Months  " & txtmonth & "  Months  " & "*" & Round((((1 / 3) * month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
            text1.Text = text1.Text + " Days  " & txtday & "  Days  " & "*" & Round((((1 / 3) * month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
        End If
    
    ElseIf txtyear >= 5 And txtyear < 10 Then
    
        If txtyear = 5 Then
                 
            total_year_m = ((1 / 3) * month_salary) * 5
            total_month_m = ((2 / 3) * month_salary / 12) * txtmonth
            total_day_m = ((2 / 3) * (month_salary) / 365) * txtday

            If SystemOptions.UserInterface = ArabicInterface Then

                text1.Text = text1.Text + "⁄œœ «·”‰Ê«   5    ”‰Ê«    Ì „ «Õ ”«»  À·À ‘Â— ⁄‰ ﬂ·Â ”‰… ÊÀ·ÀÌ‰  ⁄‰ «·«‘Â— «·“Ì«œÂ ⁄‰ «·Œ„” ”‰Ê« " & vbNewLine
                text1.Text = text1.Text + "ﬁÌ„…  «·‘Â—  ”«ÊÌ  " & (month_salary) & vbNewLine
                text1.Text = text1.Text + "ﬁÌ„… À·À «·‘Â—  ”«ÊÌ  " & ((1 / 3) * month_salary) & vbNewLine
                    
                text1.Text = text1.Text + "«Ã„«·Ì ”‰Ê«  «·Œœ„…  " & txtyear & "  ”‰…  " & "*" & Round(((1 / 3) * month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì ‘ÂÊ— «·Œœ„…  " & txtmonth & "  ‘Â—  " & "*" & Round((((2 / 3) * month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì «Ì«„ «·Œœ„…  " & txtday & "  ÌÊ„  " & "*" & Round((((2 / 3) * month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
            Else
                    
                text1.Text = text1.Text + " The number of years greater than 5 a month is calculated for each year" & vbNewLine
                text1.Text = text1.Text + "  Month Value =   " & (month_salary) & vbNewLine
                text1.Text = text1.Text + "Value equal to a One-third months =  " & ((1 / 3) * month_salary) & vbNewLine
                    
                text1.Text = text1.Text + " Years  " & txtyear & "  Years  " & "*" & Round(((1 / 3) * month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
                text1.Text = text1.Text + " Months  " & txtmonth & "  Years  " & "*" & Round((((2 / 3) * month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
                text1.Text = text1.Text + " Days " & txtday & "  Days  " & "*" & Round((((2 / 3) * month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
                     
            End If
            Total = total_year_m + total_month_m + total_day_m
        ElseIf txtyear > 5 Then
            diff_year = txtyear - 5
            interval1 = ((1 / 3) * month_salary) * 5
            interval2 = ((2 / 3) * month_salary) * diff_year
            total_year_m = interval1 + interval2
            total_month_m = ((2 / 3) * month_salary / 12) * txtmonth
            total_day_m = (((2 / 3) * month_salary) / 365) * txtday
            Total = total_year_m + total_month_m + total_day_m

            If SystemOptions.UserInterface = ArabicInterface Then
                 
                text1.Text = text1.Text + "⁄œœ «·”‰Ê«   «ﬂ»— „‰ 5   ”‰Ê«  Ê«ﬁ· „‰ 10 ”‰Ê«   " & vbNewLine
                text1.Text = text1.Text + "ﬁÌ„…  «·‘Â—  ”«ÊÌ  " & (month_salary) & vbNewLine
                text1.Text = text1.Text + "ﬁÌ„… À·ÀÌ «·‘Â—  ”«ÊÌ  " & ((2 / 3) * month_salary) & vbNewLine
                text1.Text = text1.Text + "Õ”«» «Ê· 5 ”‰Ê«  " & vbNewLine
                    
                text1.Text = text1.Text + "«Ã„«·Ì «Ê· 5 ”‰Ê«   = 5  ”‰…  " & "*" & Round(((1 / 3) * month_salary), 2) & "=" & Round(interval1, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì »«ﬁÌ «·„œÂ" & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì ”‰Ê«  «·Œœ„… «·„ »ﬁÌ…  " & diff_year & "  ”‰…  " & "*" & Round(((2 / 3) * month_salary), 2) & "=" & Round(interval2, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì ‘ÂÊ— «·Œœ„… «·„ »ﬁÌ… " & txtmonth & "  ‘Â—  " & "*" & Round((((2 / 3) * month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì «Ì«„ «·Œœ„… «·„ »ﬁÌ… " & txtday & "  ÌÊ„  " & "*" & Round((((2 / 3) * month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
            Else
                 
                text1.Text = text1.Text + " The number of years greater than 5 and less than 10 " & vbNewLine
                text1.Text = text1.Text + " Month Value=  " & (month_salary) & vbNewLine
                text1.Text = text1.Text + "Value equal to a half months   " & ((1 / 3) * month_salary) & vbNewLine
                text1.Text = text1.Text + " First 5 Years  " & vbNewLine
                    
                text1.Text = text1.Text + "First 5 Years Total= 5 Years  " & "*" & Round(((2 / 3) * month_salary), 2) & "=" & Round(interval1, 2) & vbNewLine
                text1.Text = text1.Text + "Total rest of the term " & vbNewLine
                text1.Text = text1.Text + " Total Years of service remaining  " & diff_year & "  Months  " & "*" & Round(((2 / 3) * month_salary), 2) & "=" & Round(interval2, 2) & vbNewLine
                text1.Text = text1.Text + " Total Months of service remaining " & txtmonth & "  Months  " & "*" & Round((((2 / 3) * month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
                text1.Text = text1.Text + " Total Days of service remaining  " & txtday & "  Days  " & "*" & Round((((2 / 3) * month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
            End If
End If

   ElseIf txtyear = 10 Then
                 
            total_year_m = ((2 / 3) * month_salary) * 5
            total_month_m = (month_salary / 12) * txtmonth
            total_day_m = ((month_salary) / 365) * txtday

            If SystemOptions.UserInterface = ArabicInterface Then

                text1.Text = text1.Text + "⁄œœ «·”‰Ê«   10    ”‰Ê«    Ì „ «Õ ”«» À·ÀÌ ‘Â— ⁄‰ ﬂ·Â ”‰… Ê‘Â— ⁄‰ «·«‘Â— «·“Ì«œÂ ⁄‰ «·Œ„” ”‰Ê« " & vbNewLine
                text1.Text = text1.Text + "ﬁÌ„…  «·‘Â—  ”«ÊÌ  " & (month_salary) & vbNewLine
                text1.Text = text1.Text + "ﬁÌ„… À·ÀÌ «·‘Â—  ”«ÊÌ  " & ((2 / 3) * month_salary) & vbNewLine
                    
                text1.Text = text1.Text + "«Ã„«·Ì ”‰Ê«  «·Œœ„…  " & txtyear & "  ”‰…  " & "*" & Round(((2 / 3) * month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì ‘ÂÊ— «·Œœ„…  " & txtmonth & "  ‘Â—  " & "*" & Round(((month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
                text1.Text = text1.Text + "«Ã„«·Ì «Ì«„ «·Œœ„…  " & txtday & "  ÌÊ„  " & "*" & Round(((month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
            Else
                    
                text1.Text = text1.Text + " The number of years greater than 5 a month is calculated for each year" & vbNewLine
                text1.Text = text1.Text + "  Month Value =   " & (month_salary) & vbNewLine
                text1.Text = text1.Text + "Value equal to a half months =  " & ((2 / 3) * month_salary) & vbNewLine
                    
                text1.Text = text1.Text + " Years  " & txtyear & "  Years  " & "*" & Round(((2 / 3) * month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
                text1.Text = text1.Text + " Months  " & txtmonth & "  Years  " & "*" & Round(((month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
                text1.Text = text1.Text + " Days " & txtday & "  Days  " & "*" & Round(((month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
                     
            End If
    ElseIf txtyear > 10 Then
        total_year_m = (month_salary) * txtyear
        total_month_m = (month_salary / 12) * txtmonth
        total_day_m = ((month_salary) / 365) * txtday
        Total = total_year_m + total_month_m + total_day_m
   
        If SystemOptions.UserInterface = ArabicInterface Then
                 
            text1.Text = "⁄œœ «·”‰Ê«  «ﬂ»— „‰ 10 Ì „ «Õ ”«»   ‘Â— ⁄‰ ﬂ·  ”‰…" & vbNewLine
            text1.Text = text1.Text + "ﬁÌ„…   «·‘Â—  ”«ÊÌ  " & (month_salary) & vbNewLine
   
            text1.Text = text1.Text + "«Ã„«·Ì ”‰Ê«  «·Œœ„…  " & txtyear & "  ”‰…  " & "*" & Round((month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
            text1.Text = text1.Text + "«Ã„«·Ì ‘ÂÊ— «·Œœ„…  " & txtmonth & "  ‘Â—  " & "*" & Round(((month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
            text1.Text = text1.Text + "«Ã„«·Ì «Ì«„ «·Œœ„…  " & txtday & "  ÌÊ„  " & "*" & Round(((month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
   
        Else
            text1.Text = " The number of years greater than 10 a month is calculated for each year " & vbNewLine
            text1.Text = text1.Text + " Month Vakue=  " & (month_salary) & vbNewLine
   
            text1.Text = text1.Text + " Years Total  " & txtyear & "  Years  " & "*" & Round((month_salary), 2) & "=" & Round(total_year_m, 2) & vbNewLine
            text1.Text = text1.Text + "Months Total  " & txtmonth & "  Months  " & "*" & Round(((month_salary) / 12), 2) & "=" & Round(total_month_m, 2) & vbNewLine
            text1.Text = text1.Text + " Dyas Total " & txtday & "  Dyas  " & "*" & Round(((month_salary) / 365), 2) & "=" & Round(total_day_m, 2) & vbNewLine
    
        End If
    End If
    
 
    '  month_salary = total
    '  day_salary total / 30
   ' If Not IsNumeric(txtnum.text) Then txtnum.text = 0
   ' If Not IsNumeric(txttotal.text) Then txttotal.text = 0
   ' Me.txttotal.text = Total
'
'    Select Case OPR.Caption
'
'        Case "+"
'            txtnet = val(txtnum.text) + val(txttotal.text)
'
'        Case "-"
'            txtnet = val(txttotal.text) - val(txtnum.text)
'
'        Case "*"
'            txtnet = val(txtnum.text) * val(txttotal.text)
'
'        Case "/"
'
'            If txtnum.text = 0 Then
'                '    MsgBox "·« Ì„ﬂ‰ «·ﬁ”„… ⁄·Ï ’›—", vbCritical
'                txtnet = val(txttotal.text)
'            Else
'                txtnet = val(txttotal.text) / val(txtnum.text)
'            End If
'
'    End Select
'If val(txtnet.text) = 0 Then
'txtnet.text = val(txtSal.text)
'End If
    'txttotal.text = Format(total, SystemOptions.SysDefCurrencyForamt)
  '  txtnet.text = Round(val(txtnet.text), 2)
    Me.TXTLastTotal.Text = val(txtnet.Text) - val(Me.TXTAdvanceTotal.Text) - val(Me.TxtCash.Text)
    End If
        If Not IsNumeric(txtnum.Text) Then txtnum.Text = 0
    If Not IsNumeric(txttotal.Text) Then txttotal.Text = 0
    Me.txttotal.Text = Total
txtnet.Text = val(txtnum.Text) + val(txtSal.Text) + val(txtTicketValue.Text) + val(txtCustom.Text) + val(TxtAddOther.Text)
TxtTotalDis.Text = val(TXTAdvanceTotal.Text) + val(TxtCash.Text)
If AlowAssest.value = vbChecked Then
TxtTotalDis.Text = val(TxtTotalDis.Text) + val(TxtVlueVaction.Text)
End If
TXTLastTotal.Text = val(txtnet.Text) + val(txttotal.Text) - val(TxtTotalDis.Text)
    Select Case OPR.Caption
        Case "+"
            TXTLastTotal.Text = val(TXTLastTotal.Text) + val(txtnum.Text)
        Case "-"
            TXTLastTotal.Text = val(TXTLastTotal.Text) - val(txtnum.Text)
        Case "*"
            TXTLastTotal.Text = val(txtnum.Text) * val(txttotal.Text) + val(TXTLastTotal.Text) - val(txttotal.Text)

        Case "/"

            If txtnum.Text = 0 Then
                '    MsgBox "·« Ì„ﬂ‰ «·ﬁ”„… ⁄·Ï ’›—", vbCritical
               TXTLastTotal.Text = val(TXTLastTotal.Text) + val(txtnum.Text)
            Else
                TXTLastTotal.Text = (val(txttotal.Text) / val(txtnum.Text)) + val(TXTLastTotal.Text) - val(txttotal.Text)
            End If

    End Select

    'txttotal.text = Format(total, SystemOptions.SysDefCurrencyForamt)
    TXTLastTotal.Text = Round(val(TXTLastTotal.Text), 2)
    
    
End Function

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

                'btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

Private Sub ISButton2_Click()
Dim cont As Integer
Dim Account_code As String
Dim Balance  As String
Dim Account_code2 As String
Dim Account_Code5 As String
Dim Account_code4 As String
    If Me.DcboEmp.BoundText = "" Then Exit Sub
    txttotal.Text = 0
    txtEmpCode.Text = Me.DcboEmp.BoundText
    
        Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , val(DcboEmp.BoundText), EmpCode
    txtEmpCode.Text = EmpCode
    RtriverAsse val(DcboEmp.BoundText)
     get_employee_information val(Me.DcboEmp.BoundText), Account_code, Account_code2, Account_Code5, Account_code4
   ' Me.TXTAdvanceTotal.text = Round(getEmployeeAdvance(val(Me.DcboEmp.BoundText)), 2)
   ' Me.TxtCash.text = Round(getEmployeeCash(val(Me.DcboEmp.BoundText)), 2)
      Me.TxtVlueVaction.Text = Round(getEmployeeCashAssest(val(Me.DcboEmp.BoundText)), 2)
      GetEmployeeSalaryAccordingToComponentEndservice val(Me.DcboEmp.BoundText)
      Me.TxtTicktConract.Text = GetValueTikecCont(val(Me.DcboEmp.BoundText))
      ShowComponent
      If Me.TxtModFlg.Text <> "R" Then
      WriteCustomerBalPublic Account_code, Balance, , , , , , , date2.value, 1
     ' Advance = Balance
      TXTAdvanceTotal.Text = Balance
      End If
        WriteCustomerBalPublic Account_code2, Balance, , , , , , , date2.value, 1
      txtCustom.Text = Balance
        WriteCustomerBalPublic Account_Code5, Balance, , , , , , , date2.value, 1
      TxtCusTiket.Text = Balance
        WriteCustomerBalPublic Account_code4, Balance, , , , , , , date2.value, 1
      TxtEndService.Text = Balance
    '  retcountHoliday val(Me.DcboEmp.BoundText), cont
      Me.TxtVSa.Text = GetNoDayVacWithoutSalar(val(Me.DcboEmp.BoundText))
     
      
     txtCount.Text = Day(date2.value)
     
     
      cmd_CALC_NET_Click
      
End Sub
Function GetValueTikecCont(Optional Emp_id As Double) As Double
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = "Select * from Contract where Emp_id=" & Emp_id & ""
Rs4.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetValueTikecCont = IIf(IsNull(Rs4("TicketValueTotal").value), 0, Rs4("TicketValueTotal").value)
Else
GetValueTikecCont = 0
End If
End Function
Function GetNoDayVacWithoutSalar(Optional Emp_id As Double) As Double
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = " SELECT     EmpID, SUM(NoDay) AS SumNoDay"
sql = sql & " From dbo.TblInforVacatiom"
sql = sql & " Where (EmpID = " & Emp_id & ")"
sql = sql & " GROUP BY EmpID"
sql = sql & " Having (SUM(NoDay) > 0)"
Rs4.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetNoDayVacWithoutSalar = IIf(IsNull(Rs4("SumNoDay").value), 0, Rs4("SumNoDay").value)
Else
GetNoDayVacWithoutSalar = 0
End If
End Function
Sub UpdateAdvance(Optional EmpID As Double = 0)
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim ReturnAdvanceID As Double
Dim StrSQL As String
Dim i As Integer
Dim sql As String
sql = "select AdvanceID from TblEmpAdvanceRequest where Emp_ID =" & EmpID & " "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
Rs3.MoveFirst
For i = 1 To Rs3.RecordCount
ReturnAdvanceID = IIf(IsNull(Rs3("AdvanceID").value), 0, Rs3("AdvanceID").value)
StrSQL = "Update TblEmpAdvanceRequestDetails Set Payed=1 , EndSrvID=" & val(TXTid.Text) & "    Where  AdvanceID=" & ReturnAdvanceID & " "
                Cn.Execute StrSQL, , adExecuteNoRecords
Rs3.MoveNext
Next i
Else
ReturnAdvanceID = 0
End If
End Sub


Private Sub TxtAddOther_Change()
cmd_CALC_NET_Click
End Sub

Private Sub TxtAddOther_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtAddOther.Text, 0)
End Sub

Private Sub TXTAdvanceTotal_Change()
cmd_CALC_NET_Click
End Sub

Private Sub TxtCash_Change()
'cmd_CALC_NET_Click
End Sub

Private Sub txtCount_Change()
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then

If IsNumeric(txtCount.Text) Then
    Dim Sal  As Double, value   As Double
    
    Sal = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmp.BoundText), "", 0)
    txtDayval.Text = Round((Sal / 30), 2)
    value = Round(val(txtDayval.Text) * val(txtCount.Text), 2)

    cmd_CALC_NET_Click
End If
End If
End Sub

Private Sub txtCustom_Change()
cmd_CALC_NET_Click
End Sub

Private Sub txtdate_Change()
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.Text = ""
End If
End Sub

Private Sub TxtDiscounts_Change()
cmd_CALC_NET_Click
End Sub

Private Sub txtEmpCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode txtEmpCode.Text, EmpID
        DcboEmp.BoundText = EmpID
    End If
End Sub

Private Sub TxtEndService_Change()
cmd_CALC_NET_Click
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap
   Command5.Enabled = False
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Select Case Me.TxtModFlg.Text

        Case "R"
        chkGE.Enabled = False
  Dcombos.GetEmployees Me.DcboEmp

    Command5.Enabled = True
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            Me.TXTid.locked = True

            Me.txtEmpCode.locked = True
            Me.DcboEmp.locked = True
        
            Frame4.Enabled = False
            Me.date2.Enabled = False
            txtnum.locked = True
            Frame1.Enabled = False
            
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            
            End If

        Case "N"
        Accredit.Enabled = True
        chkGE.Enabled = True
        Dcombos.ClearMyDataCombo Me.DcboEmp
      Dcombos.GetEmployees Me.DcboEmp, True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
        
            Me.TXTid.locked = False

            Me.txtEmpCode.locked = False
            Me.DcboEmp.locked = False
        
            Frame4.Enabled = True
            Me.date2.Enabled = True
            txtnum.locked = False
            Frame1.Enabled = True

        Case "E"
        chkGE.Enabled = True
Dcombos.GetEmployees Me.DcboEmp
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            Me.TXTid.locked = False

            Me.txtEmpCode.locked = False
            Me.DcboEmp.locked = False
        
            Frame4.Enabled = True
            Me.date2.Enabled = True
            txtnum.locked = False
            Frame1.Enabled = True

    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub TxtNetEnd_Change()
cmd_CALC_NET_Click
End Sub

Private Sub txtnum_Change()
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
cmd_CALC_NET_Click
End If
End Sub



Public Sub TxtReqNo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
If Me.TxtModFlg.Text <> "R" Then
EmptyTxet
If val(Me.TxtReqNo.Text) <> 0 Then
RetriveRquestEnd val(Me.TxtReqNo.Text)
End If
End If
End If
End Sub

Private Sub TxtReqNo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Unload FrmSerachRegisterEndService
Load FrmSerachRegisterEndService
FrmSerachRegisterEndService.Ind = 1
FrmSerachRegisterEndService.Show
End If
End Sub

Private Sub txtTicketValue_Change()
cmd_CALC_NET_Click
End Sub

Private Sub TxtVlueVaction_Change()
cmd_CALC_NET_Click
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
'#############################################################################################################################################################################################################
    Function fillapprovData()
    
    Dim Num As Integer
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
 
 
    StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
    StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
    StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
    StrSQL = StrSQL + " FROM         dbo.ApprovalData left JOIN"
    StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
    StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
    StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.TXTid.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
    StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsDetails.RecordCount > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
        Else
            Accredit.Caption = "Sent To approval "
        End If
        Accredit.Enabled = False
    Else
        Accredit.Enabled = True
        If SystemOptions.UserInterface = ArabicInterface Then
            Accredit.Caption = " «·«—”«· ··«⁄ „«œ"
        Else
            Accredit.Caption = "Sent To approval "
        End If
    End If
 
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        GRID2.Rows = RsDetails.RecordCount + 1
        For Num = 1 To RsDetails.RecordCount
            GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
            If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
                GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
            Else
                GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
            End If
            GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
            If SystemOptions.UserInterface = ArabicInterface Then
                GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
            Else
                GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
            End If
            If SystemOptions.UserInterface = ArabicInterface Then
                GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
                GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
            GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
            
            RsDetails.MoveNext
            
            If Num = RsDetails.RecordCount Then
                If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Label110.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                    Else
                        Label110.Caption = "Approved"
                    End If
                    Label110.BackColor = &H80FF80
                Else
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Label110.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                    Else
                        Label110.Caption = "Currently required Approve"
                    End If
                    Label110.BackColor = &HFFFFC0
                End If
            End If
        Next Num
    Else
        GRID2.Rows = 1
    End If

    RsDetails.Close
End Function
Private Sub Accredit_Click()
    
    Dim BeginTrans As Boolean
 
    SendTopost Me.Name, "End_of_service", "id", 0, val(Me.Dcbranch.BoundText), val(TXTid.Text), TXTid.Text
    
    If Me.TxtModFlg.Text <> "N" And Me.TxtModFlg.Text <> "E" Then
        rs.Resync
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
    Else
        Accredit.Caption = "Sent To approval "
    End If
   
   Retrive
End Sub

