VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmEstimationsSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ «·„Ê«“‰«  «· ﬁœÌ—Ì…"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15015
   Icon            =   "FrmEstimationsSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   15015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame lblLW 
      BackColor       =   &H00E2E9E9&
      Height          =   2775
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   4440
      Width           =   14985
      Begin VB.Frame Frame9 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         TabIndex        =   53
         Top             =   1920
         Width           =   14775
         Begin VB.TextBox txtRemarks 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   54
            Top             =   240
            Width           =   11160
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   195
            Index           =   10
            Left            =   12600
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   240
            Width           =   1560
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   1800
         TabIndex        =   48
         Top             =   1200
         Width           =   6375
         Begin VB.OptionButton Option2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”‰Ê«  «·⁄„·"
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌœÊÌ"
            Height          =   315
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ·"
            Height          =   315
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   240
            Width           =   735
         End
         Begin VB.Image Imge 
            Height          =   255
            Left            =   3120
            Picture         =   "FrmEstimationsSearch.frx":6852
            Stretch         =   -1  'True
            Top             =   240
            Width           =   255
         End
         Begin VB.Image Imgw 
            Height          =   255
            Left            =   1560
            Picture         =   "FrmEstimationsSearch.frx":712F
            Stretch         =   -1  'True
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Œ Ì«— ”‰Ê«  «·„ﬁ«—‰…"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   7
            Left            =   4560
            TabIndex        =   52
            Top             =   240
            Width           =   1725
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   1095
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   8055
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   4200
            TabIndex        =   40
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   80674819
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   600
            TabIndex        =   41
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   80674819
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   330
            Left            =   4200
            TabIndex        =   42
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   80674819
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   330
            Left            =   600
            TabIndex        =   43
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   80674819
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —… „‰"
            Height          =   195
            Index           =   9
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   240
            Width           =   1785
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —… «·Ï"
            Height          =   195
            Index           =   8
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   600
            Width           =   1500
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   1
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   600
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   0
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   8280
         TabIndex        =   32
         Top             =   1200
         Width           =   6615
         Begin VB.OptionButton check2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ìﬁ«› «·Õ”«»"
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ·"
            Height          =   315
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ–Ì— ›ﬁÿ"
            Height          =   315
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   1215
         End
         Begin VB.Image Image3 
            Height          =   255
            Left            =   3240
            Picture         =   "FrmEstimationsSearch.frx":91CA
            Stretch         =   -1  'True
            Top             =   240
            Width           =   255
         End
         Begin VB.Image Image2 
            Height          =   255
            Left            =   1560
            Picture         =   "FrmEstimationsSearch.frx":B265
            Stretch         =   -1  'True
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄‰œ „Œ«·›… «· ﬁœÌ—Ì"
            ForeColor       =   &H000000FF&
            Height          =   405
            Index           =   12
            Left            =   4560
            TabIndex        =   36
            Top             =   240
            Width           =   1965
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   1575
         Begin ImpulseButton.ISButton CmdShowMoreOptions 
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ „ ﬁœ„..."
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
            ButtonImage     =   "FrmEstimationsSearch.frx":BB42
            ColorButton     =   14871017
            ColorHoverText  =   12582912
            ButtonToggles   =   1
            DrawFocusRectangle=   0   'False
            RightToLeft     =   -1  'True
            ButtonImageToggled=   "FrmEstimationsSearch.frx":BEDC
            ColorToggledHoverText=   192
         End
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Height          =   315
         Left            =   8520
         TabIndex        =   37
         Top             =   480
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lblLL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   285
         Left            =   12840
         TabIndex        =   38
         Top             =   480
         Width           =   1965
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   3720
      Width           =   8175
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   3480
         TabIndex        =   24
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   80674819
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   600
         TabIndex        =   25
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   80674819
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   315
         Index           =   3
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   315
         Index           =   4
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·≈‰‘«¡"
         Height          =   195
         Index           =   13
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3720
      Width           =   6735
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox TxtIDTO 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   315
         Index           =   5
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   315
         Index           =   6
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·”Ì‰«—ÌÊ"
         Height          =   195
         Index           =   14
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   7800
      Width           =   15015
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   10680
         TabIndex        =   14
         Top             =   240
         Width           =   3525
         _ExtentX        =   6218
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
         ButtonImage     =   "FrmEstimationsSearch.frx":C276
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   15
         Top             =   240
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "„”Õ"
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
         ButtonImage     =   "FrmEstimationsSearch.frx":12AD8
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
      Begin ImpulseButton.ISButton Cmd 
         Cancel          =   -1  'True
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   16
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
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
         BackStyle       =   0
         ButtonImage     =   "FrmEstimationsSearch.frx":1933A
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   7200
      Width           =   15015
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·≈Ã„«·Ì"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lblL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰ ÌÃ… «·»ÕÀ"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label lblL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   10
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   3015
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   15015
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   2625
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   14835
         _cx             =   26167
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         FormatString    =   $"FrmEstimationsSearch.frx":42F5C
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
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   14985
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13800
         Picture         =   "FrmEstimationsSearch.frx":431E5
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·»ÕÀ ⁄‰ «·„Ê«“‰«  «· ﬁœÌ—Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   8040
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   5400
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E2E9E9&
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   8520
      Width           =   15015
      Begin VB.Frame Frame11 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   2280
         TabIndex        =   64
         Top             =   1080
         Width           =   12615
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   240
            Width           =   1995
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   240
            Width           =   1995
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   8760
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "›—ﬁ «·«‰Õ—«›"
            Height          =   285
            Left            =   2520
            TabIndex        =   69
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«‰Õ—«› «·„”„ÊÕ »Â"
            Height          =   285
            Left            =   6840
            TabIndex        =   66
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«‰Õ—«› «·›⁄·Ì"
            Height          =   285
            Left            =   10920
            TabIndex        =   65
            Top             =   240
            Width           =   1605
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   2280
         TabIndex        =   59
         Top             =   240
         Width           =   6615
         Begin VB.OptionButton Option7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ·"
            Height          =   315
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌœÊÌ"
            Height          =   315
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "¬·Ì"
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.Image Image5 
            Height          =   255
            Left            =   1920
            Picture         =   "FrmEstimationsSearch.frx":44992
            Stretch         =   -1  'True
            Top             =   240
            Width           =   255
         End
         Begin VB.Image Image4 
            Height          =   255
            Left            =   3600
            Picture         =   "FrmEstimationsSearch.frx":4526F
            Stretch         =   -1  'True
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· Ê“Ì⁄"
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   11
            Left            =   5040
            TabIndex        =   63
            Top             =   240
            Width           =   1365
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11280
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   240
         Width           =   1995
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   9000
         TabIndex        =   3
         Top             =   720
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton ISButton2 
         Height          =   1530
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2699
         ButtonPositionImage=   3
         Caption         =   "»ÕÀ „ ﬁœ„"
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
         ButtonImage     =   "FrmEstimationsSearch.frx":4730A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         LowerToggledContent=   0   'False
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·Õ”«»"
         Height          =   285
         Left            =   13320
         TabIndex        =   58
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·Õ”«»"
         Height          =   285
         Left            =   13320
         TabIndex        =   56
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.ComboBox DcbMoth 
      Height          =   315
      ItemData        =   "FrmEstimationsSearch.frx":4DB6C
      Left            =   18240
      List            =   "FrmEstimationsSearch.frx":4DB6E
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Dcbyear 
      Height          =   315
      ItemData        =   "FrmEstimationsSearch.frx":4DB70
      Left            =   18240
      List            =   "FrmEstimationsSearch.frx":4DB72
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "FrmEstimationsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
'Dim DCboSearch As FrmExpensespaidAdvancedSearch
Private Sub Fg_Click()
FrmEstimations.FindRec val(Fg.TextMatrix(Fg.Row, 1))
End Sub
Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 '   AdditemTocCmp
   Set Dcombos = New ClsDataCombos
   Dcombos.GetBranches Me.Dcbranch
   Dcombos.GetAccountingCodes Me.DataCombo1
   Me.Option1.value = True
   Me.Option4.value = True
   Me.Option7.value = True
   HideCulems
   Set GrdBack = New ClsBackGroundPic
   With Me.Fg
   Set .WallPaper = GrdBack.Picture
   .AutoSize 0, .Cols - 1, False
    End With
   If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
         
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo
    SetDtpickerDate Me.DTPicker1
    SetDtpickerDate Me.DTPicker2
    SetDtpickerDate Me.DTPicker3
    SetDtpickerDate Me.DTPicker4
   End Sub

   Private Sub ISButton2_Click()
      GetDataAdvince
   End Sub
  Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
        GetData
        Case 1
        clear_all Me
          Me.DtpDateFrom.value = ""
          Me.DtpDateTo.value = ""
          Me.DTPicker1.value = ""
          Me.DTPicker2.value = ""
          Me.DTPicker3.value = ""
          Me.DTPicker4.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lblL(0).Caption = "Search Results"
            End If
         Me.Option1.value = True
         Me.Option4.value = True
         Me.Option7.value = True
          Case 2
            Unload Me
    End Select
End Sub
Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FormPostion Me, SavePostion
   ' Set DCboSearch = Nothing
End Sub
Public Sub GetData()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = "SELECT     dbo.TblEstiation.transID, dbo.TblEstiation.recordDate, dbo.TblEstiation.Fromdate, dbo.TblEstiation.todate, dbo.TblEstiation.FromdateH, dbo.TblEstiation.todateH, "
    sql = sql & "                  dbo.TblEstiation.Remarks, dbo.TblEstiation.FullRemarks, dbo.TblEstiation.Percentage, dbo.TblEstiation.OperatorsID, dbo.TblEstiation.BranchId,"
   sql = sql & "                     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEstiation.Alarms, dbo.TblEstiation.ManualEntry, dbo.TblEstiation.CompYear,"
   sql = sql & "                     dbo.TblEstiation.TypeEsstame , dbo.TblEstiation.OptMethod, dbo.TblEstiation.description"
   sql = sql & "   FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
   sql = sql & "                     dbo.TblEstiation ON dbo.TblBranchesData.branch_id = dbo.TblEstiation.BranchId"
               
  
       BolBegine = False
       StrWhere = ""
    '''''''''''''''''''''''''''''''''''' id
    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblEstiation.transID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEstiation.transID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblEstiation.transID <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblEstiation.transID <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
     '''''''''''''''''''''''''''''''''''''''''''''''' date
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEstiation.recordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEstiation.recordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblEstiation.recordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEstiation.recordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Branch
    If Me.Dcbranch.text <> "" And (val(Dcbranch.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblEstiation.BranchId =" & Me.Dcbranch.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblEstiation.BranchId =" & Me.Dcbranch.BoundText & ""
       End If
     End If
         ''''' DATA SEARCH FROM DATE
     If Not IsNull(Me.DTPicker1.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEstiation.Fromdate >=" & SQLDate(Me.DTPicker1.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEstiation.Fromdate >=" & SQLDate(Me.DTPicker1.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTPicker2.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblEstiation.Fromdate <=" & SQLDate(Me.DTPicker2.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblEstiation.Fromdate <=" & SQLDate(Me.DTPicker2.value, True) & ""
        End If
    End If
    ''''''' SECAND DATE TO DATE
     If Not IsNull(Me.DTPicker3.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEstiation.todate >=" & SQLDate(Me.DTPicker3.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEstiation.todate >=" & SQLDate(Me.DTPicker3.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTPicker4.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblEstiation.todate <=" & SQLDate(Me.DTPicker4.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblEstiation.todate <=" & SQLDate(Me.DTPicker4.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''chech box 1
        If (Me.Check1.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TblEstiation.Alarms = 0 "
         Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TblEstiation.Alarms = 1 "
        End If
        End If
        
        If (Me.Check2.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TblEstiation.Alarms = 1 "
        Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TblEstiation.Alarms = 0 "
        End If
        End If
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''chech box 2
        If (Me.Option3.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TblEstiation.CompYear = 0 "
         Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TblEstiation.CompYear = 1 "
        End If
        End If
        
        If (Me.Option2.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TblEstiation.CompYear = 1 "
        Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TblEstiation.CompYear = 0 "
        End If
        End If
       '''''' SEARCH TEXT
        If Me.txtRemarks.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblEstiation.Remarks like '%" & Me.txtRemarks.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblEstiation.Remarks like '%" & Me.txtRemarks.text & "%'"
        End If
       End If
  '-----------------------------------
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblEstiation.transID"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = " ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("transID").value), "", rs("transID").value)
                 If Not (IsNull(rs("recordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("recordDate").value, "yyyy/M/d")
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                 End If
                 
                If Not (IsNull(rs("Fromdate").value)) Then
                .TextMatrix(i, .ColIndex("Fromdate")) = Format(rs("Fromdate").value, "yyyy/M/d")
                End If
                
                If Not (IsNull(rs("todate").value)) Then
                .TextMatrix(i, .ColIndex("todate")) = Format(rs("todate").value, "yyyy/M/d")
                End If
                
               .TextMatrix(i, .ColIndex("Alarms")) = IIf(IsNull(rs("Alarms").value), "", rs("Alarms").value)
               .TextMatrix(i, .ColIndex("ManualEntry")) = IIf(IsNull(rs("CompYear").value), "", rs("CompYear").value)
               .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
                rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
        End With
    End If
End Sub
Public Sub GetDataAdvince()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
    sql = "SELECT     dbo.TblEstiation.transID, dbo.TblEstiation.recordDate, dbo.TblEstiation.Fromdate, dbo.TblEstiation.todate, dbo.TblEstiation.FromdateH, dbo.TblEstiation.todateH, "
    sql = sql & "                  dbo.TblEstiation.Remarks, dbo.TblEstiation.FullRemarks, dbo.TblEstiation.Percentage, dbo.TblEstiation.OperatorsID, dbo.TblEstiation.BranchId,"
    sql = sql & "                  dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEstiation.Alarms, dbo.TblEstiation.ManualEntry, dbo.TblEstiationDetails.year1,"
    sql = sql & "                  dbo.TblEstiationDetails.year2, dbo.TblEstiationDetails.year3, dbo.TblEstiationDetails.Estimated1, dbo.TblEstiationDetails.Estimated2,"
    sql = sql & "                  dbo.TblEstiationDetails.Estimated3, dbo.TblEstiationDetails.Estimated, dbo.TblEstiationDetails.Acctual, dbo.TblEstiationDetails.Diff, dbo.TblEstiationDetails.Varance,"
    sql = sql & "                  dbo.TblEstiationDetails.AllowVariance, dbo.TblEstiationDetails.DiffVariance, dbo.TblEstiationDetails.AccountCode, dbo.ACCOUNTS.Account_Name,"
    sql = sql & "                  dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.TblEstiationDetails.id, dbo.TblEstiationDetails.Distribution,"
    sql = sql & "                  dbo.TblEstiationDetails.Auto_Manul , dbo.TblEstiation.CompYear, dbo.TblEstiation.TypeEsstame, dbo.TblEstiation.description, dbo.TblEstiation.OptMethod"
    sql = sql & "     FROM         dbo.TblEstiationDetails LEFT OUTER JOIN"
    sql = sql & "                  dbo.ACCOUNTS ON dbo.TblEstiationDetails.AccountCode = dbo.ACCOUNTS.Account_Code RIGHT OUTER JOIN"
    sql = sql & "                  dbo.TblEstiation ON dbo.TblEstiationDetails.transID = dbo.TblEstiation.transID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.TblEstiation.BranchId = dbo.TblBranchesData.branch_id"
    
       BolBegine = False
       StrWhere = ""
    '''''''''''''''''''''''''''''''''''' id
    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblEstiation.transID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEstiation.transID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblEstiation.transID <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblEstiation.transID <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
     '''''''''''''''''''''''''''''''''''''''''''''''' date
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEstiation.recordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEstiation.recordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblEstiation.recordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEstiation.recordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Branch
    If Me.Dcbranch.text <> "" And (val(Dcbranch.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblEstiation.BranchId =" & Me.Dcbranch.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblEstiation.BranchId =" & Me.Dcbranch.BoundText & ""
       End If
     End If
         ''''' DATA SEARCH FROM DATE
     If Not IsNull(Me.DTPicker1.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEstiation.Fromdate >=" & SQLDate(Me.DTPicker1.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEstiation.Fromdate >=" & SQLDate(Me.DTPicker1.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTPicker2.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblEstiation.Fromdate <=" & SQLDate(Me.DTPicker2.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblEstiation.Fromdate <=" & SQLDate(Me.DTPicker2.value, True) & ""
        End If
    End If
    ''''''' SECAND DATE TO DATE
     If Not IsNull(Me.DTPicker3.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEstiation.todate >=" & SQLDate(Me.DTPicker3.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEstiation.todate >=" & SQLDate(Me.DTPicker3.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTPicker4.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblEstiation.todate <=" & SQLDate(Me.DTPicker4.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblEstiation.todate <=" & SQLDate(Me.DTPicker4.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''chech box 1
        If (Me.Check1.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TblEstiation.Alarms = 0 "
         Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TblEstiation.Alarms = 1 "
        End If
        End If
        
        If (Me.Check2.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TblEstiation.Alarms = 1 "
        Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TblEstiation.Alarms = 0 "
        End If
        End If
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''chech box 2
        If (Me.Option3.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TblEstiation.CompYear = 0 "
         Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TblEstiation.CompYear = 1 "
        End If
        End If
        
        If (Me.Option2.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TblEstiation.CompYear = 1 "
        Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TblEstiation.CompYear = 0 "
        End If
        End If
       '''''' SEARCH TEXT
        If Me.txtRemarks.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblEstiation.Remarks like '%" & Me.txtRemarks.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblEstiation.Remarks like '%" & Me.txtRemarks.text & "%'"
        End If
       End If
       ' Advince ''''''''''''''''''''''''''''''''''''''
         '''''' SEARCH TEXT
        If Me.Text1.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblEstiationDetails.AccountCode like '%" & Me.Text1.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblEstiationDetails.AccountCode like '%" & Me.Text1.text & "%'"
        End If
       End If
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' comboooo
      If Me.DataCombo1.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND   dbo.TblEstiationDetails.AccountCode =N'" & Me.DataCombo1.BoundText & "'"
        Else:
          BolBegine = True
          StrWhere = " Where  dbo.TblEstiationDetails.AccountCode =N'" & Me.DataCombo1.BoundText & "'"
       End If
       End If
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''chech box 3
        If (Me.Option6.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TblEstiationDetails.Distribution = 1 "
         Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TblEstiationDetails.Distribution  = 2 "
        End If
        End If
        
        If (Me.Option5.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TblEstiationDetails.Distribution = 2 "
        Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TblEstiationDetails.Distribution = 1 "
        End If
        End If
             '''''' SEARCH TEXT
        If Me.Text2.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND   dbo.TblEstiationDetails.Varance like '%" & Me.Text2.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where   dbo.TblEstiationDetails.Varance like '%" & Me.Text2.text & "%'"
        End If
        End If
            '''''' SEARCH TEXT
        If Me.Text3.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND   dbo.TblEstiationDetails.AllowVariance like '%" & Me.Text3.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblEstiationDetails.AllowVariance like '%" & Me.Text3.text & "%'"
        End If
        End If
        '''''' SEARCH TEXT
        If Me.Text4.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND   dbo.TblEstiationDetails.DiffVariance like '%" & Me.Text4.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblEstiationDetails.DiffVariance like '%" & Me.Text4.text & "%'"
        End If
        End If
      '-----------------------------------
  
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblEstiation.transID"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = " ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("transID").value), "", rs("transID").value)
                 If Not (IsNull(rs("recordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("recordDate").value, "yyyy/M/d")
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                 .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                Else
                .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                 End If
                If Not (IsNull(rs("Fromdate").value)) Then
                .TextMatrix(i, .ColIndex("Fromdate")) = Format(rs("Fromdate").value, "yyyy/M/d")
                End If
                If Not (IsNull(rs("todate").value)) Then
                .TextMatrix(i, .ColIndex("todate")) = Format(rs("todate").value, "yyyy/M/d")
                End If
               .TextMatrix(i, .ColIndex("Alarms")) = IIf(IsNull(rs("Alarms").value), "", rs("Alarms").value)
               .TextMatrix(i, .ColIndex("ManualEntry")) = IIf(IsNull(rs("CompYear").value), "", rs("CompYear").value)
               .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
               '''''''''''''''''''''''''''''''''''''''''''' Advinced
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
                .TextMatrix(i, .ColIndex("Distribution")) = IIf(IsNull(rs("Distribution").value), "", rs("Distribution").value)
                .TextMatrix(i, .ColIndex("Varance")) = IIf(IsNull(rs("Varance").value), "", rs("Varance").value)
                .TextMatrix(i, .ColIndex("AllowVariance")) = IIf(IsNull(rs("AllowVariance").value), "", rs("AllowVariance").value)
                .TextMatrix(i, .ColIndex("DiffVariance")) = IIf(IsNull(rs("DiffVariance").value), "", rs("DiffVariance").value)
                rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
        End With
    End If
End Sub
Private Sub ChangeLang()
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
    lbl(2).Caption = "Total"
    lblL(0).Caption = "Search Results"
    Me.ISButton2.Caption = "Advanced Search"
    Me.Caption = "Estimations Search"
    Me.CmdShowMoreOptions.Caption = "Advanced"
    ' labell name
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(14).Caption = "Scenario No."
    Me.lbl(5).Caption = "From"
    Me.lbl(6).Caption = "To"
    Me.lbl(13).Caption = "Operation Date"
    Me.lbl(4).Caption = "From"
    Me.lbl(3).Caption = "To"
    Me.lblLL.Caption = "Branch Name"
    lbl(7).Caption = "Select Years"
    '''''''''''''''''''''''''''''''''''''''
    lbl(9).Caption = "Date From"
    lbl(0).Caption = "To"
    lbl(8).Caption = "Date To"
    lbl(1).Caption = "To"
    ''''''''''''''''''''''''''''''''''''''
    Me.lbl(12).Caption = " When The Estimated Offense "
    Me.Option1.Caption = "All"
    Me.Check1.Caption = "just Worning"
    Me.Check2.Caption = "Stop The Acount"
    '''''''''''''''''''''''''''''''''''''''''
   ' Me.lbl(7).Caption = "Last Years Entering"
    Me.Option4.Caption = "All"
    Me.Option3.Caption = "Manual"
    Me.Option2.Caption = "Years Work"
    Me.lbl(10).Caption = "Remarks"
    '''''''''''''''''''''''''''''''''''''''''''''
    Me.Label2.Caption = "Acounting Code"
    Me.Label3.Caption = "Acounting Name"
    ''''''''''''''''''''''''''''''''''
    Me.lbl(11).Caption = "Distribution"
    Me.Option7.Caption = "All"
    Me.Option6.Caption = "Manual"
    Me.Option5.Caption = "Automatic"
   '''''''''''''''''''''''''''''''''''''''''
    Me.Label4.Caption = "actual deviation"
    Me.Label5.Caption = "Allowable deviation"
    Me.Label6.Caption = "Deviation difference"
    lbl(2).Caption = "Total"
    ''''''''''''''''''''''' next
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Scenario No."
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
        .TextMatrix(0, .ColIndex("BranchID")) = "Branch Name"
        .TextMatrix(0, .ColIndex("Fromdate")) = "From date"
        .TextMatrix(0, .ColIndex("todate")) = "To date"
        .TextMatrix(0, .ColIndex("Alarms")) = "When The Estimated Offense"
        .TextMatrix(0, .ColIndex("ManualEntry")) = "Last Years Entering"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
        .TextMatrix(0, .ColIndex("AccountCode")) = "Acounting Code"
        .TextMatrix(0, .ColIndex("Account_Name")) = "Acounting Name"
        .TextMatrix(0, .ColIndex("Distribution")) = "Distribution"
        .TextMatrix(0, .ColIndex("Varance")) = "actual deviation"
        .TextMatrix(0, .ColIndex("AllowVariance")) = "Allowable deviation"
        .TextMatrix(0, .ColIndex("DiffVariance")) = "Deviation difference"
    End With
 End Sub
  Private Sub CmdShowMoreOptions_Click()
    If CmdShowMoreOptions.value = True Then
       ShowCulems
       Cmd_Click (1)
       Cmd(0).Enabled = False
       ISButton2.Enabled = True
       Me.Height = Me.Frame7.top + Frame7.Height + 400
       Else
       HideCulems
       Cmd_Click (1)
       Cmd(0).Enabled = True
       ISButton2.Enabled = False
       Me.Height = Me.Frame7.top + 400
    End If
  End Sub
  Private Sub ShowCulems()
  Fg.ColHidden(9) = False
  Fg.ColHidden(10) = False
  Fg.ColHidden(11) = False
  Fg.ColHidden(12) = False
  Fg.ColHidden(13) = False
  Fg.ColHidden(14) = False
  End Sub
  Private Sub HideCulems()
   Fg.ColHidden(9) = True
  Fg.ColHidden(10) = True
  Fg.ColHidden(11) = True
  Fg.ColHidden(12) = True
  Fg.ColHidden(13) = True
  Fg.ColHidden(14) = True
  End Sub
'''''''''''''''''''''''''''' end


