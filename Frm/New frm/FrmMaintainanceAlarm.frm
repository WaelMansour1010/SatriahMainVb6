VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMaintainanceAlarm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14550
   Icon            =   "FrmMaintainanceAlarm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   14550
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   13680
      Top             =   8160
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   8040
      Width           =   14535
      Begin ImpulseButton.ISButton Cmd 
         Height          =   492
         Index           =   6
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1368
         _ExtentX        =   2408
         _ExtentY        =   873
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
         ButtonImage     =   "FrmMaintainanceAlarm.frx":6852
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   492
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   1272
         _ExtentX        =   2249
         _ExtentY        =   873
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
         ButtonImage     =   "FrmMaintainanceAlarm.frx":30474
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2E9E9&
      Height          =   1092
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   14535
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   735
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   5895
         Begin MSComCtl2.DTPicker todate 
            Height          =   330
            Left            =   360
            TabIndex        =   8
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   94109697
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker Fromdate 
            Height          =   336
            Left            =   3120
            TabIndex        =   12
            Top             =   240
            Width           =   1692
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   94109697
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   315
            Index           =   0
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "≈«·Ï"
            Height          =   435
            Index           =   14
            Left            =   2100
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   540
         End
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   735
         Index           =   5
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   1296
         ButtonPositionImage=   1
         Caption         =   " ÕœÌÀ"
         BackColor       =   14871017
         FontSize        =   14.25
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmMaintainanceAlarm.frx":36CD6
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   12632064
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   12632064
         LowerToggledContent=   0   'False
         ColorTextShadow =   -2147483637
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
         Left            =   6960
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   600
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "HH:mm:ss"
         Format          =   25559043
         UpDown          =   -1  'True
         CurrentDate     =   37140.0034722222
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÕœÌÀ ﬂ·"
         Height          =   435
         Index           =   4
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   780
      End
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   14565
      _cx             =   25691
      _cy             =   1032
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
      Caption         =   "    «· ‰»ÌÂ«    "
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
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2520
         TabIndex        =   2
         Top             =   0
         Width           =   2205
      End
      Begin VB.Image Image1 
         Height          =   555
         Index           =   0
         Left            =   9840
         Picture         =   "FrmMaintainanceAlarm.frx":3D538
         Stretch         =   -1  'True
         Top             =   0
         Width           =   795
      End
   End
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   6372
      Left            =   0
      TabIndex        =   13
      Top             =   1560
      Width           =   14508
      _cx             =   25590
      _cy             =   11239
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
      Caption         =   " ‰»ÌÂ«  «·’Ì«‰…| ‰»ÌÂ«  «·÷„«‰"
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5955
         Left            =   15150
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   45
         Width           =   14415
         _cx             =   25426
         _cy             =   10504
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
         Begin VSFlex8UCtl.VSFlexGrid Grid2 
            Height          =   5052
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   14352
            _cx             =   25315
            _cy             =   8911
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
            BackColorAlternate=   16777152
            GridColor       =   0
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
            FormatString    =   $"FrmMaintainanceAlarm.frx":45B55
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
            Begin MSComctlLib.ProgressBar ProgressBar2 
               Height          =   612
               Left            =   3360
               TabIndex        =   20
               Top             =   2040
               Visible         =   0   'False
               Width           =   8412
               _ExtentX        =   14843
               _ExtentY        =   1085
               _Version        =   393216
               Appearance      =   0
            End
            Begin VB.Label Label2 
               Caption         =   "%"
               Height          =   375
               Index           =   1
               Left            =   10440
               TabIndex        =   21
               Top             =   -600
               Width           =   375
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   5955
         Index           =   2
         Left            =   45
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   45
         Width           =   14415
         _cx             =   25426
         _cy             =   10504
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
         Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
            Height          =   5052
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   14352
            _cx             =   25315
            _cy             =   8911
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
            BackColorAlternate=   16777152
            GridColor       =   0
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
            FormatString    =   $"FrmMaintainanceAlarm.frx":45CA7
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
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   612
               Left            =   3720
               TabIndex        =   17
               Top             =   1560
               Visible         =   0   'False
               Width           =   8412
               _ExtentX        =   14843
               _ExtentY        =   1085
               _Version        =   393216
               Appearance      =   0
            End
            Begin VB.Label Label2 
               Caption         =   "%"
               Height          =   375
               Index           =   0
               Left            =   10440
               TabIndex        =   18
               Top             =   -600
               Width           =   375
            End
         End
      End
   End
End
Attribute VB_Name = "FrmMaintainanceAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim alarm_static As Long
Dim alarm As Long
Public indexx As Integer

Private Sub Cmd_Click(Index As Integer)
Select Case Index
Case 5

ProgressBar1.Visible = True
': ProgressBar1.value = 10

FillGrid
': ProgressBar1.value = 50
ProgressBar1.Visible = False
ProgressBar1.value = 0


ProgressBar2.Visible = True
': ProgressBar2.value = 10
FillGrid2
': ProgressBar2.value = 50
ProgressBar2.Visible = False
ProgressBar2.value = 0

Case 6
Me.Hide
Case 9
        ' print_report
End Select

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


 MySQL = ""
MySQL = MySQL & "   SELECT dbo.TblQestFexed.Value, dbo.TblQestFexed.Due_Date, dbo.TblQestFexed.QestID, dbo.notes_all.NoteID, dbo.TblCustemers.CusID, dbo.notes_all.NoteSerial1,"
MySQL = MySQL & "                     dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee, dbo.notes_all.too, dbo.notes_all.NoteDateH, dbo.notes_all.NoteDate"
MySQL = MySQL & "  , inst_NO  FROM     dbo.notes_all INNER JOIN"
MySQL = MySQL & "                     dbo.TblQestFexed ON dbo.notes_all.NoteID = dbo.TblQestFexed.Ind LEFT OUTER JOIN"
MySQL = MySQL & "                     dbo.TblCustemers ON dbo.notes_all.CusID = dbo.TblCustemers.CusID"

 MySQL = MySQL & "    where 1 =1 "
  
 If IsNumeric(Text2.Text) And Text2.Text <> "" Then
    MySQL = MySQL & "  and  NoteSerial1 =  " & Text2.Text
 End If
  
 If IsNumeric(Text3.Text) And Text3.Text <> "" Then
    MySQL = MySQL & "   and too  =  " & Text3.Text
 End If
 
 If IsNumeric(Text5.Text) And Text5.Text <> "" Then
    MySQL = MySQL & "  and  value  =  " & Text5.Text
 End If
  
  If IsNumeric(Text4.Text) And Text4.Text <> "" Then
    MySQL = MySQL & "  and  Inst_No  =  " & Text4.Text
 End If
  
  
 If Not (IsNull(Me.Fromdate.value)) Then
 MySQL = MySQL + " and (dbo.TblQestFexed.Due_Date >='" & SQLDate(Fromdate.value) & "')"
 End If
 
 If Not (IsNull(Me.todate.value)) Then
 MySQL = MySQL + " and (dbo.TblQestFexed.Due_Date <='" & SQLDate(todate.value) & "')"
 End If

 'If Not (IsNull(Me.DTPicker1.value)) Then
 'MySQL = MySQL + " and (  notedate   = '" & SQLDate(todate.value) & "')"
 'End If

'If Me.DBCboClientName.text <> "" And val(Me.DBCboClientName.BoundText) <> 0 Then
'MySQL = MySQL + "and notes_all.CusID =" & val(Me.DBCboClientName.BoundText) & ""
'End If

MySQL = MySQL + "   order by  dbo.TblQestFexed.Ind "
  
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_InstallmentAlarm.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_InstallmentAlarmE.rpt"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
'        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
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
       If Not (IsNull(Me.Fromdate.value)) Then
        xReport.ParameterFields(6).AddCurrentValue Fromdate.value
       End If
      If Not (IsNull(Me.todate.value)) Then
        xReport.ParameterFields(7).AddCurrentValue todate.value
      End If
        ' xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
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

Private Sub CmdHelp_Click()
          clear_all Me
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
todate.value = Date
Fromdate.value = Date
End Sub

Private Sub DBCboClientName_Change()
'    TxtSearchCode.text = ""
''
   ' Dim DefaultSalesPersonId As Integer
'    Dim fullcode As String

   ' GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode

   ' TxtSearchCode.text = fullcode
 
End Sub
Function cahngelang()
    EleHeader.Caption = " Alarm Installment Assets "
    Me.Caption = EleHeader.Caption
    lbl(3).Caption = "Vendor"
    lbl(2).Caption = "Invoice Date"
    lbl(1).Caption = "Invoice No."
    lbl(0).Caption = "From"
    lbl(14).Caption = "To"
    Frame5.Caption = "Due Period"
    lbl(4).Caption = "Update All"
    lbl(5).Caption = "Vendor Invoice No."
    lbl(6).Caption = "Installment No."
    lbl(7).Caption = "Insatllment Value"
    
    
   CmdHelp.Caption = "Clear"
   Cmd(5).Caption = "Refresh"
   Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    With GridInstallments
    .TextMatrix(0, .ColIndex("Ser")) = "Serial"
    .TextMatrix(0, .ColIndex("QsrID")) = "Installment Num"
    .TextMatrix(0, .ColIndex("Value")) = "Installment Value"
    .TextMatrix(0, .ColIndex("Due_Date")) = "Due_Date"
    .TextMatrix(0, .ColIndex("CusName")) = "Vendor"
    .TextMatrix(0, .ColIndex("NoteSerial1")) = "Bill Num"
    .TextMatrix(0, .ColIndex("View")) = "View"
     .TextMatrix(0, .ColIndex("too")) = "Vendor invoice No."
    End With
End Function



Public Sub FillGrid(Optional str As String)
Dim cont1 As Double
Dim cont As Double
Dim typ As Integer
  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
 
 MySQL = ""
MySQL = MySQL & " select H.* , D.Itemname , D.Fullcode , m.name , m.remarks from TBLRegularMaint H , TblItems D , TblMaintenanceType M  where H.itemid = D.ItemId and H.MaintenanceIDS = M.ID"
 

  
 If Not (IsNull(Me.Fromdate.value)) Then
 MySQL = MySQL + " and DateOfRegularMaint  >='" & SQLDate(Fromdate.value) & "'"
 End If
 
 If Not (IsNull(Me.todate.value)) Then
 MySQL = MySQL + " and DateOfRegularMaint <='" & SQLDate(todate.value) & "' "
 End If


MySQL = MySQL + "   order by  H.ID  "
  
  
  
  
  
Dim ActualTotal As Double
    rs.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .Rows = 1
        .Clear flexClearScrollable
Dim j As Integer
Dim Notstr As String
j = 0
Notstr = ""
        If rs.RecordCount > 0 Then
        
        ProgressBar1.Max = rs.RecordCount
        
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
              .TextMatrix(i, .ColIndex("Ser")) = i
               ProgressBar1.value = i
               .TextMatrix(i, .ColIndex("Fullcode")) = (IIf(IsNull(rs.Fields("Fullcode").value), "", rs.Fields("Fullcode").value))
              .TextMatrix(i, .ColIndex("Itemname")) = (IIf(IsNull(rs.Fields("Itemname").value), 0, Round(rs.Fields("Itemname").value, 2)))
              .TextMatrix(i, .ColIndex("name")) = (IIf(IsNull(rs.Fields("name").value), Date, rs.Fields("name").value))
              .TextMatrix(i, .ColIndex("remarks")) = (IIf(IsNull(rs.Fields("remarks").value), "", rs.Fields("remarks").value))
              .TextMatrix(i, .ColIndex("GranteeEndDate")) = (IIf(IsNull(rs.Fields("DateOfRegularMaint").value), "", rs.Fields("DateOfRegularMaint").value))
           
             If i > 5 Then
                        .TopRow = i - 10
                End If
         DoEvents
         
        rs.MoveNext
            Next i
 
            rs.Close
        End If
 'sa .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With

End Sub



Public Sub FillGrid2(Optional str As String)

Dim cont1 As Double
Dim cont As Double
Dim typ As Integer
  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
 
MySQL = MySQL & " Select * from ( "
MySQL = MySQL & " select  H.ID ,H.SandType , H.OrderNo , H.GranteeStartDate , H.GranteeEndDate , D.Fullcode code ,D.Project_name as name , H.GranteeTypeopt , '„‘«—Ì⁄' depend"
MySQL = MySQL & " From TblWarrantyOffer H , projects  D where H.sandtype = 1  and D.ID = H.ProjectID"
MySQL = MySQL & " Union"
MySQL = MySQL & " select  H.ID ,H.SandType , H.OrderNo , H.GranteeStartDate , H.GranteeEndDate ,  '' as code  ,'' name , H.GranteeTypeopt ,'€ﬁÊœ ﬁœÌ„…' depend"
MySQL = MySQL & " From TblWarrantyOffer H , projects  D where H.sandtype = 2"
MySQL = MySQL & " Union"
MySQL = MySQL & " select HH.ID , HH.SandType , HH.orderNo , H.GranteeStartDate , H.GranteeEndDate  , D.Fullcode code , D.Itemname name  , HH.GranteeTypeopt , '„»Ì⁄« ' depend"
MySQL = MySQL & " from  TblWarrantyOffer HH, TblWarrantyOfferDet H  , TblItems D"
MySQL = MySQL & " Where hh.ID = H.WrantID"
MySQL = MySQL & " ) as tbl"

MySQL = MySQL & " where 1 = 1 "
  
 If Not (IsNull(Me.Fromdate.value)) Then
 MySQL = MySQL + " and GranteeEndDate  >='" & SQLDate(Fromdate.value) & "' "
 End If
 
 If Not (IsNull(Me.todate.value)) Then
 MySQL = MySQL + " and GranteeEndDate <='" & SQLDate(todate.value) & "' "
 End If


MySQL = MySQL + "   order by  ID  "
  
  
Dim ActualTotal As Double
    rs.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GRID2
       .Rows = 1
        .Clear flexClearScrollable
Dim j As Integer, g As Integer
Dim Notstr As String
j = 0
Notstr = ""
        If rs.RecordCount > 0 Then
         ProgressBar2.Max = rs.RecordCount
         
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
              .TextMatrix(i, .ColIndex("Ser")) = i
              ProgressBar2.value = i
              
              .TextMatrix(i, .ColIndex("Fullcode")) = (IIf(IsNull(rs.Fields("code").value), "", rs.Fields("code").value))
               .TextMatrix(i, .ColIndex("depend")) = IIf(IsNull(rs.Fields("depend").value), "", rs.Fields("depend").value)
              .TextMatrix(i, .ColIndex("Itemname")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
              .TextMatrix(i, .ColIndex("GranteeStartDate")) = (IIf(IsNull(rs.Fields("GranteeStartDate").value), Date, rs.Fields("GranteeStartDate").value))
              .TextMatrix(i, .ColIndex("GranteeEndDate")) = (IIf(IsNull(rs.Fields("GranteeEndDate").value), Date, rs.Fields("GranteeEndDate").value))
                g = IIf(IsNull(rs.Fields("GranteeType").value), 0, rs.Fields("GranteeType").value)
                
                If g = 0 Then
                                 .TextMatrix(i, .ColIndex("GranteeType")) = "»œÊ‰ «·ﬁÿ⁄ "
                ElseIf g = 1 Then
                                  .TextMatrix(i, .ColIndex("GranteeType")) = "„⁄ «·ﬁÿ⁄ "
                End If
                                
                If i > 5 Then
                        .TopRow = i - 10
                End If
         DoEvents
        rs.MoveNext
            Next i
 
            rs.Close
        End If
 'sa .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With

End Sub








Private Sub Form_Load()
Dim StrSQL As String
   Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Dim Dcombos As New ClsDataCombos
      Set Dcombos = New ClsDataCombos
    'Dcombos.GetBranches Me.Dcbranch
    'Dcombos.GetFixedAssets DcbFixed
    C1Tab1.TabVisible(0) = False
    C1Tab1.TabVisible(1) = False
    
    C1Tab1.TabVisible(indexx) = True
    
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " Select CusID,CusName From TblCustemers Where Type=2 or CustomerandVendor=1"
    Else
        StrSQL = " Select CusID,CusNamee From TblCustemers Where Type=2 or CustomerandVendor=1"
    End If


      

    If SystemOptions.UserInterface = EnglishInterface Then

        SetInterface Me
       cahngelang
    End If

Fromdate.value = Date
todate.value = Date

FillGrid

Fire_Alarm
 
End Sub

Private Sub Fire_Alarm()

Dim ss As Long, mm As Long, hh As Long, allsecond As Long

ss = Second(IntervalAll.value)
mm = Minute(IntervalAll.value)
hh = Hour(IntervalAll.value)
mm = mm * 60
hh = hh * 60 * 60
allsecond = ss + mm + hh


alarm = allsecond
alarm_static = alarm

End Sub


Private Sub FromDate_Change()
FillGrid
End Sub

Private Sub GridInstallments_Click()
On Error Resume Next

If GridInstallments.Col = GridInstallments.ColIndex("View") Then
Dim i As Integer
i = GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("NoteID"))
If i > 0 Then
FrmExpenses4.Retrive (i)
End If

End If
End Sub



Private Sub IntervalAll_Change()
Fire_Alarm
End Sub

Private Sub Timer1_Timer()

alarm = alarm - 1
If alarm <= 0 Then
       Cmd_Click (5)
       alarm = alarm_static
End If

End Sub

Private Sub ToDate_Change()
FillGrid
End Sub

Private Sub XPLbl_Click(Index As Integer)
End Sub
