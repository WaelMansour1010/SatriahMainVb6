VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form Order_no_search2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11880
   Icon            =   "Order_no_search2.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   11880
   Begin VB.TextBox TxtItemCode 
      Alignment       =   1  'Right Justify
      Height          =   345
      Index           =   1
      Left            =   8190
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   4440
      Width           =   1815
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Order_no_search2.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtremark 
      Alignment       =   1  'Right Justify
      Height          =   1020
      Left            =   14160
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   4680
      Width           =   7830
   End
   Begin VB.TextBox TxtCusID 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   4080
      Width           =   1830
   End
   Begin VB.TextBox txtorder_no 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8160
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   11865
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   3120
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
               Picture         =   "Order_no_search2.frx":0028
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search2.frx":03C2
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search2.frx":075C
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search2.frx":0AF6
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search2.frx":0E90
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search2.frx":122A
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search2.frx":15C4
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search2.frx":1B5E
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ «„— «‰ «Ã"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Index           =   2
         Left            =   6135
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   90
         Width           =   5280
      End
   End
   Begin MSDataListLib.DataCombo DBCboClientName 
      Height          =   315
      Left            =   2520
      TabIndex        =   8
      Top             =   4080
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   315
      Left            =   5520
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      Format          =   104726529
      CurrentDate     =   42221
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   315
      Left            =   10080
      TabIndex        =   10
      Top             =   8040
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "6"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2745
      Left            =   0
      TabIndex        =   14
      Top             =   720
      Width           =   11835
      _cx             =   20876
      _cy             =   4842
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"Order_no_search2.frx":1EF8
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   15
      Top             =   6240
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
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
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   1020
      TabIndex        =   16
      Top             =   6240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   1
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
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   6240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      BackStyle       =   0
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
   Begin MSDataListLib.DataCombo DCboItem 
      Height          =   315
      Left            =   480
      TabIndex        =   19
      Top             =   4440
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdItemSearch 
      Height          =   345
      Index           =   2
      Left            =   0
      TabIndex        =   20
      Top             =   4410
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "..."
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
      ButtonImage     =   "Order_no_search2.frx":20A5
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   3120
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      Format          =   104726529
      CurrentDate     =   42221
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2790
      _cx             =   4921
      _cy             =   2143
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
      Caption         =   " ÕœÌœ «·› —… «·“„‰Ì…"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   7
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
      Frame           =   0
      FrameStyle      =   5
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin MSComCtl2.DTPicker DTPickerAccFrom 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11265
            SubFormatType   =   3
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   28
         ToolTipText     =   "„‰  «—ÌŒ ﬁœÌ„"
         Top             =   240
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         _Version        =   393216
         CalendarBackColor=   -2147483624
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   104726531
         CurrentDate     =   42005
      End
      Begin MSComCtl2.DTPicker DTPickerAccTo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11265
            SubFormatType   =   3
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   29
         ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
         Top             =   600
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         _Version        =   393216
         CalendarBackColor=   -2147483624
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   104726531
         CurrentDate     =   42005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   2
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   4
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   285
         Width           =   555
      End
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ì"
      Height          =   375
      Left            =   4920
      TabIndex        =   26
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblitemid 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ «·’‰›"
      Height          =   315
      Index           =   34
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   4410
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·’‰›"
      Height          =   315
      Index           =   12
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«ﬂ » ﬂÊœ «·’‰› À„ ≈÷€ÿ ≈‰ —"
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
      Height          =   195
      Index           =   36
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   6390
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "„·ÕÊŸ…"
      Height          =   375
      Left            =   14160
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·»·œ"
      Height          =   375
      Left            =   8640
      TabIndex        =   5
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·⁄„Ì·"
      Height          =   375
      Left            =   9840
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "„‰"
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·√„—"
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "Order_no_search2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo

Private m_RetrunType As Integer
Public WithEvents FG1 As VSFlex8UCtl.VSFlexGrid
Attribute FG1.VB_VarHelpID = -1

Public WithEvents NewGrid As VSFlex8UCtl.VSFlexGrid
Attribute NewGrid.VB_VarHelpID = -1
'Public NewGrid As New ClsGrid
 
Public LngRow As Long

Public LngCol As Long

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If SystemOptions.UserInterface = ArabicInterface Then
                '   LblRes.Caption = "‰ ÌÃ… «·»ÕÀ = " & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                '   LblRes.Caption = "Search Result=" & rs.RecordCount
            End If
    
            If rs.RecordCount < 1 Then
                Fg.Clear flexClearScrollable, flexClearEverything
                Fg.Rows = 2

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    Msg = "NO Search Results Found...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation, App.title
                End If

                Exit Sub
            End If

            Retrive
            Fg.SetFocus

        Case 1
            clear_all Me
            Fg.Clear flexClearScrollable, flexClearEverything

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub

Private Sub CmdItemSearch_Click(Index As Integer)
  Load FrmItemSearch
            FrmItemSearch.RetrunType = 16
            Set FrmItemSearch.DcboItems = Me.DCboItem
            FrmItemSearch.show vbModal



End Sub

Private Sub DBCboClientName_Change()
    TxtCusID.Text = DBCboClientName.BoundText
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub fg_Click()
    On Error GoTo ErrTrap
       
    If Me.RetrunType = 1 Then
        FrmOutProductionOrder.TxtWorkOrderNO.Text = Fg.TextMatrix(Fg.Row, 1)
     
    End If
    
    If Me.RetrunType = 2 Then
        FrmExpenses3.txt_ORDER_NO = Fg.TextMatrix(Fg.Row, 1)
     
    End If
 
    If Me.RetrunType = 3 Then
        FrmExpenses5.txt_ORDER_NO = Fg.TextMatrix(Fg.Row, 1)
     
    End If
    
    If Me.RetrunType = 4 Then
        FrmProductionOrder.Retrive (val(Fg.TextMatrix(Fg.Row, 3)))
     
    End If
    
   If Me.RetrunType = 40 Then
        FrmQuality.TxtOderNo = (val(Fg.TextMatrix(Fg.Row, 1)))
     
    End If
     
     

    If Me.RetrunType = 5 Then
        FrmProductionOrder1.Retrive (val(Fg.TextMatrix(Fg.Row, 3)))
     
    End If
    
    If Me.RetrunType = 6 Then
        FrmInpoutWorkOrder.TXTOrderNO1.Text = Fg.TextMatrix(Fg.Row, 1)
     
    End If
    
    
    If Me.RetrunType = 7 Then
      FrmProductionAllocation.TxtWorkOrderNO.Text = Fg.TextMatrix(Fg.Row, 1)
      End If
      
     
   If Me.RetrunType = 8 Then
   
   
  
              If Not NewGrid Is Nothing Then
                  
                                   If Me.NewGrid.ColIndex("Code") <> -1 And Me.NewGrid.ColIndex("NProductionOrderNO") <> -1 Then
                                       NewGrid.TextMatrix(LngRow, NewGrid.ColIndex("NProductionOrderNO")) = (Fg.TextMatrix(Fg.Row, 1))
                        
                                   End If
            
             End If

            Unload Me
          
    End If
        
        
    If Me.RetrunType = 9 Then
      FrmProductionAllocation.TxtWorkOrderNOSub.Text = Fg.TextMatrix(Fg.Row, 1)
      End If
      
       If Me.RetrunType = 10 Then
      FrmShipmentRegestration.txtProductionOrder.Text = Fg.TextMatrix(Fg.Row, 1)
      End If
      If Me.RetrunType = 11 Then
      FrmPO9.Retrive val(Fg.TextMatrix(Fg.Row, 3))
      End If
        If Me.RetrunType = 12 Then
        FrmDestruction.TxtNoteSerial1.Text = Fg.TextMatrix(Fg.Row, 1)
     
    End If
    Exit Sub
ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    Fg.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        Fg.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With Fg
            If Me.RetrunType = 11 Then
                .TextMatrix(Num, .ColIndex("order_no")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
              Else
              .TextMatrix(Num, .ColIndex("order_no")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
              End If
                .TextMatrix(Num, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
            
                '    .TextMatrix(Num, .ColIndex("remark")) = IIf(IsNull(rs("remark").value), "", Trim(rs("remark").value))
                .TextMatrix(Num, .ColIndex("CusID")) = IIf(IsNull(rs("CusID").value), "", Trim(rs("CusID").value))

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                Else
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
                End If

               If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Num, .ColIndex("itemname")) = IIf(IsNull(rs("itemname").value), "", Trim(rs("itemname").value))
                Else
                    .TextMatrix(Num, .ColIndex("itemname")) = IIf(IsNull(rs("itemnamee").value), "", Trim(rs("itemnamee").value))
                End If
                

                .TextMatrix(Num, .ColIndex("Transaction_Date")) = IIf(IsNull(rs("Transaction_Date").value), "", Trim(rs("Transaction_Date").value))
                '   .TextMatrix(Num, .ColIndex("currency_code")) = IIf(IsNull(rs("Transaction_Date").value), "", Trim(rs("currency_code").value))
           
                '  .TextMatrix(Num, .ColIndex("countryid")) = IIf(IsNull(rs("countryid").value), "", (rs("countryid").value))
                '    .TextMatrix(Num, .ColIndex("CountryName")) = IIf(IsNull(rs("CountryName").value), "", Trim(rs("CountryName").value))
            
            End With

            rs.MoveNext
        Next Num

        ' Fg.AutoSize 0, Fg.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    fg_Click
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 
    Dim My_SQL As String
    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    Dcombos.GetItemsNames DCboItem, , , , True
    
    My_SQL = " select CountryID,CountryName from TblCountriesData"
 
    fill_combo Me.DataCombo4, My_SQL
 '   RetrunType = -1
 
    CenterForm Me
DTPickerAccFrom.value = Date
DTPickerAccTo.value = Date
 
    FormPostion Me, GetPostion
    Fg.WallPaper = BG.SearchWallpaper
    Set rs = New ADODB.Recordset
    DBCboClientName.BoundText = ""
    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    Set cSearchDcbo = Nothing

    FormPostion Me, SavePostion
    Set m_DcboItems = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer
 
    On Error GoTo ErrTrap

 '   StrSQL = "SELECT  dbo.Transactions.order_no,    dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.CusID, dbo.TblCustemers.CusName, "
 '   StrSQL = StrSQL + "  dbo.TblCustemers.CusNamee, dbo.Transactions.Transaction_Serial"
 '   StrSQL = StrSQL + " FROM         dbo.Transactions INNER JOIN"
 '   StrSQL = StrSQL + " dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"


StrSQL = " SELECT     dbo.Transactions.order_no, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.CusID, "
StrSQL = StrSQL + "  dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Transactions.Transaction_Serial, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode,"
StrSQL = StrSQL + "  dbo.TblItems.itemname , dbo.TblItems.ItemNamee  ,dbo.Transactions.NoteSerial1 "
StrSQL = StrSQL + "  FROM         dbo.TblItems INNER JOIN"
StrSQL = StrSQL + "  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID RIGHT OUTER JOIN"
StrSQL = StrSQL + "  dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
StrSQL = StrSQL + "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID where 1=1 "
    
    'StrSQL = StrSQL + " WHERE     (dbo.Transactions.Transaction_Type = 26)"
 
    If Me.RetrunType = 5 Then
        '     FrmProductionOrder1.Retrive (Val(FG.TextMatrix(FG.Row, 3)))
     
        StrSQL = StrSQL + " and     (dbo.Transactions.Transaction_Type = 32)"
    ElseIf Me.RetrunType = 11 Then
    StrSQL = StrSQL + " and     (dbo.Transactions.Transaction_Type = 61)"
    Else
        StrSQL = StrSQL + " and     (dbo.Transactions.Transaction_Type = 26)"

    End If
    
    If Me.DBCboClientName.BoundText <> "" And Me.DBCboClientName.Text <> "" Then
 
        StrWhere = StrWhere + " and TblCustemers.CusID =" & Me.DBCboClientName.BoundText & ""
 
    End If
 
    If Me.DCboItem.BoundText <> "" And Me.DCboItem.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.Transaction_Details.Item_ID =" & val(Me.DCboItem.BoundText)
 
    End If
    
    
    If Me.txtorder_no.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.Transactions.Transaction_Serial  like '%" & Me.txtorder_no.Text & "%'"
 
    End If
    
    
      
        If Me.DTPickerAccFrom <> Empty Or Me.DTPickerAccFrom <> Null Then
            StrWhere = StrWhere + " and    dbo.Transactions.Transaction_Date >=" & SQLDate(Me.DTPickerAccFrom, True) & ""
        End If

        If Me.DTPickerAccTo <> Empty Or Me.DTPickerAccTo <> Null Then
            StrWhere = StrWhere + " and dbo.Transactions.Transaction_Date <=" & SQLDate(Me.DTPickerAccTo, True) & ""
        End If
        
        

  '  StrWhere = StrWhere

    Build_Sql = StrSQL + StrWhere + " order by dbo.Transactions.Transaction_ID"
    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is Fg Then
            If Not Fg.TextMatrix(Fg.Row, 1) = "" Then
                fg_Click
                Unload Me
            End If

        Else
            Cmd_Click (0)
        End If
    End If

    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (2)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Public Property Get DcboItems() As DataCombo
    Set DcboItems = m_DcboItems
End Property

Public Property Set DcboItems(ByVal vNewValue As DataCombo)
    Set m_DcboItems = vNewValue
End Property

Public Property Get RetrunType() As Integer
    RetrunType = m_RetrunType
End Property

Public Property Let RetrunType(ByVal vNewValue As Integer)
    m_RetrunType = vNewValue
    ' 0 = Retrun in the Items Screen
    ' 1 = Retrun in the Data Combo
End Property

Private Sub ChangeLang()
    Me.Caption = "Search For Production Orders"
    Label1(2).Caption = Me.Caption
    Label2.Caption = "Order No"
 
    Label3.Caption = "Date"
    Label5.Caption = "Country"
    Label4.Caption = "Vendor"
    Label6.Caption = "Remark"
lbl(34).Caption = "item"

    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"
ELe(1).Caption = "From"
lbl(4).Caption = "From"
lbl(2).Caption = "To"

    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.Fg
        .TextMatrix(0, .ColIndex("order_no")) = "order no"
        '  .TextMatrix(0, .ColIndex("remark")) = "remark  "
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = " Date"
        '     .TextMatrix(0, .ColIndex("CountryName")) = "Country Name"
  
        '  .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub lblItemID_Change()
DCboItem.BoundText = val(Me.lblitemid.Caption)
End Sub

Private Sub TxtItemCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
          If KeyCode = vbKeyReturn Then
                If Trim(Me.txtItemCode(1).Text) = "" Then Exit Sub
                StrSQL = "Select ItemID From TblItems Where ItemCode='" & Trim(Me.txtItemCode(Index).Text) & "'"
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    DCboItem.BoundText = rs("ItemID").value
                Else
                    Msg = "·«ÌÊÃœ ’‰› „”Ã· »Â–« «·ﬂÊœ..!"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
            End If
            
End Sub
