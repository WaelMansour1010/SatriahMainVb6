VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frmchequesearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ ‘Ìﬂ"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14595
   Icon            =   "FrmchequeSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   14595
   Begin VB.TextBox txtbasedon 
      Alignment       =   1  'Right Justify
      Height          =   720
      Left            =   3240
      MaxLength       =   50
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Top             =   3360
      Width           =   7575
   End
   Begin VB.TextBox txtvalue 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   12360
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ComboBox CboItemCodeSearch 
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6750
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·’‰› «·„—«œ «·»ÕÀ ⁄‰Â ÌÕ ÊÏ ⁄·Ï Â–« «·’‰› ﬂ«Õœ „·Õﬁ« Â"
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
      Height          =   885
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   6030
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·’‰› «·„—«œ «·»ÕÀ ⁄‰Â ÌÕ ÊÏ ⁄·Ï Â–« «·’‰› ﬂ«Õœ „ﬂÊ‰« Â"
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
      Height          =   885
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   6300
      Width           =   6495
   End
   Begin VB.ComboBox CboArchive 
      Height          =   315
      Left            =   -1050
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   6180
      Width           =   1335
   End
   Begin VB.ComboBox CboGuar 
      Height          =   315
      Left            =   1080
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   6210
      Width           =   1305
   End
   Begin VB.ComboBox CboNameSearch 
      Height          =   315
      Left            =   -1050
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   5460
      Width           =   1515
   End
   Begin VB.ComboBox CboAttachedItem 
      Height          =   315
      Left            =   -1050
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   6540
      Width           =   1335
   End
   Begin VB.ComboBox CboAssbliedItem 
      Height          =   315
      Left            =   1080
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   6540
      Width           =   1305
   End
   Begin VB.ComboBox CboItemType 
      Height          =   315
      Left            =   3300
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   6210
      Width           =   1215
   End
   Begin VB.TextBox TxtItemID 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   12300
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2790
      Width           =   1215
   End
   Begin VB.ComboBox CboSerial 
      Height          =   315
      Left            =   -1050
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5820
      Width           =   1515
   End
   Begin VB.TextBox TxtItemName 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   3210
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2700
      Width           =   3705
   End
   Begin VB.TextBox XPTxtItemCode 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2610
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   6615
      Width           =   1395
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1950
      TabIndex        =   12
      Top             =   3795
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
      Left            =   930
      TabIndex        =   13
      Top             =   3795
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
      Left            =   30
      TabIndex        =   14
      Top             =   3795
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
   Begin MSDataListLib.DataCombo DCboGroupName 
      Height          =   315
      Left            =   1530
      TabIndex        =   5
      Top             =   5850
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid FG 
      Height          =   2565
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   14685
      _cx             =   25903
      _cy             =   4524
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
      Rows            =   50
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmchequeSearch.frx":030A
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
   Begin MSDataListLib.DataCombo Dcbobanks 
      Height          =   315
      Left            =   8760
      TabIndex        =   35
      Top             =   2760
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   1065
      Index           =   1
      Left            =   120
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   2835
      _cx             =   5001
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
         TabIndex        =   37
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
         Format          =   126615555
         CurrentDate     =   37357
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
         TabIndex        =   38
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
         Format          =   126615555
         CurrentDate     =   37357
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   15
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   285
         Width           =   555
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   12
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   600
         Width           =   555
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "»‰«¡ ⁄·Ï"
      Height          =   345
      Index           =   14
      Left            =   10620
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   3360
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌ„…"
      Height          =   345
      Index           =   13
      Left            =   13500
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   3360
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã«· «·»ÕÀ"
      Height          =   345
      Index           =   11
      Left            =   1590
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   6750
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·√—‘Ì›"
      Height          =   285
      Index           =   10
      Left            =   330
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   6210
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·÷„«‰"
      Height          =   285
      Index           =   9
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   6210
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã«· «·»ÕÀ"
      Height          =   345
      Index           =   8
      Left            =   510
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   5460
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·Õﬁ"
      Height          =   315
      Index           =   7
      Left            =   330
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   6540
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " Ã„Ì⁄"
      Height          =   285
      Index           =   6
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   6540
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·’‰›"
      Height          =   285
      Index           =   5
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   6360
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·‘Ìﬂ"
      Height          =   345
      Index           =   4
      Left            =   13560
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2790
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ÿ«„ «·”Ì—Ì«·"
      Height          =   315
      Index           =   2
      Left            =   510
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   5850
      Width           =   975
   End
   Begin VB.Label LblRes 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   3540
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   6570
      Width           =   1905
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„” ›Ìœ"
      Height          =   345
      Index           =   1
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2820
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·»‰ﬂ"
      Height          =   345
      Index           =   0
      Left            =   10620
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2790
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Ã„Ê⁄…"
      Height          =   285
      Index           =   3
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   6690
      Width           =   915
   End
End
Attribute VB_Name = "Frmchequesearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo

Private m_RetrunType As Integer

Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If SystemOptions.UserInterface = ArabicInterface Then
                LblRes.Caption = "‰ ÌÃ… «·»ÕÀ = " & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                LblRes.Caption = "Search Result=" & rs.RecordCount
            End If
    
            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.Rows = 2

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    Msg = "NO Search Results Found...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation, App.title
                End If

                Exit Sub
            End If

            FillGridWithData
            FG.SetFocus

        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything

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

Private Sub fg_Click()
    On Error GoTo ErrTrap

    If Not FG.TextMatrix(FG.Row, 1) = "" Then
        If Me.RetrunType = 0 Then
            PrintCheque.FindRec val(Me.FG.TextMatrix(Me.FG.Row, Me.FG.ColIndex("id")))
        
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Retrive()

    On Error GoTo ErrTrap
    Dim i As Integer
    'Frm2.Enabled = False

    'txtid.text = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)

    'TxtCheque_NO.text = IIf(IsNull(rs.Fields("Cheque_NO").value), "", rs.Fields("Cheque_NO").value)
    'Me.Dcbobanks.BoundText = IIf(IsNull(rs.Fields("BankID").value), "", rs.Fields("BankID").value)
    'Recorddate.value = IIf(IsNull(rs.Fields("Recorddate").value), "", rs.Fields("Recorddate").value)
    'DtpChequeDueDate.value = IIf(IsNull(rs.Fields("ChequeDate").value), "", rs.Fields("ChequeDate").value)
    'txtto.text = IIf(IsNull(rs.Fields("to").value), "", rs.Fields("to").value)
    'XPTxtVal.text = IIf(IsNull(rs.Fields("value").value), "", rs.Fields("value").value)
    'XPMTxtRemarks.text = IIf(IsNull(rs.Fields("basedOn").value), "", rs.Fields("basedOn").value)
    'txtreport_no.text = IIf(IsNull(rs.Fields("report_no").value), "", rs.Fields("report_no").value)

    ' LabCurrRec.Caption = rs.AbsolutePosition
    'LabCountRec.Caption = rs.RecordCount
    With FG

        For i = 1 To .Rows - 1

            If Trim(TxtCheque_NO.Text) = .TextMatrix(i, .ColIndex("Cheque_NO")) Then
                ' TxtSerial.text = .TextMatrix(i, .ColIndex("id"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    'Dim rs As ADODB.Recordset
    'Dim My_SQL As String

    'Set rs = New ADODB.Recordset
    'My_SQL = "select * From ChequePrintQry order by Id"
    'rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.FG
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("ID").value), "", rs.Fields("ID").value)
             
                .TextMatrix(i, .ColIndex("Cheque_NO")) = IIf(IsNull(rs.Fields("Cheque_NO").value), "", rs.Fields("Cheque_NO").value)
               
                .TextMatrix(i, .ColIndex("bankname")) = IIf(IsNull(rs.Fields("bankname").value), "", rs.Fields("bankname").value)
           
                .TextMatrix(i, .ColIndex("ChequeDate")) = IIf(IsNull(rs.Fields("ChequeDate").value), "", rs.Fields("ChequeDate").value)
           
                .TextMatrix(i, .ColIndex("ChequeDateH")) = IIf(IsNull(rs.Fields("ChequeDateH").value), "", rs.Fields("ChequeDateH").value)
            
                .TextMatrix(i, .ColIndex("To")) = IIf(IsNull(rs.Fields("To").value), "", rs.Fields("To").value)
            
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(rs.Fields("value").value), "", rs.Fields("value").value)
            
                .TextMatrix(i, .ColIndex("basedOn")) = IIf(IsNull(rs.Fields("basedOn").value), "", rs.Fields("basedOn").value)
            
                .TextMatrix(i, .ColIndex("report_no")) = IIf(IsNull(rs.Fields("report_no").value), "", rs.Fields("report_no").value)
            
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

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

    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemSGroups Me.DCboGroupName
    Set cSearchDcbo = New clsDCboSearch
    'cSearchDcbo.AllowWriting = False
    Dim My_SQL As String
    My_SQL = "select BankID,BankName From BanksData "
    fill_combo Dcbobanks, My_SQL

    Set cSearchDcbo.Client = Me.DCboGroupName

    If SystemOptions.UserInterface = ArabicInterface Then

        With Me.CboItemCodeSearch
            .Clear
            .AddItem "»ÕÀ „ÿ«»ﬁ"
            .AddItem "»ÕÀ „‰ «·»œ«Ì…"
            .AddItem "»ÕÀ „‰ «·‰Â«Ì…"
            .AddItem "»ÕÀ ›Ï «Ï „ﬂ«‰"
        End With

        With Me.CboSerial
            .Clear
            .AddItem "«·ﬂ·"
            .ItemData(0) = 0
            .AddItem "·Â ”Ì—Ì«·"
            .ItemData(1) = 1
            .AddItem "·Ì” ·Â ”Ì—Ì«·"
            .ItemData(2) = 2
        End With

        With Me.CboNameSearch
            .Clear
            .AddItem "„‰ «Ê· «·√”„"
            .AddItem "›Ï «Ï Ã“¡ „‰ «·√”„"
        End With

        With Me.CboItemType
            .Clear
            .AddItem "”·⁄…"
            .AddItem "Œœ„…"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboGuar
            .Clear
            .AddItem "·Â ÷„«‰"
            .AddItem "·Ì” ·Â ÷„«‰"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboArchive
            .Clear
            .AddItem "›Ï «·√—‘Ì›"
            .AddItem "·Ì” ›Ï «·√—‘Ì›"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboAssbliedItem
            .Clear
            .AddItem "’‰› „Ã„⁄"
            .AddItem "’‰› ⁄«œÏ"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboAttachedItem
            .Clear
            .AddItem "·Â «’‰«› „·Õﬁ…"
            .AddItem "·Ì” ·Â «’‰«› „·Õﬁ…"
            .AddItem "«·ﬂ·"
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With Me.CboItemCodeSearch
            .Clear
            .AddItem "Typical Search"
            .AddItem "From The Start"
            .AddItem "From The End"
            .AddItem "Any Where"
        End With

        With Me.CboSerial
            .Clear
            .AddItem "All"
            .ItemData(0) = 0
            .AddItem "Has Serial"
            .ItemData(1) = 1
            .AddItem "NO Serial"
            .ItemData(2) = 2
        End With

        With Me.CboNameSearch
            .Clear
            .AddItem "Start Name"
            .AddItem "Any Part of Name"
        End With

        With Me.CboItemType
            .Clear
            .AddItem "Goods"
            .AddItem "Services"
            .AddItem "All"
        End With

        With Me.CboGuar
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboArchive
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboAssbliedItem
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboAttachedItem
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

    End If

    CenterForm Me

    FormPostion Me, GetPostion
    FG.WallPaper = BG.SearchWallpaper
    Set rs = New ADODB.Recordset

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

    'On Error GoTo ErrTrap

    StrSQL = "Select * From ChequePrintQry "
    StrSQL = StrSQL + " Where id <> 0 "

    If val(Me.TxtItemID.Text) <> 0 Then
        StrSQL = StrSQL + " AND Cheque_NO LIKE'%" & Me.TxtItemID.Text & "%'"
    End If

    If Trim(Me.txtItemName.Text) <> "" Then
        StrWhere = StrWhere + " and [to] like '%" & Trim(Me.txtItemName.Text) & "%'"
    
    End If

    If Me.Dcbobanks.BoundText <> "" Then
        StrWhere = StrWhere + " and BankID =" & Me.Dcbobanks.BoundText & ""
    End If
 
    If IsNumeric(Me.TxtValue.Text) Then
          
        StrWhere = StrWhere + " and value=" & val(Me.TxtValue.Text)
    
    End If
 
    If Trim(Me.txtbasedon.Text) <> "" Then
        StrWhere = StrWhere + " and basedon like '%" & Trim(Me.txtbasedon.Text) & "%'"
    
    End If

    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    StrWhere = StrWhere + " and  ChequeDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    StrWhere = StrWhere + " and ChequeDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    'End If
 
    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is FG Then
            If Not FG.TextMatrix(FG.Row, 1) = "" Then
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
    Me.Caption = "Search For CHEQUE"
 
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    With Me.FG
        .TextMatrix(0, .ColIndex("id")) = "id"
        .TextMatrix(0, .ColIndex("Cheque_NO")) = "Cheque_NO ID"
        .TextMatrix(0, .ColIndex("bankname")) = " bankname"
        .TextMatrix(0, .ColIndex("ChequeDate")) = "ChequeDate"
        .TextMatrix(0, .ColIndex("ChequeDateH")) = "ChequeDateH Higri"

        .TextMatrix(0, .ColIndex("To")) = "To"
        .TextMatrix(0, .ColIndex("value")) = "value"
        .TextMatrix(0, .ColIndex("basedOn")) = "basedOn"

    End With

    lbl(4).Caption = "Cheque NO"
    lbl(0).Caption = "Bank"
 
    lbl(1).Caption = "To"
    lbl(13).Caption = "Value"
    lbl(14).Caption = "Base On"
 
End Sub

Private Sub TxtItemName_Change()

    If Trim$(Me.txtItemName.Text) = "" Then
        Me.lbl(8).Enabled = False
        Me.CboNameSearch.Enabled = False
    Else
        Me.lbl(8).Enabled = True
        Me.CboNameSearch.Enabled = True
    End If

End Sub

Private Sub XPTxtItemCode_Change()

    If Trim$(Me.XPTxtItemCode.Text) = "" Then
        Me.lbl(11).Enabled = False
        Me.CboItemCodeSearch.Enabled = False
    Else
        Me.lbl(11).Enabled = True
        Me.CboItemCodeSearch.Enabled = True
    End If

End Sub

