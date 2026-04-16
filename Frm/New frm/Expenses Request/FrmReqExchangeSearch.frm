VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmReqExchangeSearch 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11880
   Icon            =   "FrmReqExchangeSearch.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   5520
      Width           =   2175
      Begin XtremeSuiteControls.RadioButton Opt1 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   37
         Top             =   120
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "<"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Opt1 
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   38
         Top             =   120
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "="
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Opt1 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   120
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   ">"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
   End
   Begin VB.TextBox TxtValue 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   9840
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   5640
      Width           =   885
   End
   Begin VB.TextBox TxtTo 
      Alignment       =   1  'Right Justify
      Height          =   555
      Left            =   6000
      TabIndex        =   29
      Top             =   4920
      Width           =   4575
   End
   Begin VB.TextBox txtorder_no 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8700
      TabIndex        =   25
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "‰Ê⁄ «·„’—Ê›"
      Height          =   795
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   4800
      Width           =   5775
      Begin XtremeSuiteControls.RadioButton Opt 
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   23
         Top             =   240
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "‰ﬁœÌ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Opt 
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   24
         Top             =   240
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "‘Ìﬂ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Opt 
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   31
         Top             =   240
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ÕÊ«·Â"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   9
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
      MICON           =   "FrmReqExchangeSearch.frx":000C
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
      TabIndex        =   8
      Top             =   4680
      Width           =   7830
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   11745
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
               Picture         =   "FrmReqExchangeSearch.frx":0028
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmReqExchangeSearch.frx":03C2
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmReqExchangeSearch.frx":075C
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmReqExchangeSearch.frx":0AF6
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmReqExchangeSearch.frx":0E90
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmReqExchangeSearch.frx":122A
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmReqExchangeSearch.frx":15C4
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmReqExchangeSearch.frx":1B5E
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ÿ·» ’—›"
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
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   315
      Left            =   2880
      TabIndex        =   5
      Top             =   3600
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   98893825
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   315
      Left            =   10080
      TabIndex        =   6
      Top             =   6720
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
      TabIndex        =   10
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmReqExchangeSearch.frx":1EF8
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
      Left            =   2400
      TabIndex        =   11
      Top             =   5880
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
      Left            =   1380
      TabIndex        =   12
      Top             =   5880
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
      Left            =   480
      TabIndex        =   13
      Top             =   5880
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
   Begin MSDataListLib.DataCombo DcbEmp 
      Height          =   315
      Left            =   6000
      TabIndex        =   15
      Top             =   4440
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboEmpNameM 
      Height          =   315
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbTypeVisit1 
      Height          =   315
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbManager 
      Height          =   315
      Left            =   6000
      TabIndex        =   26
      Top             =   3960
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker ToDate 
      Height          =   315
      Left            =   480
      TabIndex        =   27
      Top             =   3600
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   98893825
      CurrentDate     =   38784
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„»·€"
      Height          =   315
      Index           =   1
      Left            =   11130
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   5670
      Width           =   765
   End
   Begin VB.Label lbltype 
      Alignment       =   1  'Right Justify
      Height          =   615
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   8760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "„’—Ê› «·Ï"
      Height          =   375
      Left            =   10380
      TabIndex        =   30
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï"
      Height          =   375
      Left            =   2280
      TabIndex        =   28
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "⁄»«—Â ⁄‰"
      Height          =   285
      Index           =   2
      Left            =   4650
      TabIndex        =   21
      Top             =   4440
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«·„œÌ— «·„»«‘—"
      Height          =   285
      Index           =   3
      Left            =   4890
      TabIndex        =   19
      Top             =   3960
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "„œÌ— «·«œ«—Â"
      Height          =   285
      Index           =   0
      Left            =   10560
      TabIndex        =   17
      Top             =   3960
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "ÿ«·» «·’—›"
      Height          =   285
      Index           =   10
      Left            =   10560
      TabIndex        =   16
      Top             =   4440
      Width           =   1125
   End
   Begin VB.Label lblitemid 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "„·ÕÊŸ…"
      Height          =   375
      Left            =   14160
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·»·œ"
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "„‰  «—ÌŒ"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   375
      Left            =   10440
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "FrmReqExchangeSearch"
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


Private Sub BtnFirst_Click()

End Sub

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
            XPDtbBill.value = ""
            ToDate.value = ""
 Opt(0).value = False
             Opt(1).value = False
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



End Sub

Private Sub DBCboClientName_Change()
    TxtCusID.Text = DBCboClientName.BoundText
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub fg_Click()
    On Error GoTo ErrTrap
If val(lbltype.Caption) = 1 Then
FrmPayments.Txtorder.Text = Fg.TextMatrix(Fg.Row, 2)

ElseIf val(lbltype.Caption) = 2 Then
FrmExpenses5.txtNoteSerial1.Text = Fg.TextMatrix(Fg.Row, 2)
FrmExpenses5.TxtorderID.Text = val(Fg.TextMatrix(Fg.Row, 1))
ElseIf val(lbltype.Caption) = 3 Then
FrmExpenses3.txtNoteSerial1.Text = Fg.TextMatrix(Fg.Row, 2)
FrmExpenses3.TxtorderID.Text = val(Fg.TextMatrix(Fg.Row, 1))
ElseIf val(lbltype.Caption) = 30 Then
 
FrmExpenses30.TxtorderID.Text = Fg.TextMatrix(Fg.Row, 2)
FrmExpenses30.Txtorder.Text = Fg.TextMatrix(Fg.Row, 2)
ElseIf val(lbltype.Caption) = 4 Then
FrmBoxDrawing.TxtOderSerial.Text = Fg.TextMatrix(Fg.Row, 2)
FrmBoxDrawing.Txtorder.Text = val(Fg.TextMatrix(Fg.Row, 1))


Else
      FrmTypeExchange.Retrive val(Fg.TextMatrix(Fg.Row, 1))
End If

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
            
               .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                .TextMatrix(Num, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                
                .TextMatrix(Num, .ColIndex("RecordDate")) = IIf(IsNull(rs("RecordDate").value), "", rs("RecordDate").value)
                .TextMatrix(Num, .ColIndex("ToPerson")) = IIf(IsNull(rs("ToPerson").value), "", rs("ToPerson").value)
                .TextMatrix(Num, .ColIndex("des")) = IIf(IsNull(rs("des").value), "", rs("des").value)
.TextMatrix(Num, .ColIndex("Price")) = val(IIf(IsNull(rs("Price").value), "", rs("Price").value))
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Num, .ColIndex("name")) = IIf(IsNull(rs("name").value), "", Trim(rs("name").value))
                    .TextMatrix(Num, .ColIndex("MangEmp_Name")) = IIf(IsNull(rs("MangEmp_Name").value), "", Trim(rs("MangEmp_Name").value))
                       .TextMatrix(Num, .ColIndex("MangEmpEmp_Name")) = IIf(IsNull(rs("MangEmpEmp_Name").value), "", Trim(rs("MangEmpEmp_Name").value))
                        .TextMatrix(Num, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", Trim(rs("Emp_Name").value))
                Else
                .TextMatrix(Num, .ColIndex("name")) = IIf(IsNull(rs("namee").value), "", Trim(rs("namee").value))
                    .TextMatrix(Num, .ColIndex("MangEmp_Name")) = IIf(IsNull(rs("MangEmp_Name").value), "", Trim(rs("MangEmp_Name").value))
                     .TextMatrix(Num, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", Trim(rs("Emp_Namee").value))
                      .TextMatrix(Num, .ColIndex("MangEmpEmp_Name")) = IIf(IsNull(rs("MangEmpEmp_Name").value), "", Trim(rs("MangEmpEmp_Name").value))
                End If

            
            End With

            rs.MoveNext
        Next Num

        Fg.AutoSize 0, Fg.Cols - 1, False
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
 XPDtbBill.value = ""
            ToDate.value = ""
            
            XPDtbBill.value = Date
             ToDate.value = Date
            Opt(0).value = False
             Opt(1).value = False
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 
    Dim My_SQL As String
    Set Dcombos = New ClsDataCombos

    
  Dcombos.GetTypeExchange Me.DcbTypeVisit1
      Dcombos.GetEmployees Me.DcboEmpNameM
        Dcombos.GetEmployees Me.DcbManager
         Dcombos.GetEmployees Me.DcbEmp
      
    RetrunType = -1
 
    CenterForm Me

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

StrSQL = " SELECT  NoteSerial1,   dbo.TblExchange.Id, dbo.TblExchange.RecordDate, dbo.TblExchange.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
 StrSQL = StrSQL & "                     dbo.TblExchange.TypeExch, dbo.TblDataTypeExchange.name, dbo.TblDataTypeExchange.namee, dbo.TblExchange.Price, dbo.TblExchange.ToPerson,"
 StrSQL = StrSQL & "                     dbo.TblExchange.EmpID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Fullcode,"
 StrSQL = StrSQL & "                     dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
 StrSQL = StrSQL & "                     dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblExchange.ManagerID,"
 StrSQL = StrSQL & "                     TblEmployee_1.Emp_Code AS ManEmp_Code, TblEmployee_1.Emp_Name AS MangEmp_Name, TblEmployee_1.Emp_Name1 AS mangEmp_Name1,"
 StrSQL = StrSQL & "                     TblEmployee_1.Emp_Name2 AS MangEmp_Name2, TblEmployee_1.Emp_Name3 AS MangEmp_Name3, TblEmployee_1.Emp_Name4 AS MangEmp_Name4,"
 StrSQL = StrSQL & "                     TblEmployee_1.Fullcode AS MangFullcode, TblEmployee_1.Emp_Namee4 AS MangEmp_Namee4, TblEmployee_1.Emp_Namee3 AS MangEmp_Namee3,"
 StrSQL = StrSQL & "                     TblEmployee_1.Emp_Namee2 AS MangEmp_Namee2, TblEmployee_1.Emp_Namee1 AS MangEmp_Namee1, TblEmployee_1.Emp_Namee AS MangEmp_Namee,"
 StrSQL = StrSQL & "                     dbo.TblExchange.MempID, TblEmployee_2.Emp_Code AS ManEmpEmp_Code, TblEmployee_2.Emp_Name AS MangEmpEmp_Name,"
 StrSQL = StrSQL & "                     TblEmployee_2.Emp_Name1 AS angEmpEmp_Name1, TblEmployee_2.Emp_Name2 AS MangEmpEmp_Name2,"
 StrSQL = StrSQL & "                     TblEmployee_2.Emp_Name3 AS MangEmpEmp_Name3, TblEmployee_2.Emp_Name4 AS MangEmpEmp_Name4, TblEmployee_2.Fullcode AS angEmpFullcode,"
 StrSQL = StrSQL & "                     TblEmployee_2.Emp_Namee4 AS MangEmpEmp_Namee4, TblEmployee_2.Emp_Namee3 AS angeEmpEmp_Namee3,"
 StrSQL = StrSQL & "                     TblEmployee_2.Emp_Namee2 AS MangEmpEmp_Namee2, TblEmployee_2.Emp_Namee1 AS angEmpEmp_Namee1,"
 StrSQL = StrSQL & "                     TblEmployee_2.Emp_Namee AS MangEMpEmp_Namee, dbo.TblExchange.Des, dbo.TblExchange.UserID, dbo.TblExchange.Type"
 StrSQL = StrSQL & " FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblExchange LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee TblEmployee_2 ON dbo.TblExchange.MempID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblExchange.ManagerID = TblEmployee_1.Emp_ID ON"
 StrSQL = StrSQL & "                     dbo.TblEmployee.Emp_ID = dbo.TblExchange.EmpID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblDataTypeExchange ON dbo.TblExchange.TypeExch = dbo.TblDataTypeExchange.Id ON dbo.TblBranchesData.branch_id = dbo.TblExchange.BranchID"
StrSQL = StrSQL & " Where (1 = 1)"
 
       If CheckAprroveScreen("FrmTypeExchange") = True And val(lbltype.Caption) <> 0 Then

  StrSQL = StrSQL & "  and Approved = 1"
  End If
  
    If SystemOptions.MonyeIssueVchrNoMust = True And val(lbltype.Caption) <> 0 Then
    StrWhere = StrWhere & "   and price >0 "
    Else
    StrWhere = ""
    
    End If
     
     

    
    If Me.Opt(0).value = True Then
 
        StrWhere = StrWhere + " and dbo.TblExchange.Type = 0 "
 
    End If
    If Me.Opt(1).value = True Then
 
        StrWhere = StrWhere + " and dbo.TblExchange.Type = 1 "
 
    End If
       If Me.Opt(2).value = True Then
 
        StrWhere = StrWhere + " and dbo.TblExchange.Type = 2 "
 
    End If
    
    If Me.DcbManager.BoundText <> "" And Me.DcbManager.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblExchange.ManagerID =" & Me.DcbManager.BoundText & ""
 
    End If
    If Me.DcboEmpNameM.BoundText <> "" And Me.DcboEmpNameM.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblExchange.MempID =" & Me.DcboEmpNameM.BoundText & ""
 
    End If
    If Me.DcbEmp.BoundText <> "" And Me.DcbEmp.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblExchange.EmpID =" & val(Me.DcbEmp.BoundText)
 
    End If
    
    If Me.DcbTypeVisit1.BoundText <> "" And Me.DcbTypeVisit1.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblExchange.TypeExch =" & val(Me.DcbTypeVisit1.BoundText)
 
    End If
    
      If Me.txtto.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblExchange.ToPerson  like '%" & Me.txtto.Text & "%'"
 
    End If
    
    If Me.TXTOrDer_no.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblExchange.NoteSerial1 ='" & Me.TXTOrDer_no.Text & "'"
 
    End If
 If Not IsNull(Me.XPDtbBill.value) Then
       
            StrWhere = StrWhere & " AND dbo.TblExchange.RecordDate >=" & SQLDate(Me.XPDtbBill.value, True) & ""
        End If
        
         If Not IsNull(Me.ToDate.value) Then
       
            StrWhere = StrWhere & " AND dbo.TblExchange.RecordDate <=" & SQLDate(Me.ToDate.value, True) & ""
        End If
        If val(lbltype.Caption) = 1 Then
 StrWhere = StrWhere + " and approved=1"
End If








 If val(Me.TxtValue.Text) > 0 Then
        If Me.opt1(1).value = True Then
 
                StrWhere = StrWhere + " AND dbo.TblExchange.Price =" & val(Me.TxtValue.Text) & ""
            

        ElseIf Me.opt1(0).value = True Then
 
                StrWhere = StrWhere + " AND dbo.TblExchange.Price >" & val(Me.TxtValue.Text) & ""
            
        Else
 
                StrWhere = StrWhere + " AND dbo.TblExchange.Price <" & val(Me.TxtValue.Text) & ""
              
        End If
    End If



    StrWhere = StrWhere + " order by dbo.TblExchange.Id"


    Build_Sql = StrSQL + StrWhere
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





Private Sub ChangeLang()
    Me.Caption = "Search For Exchange Request"
    Label1(2).Caption = Me.Caption
    Label2.Caption = "Order No"
 lbl(0).Caption = "Manager"
 lbl(3).Caption = "Manager"
     lbl(2).Caption = "That Vis"
         lbl(10).Caption = "Requser Exchange"
    Label3.Caption = "From Date"
 
    Label4.Caption = "TO"
Frame2.Caption = "Data Exchange "
   Opt(2).RightToLeft = False
Opt(0).RightToLeft = False
Opt(1).RightToLeft = False
Opt(0).Caption = "Cash"
Opt(1).Caption = "Check"
Opt(2).Caption = "Transfer"
Label7.Caption = "To"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    'OptType(0).Caption = "Start of the name"
     With Fg
     .TextMatrix(0, .ColIndex("NoteSerial1")) = "OrderNo"
     .TextMatrix(0, .ColIndex("NumIndex")) = "Ser"
                .TextMatrix(0, .ColIndex("id")) = "OrderNo"
                .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
                .TextMatrix(0, .ColIndex("MangEmp_Name")) = "Manger"
                .TextMatrix(0, .ColIndex("MangEmpEmp_Name")) = "Manger"
                .TextMatrix(0, .ColIndex("Emp_Name")) = "RequExchange"
                .TextMatrix(0, .ColIndex("ToPerson")) = "ToPerson"
                .TextMatrix(0, .ColIndex("Type")) = "Type"
                .TextMatrix(0, .ColIndex("name")) = "That Vis"
                 .TextMatrix(0, .ColIndex("Price")) = "Price "
                 .TextMatrix(0, .ColIndex("des")) = "Description"
                
                
End With
End Sub

