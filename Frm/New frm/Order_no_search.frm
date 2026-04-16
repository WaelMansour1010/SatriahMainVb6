VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form Order_no_search 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»ÕÀ  «Ê«„—  "
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9675
   Icon            =   "Order_no_search.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   9675
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox txtContainerNo 
      Height          =   345
      Left            =   6150
      TabIndex        =   32
      Top             =   5040
      Width           =   1665
   End
   Begin VB.TextBox TxtPhone 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   840
      TabIndex        =   23
      Top             =   4575
      Width           =   1650
   End
   Begin VB.TextBox TxtCashCustomerName 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   4560
      Width           =   4335
   End
   Begin VB.TextBox TxtStoreID 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   6135
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   5505
      Width           =   1665
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   8790
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
      MICON           =   "Order_no_search.frx":000C
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
      Height          =   780
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   5970
      Width           =   7830
   End
   Begin VB.TextBox TxtCusID 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   4080
      Width           =   1830
   End
   Begin VB.TextBox txtorder_no 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6120
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
      Width           =   9105
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
               Picture         =   "Order_no_search.frx":0028
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":03C2
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":075C
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":0AF6
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":0E90
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":122A
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":15C4
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":1B5E
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ  «Ê«„—  "
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
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   90
         Width           =   5280
      End
   End
   Begin MSDataListLib.DataCombo DBCboClientName 
      Height          =   315
      Left            =   720
      TabIndex        =   8
      Top             =   4080
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   315
      Left            =   3480
      TabIndex        =   9
      Top             =   3600
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      Format          =   214761473
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   3600
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
      Width           =   9075
      _cx             =   16007
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
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"Order_no_search.frx":1EF8
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
      Top             =   8010
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
      Top             =   8010
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
      Top             =   8010
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
   Begin MSDataListLib.DataCombo DCboStoreName 
      Height          =   315
      Left            =   120
      TabIndex        =   20
      Top             =   5520
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "7"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdCusSearch 
      Height          =   345
      Index           =   0
      Left            =   0
      TabIndex        =   26
      Top             =   4080
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "Order_no_search.frx":21A9
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6690
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
         Format          =   103612419
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
         Format          =   103612419
         CurrentDate     =   42005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   4
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   285
         Width           =   555
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   2
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   600
         Width           =   555
      End
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Q.Ref"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   7770
      TabIndex        =   33
      Top             =   5130
      Width           =   1590
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÌ"
      Height          =   345
      Index           =   37
      Left            =   7650
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   4575
      Width           =   1425
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ·Ì›Ê‰"
      Height          =   390
      Index           =   84
      Left            =   2535
      TabIndex        =   24
      Top             =   4575
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„Œ“‰"
      Height          =   270
      Index           =   8
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   5490
      Width           =   1065
   End
   Begin VB.Label lblSpecificsearch 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   495
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   7050
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "„·ÕÊŸ…"
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   6090
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·»·œ"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·⁄„Ì· «·„Ê—œ"
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "«· «—ÌŒ"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·”‰œ"
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "Order_no_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo

Private m_RetrunType As Integer
Public mNoteSerial As String
Public mTransactionID  As Long


Private Sub BtnFirst_Click()

End Sub

Private Sub Cmd_Click(index As Integer)
    On Error GoTo ErrTrap

    Select Case index

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
                FG.Clear flexClearScrollable, flexClearEverything
                FG.rows = 2

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    Msg = "NO Search Results Found...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation, App.Title
                End If

                Exit Sub
            End If

            Retrive
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

End Sub

Private Sub CmdCusSearch_Click(index As Integer)
              Load FrmCustemerSearch
            FrmCustemerSearch.SearchType = 100
            FrmCustemerSearch.RetrunType = 100
            Set FrmCustemerSearch.DcboCustomers = Me.DBCboClientName
            
            FrmCustemerSearch.show vbModal
End Sub

Private Sub DBCboClientName_Change()
    TxtCusID.text = DBCboClientName.BoundText
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub DCboStoreName_Change()
 TxtStoreID.text = getStoreCoding(val(DCboStoreName.BoundText))
End Sub

Private Sub fg_Click()
    On Error GoTo ErrTrap

    If Not FG.TextMatrix(FG.row, 1) = "" Then
        If Me.RetrunType = 0 Then
            FrmExpenses5.Txt_order_no.text = FG.TextMatrix(FG.row, 1)
        ElseIf Me.RetrunType = 1 Then
            FrmExpenses3.Txt_order_no.text = FG.TextMatrix(FG.row, 1)
        ElseIf Me.RetrunType = 2 Then
            FrmPayments.Txt_order_no.text = FG.TextMatrix(FG.row, 1)
        ElseIf Me.RetrunType = 70 Then
            FrmPayments.Txt_order_no.text = FG.TextMatrix(FG.row, 1)
            FrmPayments.txtTradingContractID.text = FG.TextMatrix(FG.row, FG.ColIndex("Transaction_ID"))
            FrmPayments.txtAcceptianPeriod.text = FG.TextMatrix(FG.row, FG.ColIndex("AcceptianPeriod"))
            
            FrmPayments.DcboCreditSide.BoundText = FG.TextMatrix(FG.row, FG.ColIndex("AcceptAccount_Code"))
            FrmPayments.DCproject.BoundText = val(FG.TextMatrix(FG.row, FG.ColIndex("project_id")))
            
            FrmPayments.DBCboClientName.BoundText = val(FG.TextMatrix(FG.row, FG.ColIndex("CusId")))
            
              FrmPayments.CboPayMentType.ListIndex = 4
              FrmPayments.DcbAccount.BoundText = FG.TextMatrix(FG.row, FG.ColIndex("AcceptAccount_Code"))
              
            FrmPayments.TXT_order_no_Validate True
        ElseIf Me.RetrunType = 3 Then
            FrmBillBuy.Txt_order_no.text = FG.TextMatrix(FG.row, 1)
            FrmBillBuy.Dccurrency.text = FG.TextMatrix(FG.row, 11)
            FrmBillBuy.poTransaction_ID = FG.TextMatrix(FG.row, 10)
            
        
        ElseIf Me.RetrunType = 73 Then
            FrmPO3.TxtPONo.text = FG.TextMatrix(FG.row, 1)
            FrmPO3.Dccurrency.text = FG.TextMatrix(FG.row, 11)
            FrmPO3.poTransaction_ID = FG.TextMatrix(FG.row, 10)
            
        
        
        ElseIf Me.RetrunType = 4 Then
        
        

            With FrmExpenses5.Fg_Journal
                .TextMatrix(.row, .ColIndex("Order_No")) = FG.TextMatrix(FG.row, 1)
            End With
     
        ElseIf Me.RetrunType = 5 Then

            With FrmExpenses3.Fg_Journal
                .TextMatrix(.row, .ColIndex("Order_No")) = FG.TextMatrix(FG.row, 1)
            End With
    
        ElseIf Me.RetrunType = 6 Then
            FrmProductionOrder.Txt_order_no.text = FG.TextMatrix(FG.row, 1)
        ElseIf Me.RetrunType = 38 Then
            FrmProductionOrder.Txt_order_no.text = FG.TextMatrix(FG.row, 1)
     
      ElseIf Me.RetrunType = 61 Then
            FrmProductionOrder.TxtResProductionNo.text = FG.TextMatrix(FG.row, 1)
           FrmProductionOrder.ProkerId.text = FG.TextMatrix(FG.row, FG.ColIndex("Transaction_ID"))
           
        ElseIf Me.RetrunType = 7 Then
            FrmOutProductionOrder.TxtWorkOrderNO.text = FG.TextMatrix(FG.row, 1)
     
        ElseIf Me.RetrunType = 8 Then
            frmsalebill.TXTOrDer_no.text = FG.TextMatrix(FG.row, 1)
     
        ElseIf Me.RetrunType = 9 Then
            FrmInpout.Txt_order_no.text = FG.TextMatrix(FG.row, 1)
     
     
          ElseIf Me.RetrunType = 10 Then
            FrmPO6.Txt_order_no.text = FG.TextMatrix(FG.row, 1)
            
          ElseIf Me.RetrunType = 89 Then
            FrmOut.Txt_order_no.text = FG.TextMatrix(FG.row, 1)
            
          ElseIf Me.RetrunType = 99 Then
            FrmOut.Txt_order_no.text = FG.TextMatrix(FG.row, 1)
             ElseIf Me.RetrunType = 98 Then
            FrmOut.Txt_order_no.text = FG.TextMatrix(FG.row, 1)
            
          ElseIf Me.RetrunType = 11 Then
            FrmOut.Txt_order_no.text = FG.TextMatrix(FG.row, 1)
                 
                 ElseIf Me.RetrunType = 12 Then
            FrmShipmentOrder.TxtPONo.text = FG.TextMatrix(FG.row, 1)
                         ElseIf Me.RetrunType = 14 Then
            FrmPO9.TxtPONo.text = FG.TextMatrix(FG.row, 1)
                   
                         ElseIf Me.RetrunType = 15 Then
            FrmTypeExchange.TxtOrderNo.text = FG.TextMatrix(FG.row, 1)
                   FrmTypeExchange.txtTransaction_ID.text = FG.TextMatrix(FG.row, 9)
                   
                   FrmTypeExchange.DCboCashType121.ListIndex = 1
                  FrmTypeExchange.DBCboClientName.BoundText = FG.TextMatrix(FG.row, FG.ColIndex("CusID"))
                  
                   
                             ElseIf Me.RetrunType = 16 Then
            FrmProductionPlan.TxtNoteSerial.text = FG.TextMatrix(FG.row, 1)
                  
               ElseIf Me.RetrunType = 17 Then
             FrmShowPrice.Txt_order_no.text = FG.TextMatrix(FG.row, 1)
                 ElseIf Me.RetrunType = 18 Then
             FrmPO3.Retrive FG.TextMatrix(FG.row, FG.ColIndex("Transaction_ID"))
                   
                 ElseIf Me.RetrunType = 19 Then
             Projects.TXTOrDer_no.text = FG.TextMatrix(FG.row, 1)
            
                             ElseIf Me.RetrunType = 20 Then
             FrmPO10.TxtPO6.text = FG.TextMatrix(FG.row, 1) ' Fg.TextMatrix(Fg.Row, Fg.ColIndex("Transaction_ID"))
   
   
        End If
    
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    Dim Transaction_Type As Integer
    
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
            Transaction_Type = IIf(IsNull(rs("Transaction_Type").value), "", rs("Transaction_Type").value)
           
            If Transaction_Type = 70 Then
                .TextMatrix(Num, .ColIndex("project_id")) = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)
                .TextMatrix(Num, .ColIndex("AcceptAccount_Code")) = IIf(IsNull(rs("AcceptAccount_Code").value), "", rs("AcceptAccount_Code").value)
                .TextMatrix(Num, .ColIndex("AcceptianPeriod")) = IIf(IsNull(rs("AcceptianPeriod").value), "", rs("AcceptianPeriod").value)
            End If
              If Transaction_Type = 21 Or Transaction_Type = 42 Or Transaction_Type = 70 Or Transaction_Type = 26 Or Transaction_Type = 28 Or Transaction_Type = 38 Then
                .TextMatrix(Num, .ColIndex("order_no")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
            Else
                 .TextMatrix(Num, .ColIndex("order_no")) = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
            End If
            
            .TextMatrix(Num, .ColIndex("CashCustomerName")) = IIf(IsNull(rs("CashCustomerName").value), "", (rs("CashCustomerName").value))
                .TextMatrix(Num, .ColIndex("remark")) = IIf(IsNull(rs("remark").value), "", Trim(rs("remark").value))
                .TextMatrix(Num, .ColIndex("CusID")) = IIf(IsNull(rs("CusID").value), "", Trim(rs("CusID").value))
                      .TextMatrix(Num, .ColIndex("ContainerNo")) = IIf(IsNull(rs("ContainerNo").value), "", Trim(rs("ContainerNo").value))

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                Else
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
                End If


'       If SystemOptions.UserInterface = ArabicInterface Then
'                    .TextMatrix(Num, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", Trim(rs("StoreName").value))
'                Else
'                    .TextMatrix(Num, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreNamee").value), "", Trim(rs("StoreNamee").value))
'                End If
'
                
                .TextMatrix(Num, .ColIndex("Transaction_Date")) = IIf(IsNull(rs("Transaction_Date").value), "", Trim(rs("Transaction_Date").value))
                
                '.TextMatrix(Num, .ColIndex("currency_code")) = IIf(IsNull(rs("currency_code").value), "", Trim(rs("currency_code").value))
           
                .TextMatrix(Num, .ColIndex("countryid")) = IIf(IsNull(rs("countryid").value), "", (rs("countryid").value))
                .TextMatrix(Num, .ColIndex("CountryName")) = IIf(IsNull(rs("CountryName").value), "", Trim(rs("CountryName").value))
                .TextMatrix(Num, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("Transaction_ID").value), "", (rs("Transaction_ID").value))
                
            'Transaction_ID
            End With

            rs.MoveNext
        Next Num

  FG.AutoSize 0, FG.Cols - 1, False
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
DTPickerAccFrom.value = Date
DTPickerAccTo.value = Date

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 
    Dim My_SQL As String
    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
Dcombos.GetStores Me.DCboStoreName
XPDtbBill.value = Date


    My_SQL = " select CountryID,CountryName from TblCountriesData"
 
    fill_combo Me.DataCombo4, My_SQL
    RetrunType = -1
 
    CenterForm Me

    FormPostion Me, GetPostion
    FG.WallPaper = BG.SearchWallpaper
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


StrSQL = " SELECT  distinct CashCustomerName,   dbo.Transactions.Transaction_ID, dbo.Transactions.NoteSerial1,dbo.Transactions.Transaction_Type,  dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.CusID, dbo.Transactions.countryid, "
StrSQL = StrSQL & "  dbo.Transactions.order_no, dbo.Transactions.remark, dbo.TblCountriesData.CountryName, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
StrSQL = StrSQL & "  dbo.Transactions.Closed,Transactions.ContainerNo, dbo.Transactions.Currency_id, dbo.currency.code AS currency_code, dbo.Transactions.StoreID, dbo.TblStore.StoreName,"
StrSQL = StrSQL & "  dbo.TblStore.StoreNamee , dbo.TblStore.code"
StrSQL = StrSQL & " FROM         dbo.Transactions LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.currency ON dbo.Transactions.Currency_id = dbo.currency.id LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
StrSQL = StrSQL & " dbo.TblCountriesData ON dbo.Transactions.countryid = dbo.TblCountriesData.CountryID"
StrSQL = StrSQL & " Inner join Transaction_Details On Transaction_Details.Transaction_ID =Transactions.Transaction_ID"
    If Me.RetrunType = 3 Then
      '  StrSQL = "Select * From Order_no_details "

        If lblSpecificsearch.Caption = "1" Then
            StrSQL = StrSQL + " Where   ( Transaction_Type=29  )"
            
'                StrSQL = StrSQL + " and (Transactions.NoteSerial1 Not In (Select T.order_no from Transactions T where T.Transaction_Type = 22"
'                If mTransactionID <> 0 Then
'                    StrSQL = StrSQL + " And T.Transaction_ID <> " & mTransactionID
'                End If
                
                StrSQL = StrSQL & " and dbo.[GetBalanceQtyPO3] (Transaction_Details.Item_ID,Transactions.NoteSerial1," & mTransactionID & ") <> 0"
                'StrSQL = StrSQL + " ))"
                
        ElseIf lblSpecificsearch.Caption = "2" Then
            StrSQL = StrSQL + " Where   ( Transaction_Type=17)"
        End If

      If Me.DTPickerAccFrom <> Empty Or Me.DTPickerAccFrom <> Null Then
            StrWhere = StrWhere + " and    dbo.Transactions.Transaction_Date >=" & SQLDate(Me.DTPickerAccFrom, True) & ""
        End If
        
If CheckAprroveScreen("FrmPO10") = True Then
StrSQL = StrSQL + " and approved=1"
End If
ElseIf Me.RetrunType = 98 Then
      '  StrSQL = "Select * From Order_no_details "
      
        StrSQL = StrSQL + " Where   ( Transaction_Type=28)"
ElseIf Me.RetrunType = 38 Then
      '  StrSQL = "Select * From Order_no_details "
      
        StrSQL = StrSQL + " Where   ( Transaction_Type=38)"
ElseIf Me.RetrunType = 89 Or Me.RetrunType = 73 Then
      '  StrSQL = "Select * From Order_no_details "
      
        StrSQL = StrSQL + " Where   ( Transaction_Type=29)"

    ElseIf Me.RetrunType = 8 Or Me.RetrunType = 16 Then
      '  StrSQL = "Select * From Order_no_details "
      
        StrSQL = StrSQL + " Where   ( Transaction_Type=" & val(lblSpecificsearch.Caption) & ")"

    ElseIf Me.RetrunType = 7 Then
      '  StrSQL = "Select * From Order_no_details "
        StrSQL = StrSQL + " Where  Transaction_Type=26"
ElseIf Me.RetrunType = 99 Then
      '  StrSQL = "Select * From Order_no_details "
        StrSQL = StrSQL + " Where  Transaction_Type=26"

    ElseIf Me.RetrunType = 10 Or Me.RetrunType = 12 Or Me.RetrunType = 13 Or Me.RetrunType = 14 Or Me.RetrunType = 60 Or Me.RetrunType = 61 Then
      '  StrSQL = "Select * From Order_no_details "
      
        StrSQL = StrSQL + " Where   ( Transaction_Type=" & val(lblSpecificsearch.Caption) & ")"
     ElseIf Me.RetrunType = 18 Or Me.RetrunType = 19 Then
 
  StrSQL = StrSQL + " Where ( Transaction_Type=6  )"
    Else
      '  StrSQL = "Select * From Order_no_details "
        StrSQL = StrSQL + " Where ( Transaction_Type=6  or Transaction_Type=29 or    Transaction_Type=17)"

    End If

 If Me.RetrunType = 11 Or Me.RetrunType = 12 Or Me.RetrunType = 99 Then
     
     If (Me.TXTOrDer_no.text) <> "" Then
        StrSQL = StrSQL + " AND NoteSerial1 like'%" & Me.TXTOrDer_no.text & "%'"
    End If
    
 Else
 
 
    If (Me.TXTOrDer_no.text) <> "" Then
        StrSQL = StrSQL + " AND Transactions.order_no like'%" & Me.TXTOrDer_no.text & "%'"
    End If

End If

    If DataCombo4.BoundText <> "" Then
        StrWhere = StrWhere + " and countryid =" & DataCombo4.BoundText & " "
    
    End If
    

    If Me.DBCboClientName.BoundText <> "" Then
 
        StrWhere = StrWhere + " and Transactions.CusID =" & Me.DBCboClientName.BoundText & ""
 
    End If

    If Me.DCboStoreName.BoundText <> "" Then
 
        'StrWhere = StrWhere + " and dbo.Transactions.storeid  =" & val(Me.DCboStoreName.BoundText)
 
    End If
    
    If Trim(Me.txtRemark.text) <> "" Then
    
        StrWhere = StrWhere + " and Transactions.remark like '%" & Trim(Me.txtRemark.text) & "%'"
     
    End If


    If Trim(Me.txtContainerNo.text) <> "" Then
    
        StrWhere = StrWhere + " and Transactions.ContainerNo like '%" & Trim(Me.txtContainerNo.text) & "%'"
     
    End If
    
    'StrSQL = StrSQL & "  order by Transaction_ID "
If Me.RetrunType = 15 Then
 StrWhere = StrWhere & " and requestOrOrder=0 "
 End If


 If TxtCashCustomerName.text <> "" Then
           
                StrWhere = StrWhere + " and dbo.Transactions.CashCustomerName like '%" & (TxtCashCustomerName.text) & "%'"
           
      End If
        If TxtPhone.text <> "" Then
         
                StrWhere = StrWhere + " and dbo.Transactions.CashCustomerPhone like '%" & (TxtCashCustomerPhone.text) & "%'"
              
      End If
      

      
        If Me.DTPickerAccFrom <> Empty Or Me.DTPickerAccFrom <> Null Then
            StrWhere = StrWhere + " and    dbo.Transactions.Transaction_Date >=" & SQLDate(Me.DTPickerAccFrom, True) & ""
        End If

        If Me.DTPickerAccTo <> Empty Or Me.DTPickerAccTo <> Null Then
            StrWhere = StrWhere + " and dbo.Transactions.Transaction_Date <=" & SQLDate(Me.DTPickerAccTo, True) & ""
        End If
        
 


    Build_Sql = StrSQL + StrWhere & "  order by dbo.Transactions.Transaction_ID "
    
    If Me.RetrunType = 70 Then Build_Sql = Build_Sql2
    Exit Function
ErrTrap:
End Function



Private Function Build_Sql2()
    Dim StrSQL As String
    Dim MySQL As String
    
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer
 
    On Error GoTo ErrTrap

    
                


MySQL = "  SELECT BanksData_1.BankName,TblLC.FromDate Transaction_Date,ContainerNo = "",  dbo.currency.name AS CurrencyName, dbo.LCTypes.name AS TypeName, dbo.LCTypes.namee AS TypeNameE, dbo.TblLC.CountryId,TblLC.prifix,TblLC.PercentV,TblLC.PrimaryInvoiceNo,  projects.Fullcode AS ProjectCode, projects.Project_name,projects.Project_nameE,"
  MySQL = MySQL & "                 TblLC.*, dbo.TblCountriesData.CountryName,Transaction_Type  = 70,TblLC.Bank2 as CashCustomerName,TblLC.Remarks as  remark,TblLC.AcceptianPeriod,"
 MySQL = MySQL & "                          dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,TblLC.LCNO as NoteSerial1 ,TblCustemers.CusID,TblLC.TblLCID Transaction_ID ,"
   MySQL = MySQL & "                       dbo.TblLC.Remarks, dbo.TblLC.NoOfParcil, dbo.TblLC.PaymentTypeID, dbo.TblLC.ChequeNumber, dbo.TblLC.ChequeDueDate, dbo.TblBoxesData.BoxName,"
   MySQL = MySQL & "                       dbo.BanksData.BankName AS BankName2, dbo.currency.nameE AS CurrencyNameE, dbo.BanksData.BankNamee AS BankNameE2,"
 MySQL = MySQL & "                         BanksData_1.BankNamee AS BankNameE, dbo.TblBoxesData.BoxNameE, projects.Fullcode as ProjectCode"
 MySQL = MySQL & "       FROM     dbo.TblCountriesData RIGHT OUTER JOIN"
     MySQL = MySQL & "                     dbo.TblCustemers RIGHT OUTER JOIN"
     MySQL = MySQL & "                     dbo.TblBoxesData RIGHT OUTER JOIN"
    MySQL = MySQL & "                      dbo.TblLC LEFT OUTER JOIN"
   MySQL = MySQL & "                       dbo.BanksData ON dbo.TblLC.BankID2 = dbo.BanksData.BankID ON dbo.TblBoxesData.BoxID = dbo.TblLC.BoxID ON dbo.TblCustemers.CusID = dbo.TblLC.VendorId ON"
  MySQL = MySQL & "                        dbo.TblCountriesData.CountryID = dbo.TblLC.CountryId LEFT OUTER JOIN"
     MySQL = MySQL & "                     dbo.currency ON dbo.TblLC.CurrencyId = dbo.currency.id LEFT OUTER JOIN"
    MySQL = MySQL & "                      dbo.LCTypes ON dbo.TblLC.LCTyperId = dbo.LCTypes.id LEFT OUTER JOIN"
      MySQL = MySQL & "                    dbo.BanksData AS BanksData_1 ON dbo.TblLC.BankId = BanksData_1.BankID"
      
      
    MySQL = MySQL & "                       LEFT OUTER JOIN"
      MySQL = MySQL & "                    dbo.projects ON dbo.TblLC.project_id = projects.Id"
    
      
      


      If Me.DTPickerAccFrom <> Empty Or Me.DTPickerAccFrom <> Null Then
            StrWhere = StrWhere + " and    dbo.TblLC.FromDate >=" & SQLDate(Me.DTPickerAccFrom, True) & ""
        End If
        
        If Me.DTPickerAccTo <> Empty Or Me.DTPickerAccTo <> Null Then
            StrWhere = StrWhere + " and    dbo.TblLC.ToDate <=" & SQLDate(Me.DTPickerAccTo, True) & ""
        End If
        

     
     If (Me.TXTOrDer_no.text) <> "" Then
        StrSQL = StrSQL + " AND LCNO like'%" & Me.TXTOrDer_no.text & "%'"
    End If
    

    If DataCombo4.BoundText <> "" Then
        StrWhere = StrWhere + " and countryid =" & DataCombo4.BoundText & " "
    
    End If
    

    If Me.DBCboClientName.BoundText <> "" Then
 
        StrWhere = StrWhere + " and TblCustemers.CusID =" & Me.DBCboClientName.BoundText & ""
 
    End If

    If Me.DCboStoreName.BoundText <> "" Then
 
        'StrWhere = StrWhere + " and dbo.Transactions.storeid  =" & val(Me.DCboStoreName.BoundText)
 
    End If
    
    If Trim(Me.txtRemark.text) <> "" Then
    
        StrWhere = StrWhere + " and remark like '%" & Trim(Me.txtRemark.text) & "%'"
     
    End If

    'StrSQL = StrSQL & "  order by Transaction_ID "


 If TxtCashCustomerName.text <> "" Then
           
                StrWhere = StrWhere + " and TblLC.Bank2 like '%" & (TxtCashCustomerName.text) & "%'"
           
      End If
        
 


    Build_Sql2 = MySQL + StrWhere & "  order by dbo.TblLC.LCNO "
    Exit Function
ErrTrap:
End Function


Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is FG Then
            If Not FG.TextMatrix(FG.row, 1) = "" Then
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
    Me.Caption = "Search For Purchase Orders"
    Label1(2).Caption = Me.Caption
    Label2.Caption = "Order No"
 
    Label3.Caption = "Date"
    Label5.Caption = "Country"
    Label4.Caption = "Vendor"
    Label6.Caption = "Remark"
lbl(8).Caption = "Store"

    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.FG
        .TextMatrix(0, .ColIndex("order_no")) = "order no"
        .TextMatrix(0, .ColIndex("remark")) = "remark  "
        .TextMatrix(0, .ColIndex("CusName")) = "Vendor Name"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = " Date"
        .TextMatrix(0, .ColIndex("CountryName")) = "Country Name"
        .TextMatrix(0, .ColIndex("STORENAME")) = "Store Name"
        
  
        '  .AutoSize 0, .Cols - 1, False
    End With

End Sub

