VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmItemShowDet 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·«’‰«›"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10710
   Icon            =   "FrmitemShowDet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtStillQty 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3150
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   7230
      Width           =   1095
   End
   Begin VB.TextBox txtTotalQty 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5670
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   7200
      Width           =   1095
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   5835
      Left            =   30
      TabIndex        =   0
      Top             =   1290
      Width           =   10515
      _cx             =   18547
      _cy             =   10292
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
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmitemShowDet.frx":038A
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   1
      Top             =   7110
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   810
      TabIndex        =   2
      Top             =   7110
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   30
      TabIndex        =   3
      Top             =   7110
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonPositionImage=   1
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
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   825
      Index           =   13
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   10455
      _cx             =   18441
      _cy             =   1455
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
      Begin VB.TextBox TxtItemQty 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   3
         Left            =   4095
         MaxLength       =   10
         TabIndex        =   26
         Top             =   315
         Width           =   885
      End
      Begin VB.TextBox TxtItemQty 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   1
         Left            =   2295
         MaxLength       =   10
         TabIndex        =   20
         Top             =   300
         Width           =   870
      End
      Begin VB.TextBox TxtItemQty 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   0
         Left            =   3165
         MaxLength       =   10
         TabIndex        =   18
         Top             =   315
         Width           =   885
      End
      Begin VB.TextBox TxtCodeAother 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   9225
         TabIndex        =   6
         Top             =   315
         Width           =   1125
      End
      Begin VB.TextBox TxtItemQty 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   2
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   5
         Top             =   315
         Width           =   870
      End
      Begin MSDataListLib.DataCombo Dcbiteem 
         Height          =   315
         Left            =   6555
         TabIndex        =   7
         Top             =   315
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   360
         Index           =   24
         Left            =   645
         TabIndex        =   8
         Top             =   285
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   635
         Caption         =   "≈÷«›…"
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
         ColorButton     =   14871017
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   360
         Index           =   25
         Left            =   105
         TabIndex        =   9
         Top             =   285
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   635
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
         ColorButton     =   14871017
      End
      Begin MSDataListLib.DataCombo Dcbuniit 
         Height          =   315
         Left            =   5070
         TabIndex        =   10
         Top             =   300
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «· ”·Ì„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   2
         Left            =   4200
         TabIndex        =   27
         Top             =   0
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”⁄—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   1
         Left            =   1005
         TabIndex        =   21
         Top             =   0
         Width           =   1995
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂ„Ì…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   0
         Left            =   1740
         TabIndex        =   19
         Top             =   0
         Width           =   1995
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·’‰›"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   6315
         TabIndex        =   16
         Top             =   0
         Width           =   645
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·’‰›"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   9450
         TabIndex        =   15
         Top             =   0
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   38
         Left            =   105
         TabIndex        =   14
         ToolTipText     =   "⁄œœ «·√’‰«› «·„ﬂÊ‰… ·Â–« «·’‰› «·„Ã„⁄"
         Top             =   30
         Width           =   165
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·√’‰«›"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   39
         Left            =   1170
         TabIndex        =   13
         ToolTipText     =   "⁄œœ «·√’‰«› «·„ﬂÊ‰… ·Â–« «·’‰› «·„Ã„⁄"
         Top             =   -30
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬂ„Ì…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   41
         Left            =   960
         TabIndex        =   12
         Top             =   -390
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊÕœÂ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   51
         Left            =   5400
         TabIndex        =   11
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«·ﬂ„Ì… «·„ »ﬁÌ…"
      Height          =   225
      Left            =   4260
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   7260
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì «·ﬂ„Ì…"
      Height          =   225
      Left            =   6780
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   7230
      Width           =   1305
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘… «’‰«› «·Œ’„"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   45
      TabIndex        =   17
      Top             =   0
      Width           =   6420
   End
End
Attribute VB_Name = "FrmItemShowDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public mIndex As Long
Public mItemId As Long
Public mUnitId As Long
Public strDisplay As String
Public txtdate As Date
Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
    save
    
' GetData
           
      '  Case 1
           ' clear_all Me
'Me.DtpDateFrom.value = ""
'Me.DtpDateTo.value = ""
      '      If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
      '      Else
               ' Me.lbl(0).Caption = "Search Results"
      '      End If

      '  Case 2
      '      Unload Me
       Case 24
       AddNewFgRowother
       Case 25
            DeleteFgRowAther
    End Select

End Sub
Sub save()
Dim str As String
Dim i As Integer
str = ""
Dim mDeferred As Double
Dim mTotalQty As Double
With Me.Fg
For i = 1 To .rows - 1
 If .TextMatrix(i, .ColIndex("ItemName")) <> "" Then
 str = str & .TextMatrix(i, .ColIndex("ItemID")) & "#"
 str = str & .TextMatrix(i, .ColIndex("unitid")) & "#"
 str = str & .TextMatrix(i, .ColIndex("ItemQty")) & "#"
 str = str & .TextMatrix(i, .ColIndex("ItemName")) & "#"
 str = str & .TextMatrix(i, .ColIndex("Qty")) & "#"
 str = str & .TextMatrix(i, .ColIndex("UnitPrice")) & "#"
 str = str & .TextMatrix(i, .ColIndex("RecNo")) & "#"

 
 


 str = str & "@"
  str = str & CHR(13)
  mDeferred = mDeferred + val(.TextMatrix(i, .ColIndex("ItemQty")))
  mTotalQty = mTotalQty + val(.TextMatrix(i, .ColIndex("Qty")))
 End If
Next
If mIndex = 0 Then
    Frmovers.FgItemPloice.TextMatrix(Frmovers.LngRow, Frmovers.FgItemPloice.ColIndex("itemssh")) = str
ElseIf mIndex = 1 Then
    Dim mTotalGrid As Double
    mTotalGrid = val(frmsalebill6.Fg.TextMatrix(frmsalebill6.LngRow, frmsalebill6.Fg.ColIndex("VisaQty"))) + val(frmsalebill6.Fg.TextMatrix(frmsalebill6.LngRow, frmsalebill6.Fg.ColIndex("CashQty"))) + val(frmsalebill6.Fg.TextMatrix(frmsalebill6.LngRow, frmsalebill6.Fg.ColIndex("MadaQty"))) + mTotalQty
    
    If mTotalGrid > val(frmsalebill6.Fg.TextMatrix(frmsalebill6.LngRow, frmsalebill6.Fg.ColIndex("Count"))) Then
        MsgBox "·« Ì„ﬂ‰  ŒÿÏ ﬂ„Ì… ”ÿ— «·„÷Œ… Ì—ÃÏ  ⁄œÌ· «·ﬂ„Ì« "
        Exit Sub
    End If
    With frmsalebill6.Fg
        mTotalGrid = val(.TextMatrix(frmsalebill6.LngRow, .ColIndex("CashQty"))) + val(.TextMatrix(frmsalebill6.LngRow, .ColIndex("MadaQty"))) + val(.TextMatrix(frmsalebill6.LngRow, .ColIndex("VisaQty"))) + val(.TextMatrix(frmsalebill6.LngRow, .ColIndex("DeferredQty")))
        .TextMatrix(frmsalebill6.LngRow, .ColIndex("DetailsPump")) = str
    
        .TextMatrix(frmsalebill6.LngRow, .ColIndex("StillPumbQty")) = val(.TextMatrix(frmsalebill6.LngRow, .ColIndex("Count"))) - mTotalGrid
        
    End With
    frmsalebill6.Fg.TextMatrix(frmsalebill6.LngRow, frmsalebill6.Fg.ColIndex("DetailsPump")) = str
    frmsalebill6.Fg.TextMatrix(frmsalebill6.LngRow, frmsalebill6.Fg.ColIndex("DeferredQty")) = mTotalQty
    
    frmsalebill6.Fg.TextMatrix(frmsalebill6.LngRow, frmsalebill6.Fg.ColIndex("Deferred")) = mDeferred
    
    frmsalebill6.NewGrid.Grid_AfterEdit frmsalebill6.LngRow, frmsalebill6.NewGrid.Grid.ColIndex("VisaQty")
End If

End With
Unload Me
End Sub
Public Sub DisplayGrid()
Dim st As String
Dim i As Integer
  Dim astrSplitItems()  As String
    Dim astrSplitItems2() As String
    Dim nElements         As Integer
    Dim j                 As Integer
    
st = Trim(strDisplay)
Dim mDeferred As Double


            If st <> "" Then
                        'st = Fg.TextMatrix(RowNum, Fg.ColIndex("DetailsPump"))
                        astrSplitItems = Split(st, "@")
     
                        nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
                        Fg.rows = UBound(astrSplitItems) + 1
                        For j = 0 To nElements - 1
                            
                            astrSplitItems2 = Split(astrSplitItems(j), "#")
                            Fg.TextMatrix(j + 1, Fg.ColIndex("ItemID")) = astrSplitItems2(0)
                            Fg.TextMatrix(j + 1, Fg.ColIndex("unitid")) = astrSplitItems2(1)
                            Fg.TextMatrix(j + 1, Fg.ColIndex("ItemQty")) = astrSplitItems2(2)
                            Fg.TextMatrix(j + 1, Fg.ColIndex("ItemName")) = astrSplitItems2(3)
                            
                            Fg.TextMatrix(j + 1, Fg.ColIndex("Qty")) = astrSplitItems2(4)
                            Fg.TextMatrix(j + 1, Fg.ColIndex("UnitPrice")) = astrSplitItems2(5)
                            

'
'                                   .TextMatrix(LngNewRow, .ColIndex("ItemID")) = astrSplitItems2(0)
'
'        .TextMatrix(LngNewRow, .ColIndex("ItemCode")) = Trim$(Me.TxtCodeAother.text)
'        .TextMatrix(LngNewRow, .ColIndex("ItemName")) = Me.Dcbiteem.text
'
'        .TextMatrix(LngNewRow, .ColIndex("UnitId")) = Me.Dcbuniit.BoundText
'        .TextMatrix(LngNewRow, .ColIndex("UnitName")) = Me.Dcbuniit.text
'        .TextMatrix(LngNewRow, .ColIndex("Qty")) = val(Me.TxtItemQty(0).text)
'
'        .TextMatrix(LngNewRow, .ColIndex("UnitPrice")) = val(Me.TxtItemQty(1).text)
'        .TextMatrix(LngNewRow, .ColIndex("ItemQty")) = val(Me.TxtItemQty(2).text)
        
'                            RsDetails1("Transaction_ID").value = val(XPTxtBillID.text)
'                          '  RsDetails1("TransType").value = val(mIndex)
'                            RsDetails1("LineID").value = RowNum ' RSTransDetails("ID").value
'                            'RsDetails1("TransType").value = val(mIndex)
'                            RsDetails1("Cusid").value = astrSplitItems2(0)
'                            'RsDetails1("UnitID").value = astrSplitItems2(1)
'                            RsDetails1("Amount").value = astrSplitItems2(2)
'                            RsDetails1.update
                        Next j
          
                    End If


'With Me.Fg
'For i = 1 To .rows - 1
' If .TextMatrix(i, .ColIndex("ItemName")) <> "" Then
' str = str & .TextMatrix(i, .ColIndex("ItemID")) & "#"
' str = str & .TextMatrix(i, .ColIndex("unitid")) & "#"
' str = str & .TextMatrix(i, .ColIndex("ItemQty")) & "#"
' str = str & "@"
'  str = str & CHR(13)
'  mDeferred = mDeferred + val(.TextMatrix(i, .ColIndex("ItemQty")))
' End If
'Next
'If mIndex = 0 Then
'    Frmovers.FgItemPloice.TextMatrix(Frmovers.LngRow, Frmovers.FgItemPloice.ColIndex("itemssh")) = str
'ElseIf mIndex = 1 Then
'    frmsalebill6.Fg.TextMatrix(frmsalebill6.LngRow, frmsalebill6.Fg.ColIndex("DetailsPump")) = str
'    frmsalebill6.Fg.TextMatrix(frmsalebill6.LngRow, frmsalebill6.Fg.ColIndex("Deferred")) = mDeferred
'End If

'End With
End Sub
Private Sub DeleteFgRowAther()

    With Me.Fg

        If .Row = -1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        .RemoveItem .Row
        '.AutoSize 0, .Cols - 1, False
       ' Me.lbl(21).Caption = ModFgLib.GetItemsInFg(Fg, Fg.ColIndex("ItemID"))
    End With

End Sub
Private Sub AddNewFgRowother()

    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long

    If val(Me.Dcbiteem.BoundText) = 0 Then
        Msg = "  ÌÃ»  ÕœÌœ «”„ «·’‰›"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.Dcbiteem.SetFocus
        Exit Sub
    End If

   ' If Me.TxtModFlg.text = "E" Then
   '     If val(Me.DcboItems.BoundText) = val(Me.XPTxtID.text) Then
   '         Msg = "?????? ?? ???? ????? ??? ?? ????....!!!"
   '         MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
   '         Me.DcboItems.SetFocus
   '         Exit Sub
   '     End If
   ' End If
    
    If val(Me.TxtItemQty(2).text) = 0 Then
        Msg = " ÌÃ»  ÕœÌœ «·ﬁÌ„…"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.TxtItemQty(2).SetFocus
        Exit Sub
    End If

 
    If mIndex = 0 Then
        If val(Me.Dcbuniit.BoundText) = 0 Then
            Msg = " ÌÃ»  ÕœÌœ ÊÕœ… «·’‰›"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Me.Dcbuniit.SetFocus
            Exit Sub
        End If
    End If
'    With Me.Fg
'        LngFindRow = .FindRow(val(Me.Dcbiteem.BoundText), .FixedRows, .ColIndex("ItemID"), False, True)
'
'        If LngFindRow <> -1 Then
'            Msg = "Â–« «·„·› „ÊÃÊœ ›⁄·«"
'            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            .SetFocus
'            Exit Sub
'        End If
'
'    End With



Dim i As Long
Dim recNoVal As Long
Dim itemIDVal As String

' ﬁ—«¡… «·ﬁÌ„ «·„ÿ·Ê»… „‰ ⁄‰«’— «·Ê«ÃÂ…
recNoVal = val(Me.TxtItemQty(3).text)
itemIDVal = Me.Dcbiteem.BoundText

With Me.Fg
    ' ‰› —÷ √‰ «·’›Ê› «·À«» … (FixedRows)  ÕÊÌ ⁄‰«ÊÌ‰ «·√⁄„œ…° ·–·ﬂ ‰»œ√ „‰Â«
    For i = .FixedRows To .rows - 1
        '  Õﬁﬁ „‰ «·‘—ÿÌ‰:
        ' 1- «· Õﬁﬁ „‰ ﬁÌ„… "RecNo" (Ì „  ÕÊÌ· «·‰’ ≈·Ï —ﬁ„ ≈–« ·“„ «·√„—)
        ' 2- «· Õﬁﬁ „‰ ﬁÌ„… "ItemID"
        If val(.TextMatrix(i, .ColIndex("RecNo"))) = recNoVal And _
           .TextMatrix(i, .ColIndex("ItemID")) = itemIDVal Then
           
            MsgBox "Â–« «·„·› „ÊÃÊœ ›⁄·«", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            .SetFocus
            Exit Sub
        End If
    Next i
End With


    LngNewRow = ModFgLib.SetFgForNewRow(Fg, Fg.ColIndex("ItemID"))

    With Me.Fg
    .TextMatrix(LngNewRow, .ColIndex("Serial")) = LngNewRow
        .TextMatrix(LngNewRow, .ColIndex("ItemID")) = Me.Dcbiteem.BoundText
    
        .TextMatrix(LngNewRow, .ColIndex("ItemCode")) = Trim$(Me.TxtCodeAother.text)
        .TextMatrix(LngNewRow, .ColIndex("ItemName")) = Me.Dcbiteem.text
    
        .TextMatrix(LngNewRow, .ColIndex("UnitId")) = Me.Dcbuniit.BoundText
        .TextMatrix(LngNewRow, .ColIndex("UnitName")) = Me.Dcbuniit.text
        .TextMatrix(LngNewRow, .ColIndex("Qty")) = val(Me.TxtItemQty(0).text)
        
        .TextMatrix(LngNewRow, .ColIndex("UnitPrice")) = val(Me.TxtItemQty(1).text)
        .TextMatrix(LngNewRow, .ColIndex("ItemQty")) = val(Me.TxtItemQty(2).text)
        .TextMatrix(LngNewRow, .ColIndex("RecNo")) = val(Me.TxtItemQty(3).text)
       
        txtTotalQty = val(txtTotalQty) + val(.TextMatrix(LngNewRow, .ColIndex("Qty")))
        txtStillQty = val(txtStillQty) - val(txtTotalQty)
       
        '.AutoSize 0, .Cols - 1, False
    End With

    Me.lbl(38).Caption = ModFgLib.GetItemsInFg(Fg, Fg.ColIndex("ItemID"))

    Me.TxtCodeAother.text = ""
    Me.Dcbiteem.BoundText = ""
    Me.TxtItemQty(2).text = ""

    Me.TxtCodeAother.SetFocus
End Sub



'Private Sub Fg_Click()
'
 
'   On Error GoTo ErrTrap
'  '  FrmModels.FindRec
'   FrmSystemUnites.FindRec val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id")))
'ErrTrap:
'End Sub

Sub Retrive(Optional ID As Integer = 0, Optional IDDet As Integer = 0)

 Dim RsDetails As ADODB.Recordset
 Dim StrSQL As String
 Dim i As Integer


If mIndex = 0 Then

   Set RsDetails = New ADODB.Recordset
StrSQL = " SELECT     dbo.TblItemShInfo.ID, dbo.TblItemShInfo.ID2, dbo.TblItemShInfo.Ind, dbo.TblItemShInfo.Qntity, dbo.TblItemShInfo.ItemID, dbo.TblItems.ItemCode,"
StrSQL = StrSQL & "                      dbo.TblItems.ItemName , dbo.TblItems.ItemNamee, dbo.TblItemShInfo.unitid, dbo.TblUnites.Unitname, dbo.TblUnites.UnitNamee"
StrSQL = StrSQL & " FROM         dbo.TblItemShInfo LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.TblItemShInfo.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItems ON dbo.TblItemShInfo.ItemID = dbo.TblItems.ItemID"
StrSQL = StrSQL & "  Where (dbo.TblItemShInfo.ID2 =" & IDDet & ") And (dbo.TblItemShInfo.ind = " & ID & ")"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails.RecordCount > 0 Then
   With Me.Fg
   .rows = .rows + RsDetails.RecordCount
   RsDetails.MoveFirst
   For i = 1 To .rows - 1
   .TextMatrix(i, .ColIndex("Serial")) = i
   .TextMatrix(i, .ColIndex("ItemQty")) = IIf(IsNull(RsDetails("Qntity").value), "", RsDetails("Qntity").value)
   .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsDetails("ItemCode").value), "", RsDetails("ItemCode").value)
   .TextMatrix(i, .ColIndex("unitid")) = IIf(IsNull(RsDetails("UnitID").value), "", RsDetails("UnitID").value)
    .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsDetails("ItemID").value), "", RsDetails("ItemID").value)
    If SystemOptions.UserInterface = EnglishInterface Then
     .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDetails("ItemNamee").value), "", RsDetails("ItemNamee").value)
    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDetails("UnitNamee").value), "", RsDetails("UnitNamee").value)
      Else
    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDetails("ItemName").value), "", RsDetails("ItemName").value)
    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDetails("UnitName").value), "", RsDetails("UnitName").value)
        
    End If
    RsDetails.MoveNext
   Next i
  
   End With
    End If
   
End If
End Sub



Private Sub Dcbiteem_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 22222
        FrmCustemerSearch.show vbModal

    End If
 
End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
      Dcombos.GetItemsUnits Me.Dcbuniit



    Set DCboSearch = New clsDCboSearch
   
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    
    
    If mIndex = 0 Then
        Dcombos.GetItemsNames Me.Dcbiteem
        If val(Frmovers.XPTxtID.text) <> 0 And val(Frmovers.FgItemPloice.TextMatrix(Frmovers.LngRow, Frmovers.FgItemPloice.ColIndex("IDD"))) <> 0 Then
            Retrive val(Frmovers.XPTxtID.text), val(Frmovers.FgItemPloice.TextMatrix(Frmovers.LngRow, Frmovers.FgItemPloice.ColIndex("IDD")))
        End If
    ElseIf mIndex = 1 Then
        Label5.Caption = "«·⁄„·«¡"
        FrmItemShowDet.Caption = "«·⁄„·«¡"
        Dcbuniit.Visible = False
        'TxtItemQty(2).Visible = False
        lbl(41).Caption = "«·ﬁÌ„…"
        lbl(51).Visible = False
        Fg.TextMatrix(0, Fg.ColIndex("ItemCode")) = "ﬂÊœ «·⁄„Ì·"
        Fg.TextMatrix(0, Fg.ColIndex("ItemName")) = "«”„ «·⁄„Ì·"
        Fg.TextMatrix(0, Fg.ColIndex("ItemQty")) = "«·„»·€"
        Fg.ColHidden(Fg.ColIndex("UnitName")) = True
        
        Dcombos.GetCustomersSuppliers 1, Me.Dcbiteem, , , 1
    End If
    
    Set GrdBack = New ClsBackGroundPic

'    With Me.Fg
'        Set .WallPaper = GrdBack.Picture
'        .AutoSize 0, .Cols - 1, False
'    End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
'

Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
    
  Me.Caption = "Search CardAuthorizationReform"

'Me.LblClientName.Caption = "ClientName"
'lbl(4).Caption = "From"
'lbl(3).Caption = "To"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(0).Caption = "Unit Name Arb"
lbl(8).Caption = "Unit Name ENG"
lbl(2).Caption = "Total"
'Me.lbreg.Caption = "Date Registration"
'Me.lbprocess.Caption = "Process No"
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "Code"
       ' .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("ClientName")) = "Unit Name Arb"
        .TextMatrix(0, .ColIndex("Telephone")) = "Unit Name ENG"
       '.TextMatrix(0, .ColIndex("PlateNo")) = "PlateNo"
    End With
  '
End Sub
Private Function GetPriceItem(mCusId As Long) As Double
       Dim sql  As String
        If val(Dcbiteem.BoundText) = 0 Then Exit Function
       If mIndex = 1 Then
            txtdate = frmsalebill6.XPDtbBill.value
        Else
        txtdate = Date
       End If
        sql = "Select TblCustomerContractDetails.Price From TblCustomerContract "
      sql = sql & " Inner Join TblCustomerContractDetails "
      sql = sql & " On TblCustomerContract.TblCustomerContractD = TblCustomerContractDetails.TblCustomerContractD " & _
      sql = sql & " Where TblCustomerContract.CustomerId = " & val(Dcbiteem.BoundText) & " and  '" & txtdate & "' Between FromDate And ToDate "
      sql = sql & " and TblCustomerContractDetails.ItemID = " & mItemId
      sql = sql & " and TblCustomerContractDetails.UnitId = " & mUnitId
      
   Dim formattedDate As String
formattedDate = Format(txtdate, "yyyy-MM-dd")
   
      
    sql = "Select TblCustomerContractDetails.Price From TblCustomerContract "
    sql = sql & "Inner Join TblCustomerContractDetails "
    sql = sql & "On TblCustomerContract.TblCustomerContractD = TblCustomerContractDetails.TblCustomerContractD "
    sql = sql & "Where TblCustomerContract.CustomerId = " & val(Dcbiteem.BoundText) & " "
    sql = sql & "and '" & formattedDate & "' Between FromDate And ToDate "
    sql = sql & "and TblCustomerContractDetails.ItemID = " & mItemId & " "
    sql = sql & "and TblCustomerContractDetails.UnitId = " & mUnitId


      Dim rsDummy As New ADODB.Recordset
      rsDummy.Open sql, Cn, adOpenStatic, adLockReadOnly
      If Not rsDummy.EOF Then
            GetPriceItem = val(rsDummy!Price & "")
      End If
      rsDummy.Close
      If GetPriceItem = 0 Then
      
        sql = "Select UnitSalesPrice From TblItemsUnits where ItemID=" & mItemId & " and UnitID =" & mUnitId
        rsDummy.Open sql, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then
            GetPriceItem = val(rsDummy!UnitSalesPrice & "")
            
            
           End If
      End If
End Function



Private Sub TxtCodeAother_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        If TxtCodeAother.text = "" Then
            Me.Dcbiteem.BoundText = ""
        Else
            Me.Dcbiteem.BoundText = GetItemID(Trim$(Me.TxtCodeAother.text))
        End If
    End If
End Sub




Private Sub Dcbiteem_Change()
 Dim UnitID As Long
    Dim UnitName As String
    Me.TxtCodeAother.text = GetItemCode(val(Me.Dcbiteem.BoundText))
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsUnits·byitemid Me.Dcbuniit, val(Me.Dcbiteem.BoundText)
  
    GetDefaultItemUnit val(Me.Dcbiteem.BoundText), UnitID, UnitName
    Dcbuniit.text = UnitName
    Dcbuniit.BoundText = UnitID
    
     TxtItemQty(1) = GetPriceItem(val(Dcbiteem.BoundText))

End Sub

Private Sub Dcbiteem_Click(Area As Integer)
 Dcbiteem_Change
End Sub

Private Sub TxtItemQty_Change(Index As Integer)
If Index <> 2 Then
    
    TxtItemQty(2) = val(TxtItemQty(0)) * val(IIf(val(Me.TxtItemQty(1).text) <> 0, val(Me.TxtItemQty(1).text), 1))
End If
End Sub

Private Sub TxtItemQty_KeyPress(Index As Integer, KeyAscii As Integer)
   
  ' KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtItemQty.Item(2).text, 1)
End Sub
