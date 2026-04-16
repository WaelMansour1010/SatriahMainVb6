VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form ArrowsCurrentValue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·ﬁÌ„… «·”ÊﬁÌ… ·Ã„Ì⁄ «·«”Â„"
   ClientHeight    =   8400
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   18735
   Icon            =   "ArrowsCurrentValue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   18735
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
      Caption         =   "Frame1"
      Height          =   495
      Left            =   840
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   6960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   4740
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   18435
      _cx             =   32517
      _cy             =   8361
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
      BackColorBkg    =   16777215
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
      Cols            =   23
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"ArrowsCurrentValue.frx":000C
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
   Begin VB.CommandButton Command1 
      Caption         =   "«‰‘«¡ ﬁÌœ «·«—»«Õ Ê «·Œ”«∆—"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   7920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   6240
      Width           =   14895
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "‰”»… «·«—»«Õ Ê «·Œ”«∆—"
         Height          =   255
         Index           =   8
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ã„Ê⁄ «·«—»«Õ Ê «·Œ”«∆—"
         Height          =   255
         Index           =   7
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ã„Ê⁄ «·«—’œ…"
         Height          =   375
         Index           =   6
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "‰”»… «·«—»«Õ Ê «·Œ”«∆—"
         Height          =   255
         Index           =   5
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ã„Ê⁄ «·«—»«Õ Ê «·Œ”«∆—"
         Height          =   255
         Index           =   4
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ã„Ê⁄ «·«—’œ…"
         Height          =   375
         Index           =   3
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "‰”»… «·«—»«Õ Ê «·Œ”«∆—"
         Height          =   255
         Index           =   2
         Left            =   11760
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ã„Ê⁄ «·«—»«Õ Ê «·Œ”«∆—"
         Height          =   255
         Index           =   1
         Left            =   11760
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ã„Ê⁄ «·«—’œ…"
         Height          =   375
         Index           =   0
         Left            =   11640
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
      Begin VB.Line Line3 
         Index           =   3
         X1              =   2280
         X2              =   2280
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   4680
         X2              =   4680
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   11400
         X2              =   11400
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Line Line4 
         X1              =   6600
         X2              =   6600
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   9600
         X2              =   9600
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   14760
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   14760
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid11 
      Height          =   1500
      Left            =   0
      TabIndex        =   6
      Top             =   7920
      Visible         =   0   'False
      Width           =   15075
      _cx             =   26591
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
      BackColorBkg    =   16777215
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
      Cols            =   26
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"ArrowsCurrentValue.frx":0367
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
   Begin VB.CommandButton Cmd 
      Caption         =   " ’œÌ— «·Ï «·«ﬂ”Ì·"
      Height          =   315
      Index           =   3
      Left            =   8280
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   10215
      ExtentX         =   18018
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   4320
      Width           =   10815
      ExtentX         =   19076
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin C1SizerLibCtl.C1Elastic EleTop 
      Height          =   660
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   18735
      _cx             =   33046
      _cy             =   1164
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   20.25
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
      BackColor       =   12648447
      ForeColor       =   8421376
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "«·ﬁÌ„… «·”ÊﬁÌ… ·Ã„Ì⁄ «·«”Â„"
      Align           =   1
      AutoSizeChildren=   7
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
      Style           =   0
      TagSplit        =   2
      PicturePos      =   7
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
   Begin MSDataListLib.DataCombo DcboGovernmentID 
      Height          =   315
      Left            =   3720
      TabIndex        =   5
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   840
      Visible         =   0   'False
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboFinMarketId 
      Height          =   315
      Left            =   13680
      TabIndex        =   18
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   840
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser3 
      Height          =   975
      Left            =   0
      TabIndex        =   21
      Top             =   5040
      Width           =   10815
      ExtentX         =   19076
      ExtentY         =   1720
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õœœ «·»Ê—’Â"
      Height          =   255
      Index           =   4
      Left            =   16680
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·„Õ›Ÿ…"
      Height          =   255
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "ArrowsCurrentValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim path As String
Dim NEW_interface As Boolean

Private Sub Cmd_Click(Index As Integer)

    If DcboFinMarketId.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«Œ — »Ê—’… «Ê·« "
            DcboFinMarketId.SetFocus
            SendKeys ("{F4}")
        Else
            MsgBox "ÚÚSelect Market"
            DcboFinMarketId.SetFocus
            SendKeys ("{F4}")

        End If

    End If

    Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    Me.VSFlexGrid1.Rows = 2

    Select Case DcboFinMarketId.BoundText

        Case 1
            Frame1.Visible = False
            NEW_interface = True
            path = "http://www.tadawul.com.sa/wps/portal/!ut/p/c1/04_SB8K8xLLM9MSSzPy8xBz9CP0os3g_A-ewIE8TIwMLj2AXA0_vQGNzY18g18cQKB-JJO8eEGZq4GniE2wUHOBlbOBpREB3cGKRvp9Hfm6qfkFuRDkAgpcLJw!!/dl2/d1/L2dJQSEvUUt3QS9ZQnB3LzZfTjBDVlJJNDIwMFM1MDBJNExWVENMRzMwMjY!/"
            WebBrowser1.Navigate2 path
            path = "http://www.tadawul.com.sa/wps/portal/!ut/p/c1/04_SB8K8xLLM9MSSzPy8xBz9CP0os3g_A-ewIE8TIwODYFMDA08Tn7AQZx93YwMjM6B8JG55AwOSdLsHhJmC5IONggO8jA08jQjoDk4s0vfzyM9N1S_IDY0od1RUBAD6Iu2e/dl2/d1/L2dJQSEvUUt3QS9ZQnB3LzZfTjBDVlJJNDIwR05QOTBJSzZFSUlEUjAwVDY!/"
            WebBrowser2.Navigate2 path
    End Select

    If val(DcboFinMarketId.BoundText) > 1 Then

        Frame1.Visible = True

        get_Financial_market_data val(DcboFinMarketId.BoundText), , , , , , , , path
        WebBrowser3.Navigate path

    End If

End Sub

Private Sub DcboFinMarketId_Change()
    Cmd_Click (0)
End Sub

Private Sub Form_Load()
    Resize_Form Me
    NEW_interface = False

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.getFinMarkets DcboFinMarketId

    With VSFlexGrid1
        .MergeCells = flexMergeFixedOnly
    
        .MergeRow(0) = True
        .MergeCol(.ColIndex("Symbol")) = True
        .MergeCol(.ColIndex("Name")) = True
    
        .Cell(flexcpText, 0, .ColIndex("Symbol"), 1, .ColIndex("Symbol")) = "—„“ «·‘—ﬂÂ  "

        .Cell(flexcpText, 0, .ColIndex("Name"), 1, .ColIndex("Name")) = "«”„ «·‘—ﬂÂ"
  
        .Cell(flexcpText, 0, .ColIndex("LastPrice"), 0, .ColIndex("ChangePercentage")) = " «Œ— ’›ﬁ…"
        .Cell(flexcpAlignment, 0, .ColIndex("LastPrice"), 0, .ColIndex("ChangePercentage")) = flexAlignCenterCenter

        .Cell(flexcpText, 0, .ColIndex("BestOrderPrice"), 0, .ColIndex("BestOrderQty")) = " «›÷· ÿ·»  "
        .Cell(flexcpAlignment, 0, .ColIndex("BestOrderPrice"), 0, .ColIndex("BestOrderQty")) = flexAlignCenterCenter

        .Cell(flexcpText, 0, .ColIndex("BestViewrPrice"), 0, .ColIndex("BestViewrQty")) = " «›÷· ⁄—÷"
        .Cell(flexcpAlignment, 0, .ColIndex("BestViewrPrice"), 0, .ColIndex("BestViewrQty")) = flexAlignCenterCenter

        .Cell(flexcpText, 0, .ColIndex("NoOfDeals"), 0, .ColIndex("Qty")) = "  —«ﬂ„Ì  "
        .Cell(flexcpAlignment, 0, .ColIndex("NoOfDeals"), 0, .ColIndex("Qty")) = flexAlignCenterCenter

        .Cell(flexcpText, 0, .ColIndex("Opening"), 0, .ColIndex("Min")) = " «·ÌÊ„  "
        .Cell(flexcpAlignment, 0, .ColIndex("Opening"), 0, .ColIndex("Min")) = flexAlignCenterCenter

        .Cell(flexcpText, 0, .ColIndex("avg"), 0, .ColIndex("Resd2")) = " »Ì«‰«   Õ·Ì·Ì…  "
        .Cell(flexcpAlignment, 0, .ColIndex("avg"), 0, .ColIndex("Resd2")) = flexAlignCenterCenter
 
    End With

    'WebBrowser1.Navigate2 "http://www.tadawul.com.sa/Resources/Reports/DetailedDaily_ar.html"

End Sub

Private Sub VSFlexGrid1_Click()

    With VSFlexGrid1

        If Not .TextMatrix(.Row, .ColIndex("HyperLink")) = "" Then
            ArrowsCompanyDetails.show
            ArrowsCompanyDetails.LoadPage .TextMatrix(.Row, .ColIndex("HyperLink")), .TextMatrix(.Row, .ColIndex("Symbol")), .TextMatrix(.Row, .ColIndex("Name"))
        End If

    End With

End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, _
                                         URL As Variant)
    On Error GoTo ErrTrap

    If NEW_interface = False Then Exit Sub
    Dim i As Integer

    Dim objTable As Object
    Dim maxPrice As Double
    Dim MinPrice As Double
    Dim LastPrice As Double
    Dim Point1 As Double
    Dim Point2 As Double
    Dim Resd1 As Double
    Dim Resd2 As Double
    Dim avg As Double
    Dim result1 As Double
    Dim result2 As Double
    Dim TodayValue As String
    Dim CompanyId As Integer
    'The ninth table in the page is the Companies List
    Dim startLoad As Integer
    Dim Cols As Integer
    'On Error Resume Next
    startLoad = 77 + 17
    Set objTable = WebBrowser1.Document.getElementsByTagName("table").Item(13)

    With Me.VSFlexGrid1
 
        .Rows = objTable.getElementsByTagName("tr").Length - 1
 
        For i = startLoad To .Rows
            Cols = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Length
            Dim HyperLink  As String
            Dim SymbolNo As String

            If Cols >= 2 Then
                .TextMatrix((i - startLoad) + 1, .ColIndex("LineNo")) = (i - startLoad) + 1
                .TextMatrix((i - startLoad) + 1, .ColIndex("Name")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(0).innerText
      
            End If
    
            If Cols = 14 Then
                HyperLink = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("a")
                SymbolNo = right(HyperLink, 4)
                .TextMatrix((i - startLoad) + 1, .ColIndex("Symbol")) = SymbolNo
                .TextMatrix((i - startLoad) + 1, .ColIndex("HyperLink")) = HyperLink
                .TextMatrix((i - startLoad) + 1, .ColIndex("LastPrice")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(1).innerText
 
                TodayValue = val(.TextMatrix((i - startLoad) + 1, .ColIndex("LastPrice")))
     
                CompanyId = 0
     
                GetArrowsCompanyData CompanyId, SymbolNo

                If CompanyId > 0 Then
                    update_record_to_table "ArrowsCompanies", "TodayValue", TodayValue, "CompanyId", CompanyId
                End If
     
                .TextMatrix((i - startLoad) + 1, .ColIndex("Qty1")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(2).innerText
     
                .TextMatrix((i - startLoad) + 1, .ColIndex("Change")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(3).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("ChangePercentage")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(4).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("NoOfDeals")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(5).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("Qty")) = val(objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(6).innerText)
     
                .TextMatrix((i - startLoad) + 1, .ColIndex("BestOrderPrice")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(7).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("BestOrderQty")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(8).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("BestViewrPrice")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(9).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("BestViewrQty")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(10).innerText
     
                .TextMatrix((i - startLoad) + 1, .ColIndex("Opening")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(11).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("Max")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(12).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("Min")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(13).innerText
                Dim ChangePercentage As Double
                ChangePercentage = val(.TextMatrix((i - startLoad) + 1, .ColIndex("ChangePercentage")))

                If ChangePercentage > 0 Then
                    .Cell(flexcpBackColor, (i - startLoad) + 1, 6, (i - startLoad) + 1, 6) = vbGreen
                ElseIf ChangePercentage < 0 Then
                    .Cell(flexcpBackColor, (i - startLoad) + 1, 6, (i - startLoad) + 1, 6) = vbRed
                ElseIf ChangePercentage = 0 Then
                    .Cell(flexcpBackColor, (i - startLoad) + 1, 6, (i - startLoad) + 1, 6) = vbYellow
                End If
 
                LastPrice = val(.TextMatrix((i - startLoad) + 1, .ColIndex("LastPrice")))
                maxPrice = val(.TextMatrix((i - startLoad) + 1, .ColIndex("Max")))
                MinPrice = val(.TextMatrix((i - startLoad) + 1, .ColIndex("Min")))
                avg = (LastPrice + maxPrice + MinPrice) / 3 ' «·«— ﬂ«“
                result1 = avg
                result2 = avg * 2
                Point1 = result2 - maxPrice '«·œ⁄„ 1
                Resd1 = result2 - MinPrice '„ﬁ«Ê„Â 1
                Point2 = Resd1 - Point1 - result1 'œ⁄„ 2
                Resd2 = result1 - Point1 - Resd1 '„ﬁ«Ê„Â 2
                .TextMatrix((i - startLoad) + 1, .ColIndex("avg")) = Round(avg, 3)
                .TextMatrix((i - startLoad) + 1, .ColIndex("Point1")) = Round(Point1, 3)
                .TextMatrix((i - startLoad) + 1, .ColIndex("Resd1")) = Round(Resd1, 3)
                .TextMatrix((i - startLoad) + 1, .ColIndex("Point2")) = Round(Point2, 3)
                .TextMatrix((i - startLoad) + 1, .ColIndex("Resd2")) = Round(Resd2, 3)
 
            End If

        Next i

        .AutoSize 0, .Cols - 1, False
        Dim j As Integer
        Dim lastindex As Integer

        For j = .Rows - 1 To 2 Step -1

            If .TextMatrix(j, .ColIndex("Name")) <> "" Then
                lastindex = j + 1
                GoTo LL
            End If

        Next j

LL:
        .Rows = lastindex + 1
    End With

    Set objTable = Nothing
    Exit Sub
ErrTrap:
    MsgBox "·«»œ „‰ «·« ’«· »«·«‰ —‰  «Ê·«"
 
End Sub

Private Sub WebBrowser2_DocumentComplete(ByVal pDisp As Object, _
                                         URL As Variant)
    On Error GoTo ErrTrap
    Exit Sub
    Dim i As Integer
    Dim objTable As Object

    If NEW_interface = False Then Exit Sub
    'The ninth table in the page is the Companies List
    Set objTable = WebBrowser2.Document.getElementsByTagName("table").Item(7)

    'Now enumerate all TR tags within the table

    Set objTable = Nothing
    Exit Sub
ErrTrap:
    MsgBox "·«»œ „‰ «·« ’«· »«·«‰ —‰  «Ê·«"

End Sub

