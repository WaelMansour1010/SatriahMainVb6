VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Arrows 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Õ„Ì· «·«”⁄«— «·Œ«’… »«·«”Â„"
   ClientHeight    =   9180
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   18270
   Icon            =   "ArrowsFollow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   18270
   Begin VB.Frame Frame3 
      Caption         =   "œ·«·«  «·«·Ê«‰"
      Height          =   975
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   8160
      Width           =   2655
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "·«ÌÊÃœ  €ÌÌ—"
         Height          =   255
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Height          =   255
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "—»Õ"
         Height          =   255
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ”«—…"
         Height          =   255
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "»Ì«‰«   Õ·Ì·Ì…"
      Height          =   975
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   8160
      Width           =   13095
      Begin VB.Label txtPoint2 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label txtPoint1 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label txtResd2 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label txtResd1 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Txtavg 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   9960
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "‰ﬁÿÂ  „ﬁ«Ê„Â  2"
         Height          =   255
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "‰ﬁÿÂ œ⁄„ 2"
         Height          =   255
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "‰ﬁÿÂ „ﬁ«Ê„Â 1"
         Height          =   375
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "‰ﬁÿÂ œ⁄„ 1"
         Height          =   375
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«— ﬂ«“"
         Height          =   375
         Left            =   11280
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   -2400
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   18615
      Begin SHDocVwCtl.WebBrowser WebBrowser3 
         Height          =   5895
         Left            =   -120
         TabIndex        =   12
         Top             =   360
         Width           =   18615
         ExtentX         =   32835
         ExtentY         =   10398
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
   End
   Begin VB.CommandButton Cmd 
      Caption         =   " ÕÊÌ· «·Ï «·«ﬂ”Ì·"
      Height          =   315
      Index           =   3
      Left            =   3120
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Õ›Ÿ ›Ï «·»—‰«„Ã"
      Height          =   315
      Index           =   2
      Left            =   5040
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Cmd 
      Caption         =   " Õ„Ì· «·«”⁄«— „‰ „·›"
      Height          =   315
      Index           =   1
      Left            =   7200
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Cmd 
      Caption         =   " Õ„Ì· «·«”⁄«— „‰ «·«‰ —‰ "
      Height          =   315
      Index           =   0
      Left            =   9240
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   6540
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   18195
      _cx             =   32094
      _cy             =   11536
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
      Cols            =   24
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"ArrowsFollow.frx":000C
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
      TabIndex        =   7
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   18270
      _cx             =   32226
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
      Caption         =   " Õ„Ì· «·«”⁄«— «·Œ«’… »«·«”Â„"
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
   Begin MSDataListLib.DataCombo DcboFinMarketId 
      Height          =   315
      Left            =   11520
      TabIndex        =   9
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   720
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdPrint 
      Height          =   390
      Left            =   120
      TabIndex        =   31
      Top             =   8760
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄…"
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
      ButtonImage     =   "ArrowsFollow.frx":0390
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õœœ «·»Ê—’Â"
      Height          =   255
      Index           =   4
      Left            =   16440
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·„ƒ‘— «·⁄«„"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   11775
   End
End
Attribute VB_Name = "Arrows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim path As String
Dim NEW_interface As Boolean
Dim HyperLinkGeneral As String

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
    With VSFlexGrid1
        .MergeCells = flexMergeFixedOnly
    
        .MergeRow(0) = True
        .MergeCol(.ColIndex("Symbol")) = True
        .MergeCol(.ColIndex("Name")) = True
    
        .Cell(flexcpText, 0, .ColIndex("Symbol"), 1, .ColIndex("Symbol")) = "—„“ «·‘—ﬂÂ  "
        '    .ColWidth(.ColIndex("Symbol")) = 1200

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
  
        '    Set .WallPaper = GrdBck.Picture
    End With

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

Private Sub CmdPrint_Click()
    On Error Resume Next
    Dim GrdBack As ClsBackGroundPic
    'Grid.ExtendLastCol = True
    VSFlexGrid1.WallPaper = Nothing
    'Grid.AutoSize  0, Grid.Cols - 1, False
    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
    'Printer.RightToLeft = True
    'Printer.Print ("Employee Salary Report")

    Me.VSFlexGrid1.PrintGrid " ﬁ—Ì— »ﬂ· »Ì«‰«  «·«”Â„", True, 2, 1, 1500

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
        '    .ColWidth(.ColIndex("Symbol")) = 1200

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
  
        '    Set .WallPaper = GrdBck.Picture
    End With

    'WebBrowser1.Navigate2 "http://www.tadawul.com.sa/Resources/Reports/DetailedDaily_ar.html"
End Sub

Private Sub VSFlexGrid1_Click()
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

    With VSFlexGrid1

        If Not .TextMatrix(.Row, .ColIndex("HyperLink")) = "" Then
 
            LastPrice = val(.TextMatrix(.Row, .ColIndex("LastPrice")))
            maxPrice = val(.TextMatrix(.Row, .ColIndex("Max")))
            MinPrice = val(.TextMatrix(.Row, .ColIndex("Min")))
            avg = (LastPrice + maxPrice + MinPrice) / 3 ' «·«— ﬂ«“
            result1 = avg
            result2 = avg * 2
            Point1 = result2 - maxPrice '«·œ⁄„ 1
            Resd1 = result2 - MinPrice '„ﬁ«Ê„Â 1
            Point2 = Resd1 - Point1 - result1 'œ⁄„ 2
            Resd2 = result1 - Point1 - Resd1 '„ﬁ«Ê„Â 2
            Txtavg = avg
            txtPoint1 = Point1
            txtResd1 = Resd1
            txtPoint2 = Point2
            txtResd2 = Resd2
        End If

    End With

End Sub

Private Sub VSFlexGrid1_DblClick()
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

    With VSFlexGrid1

        If Not .TextMatrix(.Row, .ColIndex("HyperLink")) = "" Then
            ArrowsCompanyDetails.show
            ArrowsCompanyDetails.LoadPage .TextMatrix(.Row, .ColIndex("HyperLink")), .TextMatrix(.Row, .ColIndex("Symbol")), .TextMatrix(.Row, .ColIndex("Name"))
            LastPrice = val(.TextMatrix(.Row, .ColIndex("LastPrice")))
            maxPrice = val(.TextMatrix(.Row, .ColIndex("Max")))
            MinPrice = val(.TextMatrix(.Row, .ColIndex("Min")))
            avg = (LastPrice + maxPrice + MinPrice) / 3 ' «·«— ﬂ«“
            result1 = avg
            result2 = avg * 2
            Point1 = result2 - maxPrice '«·œ⁄„ 1
            Resd1 = result2 - MinPrice '„ﬁ«Ê„Â 1
            Point2 = Resd1 - Point1 - result1 'œ⁄„ 2
            Resd2 = result1 - Point1 - Resd1 '„ﬁ«Ê„Â 2
            Txtavg = avg
            txtPoint1 = Point1
            txtResd1 = Resd1
            txtPoint2 = Point2
            txtResd2 = Resd2
        End If

    End With

End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, _
                                         URL As Variant)

    'On Error GoTo ErrTrap
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

    'The ninth table in the page is the Companies List
    Dim startLoad As Integer
    Dim Cols As Integer
    On Error Resume Next

    DoEvents
    startLoad = 77 + 17
    Set objTable = WebBrowser1.Document.getElementsByTagName("table").Item(13)


    With VSFlexGrid1
        .MergeCells = flexMergeFixedOnly
    
        .MergeRow(0) = True
        .MergeCol(.ColIndex("Symbol")) = True
        .MergeCol(.ColIndex("Name")) = True
    
        .Cell(flexcpText, 0, .ColIndex("Symbol"), 1, .ColIndex("Symbol")) = "—„“ «·‘—ﬂÂ  "
        '    .ColWidth(.ColIndex("Symbol")) = 1200

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
  
        '    Set .WallPaper = GrdBck.Picture
    End With

    With Me.VSFlexGrid1
 
        .Rows = objTable.getElementsByTagName("tr").Length - 1
 
        For i = startLoad To .Rows
            Cols = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Length
            Dim HyperLink  As String
            Dim SymbolNo As Integer

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
                    .Cell(flexcpBackColor, (i - startLoad) + 1, 5, (i - startLoad) + 1, 5) = vbGreen
                ElseIf ChangePercentage < 0 Then
                    .Cell(flexcpBackColor, (i - startLoad) + 1, 6, (i - startLoad) + 1, 6) = vbRed
                    .Cell(flexcpBackColor, (i - startLoad) + 1, 5, (i - startLoad) + 1, 5) = vbRed
                    
                ElseIf ChangePercentage = 0 Then
                    .Cell(flexcpBackColor, (i - startLoad) + 1, 6, (i - startLoad) + 1, 6) = vbYellow
                    .Cell(flexcpBackColor, (i - startLoad) + 1, 5, (i - startLoad) + 1, 5) = vbYellow
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
    Exit Sub
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim objTable As Object

    If NEW_interface = False Then Exit Sub
    'The ninth table in the page is the Companies List
    Set objTable = WebBrowser2.Document.getElementsByTagName("table").Item(7)

    'Now enumerate all TR tags within the table
 
    Label1.Caption = objTable.getElementsByTagName("tr").Item(0).getElementsByTagName("td").Item(1).innerText & vbCrLf

    Set objTable = Nothing
    Exit Sub
ErrTrap:
    MsgBox "·«»œ „‰ «·« ’«· »«·«‰ —‰  «Ê·«"

End Sub

Private Sub WebBrowser3_DocumentComplete(ByVal pDisp As Object, _
                                         URL As Variant)
    'On Error Resume Next
End Sub

