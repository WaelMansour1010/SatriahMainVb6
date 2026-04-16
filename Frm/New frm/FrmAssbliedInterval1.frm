VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmAssbliedInterval1 
   Caption         =   " ﬁ—Ì— ⁄‰ › —… „Ã„⁄…"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   9855
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8160
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9855
      _cx             =   17383
      _cy             =   14393
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
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   1
      ChildSpacing    =   1
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmAssbliedInterval1.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   7185
         Left            =   15
         TabIndex        =   2
         Top             =   960
         Width           =   9825
         _cx             =   17330
         _cy             =   12674
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   930
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   15
         Width           =   9825
         _cx             =   17330
         _cy             =   1640
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
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   375
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   450
            Visible         =   0   'False
            Width           =   1305
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   375
            Left            =   90
            TabIndex        =   9
            Top             =   60
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   661
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
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   915
            Left            =   7650
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   -30
            Width           =   2145
            _cx             =   3784
            _cy             =   1614
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
            Caption         =   " ÕœÌœ «·› —…"
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
            Begin MSComCtl2.DTPicker DtpFrom 
               Height          =   345
               Left            =   60
               TabIndex        =   5
               Top             =   180
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   105709569
               CurrentDate     =   39561
            End
            Begin MSComCtl2.DTPicker DtpTo 
               Height          =   345
               Left            =   60
               TabIndex        =   6
               Top             =   540
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   105709569
               CurrentDate     =   39561
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   315
               Index           =   1
               Left            =   1590
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   510
               Width           =   405
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   315
               Index           =   0
               Left            =   1590
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   210
               Width           =   405
            End
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   420
            Left            =   90
            TabIndex        =   3
            Top             =   450
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   741
            Caption         =   "⁄—÷"
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   12632256
         End
      End
   End
End
Attribute VB_Name = "FrmAssbliedInterval1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TransactionsValues
    TotalCash As Single
    TotalDue As Single
    TotalNet As Single
End Type

Private Sub Command1_Click()
    Dim StrFileName As String
    StrFileName = App.path & "\Temp1.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If

    Me.Fg.SaveGrid StrFileName, flexFileExcel, True
    OpenFile StrFileName
End Sub

Private Sub Fg_Click()
    Dim LngMouseRow As Long
    Dim LngMouseCol As Long

    Dim FgNode As VSFlex8UCtl.VSFlexNode
    LngMouseRow = Me.Fg.MouseRow

    If (LngMouseRow < 0 Or Fg.Rows = 0) Then
        Exit Sub
    End If

    If (Fg.IsSubtotal(LngMouseRow) = True And LngMouseCol = 0) Then
        Set FgNode = Fg.GetNode(LngMouseRow)
        FgNode.Expanded = Not FgNode.Expanded
    
    End If

End Sub

Private Sub Fg_MouseMove(Button As Integer, _
                         Shift As Integer, _
                         X As Single, _
                         Y As Single)
    Dim LngMouseRow As Long
    Dim LngMouseCol As Long

    LngMouseRow = Me.Fg.MouseRow
    LngMouseCol = Me.Fg.MouseCol

    If (LngMouseRow < 0) Or (LngMouseCol < 0) Or Fg.Rows = 0 Then
        Fg.MousePointer = flexDefault
        Exit Sub
    End If

    If Fg.IsSubtotal(LngMouseRow) = True Then
        Fg.MousePointer = flexHand
    Else
        Fg.MousePointer = flexDefault
    End If

    WriteToolTip LngMouseRow, LngMouseCol

End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic

    With Me.Fg
        .Clear flexClearScrollable, flexClearEverything
        .Cols = 10
        .Rows = 0
        .FixedCols = 0
        .FixedRows = 0
        .MergeCells = flexMergeOutline
        .ColAlignment(0) = flexAlignRightCenter
        .RowHeightMin = 320
        .ExplorerBar = flexExNone
        ' appearance
        .GridLines = flexGridNone
        .BackColorBkg = .backcolor
        .SheetBorder = .backcolor
        .ExtendLastCol = True
        .Redraw = flexRDBuffered ' << new setting
        .OutlineCol = 0
        .OutlineBar = flexOutlineBarCompleteLeaf
        '.NodeClosedPicture = MDIFrmMain.ImgLstTree.ListImages("Close").Picture
        '.NodeOpenPicture = MDIFrmMain.ImgLstTree.ListImages("OpenFolder").Picture
        .Ellipsis = flexEllipsisEnd
        'Set Grdback = New ClsBackGroundPic
        'Set .WallPaper = Grdback.Picture
        ' behavior
        .AllowSelection = False
        .Highlight = flexHighlightWithFocus
        .ScrollTrack = True
        .AutoSearch = flexSearchFromCursor
    End With

    SetDtpickerDate Me.DTPFrom
    SetDtpickerDate Me.DTPTo
    Me.Width = 11000
    Me.Height = 9000
    Resize_Form Me
    Cn.CommandTimeout = 180
End Sub

Private Sub ISButton1_Click()
    LoadData
End Sub

Private Sub LoadData()
    Dim XFont As IFontDisp
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer, J As Integer
    Dim IntStartSelect As Integer, IntEndSelect As Integer
    Dim SngTempValue As Single
    Dim StrOneRowData As String
    Dim SngHeaderBackColor As Single
    Dim SngDataBackColor As Single
    Dim StrStartDate As String
    Dim SngTemp1 As Single, SngTemp2 As Single, SngTemp3 As Single
    Dim TransValues As TransactionsValues
    Dim LngCustomersCount As Long, LngTempRow As Long

    SngHeaderBackColor = &HC0C0C0
    SngDataBackColor = &HE2E9E9

    With Me.Fg
        .Clear flexClearScrollable, flexClearEverything
        .Rows = 0
        .OutlineBar = flexOutlineBarComplete
        .MergeCells = flexMergeOutline
        StrOneRowData = "«· ﬁ—Ì— «·„Ã„⁄ ⁄‰ «·› —… "

        If Not IsNull(Me.DTPFrom.value) Then
            StrOneRowData = StrOneRowData & "„‰ " & DisplayDate(Me.DTPFrom.value)
        End If

        If Not IsNull(Me.DTPTo.value) Then
            StrOneRowData = StrOneRowData & " ≈·Ï " & DisplayDate(Me.DTPTo.value)
        End If

        .AddItem StrOneRowData
        .RowOutlineLevel(0) = 1
        .IsSubtotal(.Rows - 1) = True
        .RowHeight(.Rows - 1) = 450
        .Cell(flexcpFontBold, .Rows - 1, 0) = True
        Set XFont = Me.Font
        XFont.name = "Tahoma"
        XFont.size = 12
        XFont.Charset = 178
        .Cell(flexcpFont, .Rows - 1, 0) = XFont
        '-------------------------------------------------
        .AddItem "«·„»Ì⁄« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        TransValues = GetTransactionsValues("2 Or transaction_type = 21")
        .AddItem "≈Ã„«·Ï «·„»Ì⁄«  ›Ï «·› —…:" & vbTab & FormatNumber(TransValues.TotalNet, 2, vbUseDefault, , vbTrue) & " " '& WriteNo(CStr(TransValues.TotalNet), 0)
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Items1"
    
        .AddItem "«·„»Ì⁄«  «·‰ﬁœÌ…:" & vbTab & FormatNumber(TransValues.TotalCash, 2, vbUseDefault, , vbTrue) & " " '& WriteNo(CStr(TransValues.TotalCash), 0)
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "«·„»Ì⁄«  «·√Ã·…:" & vbTab & FormatNumber(TransValues.TotalDue, 2, vbUseDefault, , vbTrue) & " " '& WriteNo(CStr(TransValues.TotalDue), 0)
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "«·⁄„·«¡„‰ ÕÌÀ «·„»Ì⁄«  «·‰ﬁœÌ… Ê«·√Ã·…"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        LngTempRow = .Rows - 1
        LngCustomersCount = LoadTransCustomers(2, TransValues.TotalNet)
        .TextMatrix(LngTempRow, 0) = .TextMatrix(LngTempRow, 0) & ":" & LngCustomersCount
            
        LoadTransItems 2
            
        TransValues = LoadSalTypeTrans
        .AddItem "«·„»Ì⁄«  «·ﬁÿ«⁄Ï:" & vbTab & TransValues.TotalCash
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        .AddItem "»Ì«‰«  «·⁄„·«¡ «·ﬁÿ«⁄Ï Ê„”ÕÊ»« Â„" & ""
        .RowOutlineLevel(.Rows - 1) = 4
        .IsSubtotal(.Rows - 1) = True
            
        .AddItem "«·„»Ì⁄«  «· Ã«—Ï:" & vbTab & TransValues.TotalDue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        .AddItem "»Ì«‰«  «·⁄„·«¡ «· Ã«—Ï Ê„”ÕÊ»« Â„" & ""
        .RowOutlineLevel(.Rows - 1) = 4
        .IsSubtotal(.Rows - 1) = True
            
        .AddItem "«·„‘ —Ì« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        TransValues = GetTransactionsValues("1 Or transaction_type = 22")
        .AddItem "≈Ã„«·Ï «·„‘ —Ì«  ›Ï «·› —…:" & vbTab & TransValues.TotalNet
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Items1"
    
        .AddItem "«·„‘ —Ì«  «·‰ﬁœÌ…:" & vbTab & TransValues.TotalCash
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "«·„‘ —Ì«  «·√Ã·…:" & vbTab & TransValues.TotalDue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "«·„Ê—œÌ‰"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        LngTempRow = .Rows - 1
        LngCustomersCount = LoadTransCustomers(1, TransValues.TotalNet)
        .TextMatrix(LngTempRow, 0) = .TextMatrix(LngTempRow, 0) & ":" & LngCustomersCount
            
        LoadTransItems 1
            
        .AddItem "„— Ã⁄ «·„»Ì⁄« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        TransValues = GetTransactionsValues(9)
        .AddItem "≈Ã„«·Ï „— Ã⁄ «·„»Ì⁄«  ›Ï «·› —…:" & vbTab & TransValues.TotalNet
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Items1"
    
        .AddItem "„— Ã⁄ «·„»Ì⁄«  «·‰ﬁœÌ…:" & vbTab & TransValues.TotalCash
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "„— Ã⁄ «·„»Ì⁄«  «·√Ã·…:" & vbTab & TransValues.TotalDue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "«·⁄„·«¡"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        LngTempRow = .Rows - 1
        LngCustomersCount = LoadTransCustomers(9, TransValues.TotalNet)
        .TextMatrix(LngTempRow, 0) = .TextMatrix(LngTempRow, 0) & ":" & LngCustomersCount
        LoadTransItems 9
        '-----------------------------------------------------------------------------------------
        .AddItem "„— Ã⁄ «·„‘ —Ì« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        TransValues = GetTransactionsValues(5)
        .AddItem "≈Ã„«·Ï „— Ã⁄ «·„‘ —Ì«  ›Ï «·› —…:" & vbTab & TransValues.TotalNet
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Items1"
    
        .AddItem "„— Ã⁄ «·„‘ —Ì«  «·‰ﬁœÌ…:" & vbTab & TransValues.TotalCash
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "„— Ã⁄ «·„‘ —Ì«  «·√Ã·…:" & vbTab & TransValues.TotalDue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .AddItem "«·„Ê—œÌ‰"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        LngTempRow = .Rows - 1
        LngCustomersCount = LoadTransCustomers(5, TransValues.TotalNet)
        .TextMatrix(LngTempRow, 0) = .TextMatrix(LngTempRow, 0) & ":" & LngCustomersCount
        LoadTransItems 5
        '------------------------------------------------------------------------------------------
        '        .AddItem "„Œ“Ê‰ «·»÷«⁄…"
        '            .RowOutlineLevel(.Rows - 1) = 2
        '            .IsSubtotal(.Rows - 1) = True
        '
        '            .AddItem "ﬁÌ„… «·„Œ“Ê‰ «Ê· «·› —…"
        '            .RowOutlineLevel(.Rows - 1) = 3
        '            .IsSubtotal(.Rows - 1) = False
        '
        '            .AddItem "ﬁÌ„… «·„Œ“Ê‰ ‰Â«Ì… «·› —…"
        '            .RowOutlineLevel(.Rows - 1) = 3
        '            .IsSubtotal(.Rows - 1) = False
        '
        '            .AddItem "«·√’‰«› «·ÃœÌœ… «· Ï ≈÷Ì›  Œ·«· «·› —…"
        '            .RowOutlineLevel(.Rows - 1) = 3
        '            .IsSubtotal(.Rows - 1) = False
            
        '-------------------------------------------------------------------------------
        .AddItem "«·Œ“‰"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        Set rs = New ADODB.Recordset

        If Not (IsNull(Me.DTPFrom.value)) Then
            StrSQL = " SELECT SUM(Note_Value * TransDir) AS BoxAccount "
            StrSQL = StrSQL & " FROM dbo.QryBoxBalance() QryBoxBalance "
            StrSQL = StrSQL + " Where QryBoxBalance.NoteDate <" & SQLDate(Me.DTPFrom.value, True)
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                SngTempValue = IIf(IsNull(rs("BoxAccount").value), 0, rs("BoxAccount").value)
            End If

        Else
            SngTempValue = 0
        End If

        .AddItem "≈Ã„«·Ï —’Ìœ «·Œ“‰ «Ê· «·› —… :" & SngTempValue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        StrSQL = " SELECT SUM(Note_Value * TransDir) AS BoxAccount "
        StrSQL = StrSQL & " FROM dbo.QryBoxBalance() QryBoxBalance "

        If Not (IsNull(Me.DTPTo.value)) Then
            StrSQL = StrSQL + " Where QryBoxBalance.NoteDate <" & SQLDate(Me.DTPTo.value, True)
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        SngTempValue = 0

        If Not (rs.BOF Or rs.EOF) Then
            SngTempValue = IIf(IsNull(rs("BoxAccount").value), 0, rs("BoxAccount").value)
        End If

        .AddItem "≈Ã„«·Ï —’Ìœ «·Œ“‰ ‰Â«Ì… «·› —… :" & vbTab & SngTempValue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If IsNull(Me.DTPFrom.value) Then
            StrStartDate = SQLDate(CDate("01/01/1900"), True)
        Else
            StrStartDate = SQLDate(Me.DTPFrom.value, True)
        End If

        StrSQL = "SELECT dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, "
        StrSQL = StrSQL + " dbo.QryBoxCreditUptoDate(dbo.TblBoxesData.BoxID," & StrStartDate & ") AS StartBal,"
        StrSQL = StrSQL + " Convert(Decimal(38,2),SUM(CASE TransDir WHEN 1 THEN  Note_Value ELSE 0 END)) AS SumIn "
        StrSQL = StrSQL + ",Convert(Decimal(38,2),SUM(CASE TransDir WHEN -1 THEN  Note_Value ELSE 0 END)) AS SumOut"
        StrSQL = StrSQL + " FROM         dbo.TblBoxesData INNER JOIN dbo.QryBoxBalance() QryBoxBalance ON " & "dbo.TblBoxesData.BoxID = QryBoxBalance.BoxID "
        StrSQL = StrSQL + " Where dbo.TblBoxesData.BoxID <> 0"

        If Not IsNull(Me.DTPFrom.value) Then
            StrSQL = StrSQL + " AND  QryBoxBalance.NoteDate >=" & SQLDate(Me.DTPFrom, True) & ""
        End If

        If Not IsNull(Me.DTPTo.value) Then
            StrSQL = StrSQL + " AND  QryBoxBalance.NoteDate <=" & SQLDate(Me.DTPTo.value, True) & ""
        End If

        StrSQL = StrSQL + " Group BY dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName "
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        .AddItem " ›«’Ì· √—’œ… «·Œ“‰"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "«”„ «·Œ“‰…" & vbTab & "«·—’Ìœ «·≈›  «ÕÏ" & vbTab & "≈Ã„«·Ï Ê«—œ" & vbTab & "≈Ã„«·Ï ’«œ—"
            StrOneRowData = StrOneRowData & vbTab & "—’Ìœ ‰Â«Ì… «·› —…"
            .AddItem StrOneRowData
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 4) = SngHeaderBackColor
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
                
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                StrOneRowData = ""
                SngTemp1 = IIf(IsNull(rs("StartBal").value), 0, rs("StartBal").value)
                SngTemp2 = IIf(IsNull(rs("SumIn").value), 0, rs("SumIn").value)
                SngTemp3 = IIf(IsNull(rs("SumOut").value), 0, rs("SumOut").value)
                StrOneRowData = rs("BoxName").value & vbTab & SngTemp1 & vbTab & SngTemp2 & vbTab & SngTemp3 & vbTab & ((SngTemp1 + SngTemp2) - SngTemp3)
                .AddItem StrOneRowData
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 4) = SngDataBackColor
        End If

        '-------------------------------------------------------------------------------
        .AddItem "«·„’—Ê›« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        StrSQL = "SELECT SUM(Note_Value) AS SumX "
        StrSQL = StrSQL + " From ExpensesReport  Where NoteID <> 0 "

        If Not (IsNull(Me.DTPFrom.value)) Then
            StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(Me.DTPFrom.value, True)
        End If

        If Not (IsNull(Me.DTPTo.value)) Then
            StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(Me.DTPTo.value, True)
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        SngTempValue = 0

        If Not (rs.BOF Or rs.EOF) Then
            SngTempValue = IIf(IsNull(rs("SumX").value), 0, rs("SumX").value)
        End If

        .AddItem "≈Ã„«·Ï „’—Ê›«  «·› —… : " & vbTab & SngTempValue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .MergeRow(.Rows - 1) = True
        StrSQL = "SELECT SUM(ExpensesReport.Note_Value) AS SumX, COUNT(ExpensesReport.NoteID) AS CountX," & "ExpensesType.Name "
        StrSQL = StrSQL + " FROM ExpensesReport INNER JOIN ExpensesType ON ExpensesReport.ExpensesID =" & "ExpensesType.ID Where (ExpensesReport.NoteID <> 0)"

        If Not (IsNull(Me.DTPFrom.value)) Then
            StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(Me.DTPFrom.value, True)
        End If

        If Not (IsNull(Me.DTPTo.value)) Then
            StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(Me.DTPTo.value, True)
        End If

        StrSQL = StrSQL + " GROUP BY ExpensesType.Name "
        StrSQL = StrSQL + " ORDER BY SUM(ExpensesReport.Note_Value) DESC "
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        .AddItem " ›«’Ì· «·„’—Ê›«  ›Ï «·› —…"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "‰Ê⁄ «·„’—Ê›« " & vbTab & "⁄œœ „—«  «·’—›" & vbTab & "≈Ã„«·Ï" & vbTab & "«·‰”»… „‰ ≈Ã„«·Ï „’«—Ì› «·› —…"
            .AddItem StrOneRowData
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 4) = SngHeaderBackColor
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
                
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                StrOneRowData = ""
                StrOneRowData = rs("Name").value & vbTab & rs("CountX").value & vbTab & rs("SumX").value
                .AddItem StrOneRowData
                .Cell(flexcpFloodPercent, .Rows - 1, 3) = 100 * val(.TextMatrix(.Rows - 1, 2)) / SngTempValue
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 4) = SngDataBackColor
        End If

        '-------------------------------------------------------------------------------
        LoadCustomersAccounts
        
        '--------------------------------------------------------------
        If SystemOptions.UserInterface = ArabicInterface Then

            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignRightCenter
                .FixedAlignment(i) = flexAlignRightCenter
            Next i

        ElseIf SystemOptions.UserInterface = EnglishInterface Then

            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignLeftCenter
                .FixedAlignment(i) = flexAlignLeftCenter
            Next i

        End If

        FormatGrid
        '--------------------------------------------------------------
        .AutoSize 0, .Cols - 1, False
        .Outline 2
    End With

End Sub

Private Sub LoadData1()
    Dim XFont As IFontDisp
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer, J As Integer
    Dim IntStartSelect As Integer, IntEndSelect As Integer
    Dim SngTempValue As Single
    Dim StrOneRowData As String
    Dim SngHeaderBackColor As Single
    Dim SngDataBackColor As Single
    Dim StrStartDate As String
    Dim SngTemp1 As Single, SngTemp2 As Single, SngTemp3 As Single
    Dim TransValues As TransactionsValues
    Dim LngCustomersCount As Long, LngTempRow As Long

    SngHeaderBackColor = &HC0C0C0
    SngDataBackColor = &HE2E9E9

    With Me.Fg
        .Clear flexClearScrollable, flexClearEverything
        .Rows = 0
        .OutlineBar = flexOutlineBarComplete
        .MergeCells = flexMergeOutline
        StrOneRowData = "«·„Ì“«‰ÌÂ ⁄‰ «·› —… "

        If Not IsNull(Me.DTPFrom.value) Then
            StrOneRowData = StrOneRowData & "„‰ " & DisplayDate(Me.DTPFrom.value)
        End If

        If Not IsNull(Me.DTPTo.value) Then
            StrOneRowData = StrOneRowData & " ≈·Ï " & DisplayDate(Me.DTPTo.value)
        End If
    
        .AddItem StrOneRowData
        .RowOutlineLevel(0) = 1
        .IsSubtotal(.Rows - 1) = True
        .RowHeight(.Rows - 1) = 450
        .Cell(flexcpFontBold, .Rows - 1, 0) = True
        Set XFont = Me.Font
        XFont.name = "Tahoma"
        XFont.size = 12
        XFont.Charset = 178
        .Cell(flexcpFont, .Rows - 1, 0) = XFont
        '-------------------------------------------------
        .AddItem "«·„»Ì⁄« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        TransValues = GetTransactionsValues("2 Or transaction_type = 21")
        .AddItem "≈Ã„«·Ï «·„»Ì⁄«  ›Ï «·› —…:" & vbTab & FormatNumber(TransValues.TotalNet, 2, vbUseDefault, , vbTrue) & " " '& WriteNo(CStr(TransValues.TotalNet), 0)
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Items1"
    
        .AddItem "«·„»Ì⁄«  «·‰ﬁœÌ…:" & vbTab & FormatNumber(TransValues.TotalCash, 2, vbUseDefault, , vbTrue) & " " '& WriteNo(CStr(TransValues.TotalCash), 0)
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "«·„»Ì⁄«  «·√Ã·…:" & vbTab & FormatNumber(TransValues.TotalDue, 2, vbUseDefault, , vbTrue) & " " '& WriteNo(CStr(TransValues.TotalDue), 0)
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "«·⁄„·«¡„‰ ÕÌÀ «·„»Ì⁄«  «·‰ﬁœÌ… Ê«·√Ã·…"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        LngTempRow = .Rows - 1
        LngCustomersCount = LoadTransCustomers(2, TransValues.TotalNet)
        .TextMatrix(LngTempRow, 0) = .TextMatrix(LngTempRow, 0) & ":" & LngCustomersCount
            
        LoadTransItems 2
            
        TransValues = LoadSalTypeTrans
        .AddItem "«·„»Ì⁄«  «·ﬁÿ«⁄Ï:" & vbTab & TransValues.TotalCash
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        .AddItem "»Ì«‰«  «·⁄„·«¡ «·ﬁÿ«⁄Ï Ê„”ÕÊ»« Â„" & ""
        .RowOutlineLevel(.Rows - 1) = 4
        .IsSubtotal(.Rows - 1) = True
            
        .AddItem "«·„»Ì⁄«  «· Ã«—Ï:" & vbTab & TransValues.TotalDue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        .AddItem "»Ì«‰«  «·⁄„·«¡ «· Ã«—Ï Ê„”ÕÊ»« Â„" & ""
        .RowOutlineLevel(.Rows - 1) = 4
        .IsSubtotal(.Rows - 1) = True
            
        .AddItem "«·„‘ —Ì« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        TransValues = GetTransactionsValues("1 Or transaction_type = 22")
        .AddItem "≈Ã„«·Ï «·„‘ —Ì«  ›Ï «·› —…:" & vbTab & TransValues.TotalNet
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Items1"
    
        .AddItem "«·„‘ —Ì«  «·‰ﬁœÌ…:" & vbTab & TransValues.TotalCash
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "«·„‘ —Ì«  «·√Ã·…:" & vbTab & TransValues.TotalDue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "«·„Ê—œÌ‰"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        LngTempRow = .Rows - 1
        LngCustomersCount = LoadTransCustomers(1, TransValues.TotalNet)
        .TextMatrix(LngTempRow, 0) = .TextMatrix(LngTempRow, 0) & ":" & LngCustomersCount
            
        LoadTransItems 1
            
        .AddItem "„— Ã⁄ «·„»Ì⁄« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        TransValues = GetTransactionsValues(9)
        .AddItem "≈Ã„«·Ï „— Ã⁄ «·„»Ì⁄«  ›Ï «·› —…:" & vbTab & TransValues.TotalNet
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Items1"
    
        .AddItem "„— Ã⁄ «·„»Ì⁄«  «·‰ﬁœÌ…:" & vbTab & TransValues.TotalCash
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "„— Ã⁄ «·„»Ì⁄«  «·√Ã·…:" & vbTab & TransValues.TotalDue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "«·⁄„·«¡"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        LngTempRow = .Rows - 1
        LngCustomersCount = LoadTransCustomers(9, TransValues.TotalNet)
        .TextMatrix(LngTempRow, 0) = .TextMatrix(LngTempRow, 0) & ":" & LngCustomersCount
        LoadTransItems 9
        '-----------------------------------------------------------------------------------------
        .AddItem "„— Ã⁄ «·„‘ —Ì« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        TransValues = GetTransactionsValues(5)
        .AddItem "≈Ã„«·Ï „— Ã⁄ «·„‘ —Ì«  ›Ï «·› —…:" & vbTab & TransValues.TotalNet
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Items1"
    
        .AddItem "„— Ã⁄ «·„‘ —Ì«  «·‰ﬁœÌ…:" & vbTab & TransValues.TotalCash
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "„— Ã⁄ «·„‘ —Ì«  «·√Ã·…:" & vbTab & TransValues.TotalDue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .AddItem "«·„Ê—œÌ‰"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        LngTempRow = .Rows - 1
        LngCustomersCount = LoadTransCustomers(5, TransValues.TotalNet)
        .TextMatrix(LngTempRow, 0) = .TextMatrix(LngTempRow, 0) & ":" & LngCustomersCount
        LoadTransItems 5
        '------------------------------------------------------------------------------------------
        '        .AddItem "„Œ“Ê‰ «·»÷«⁄…"
        '            .RowOutlineLevel(.Rows - 1) = 2
        '            .IsSubtotal(.Rows - 1) = True
        '
        '            .AddItem "ﬁÌ„… «·„Œ“Ê‰ «Ê· «·› —…"
        '            .RowOutlineLevel(.Rows - 1) = 3
        '            .IsSubtotal(.Rows - 1) = False
        '
        '            .AddItem "ﬁÌ„… «·„Œ“Ê‰ ‰Â«Ì… «·› —…"
        '            .RowOutlineLevel(.Rows - 1) = 3
        '            .IsSubtotal(.Rows - 1) = False
        '
        '            .AddItem "«·√’‰«› «·ÃœÌœ… «· Ï ≈÷Ì›  Œ·«· «·› —…"
        '            .RowOutlineLevel(.Rows - 1) = 3
        '            .IsSubtotal(.Rows - 1) = False
            
        '-------------------------------------------------------------------------------
        .AddItem "«·Œ“‰"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        Set rs = New ADODB.Recordset

        If Not (IsNull(Me.DTPFrom.value)) Then
            StrSQL = " SELECT SUM(Note_Value * TransDir) AS BoxAccount "
            StrSQL = StrSQL & " FROM dbo.QryBoxBalance() QryBoxBalance "
            StrSQL = StrSQL + " Where QryBoxBalance.NoteDate <" & SQLDate(Me.DTPFrom.value, True)
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                SngTempValue = IIf(IsNull(rs("BoxAccount").value), 0, rs("BoxAccount").value)
            End If

        Else
            SngTempValue = 0
        End If

        .AddItem "≈Ã„«·Ï —’Ìœ «·Œ“‰ «Ê· «·› —… :" & SngTempValue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        StrSQL = " SELECT SUM(Note_Value * TransDir) AS BoxAccount "
        StrSQL = StrSQL & " FROM dbo.QryBoxBalance() QryBoxBalance "

        If Not (IsNull(Me.DTPTo.value)) Then
            StrSQL = StrSQL + " Where QryBoxBalance.NoteDate <" & SQLDate(Me.DTPTo.value, True)
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        SngTempValue = 0

        If Not (rs.BOF Or rs.EOF) Then
            SngTempValue = IIf(IsNull(rs("BoxAccount").value), 0, rs("BoxAccount").value)
        End If

        .AddItem "≈Ã„«·Ï —’Ìœ «·Œ“‰ ‰Â«Ì… «·› —… :" & vbTab & SngTempValue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If IsNull(Me.DTPFrom.value) Then
            StrStartDate = SQLDate(CDate("01/01/1900"), True)
        Else
            StrStartDate = SQLDate(Me.DTPFrom.value, True)
        End If

        StrSQL = "SELECT dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, "
        StrSQL = StrSQL + " dbo.QryBoxCreditUptoDate(dbo.TblBoxesData.BoxID," & StrStartDate & ") AS StartBal,"
        StrSQL = StrSQL + " Convert(Decimal(38,2),SUM(CASE TransDir WHEN 1 THEN  Note_Value ELSE 0 END)) AS SumIn "
        StrSQL = StrSQL + ",Convert(Decimal(38,2),SUM(CASE TransDir WHEN -1 THEN  Note_Value ELSE 0 END)) AS SumOut"
        StrSQL = StrSQL + " FROM         dbo.TblBoxesData INNER JOIN dbo.QryBoxBalance() QryBoxBalance ON " & "dbo.TblBoxesData.BoxID = QryBoxBalance.BoxID "
        StrSQL = StrSQL + " Where dbo.TblBoxesData.BoxID <> 0"

        If Not IsNull(Me.DTPFrom.value) Then
            StrSQL = StrSQL + " AND  QryBoxBalance.NoteDate >=" & SQLDate(Me.DTPFrom, True) & ""
        End If

        If Not IsNull(Me.DTPTo.value) Then
            StrSQL = StrSQL + " AND  QryBoxBalance.NoteDate <=" & SQLDate(Me.DTPTo.value, True) & ""
        End If

        StrSQL = StrSQL + " Group BY dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName "
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        .AddItem " ›«’Ì· √—’œ… «·Œ“‰"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "«”„ «·Œ“‰…" & vbTab & "«·—’Ìœ «·≈›  «ÕÏ" & vbTab & "≈Ã„«·Ï Ê«—œ" & vbTab & "≈Ã„«·Ï ’«œ—"
            StrOneRowData = StrOneRowData & vbTab & "—’Ìœ ‰Â«Ì… «·› —…"
            .AddItem StrOneRowData
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 4) = SngHeaderBackColor
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
                
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                StrOneRowData = ""
                SngTemp1 = IIf(IsNull(rs("StartBal").value), 0, rs("StartBal").value)
                SngTemp2 = IIf(IsNull(rs("SumIn").value), 0, rs("SumIn").value)
                SngTemp3 = IIf(IsNull(rs("SumOut").value), 0, rs("SumOut").value)
                StrOneRowData = rs("BoxName").value & vbTab & SngTemp1 & vbTab & SngTemp2 & vbTab & SngTemp3 & vbTab & ((SngTemp1 + SngTemp2) - SngTemp3)
                .AddItem StrOneRowData
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 4) = SngDataBackColor
        End If

        '-------------------------------------------------------------------------------
        .AddItem "«·„’—Ê›« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        StrSQL = "SELECT SUM(Note_Value) AS SumX "
        StrSQL = StrSQL + " From ExpensesReport  Where NoteID <> 0 "

        If Not (IsNull(Me.DTPFrom.value)) Then
            StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(Me.DTPFrom.value, True)
        End If

        If Not (IsNull(Me.DTPTo.value)) Then
            StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(Me.DTPTo.value, True)
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        SngTempValue = 0

        If Not (rs.BOF Or rs.EOF) Then
            SngTempValue = IIf(IsNull(rs("SumX").value), 0, rs("SumX").value)
        End If

        .AddItem "≈Ã„«·Ï „’—Ê›«  «·› —… : " & vbTab & SngTempValue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .MergeRow(.Rows - 1) = True
        StrSQL = "SELECT SUM(ExpensesReport.Note_Value) AS SumX, COUNT(ExpensesReport.NoteID) AS CountX," & "ExpensesType.Name "
        StrSQL = StrSQL + " FROM ExpensesReport INNER JOIN ExpensesType ON ExpensesReport.ExpensesID =" & "ExpensesType.ID Where (ExpensesReport.NoteID <> 0)"

        If Not (IsNull(Me.DTPFrom.value)) Then
            StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(Me.DTPFrom.value, True)
        End If

        If Not (IsNull(Me.DTPTo.value)) Then
            StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(Me.DTPTo.value, True)
        End If

        StrSQL = StrSQL + " GROUP BY ExpensesType.Name "
        StrSQL = StrSQL + " ORDER BY SUM(ExpensesReport.Note_Value) DESC "
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        .AddItem " ›«’Ì· «·„’—Ê›«  ›Ï «·› —…"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "‰Ê⁄ «·„’—Ê›« " & vbTab & "⁄œœ „—«  «·’—›" & vbTab & "≈Ã„«·Ï" & vbTab & "«·‰”»… „‰ ≈Ã„«·Ï „’«—Ì› «·› —…"
            .AddItem StrOneRowData
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 4) = SngHeaderBackColor
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
                
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                StrOneRowData = ""
                StrOneRowData = rs("Name").value & vbTab & rs("CountX").value & vbTab & rs("SumX").value
                .AddItem StrOneRowData
                .Cell(flexcpFloodPercent, .Rows - 1, 3) = 100 * val(.TextMatrix(.Rows - 1, 2)) / SngTempValue
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 4) = SngDataBackColor
        End If

        '-------------------------------------------------------------------------------
        LoadCustomersAccounts
        
        '--------------------------------------------------------------
        If SystemOptions.UserInterface = ArabicInterface Then

            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignRightCenter
                .FixedAlignment(i) = flexAlignRightCenter
            Next i

        ElseIf SystemOptions.UserInterface = EnglishInterface Then

            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignLeftCenter
                .FixedAlignment(i) = flexAlignLeftCenter
            Next i

        End If

        FormatGrid
        '--------------------------------------------------------------
        .AutoSize 0, .Cols - 1, False
        .Outline 2
    End With

End Sub

Private Function GetTransactionsValues(IntTransType As String) As TransactionsValues
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "Select Isnull(TotalCash,0)as TotalCash, Isnull(TotalDue,0)as TotalDue,"
        StrSQL = StrSQL + " Isnull(TotalCash,0)+Isnull(TotalDue,0)as  NET"
        StrSQL = StrSQL + " From"
        StrSQL = StrSQL + "("
        StrSQL = StrSQL + " SELECT   Convert(Decimal(38,2), SUM(CASE WHEN PaymentType = 0 THEN TotalAfterTax ELSE 0 END)) AS TotalCash,"
        StrSQL = StrSQL + " Convert(Decimal(38,2), SUM(CASE WHEN PaymentType =1 THEN TotalAfterTax ELSE 0 END)) AS TotalDue"
        StrSQL = StrSQL + " FROM         dbo.QryTransactionsTotal() QryTransactionsTotal "
        StrSQL = StrSQL + " Where (Transaction_Type = " & IntTransType & ")"

        If Not (IsNull(Me.DTPFrom.value)) Then
            StrSQL = StrSQL + " AND Transaction_Date >=" & SQLDate(Me.DTPFrom.value, True)
        End If

        If Not (IsNull(Me.DTPTo.value)) Then
            StrSQL = StrSQL + " AND Transaction_Date <=" & SQLDate(Me.DTPTo.value, True)
        End If

        StrSQL = StrSQL + ") as xTable"
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        GetTransactionsValues.TotalCash = IIf(IsNull(rs("TotalCash").value), 0, rs("TotalCash").value)
        GetTransactionsValues.TotalDue = IIf(IsNull(rs("TotalDue").value), 0, rs("TotalDue").value)
        GetTransactionsValues.TotalNet = IIf(IsNull(rs("NET").value), 0, rs("NET").value)
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Function LoadTransCustomers(IntTransType As Integer, _
                                    SngTransTotals As Single) As Long
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrOneRowData As String
    Dim IntStartSelect As Integer, IntEndSelect As Integer
    Dim SngTempValue As Single
    Dim SngHeaderBackColor As Single
    Dim SngDataBackColor As Single
    Dim StrStartDate As String
    Dim SngTemp1 As Single, SngTemp2 As Single, SngTemp3 As Single, SngTemp4 As Single

    Dim SngCashCount As Single
    Dim SngDueCount As Single
    Dim SngCashTotal As Single
    Dim SngDueTotal As Single
    Dim LngCustomersCount As Long

    SngHeaderBackColor = &HC0C0C0
    SngDataBackColor = &HE2E9E9

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT Convert(Decimal(38,2),SUM(Case WHEN PaymentType=0 THEN QryTransactionsTotal.TotalAfterTax ELSE 0 END)) AS TotalCash,"
        StrSQL = StrSQL + " Count(Case WHEN PaymentType=0 THEN QryTransactionsTotal.TotalAfterTax END) AS CountCash,"
        StrSQL = StrSQL + " Convert(Decimal(38,2),SUM(Case WHEN PaymentType=1 THEN QryTransactionsTotal.TotalAfterTax ELSE 0 END)) AS TotalDue,"
        StrSQL = StrSQL + " Count(Case WHEN PaymentType=1 THEN QryTransactionsTotal.TotalAfterTax  END) AS CountDue,"
        StrSQL = StrSQL + " Convert(Decimal(38,2),SUM( QryTransactionsTotal.TotalAfterTax))as Total,"
        StrSQL = StrSQL + " dbo.TblCustemers.CusName"
        StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal INNER JOIN "
        StrSQL = StrSQL + " dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID"
        StrSQL = StrSQL + " Where (QryTransactionsTotal.Transaction_Type = " & IntTransType & ")"

        If Not (IsNull(Me.DTPFrom.value)) Then
            StrSQL = StrSQL + " AND Transaction_Date >=" & SQLDate(Me.DTPFrom.value, True)
        End If

        If Not (IsNull(Me.DTPTo.value)) Then
            StrSQL = StrSQL + " AND Transaction_Date <=" & SQLDate(Me.DTPTo.value, True)
        End If

        StrSQL = StrSQL + " GROUP BY  dbo.TblCustemers.CusName"
        StrSQL = StrSQL + " Order By SUM( QryTransactionsTotal.TotalAfterTax) DESC"
    End If

    With Me.Fg
        StrOneRowData = "«”„ «·⁄„Ì·" & vbTab & "‰ﬁœÏ(⁄œœ ---≈Ã„«·Ï)" & vbTab & "√Ã·(⁄œœ ---≈Ã„«·Ï)" & vbTab & "≈Ã„«·Ï " & vbTab & "«·‰”»…"
        .AddItem StrOneRowData
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 4) = SngHeaderBackColor
    End With

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            
    If Not (rs.BOF Or rs.EOF) Then

        With Me.Fg
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                LngCustomersCount = LngCustomersCount + 1
                SngTemp1 = IIf(IsNull(rs("TotalCash").value), 0, rs("TotalCash").value)
                SngTemp2 = IIf(IsNull(rs("CountCash").value), 0, rs("CountCash").value)
                SngTemp3 = IIf(IsNull(rs("TotalDue").value), 0, rs("TotalDue").value)
                SngTemp4 = IIf(IsNull(rs("CountDue").value), 0, rs("CountDue").value)
            
                SngCashCount = SngCashCount + SngTemp2
                SngDueCount = SngDueCount + SngTemp4
                SngCashTotal = SngCashTotal + SngTemp1
                SngDueTotal = SngDueTotal + SngTemp3
            
                StrOneRowData = rs("CusName").value & vbTab
                StrOneRowData = StrOneRowData & "" & SngTemp2 & " --- " & SngTemp1 & vbTab
            
                StrOneRowData = StrOneRowData & " " & SngTemp4 & " --- " & SngTemp3 & vbTab
                StrOneRowData = StrOneRowData & " " & (SngTemp2 + SngTemp4) & " --- " & (SngTemp1 + SngTemp3)
                .AddItem StrOneRowData
                .RowOutlineLevel(.Rows - 1) = 4
                .IsSubtotal(.Rows - 1) = False

                If SngTransTotals <> 0 Then
                    .TextMatrix(.Rows - 1, 4) = Format((100 * (SngTemp1 + SngTemp3)) / SngTransTotals, SystemOptions.SysDefCurrencyForamt)
                    .Cell(flexcpFloodPercent, .Rows - 1, 4) = (100 * (SngTemp1 + SngTemp3)) / SngTransTotals
                End If

                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 4) = SngDataBackColor
            StrOneRowData = ""
            StrOneRowData = "⁄œœ «·⁄„·«¡ «Ê «·„Ê—œÌ‰ : " & LngCustomersCount
            StrOneRowData = StrOneRowData & vbTab & SngCashCount & "---" & SngCashTotal
            StrOneRowData = StrOneRowData & vbTab & SngDueCount & "---" & SngDueTotal
            StrOneRowData = StrOneRowData & vbTab & (SngCashCount + SngDueCount) & "---" & (SngCashTotal + SngDueTotal)
            .AddItem StrOneRowData
            .RowOutlineLevel(.Rows - 1) = 4
            .IsSubtotal(.Rows - 1) = False
            .Cell(flexcpForeColor, .Rows - 1, 1, .Rows - 1, 4) = vbRed
            '.Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, 5) = SngHeaderBackColor
            LoadTransCustomers = LngCustomersCount
        End With

    Else
        Exit Function
    End If

End Function

Private Sub FormatGrid()
    Dim XFont As IFontDisp
    Dim i As Long

    With Me.Fg
        Set XFont = Me.Font
        XFont.name = "Tahoma"
        XFont.Charset = 178
        XFont.Bold = True
        XFont.Underline = True
        XFont.size = 10

        For i = 1 To .Rows - 1

            If .IsSubtotal(i) = True Then
                .RowHeight(i) = 450
                XFont.size = (14 - (.RowOutlineLevel(i) + 1))
                Set .Cell(flexcpFont, i, 0, i, 0) = XFont

                If .RowOutlineLevel(i) = 2 Then
                    .Cell(flexcpForeColor, i, 0, i, 0) = vbBlue
                ElseIf .RowOutlineLevel(i) = 3 Then
                    .Cell(flexcpForeColor, i, 0, i, 0) = vbRed
                ElseIf .RowOutlineLevel(i) = 4 Then
                    .Cell(flexcpForeColor, i, 0, i, 0) = vbGreen
                Else
                    .Cell(flexcpForeColor, i, 0, i, 0) = vbBlack
                End If
            End If

        Next i

    End With

End Sub

Private Sub ISButton2_Click()
    PrintData
End Sub

Private Sub PrintData()
    On Error Resume Next
    Dim Frm As FrmViewListPrint

    Set Frm = New FrmViewListPrint
    Frm.VSPrinter1.Zoom = 100
    Frm.VSPrinter1.StartDoc
    Frm.VSPrinter1.MarginLeft = 100
    Frm.VSPrinter1.MarginRight = 100
    Frm.VSPrinter1.CurrentX = 100
    Frm.VSPrinter1.CurrentY = 100
    Frm.VSPrinter1.text = "»«Ì  ··»—„ÃÌ« "
    'Frm.VSPrinter1.CurrentX = 100
    'Frm.VSPrinter1.CurrentY = 500
    Frm.Caption = " ﬁ—Ì— „Ã„⁄"
    Frm.VSPrinter1.RenderControl = Fg.hwnd
    Frm.VSPrinter1.EndDoc
    Set Frm.Fg = Me.Fg
    Frm.show
End Sub

Private Function LoadTransItems(IntTransType As Integer) As Long
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrOneRowData As String
    Dim IntStartSelect As Integer, IntEndSelect As Integer
    Dim SngTempValue As Single
    Dim SngHeaderBackColor As Single
    Dim SngDataBackColor As Single
    Dim StrStartDate As String
    Dim SngTemp1 As Single, SngTemp2 As Single, SngTemp3 As Single, SngTemp4 As Single
    Dim SngItemsTotal As Single
    Dim LngItemsCount As Long

    SngHeaderBackColor = &HC0C0C0
    SngDataBackColor = &HE2E9E9

    If IntTransType = 1 Or 22 Then
        StrOneRowData = "«·√’‰«› «· Ï œŒ·  ›Ï «·„‘ —Ì« "
    ElseIf IntTransType = 2 Or 21 Then
        StrOneRowData = "«·√’‰«› «· Ï œŒ·  ›Ï «·„»Ì⁄« "
    ElseIf IntTransType = 5 Then
        StrOneRowData = "«·√’‰«› «· Ï œŒ·  ›Ï „— Ã⁄ «·„‘ —Ì« "
    ElseIf IntTransType = 9 Then
        StrOneRowData = "«·√’‰«› «· Ï œŒ·  ›Ï „— Ã⁄ «·„»Ì⁄« "
    End If

    With Me.Fg
        .AddItem StrOneRowData
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        StrOneRowData = "«”„ «·’‰›" & vbTab & "«·ﬂ„Ì…" & vbTab & "„ Ê”ÿ «·”⁄—" & vbTab & "≈Ã„«·Ï " & vbTab & "«·‰”»…"
        .AddItem StrOneRowData
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 4) = SngHeaderBackColor
    
        StrSQL = "SELECT  Item_ID, ItemCode, ItemName, SUM(Quantity) AS SumQty, AVG(Price) AS AvgPrice"
        StrSQL = StrSQL + " From dbo.ItemsTrans "
        StrSQL = StrSQL + " Where ItemsTrans.Transaction_Type=" & IntTransType & ""

        If Not (IsNull(Me.DTPFrom.value)) Then
            StrSQL = StrSQL + " AND ItemsTrans.Transaction_Date >=" & SQLDate(Me.DTPFrom.value, True)
        End If

        If Not (IsNull(Me.DTPTo.value)) Then
            StrSQL = StrSQL + " AND ItemsTrans.Transaction_Date <=" & SQLDate(Me.DTPTo.value, True)
        End If

        StrSQL = StrSQL + " GROUP BY Item_ID, ItemCode, ItemName"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            
        If Not (rs.BOF Or rs.EOF) Then

            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                LngItemsCount = LngItemsCount + 1
            
                StrOneRowData = rs("ItemName").value & vbTab
                StrOneRowData = StrOneRowData & rs("SumQty").value & vbTab
            
                StrOneRowData = StrOneRowData & IIf(IsNull(rs("AvgPrice").value), 0, rs("AvgPrice").value) & vbTab
                SngTemp1 = (rs("SumQty").value * IIf(IsNull(rs("AvgPrice").value), 0, rs("AvgPrice").value))
                StrOneRowData = StrOneRowData & SngTemp1
                SngItemsTotal = SngItemsTotal + SngTemp1
                .AddItem StrOneRowData
                .RowOutlineLevel(.Rows - 1) = 4
                .IsSubtotal(.Rows - 1) = False
                '            If SngTransTotals <> 0 Then
                '                .TextMatrix(.Rows - 1, 4) = Format((100 * (SngTemp1 + SngTemp3)) / SngTransTotals, SystemOptions.SysDefCurrencyForamt)
                '                .Cell(flexcpFloodPercent, .Rows - 1, 4) = (100 * (SngTemp1 + SngTemp3)) / SngTransTotals
                '            End If
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 4) = SngDataBackColor

            For i = IntStartSelect To IntEndSelect
                .Cell(flexcpFloodPercent, i, 4, i, 4) = 100 * val(.TextMatrix(i, 3)) / SngItemsTotal
                .TextMatrix(i, 4) = Format(100 * val(.TextMatrix(i, 3)) / SngItemsTotal, SystemOptions.SysDefCurrencyForamt)
                .Cell(flexcpFontBold, i, 4, i, 4) = True
            Next i

            StrOneRowData = ""
            StrOneRowData = "⁄œœ «·√’‰«›: " & LngItemsCount
            .AddItem StrOneRowData
            .RowOutlineLevel(.Rows - 1) = 4
            .IsSubtotal(.Rows - 1) = False
            .Cell(flexcpForeColor, .Rows - 1, 1, .Rows - 1, 4) = vbRed
            '.Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, 5) = SngHeaderBackColor
            LoadTransItems = LngItemsCount
    
        Else
            Exit Function
        End If

    End With

End Function

Private Sub WriteToolTip(LngMouseRow As Long, _
                         LngMouseCol As Long)
    Dim StrTemp As String
    Dim StrToolTip As String
    Dim VarTemp  As Variant
    On Error GoTo hErr

    With Fg
        StrTemp = Trim$(.TextMatrix(LngMouseRow, LngMouseCol))

        If StrTemp = "" Then
            Fg.ToolTipText = ""
            Exit Sub
        ElseIf InStr(1, StrTemp, "---", vbTextCompare) <> 0 Then
            VarTemp = Split(StrTemp, "---", , vbTextCompare)

            If val(VarTemp(0)) <> 0 Then
                StrToolTip = WriteNo(CStr(VarTemp(0)), 0, False)
            End If

            If val(VarTemp(1)) <> 0 Then
                StrToolTip = StrToolTip & "     " & WriteNo(CStr(VarTemp(1)), 0, False)
            End If

        ElseIf val(StrTemp) <> 0 Then 'Â–Â «·ﬁÌ„…  Õ ÊÏ ⁄·Ï «—ﬁ«„

            If IsDblValue(StrTemp) Then
                StrToolTip = WriteNo(IsDblValue(StrTemp), 0, False)
            Else
                StrToolTip = WriteNo(val(StrTemp), 0, False)
            End If
        End If

        Fg.ToolTipText = StrToolTip
    End With

    Exit Sub
hErr:
    Fg.ToolTipText = ""
End Sub

Private Function IsDblValue(strValue As String) As Double
    On Error GoTo hErr
    'IsDblValue = CDbl(strValue)
    Exit Function
hErr:
    IsDblValue = 0
End Function

Private Sub LoadCustomersAccounts()
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrOneRowData As String

    With Me.Fg
        .AddItem "«·œ«∆‰Ì‰"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True

        If IsNull(Me.DTPFrom.value) Then
            StrOneRowData = "—’Ìœ «·œ«∆‰Ì‰ «Ê· «·› —…:" & vbTab & "0"
        Else
            StrSQL = CustomersAccountsSQL(Me.DTPFrom.value, 1)
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                StrOneRowData = "—’Ìœ «·œ«∆‰Ì‰ «Ê· «·› —…:" & vbTab & rs("SumX").value
            Else
                StrOneRowData = "—’Ìœ «·œ«∆‰Ì‰ «Ê· «·› —…:" & vbTab & "0"
            End If
        End If

        .AddItem StrOneRowData
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
     
        StrSQL = CustomersAccountsSQL(Me.DTPTo.value, 1)
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "—’Ìœ «·œ«∆‰Ì‰ ‰Â«Ì… «·› —…:" & vbTab & rs("SumX").value
        Else
            StrOneRowData = "—’Ìœ «·œ«∆‰Ì‰ ‰Â«Ì… «·› —…:" & vbTab & "0"
        End If

        .AddItem StrOneRowData
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        '-------------------------------------------------------------------------------
        .AddItem "«·„œ‰Ì‰"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True

        If IsNull(Me.DTPFrom.value) Then
            StrOneRowData = "—’Ìœ «·„œ‰Ì‰ «Ê· «·› —…:" & vbTab & "0"
        Else
            StrSQL = CustomersAccountsSQL(Me.DTPFrom.value, 0)
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                StrOneRowData = "—’Ìœ «·„œ‰Ì‰ «Ê· «·› —…:" & vbTab & Abs(rs("SumX").value)
            Else
                StrOneRowData = "—’Ìœ «·„œ‰Ì‰ «Ê· «·› —…:" & vbTab & "0"
            End If
        End If
        
        .AddItem StrOneRowData
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        
        StrSQL = CustomersAccountsSQL(Me.DTPTo.value, 0)
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "—’Ìœ «·„œ‰Ì‰ ‰Â«Ì… «·› —…:" & vbTab & Abs(rs("SumX").value)
        Else
            StrOneRowData = "—’Ìœ «·„œ‰Ì‰ ‰Â«Ì… «·› —…:" & vbTab & "0"
        End If
        
        .AddItem StrOneRowData
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

    End With

End Sub

Private Function CustomersAccountsSQL(toDate As Variant, _
                                      IntAccountType As Integer) As String

    Dim StrSQL As String
    StrSQL = "Select Cast(Sum(CurrentAccount) as decimal(38,2) )as SumX "
    StrSQL = StrSQL + " From"
    StrSQL = StrSQL + "( "
    StrSQL = StrSQL + "Select "
    StrSQL = StrSQL + "IsNull"
    StrSQL = StrSQL + "("
    StrSQL = StrSQL + "Case OpenBalanceType WHEN 0 THEN (CustomerAccount )+ (-1*OpenBalance)"
    StrSQL = StrSQL + " WHEN 1 THEN (CustomerAccount)+ (OpenBalance)"
    StrSQL = StrSQL + " Else"
    StrSQL = StrSQL + "    CustomerAccount"
    StrSQL = StrSQL + " End"
    StrSQL = StrSQL + ",0)as CurrentAccount"
    StrSQL = StrSQL + " From"
    StrSQL = StrSQL + "("
    StrSQL = StrSQL + "SELECT TOP 100 PERCENT dbo.TblCustemers.CusID, dbo.TblCustemers.CusName," & "dbo.TblCustemers.Type,"
    StrSQL = StrSQL + "dbo.TblCustemers.OpenBalance, dbo.TblCustemers.OpenBalanceType," & "dbo.TblCustemers.OpenBalanceDate,"
    StrSQL = StrSQL + "SUM(dbo.QryCustomerBalance.Note_Value * dbo.QryCustomerBalance.CreditOrDebit)" & "As CustomerAccount"
    StrSQL = StrSQL + " FROM dbo.QryCustomerBalance(0)  RIGHT JOIN dbo.TblCustemers ON " & "dbo.QryCustomerBalance.CusID = dbo.TblCustemers.CusID"

    If Not (IsNull(toDate)) Then
        StrSQL = StrSQL + " Where dbo.QryCustomerBalance.NoteDate < " & SQLDate(CDate(toDate), True) & ""
    End If

    StrSQL = StrSQL + " GROUP BY dbo.TblCustemers.CusID, dbo.TblCustemers.CusName,dbo.TblCustemers.OpenBalance ,"
    StrSQL = StrSQL + " dbo.TblCustemers.OpenBalanceType , dbo.TblCustemers.OpenBalanceDate, dbo.TblCustemers.Type"
    StrSQL = StrSQL + " ORDER BY dbo.TblCustemers.CusID"
    StrSQL = StrSQL + ")"
    StrSQL = StrSQL + "XTable"
    StrSQL = StrSQL + ")XXTable"

    If IntAccountType = 0 Then
        '«·√—’œ… «·„œÌ‰…
        StrSQL = StrSQL + " Where XXTable.CurrentAccount < 0"
    ElseIf IntAccountType = 1 Then
        '«·√—’œ… «·œ«∆‰…
        StrSQL = StrSQL + " Where XXTable.CurrentAccount > 0"
    End If

    CustomersAccountsSQL = StrSQL
End Function

Private Function LoadSalTypeTrans() As TransactionsValues
    'Â‰« ‰ﬁÊ„ »«·√” ⁄·«„ ⁄‰
    '⁄‰ «·„»Ì⁄«  «· Ã«—Ï Ê«·ﬁÿ«⁄Ï
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT SUM(CASE WHEN SAleType=0 THEN QryTransactionsTotal.TotalAfterTax ELSE 0 END) AS SumSaleType0 "
        StrSQL = StrSQL + ",SUM(CASE WHEN SAleType=1 THEN QryTransactionsTotal.TotalAfterTax ELSE 0 END) AS SumSaleType1"
        StrSQL = StrSQL + " FROM         dbo.Transactions INNER JOIN "
        StrSQL = StrSQL + "dbo.QryTransactionsTotal() QryTransactionsTotal ON dbo.Transactions.Transaction_ID = QryTransactionsTotal.Transaction_ID"
        StrSQL = StrSQL + " WHERE     (QryTransactionsTotal.Transaction_Type=2) or (QryTransactionsTotal.Transaction_Type=21) AND ((dbo.Transactions.SaleType = 0) OR"
        StrSQL = StrSQL + " (dbo.Transactions.SaleType = 1))"

        If Not (IsNull(Me.DTPFrom.value)) Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date >=" & SQLDate(Me.DTPFrom.value, True)
        End If

        If Not (IsNull(Me.DTPTo.value)) Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date <=" & SQLDate(Me.DTPTo.value, True)
        End If
    
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        LoadSalTypeTrans.TotalCash = Format(IIf(IsNull(rs("SumSaleType0").value), 0, rs("SumSaleType0").value), SystemOptions.SysDefCurrencyForamt)
        LoadSalTypeTrans.TotalDue = Format(IIf(IsNull(rs("SumSaleType1").value), 0, rs("SumSaleType1").value), SystemOptions.SysDefCurrencyForamt)
        LoadSalTypeTrans.TotalNet = LoadSalTypeTrans.TotalCash + LoadSalTypeTrans.TotalDue
    End If

    rs.Close
    Set rs = Nothing
End Function
