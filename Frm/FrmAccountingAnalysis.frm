VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmAccountingAnalysis 
   Caption         =   "⁄—÷ √—’œ… «·Õ”«»« "
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   9630
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   7620
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9630
      _cx             =   16986
      _cy             =   13441
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
      _GridInfo       =   $"FrmAccountingAnalysis.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   660
         Left            =   15
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   6945
         Width           =   9600
         _cx             =   16933
         _cy             =   1164
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
         Appearance      =   5
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   2
            Left            =   30
            TabIndex        =   9
            Top             =   120
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Œ—ÊÃ"
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
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   6075
         Left            =   15
         TabIndex        =   2
         Top             =   855
         Width           =   9600
         _cx             =   16933
         _cy             =   10716
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmAccountingAnalysis.frx":0081
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   825
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   15
         Width           =   9600
         _cx             =   16933
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
         Appearance      =   5
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   0
            Left            =   30
            TabIndex        =   7
            Top             =   30
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   " ÕœÌÀ «·»Ì«‰« "
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
         Begin VB.Frame Fra 
            Caption         =   "›Ï Œ·«· «·› —…"
            Height          =   765
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   30
            Width           =   2625
         End
         Begin VB.ComboBox CboAccountsIntervals 
            Height          =   315
            Left            =   5490
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   150
            Width           =   2535
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   30
            TabIndex        =   8
            Top             =   420
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "ÿ»«⁄… "
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·› —… «·„Õ«”»Ì…"
            Height          =   345
            Left            =   8040
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   180
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "FrmAccountingAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Click(Index As Integer)
    LoadData
End Sub

Private Sub Form_Load()
    Resize_Form Me, ReportSize
    LoadData
End Sub

Private Sub LoadData()
    Dim StrSQL As String
    Dim Rs_Accounts As ADODB.Recordset
    Dim LngParentRow As Long
    Dim i As Long
    Dim GrdPic As ClsBackGroundPic
    Dim IntColName As Integer
    Dim BolRtl As Boolean
    Dim LngNewRowPos As Long
    Dim XNode As VSFlex8UCtl.VSFlexNode
    Dim YNode As VSFlex8UCtl.VSFlexNode
    Dim StrTemp As String

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    Set GrdPic = New ClsBackGroundPic

    With Me.FG
        .RowHeightMin = 300
        .ScrollTrack = False
        .ScrollTips = True
        .SheetBorder = vbWhite
        .Rows = 1
        .ExtendLastCol = True
        .AutoSize 0, .Cols - 1, False

        If BolRtl = True Then
            IntColName = 2
            .AddItem "«·‹‹œ·‹‹Ì‹‹· «·„‹‹Õ‹‹«”‹‹‹»‹‹Ï"
        Else
            .AddItem "Charts Of Accounts"
            IntColName = 9
        End If

        .Rowdata(.Rows - 1) = "r"
        .IsSubtotal(.Rows - 1) = True
        .Cell(flexcpFontBold, .Rows - 1, 1) = True
        .GridLines = flexGridFlat
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarComplete
        .AllowUserResizing = flexResizeColumns
        '.NodeClosedPicture = ImgLst.ListImages("Folder").Picture
        '.NodeOpenPicture = ImgLst.ListImages("OpenFolder").Picture
        Set .WallPaper = GrdPic.Picture

        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = " SELECT ACCOUNTS.* "
            StrSQL = StrSQL + " From ACCOUNTS "
            StrSQL = StrSQL + " WHERE (((ACCOUNTS.last_account)=False)and(Account_Code <> 'r')" & " And Parent_Account_Code = 'r'); "
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = " SELECT ACCOUNTS.* "
            StrSQL = StrSQL + " From ACCOUNTS "
            StrSQL = StrSQL + " WHERE (((ACCOUNTS.last_account)=0)and(Account_Code <> 'r') " & " And Parent_Account_Code = 'r'); "
        End If

        Set Rs_Accounts = New ADODB.Recordset
        Rs_Accounts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If SystemOptions.SysDataBaseType = AccessDataBase Then
            Call LoadGridTree("r", Rs_Accounts, FG, "ACCOUNTS", "Parent_Account_Code", " (ACCOUNTS.last_account)=False ", , IntColName, vbBlue)
        Else
            Call LoadGridTree("r", Rs_Accounts, FG, "ACCOUNTS", "Parent_Account_Code", " (ACCOUNTS.last_account)=0", , IntColName, vbBlue)
        End If
    
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = " SELECT ACCOUNTS.* "
            StrSQL = StrSQL + " From ACCOUNTS "
            StrSQL = StrSQL + " WHERE (((ACCOUNTS.last_account)=TRUE)and( Account_Code <> 'r')) " & "ORDER By Account_Code DESC;"
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = " SELECT ACCOUNTS.* "
            StrSQL = StrSQL + " From ACCOUNTS "
            StrSQL = StrSQL + " WHERE (((ACCOUNTS.last_account)=1) and( Account_Code <> 'r')) " & " ORDER By Account_Code ASC;"
        End If

        Set Rs_Accounts = New ADODB.Recordset
        Rs_Accounts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (Rs_Accounts.EOF Or Rs_Accounts.BOF) Then
            Rs_Accounts.MoveFirst

            Do While Not Rs_Accounts.EOF

                If Rs_Accounts("Account_Code").value = "a1a1a2" Then
                    'Stop
                End If

                LngParentRow = FG.FindRow(CStr(Rs_Accounts("Parent_Account_Code").value), 0, , False, True)

                If LngParentRow > 0 Then
                    Set XNode = .GetNode(LngParentRow)
                    Set YNode = XNode.AddNode(flexNTLastChild, Rs_Accounts("Account_Name").value, Rs_Accounts("Account_Code").value)
                    .Rowdata(YNode.Row) = Rs_Accounts("Account_Code").value
                
                Else
                    Stop
                End If

                Rs_Accounts.MoveNext
            Loop

        End If

        '--------------------------------------------------------------------------
        StrSQL = "Select Account_ID,Account_Code,Account_Name"
        StrSQL = StrSQL + ",Sum(Debit1) as SumDebit,Sum(Credit1) as SumCredit"
        StrSQL = StrSQL + " From"
        StrSQL = StrSQL + "("
        StrSQL = StrSQL + " SELECT     dbo.ACCOUNTS.Account_ID, dbo.ACCOUNTS.Account_Code, dbo.ACCOUNTS.Account_Name"
        StrSQL = StrSQL + " ,'Debit1'=Case"
        StrSQL = StrSQL + " When dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 then"
        StrSQL = StrSQL + "     dbo.DOUBLE_ENTREY_VOUCHERS.[Value]"
        StrSQL = StrSQL + " Else "
        StrSQL = StrSQL + " 0 "
        StrSQL = StrSQL + " END,"
        StrSQL = StrSQL + " 'Credit1'=Case"
        StrSQL = StrSQL + " When dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 then"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.[Value]"
        StrSQL = StrSQL + " Else"
        StrSQL = StrSQL + " 0 "
        StrSQL = StrSQL + " End "
        StrSQL = StrSQL + " FROM dbo.ACCOUNTS LEFT OUTER JOIN"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
        StrSQL = StrSQL + " )XTable"
        StrSQL = StrSQL + " GROUP BY Account_ID, Account_Code,Account_Name"
        StrSQL = StrSQL + " Order BY Account_Code"
        Set Rs_Accounts = New ADODB.Recordset
        Rs_Accounts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (Rs_Accounts.EOF Or Rs_Accounts.BOF) Then
            Rs_Accounts.MoveFirst

            Do While Not Rs_Accounts.EOF
                LngNewRowPos = FG.FindRow(CStr(Rs_Accounts("Account_Code").value), 0, , False, True)

                If LngNewRowPos > 0 Then
                    .TextMatrix(LngNewRowPos, .ColIndex("DebitValue")) = Rs_Accounts("SumDebit").value
                    .TextMatrix(LngNewRowPos, .ColIndex("CreditValue")) = Rs_Accounts("SumCredit").value
                    '.Cell(flexcpChecked, LngNewRowPos, .ColIndex("Last_Account"), _
                     LngNewRowPos, .ColIndex("Last_Account")) = flexChecked
                Else
                    'Stop
                End If

                Rs_Accounts.MoveNext
            Loop

        End If

        '--------------------------------------------------------------------------
        '    For I = .FixedRows To .Rows - 1
        '        Set XNode = .GetNode(I)
        '        StrTemp = ModFgLib.GetNodeChildTotal(Fg, XNode, flexSTSum, Fg.ColIndex("DebitValue"), True)
        '        .TextMatrix(I, .ColIndex("DebitValue")) = StrTemp
        '    Next I
        .AutoSize 0, .Cols - 1, False
        .Outline 1
    End With

End Sub
