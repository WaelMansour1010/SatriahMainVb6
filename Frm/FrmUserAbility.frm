VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmUserAbility 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "’·«ÕÌ«  «·„” Œœ„Ì‰"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   HelpContextID   =   270
   Icon            =   "FrmUserAbility.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   4395
      Left            =   0
      TabIndex        =   12
      Top             =   1080
      Width           =   4515
      _cx             =   7964
      _cy             =   7752
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
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   12648447
      ForeColor       =   -2147483630
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "’·«ÕÌ«  ⁄«„…|’·«ÕÌ«  „Œ’’…"
      Align           =   0
      CurrTab         =   1
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
      Picture(0)      =   "FrmUserAbility.frx":038A
      Picture(1)      =   "FrmUserAbility.frx":0724
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   3930
         Left            =   45
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   45
         Width           =   4425
         _cx             =   7805
         _cy             =   6932
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
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   825
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   810
            Width           =   4155
            Begin VB.CheckBox ChkInvProfit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·Â Õﬁ „‘«Âœ… ’«›Ï «·—»Õ ›Ï «·›« Ê—…"
               Height          =   285
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   18
               Top             =   450
               Width           =   4005
            End
            Begin VB.CheckBox ChkInvAbility 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·Â «·ﬁœ—… ⁄·Ï  ⁄œÌ· «·”⁄— ›Ï ›« Ê—… «·»Ì⁄"
               Height          =   285
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   150
               Width           =   4005
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·ÕÊŸ…:-"
            Height          =   255
            Index           =   1
            Left            =   2970
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   60
            Width           =   1305
         End
         Begin VB.Image Image1 
            Height          =   495
            Left            =   2910
            Top             =   2250
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H000000C0&
            Height          =   435
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   360
            Width           =   3555
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   3930
         Left            =   -5070
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   45
         Width           =   4425
         _cx             =   7805
         _cy             =   6932
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
         Begin VSFlex8UCtl.VSFlexGrid Fg 
            Height          =   3855
            Left            =   30
            TabIndex        =   15
            Top             =   30
            Width           =   4350
            _cx             =   7673
            _cy             =   6800
            Appearance      =   2
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
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   20
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmUserAbility.frx":0ABE
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
      End
   End
   Begin VB.Frame Fram 
      BackColor       =   &H00E2E9E9&
      Height          =   1095
      Index           =   1
      Left            =   90
      TabIndex        =   6
      Top             =   5550
      Width           =   1935
      Begin ImpulseButton.ISButton CmdEdit 
         Height          =   375
         Left            =   1020
         TabIndex        =   7
         Top             =   180
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   " ⁄œÌ·"
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
         ButtonImage     =   "FrmUserAbility.frx":0B26
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
      Begin ImpulseButton.ISButton CmdUndo 
         Height          =   375
         Left            =   1020
         TabIndex        =   8
         Top             =   600
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   " —«Ã⁄"
         BackColor       =   14871017
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmUserAbility.frx":0EC0
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
      Begin ImpulseButton.ISButton CmdExit 
         Height          =   375
         Left            =   90
         TabIndex        =   9
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
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
         ButtonImage     =   "FrmUserAbility.frx":125A
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
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   375
         Left            =   90
         TabIndex        =   10
         Top             =   180
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "„”«⁄œ…"
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
         ButtonImage     =   "FrmUserAbility.frx":15F4
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
   Begin VB.Frame Fram 
      BackColor       =   &H00E2E9E9&
      Height          =   1095
      Index           =   0
      Left            =   2430
      TabIndex        =   5
      Top             =   5580
      Width           =   2025
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ’·«ÕÌ…  «„…"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   660
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   150
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Œ’Ì’"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   1005
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·€«¡ Ã„⁄ «·’·«ÕÌ«  "
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   435
         Width           =   1875
      End
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   540
      TabIndex        =   0
      Top             =   690
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„” Œœ„"
      Height          =   315
      Index           =   4
      Left            =   3090
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   690
      Width           =   975
   End
   Begin VB.Label LblCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " ’·«ÕÌ«  «·„” Œœ„Ì‰"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   615
      Index           =   7
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   4515
   End
End
Attribute VB_Name = "FrmUserAbility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CmdType As Boolean
Dim cSearchDcbo As clsDCboSearch

Private Sub CmdEdit_Click()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RowNum As Integer
    Dim Msg As String

    If CmdType = False Then
        If user_id <> 1 Then
            Msg = "·« Ì„ﬂ‰ﬂ «· ⁄œÌ· ›Ì ’·«ÕÌ«  «·„” Œœ„Ì‰" & Chr(13)
            Msg = Msg + "Â–Â «·⁄„·Ì… Œ«’… »—∆Ì” «·‘—ﬂ… √Ê «·„‰‘√…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If

        If DCboUserName.text = "" Then
            Msg = "Õœœ «”„ «·„” Œœ„ " & Chr(13)
            Msg = Msg + "«·–Ì  —€» ›Ì  ⁄œÌ· ’·«ÕÌ« Â"
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboUserName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        FG.Editable = flexEDKbdMouse
        CmdUndo.Enabled = True
        Opt(0).Enabled = True
        Opt(1).Enabled = True
        Opt(2).Enabled = True
        Opt(2).value = True
        Opt(1).value = False
        Opt(0).value = False
        Fra.Enabled = True
        Set CmdEdit.ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
        CmdEdit.Caption = "Õ›Ÿ"
    ElseIf CmdType = True Then
        FG.Editable = flexEDNone
        CmdUndo.Enabled = False
        Opt(0).Enabled = False
        Opt(1).Enabled = False
        Opt(2).Enabled = False

        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "delete * From ScreenJuncUser where [User-ID]=" & DCboUserName.BoundText
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "delete  From ScreenJuncUser where [User-ID]=" & DCboUserName.BoundText
        End If

        Cn.Execute StrSQL
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open "ScreenJuncUser", Cn, adOpenStatic, adLockOptimistic, adCmdTable

        For RowNum = 1 To FG.Rows - 1

            If left(FG.Rowdata(RowNum), 3) = "Frm" Then
                RsTemp.AddNew
                RsTemp("JuncID").value = new_id("ScreenJuncUser", "JuncID", "", True)
                RsTemp("ScreenName").value = FG.Rowdata(RowNum)
                RsTemp("User-ID").value = DCboUserName.BoundText
                RsTemp("Ability").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Apility")) = "", False, FG.TextMatrix(RowNum, FG.ColIndex("Apility")))
                RsTemp.update
            End If

        Next RowNum

        RsTemp.Close

        If Me.ChkInvAbility.value = vbChecked Then
            StrSQL = "Update TblUsers Set InvPrices=1 where UserID=" & Me.DCboUserName.BoundText
        Else
            StrSQL = "Update TblUsers Set InvPrices=0 where UserID=" & Me.DCboUserName.BoundText
        End If

        Cn.Execute StrSQL, , adExecuteNoRecords

        If Me.ChkInvProfit.value = vbChecked Then
            StrSQL = "Update TblUsers Set ShowInvProfit=1 where UserID=" & Me.DCboUserName.BoundText
        Else
            StrSQL = "Update TblUsers Set ShowInvProfit=0 where UserID=" & Me.DCboUserName.BoundText
        End If

        Cn.Execute StrSQL, , adExecuteNoRecords
    
        Set CmdEdit.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
        Fra.Enabled = False
        CmdEdit.Caption = " ⁄œÌ·"
        Msg = " „ Õ›Ÿ ’·«ÕÌ«  «·„” Œœ„"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If

    CmdType = Not CmdType
    Exit Sub
ErrTrap:
    Stop
    'Resume
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdUndo_Click()
    On Error GoTo ErrTrap
    'CmdEdit.ButtonImage = ImgLstTree.ListImages("Edit").Picture
    CmdEdit.Caption = " ⁄œÌ·"
    Opt(0).Enabled = False
    Opt(1).Enabled = False
    Opt(2).Enabled = False
    CmdUndo.Enabled = False
    FG.Editable = flexEDNone
    CmdType = False
    Exit Sub
ErrTrap:
End Sub

Private Sub DCboUserName_Change()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RowNum As Integer
    Dim RsTemp As New ADODB.Recordset

    FG.Clear flexClearScrollable, flexClearEverything
    ShowData

    For RowNum = 1 To FG.Rows - 1

        If left(FG.Rowdata(RowNum), 3) = "Frm" Then
            If DCboUserName.BoundText <> "" Then
                StrSQL = "select * From ScreenJuncUser where [User-ID]=" & DCboUserName.BoundText
                StrSQL = StrSQL + " and ScreenName='" & FG.Rowdata(RowNum) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.EOF Or RsTemp.BOF Then Exit Sub
                FG.TextMatrix(RowNum, FG.ColIndex("Apility")) = IIf(IsNull(RsTemp("Ability").value), False, RsTemp("Ability").value)
                RsTemp.Close
            End If
        End If

    Next RowNum

    StrSQL = "Select InvPrices,ShowInvProfit From TblUsers  where UserID=" & Me.DCboUserName.BoundText
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsTemp.BOF Or RsTemp.EOF) Then
        Me.ChkInvAbility.value = IIf(IsNull(RsTemp("InvPrices").value), vbUnchecked, RsTemp("InvPrices").value)
        Me.ChkInvProfit.value = IIf(IsNull(RsTemp("ShowInvProfit").value), vbUnchecked, RsTemp("ShowInvProfit").value)
    End If

    RsTemp.Close
    Set RsTemp = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

    With FG

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If .Rowdata(Row) = "FrmSaleBill" Then
            If .Cell(flexcpChecked, Row, .ColIndex("Apility")) = flexChecked Then
                Me.ChkInvAbility.Enabled = True
            Else
                Me.ChkInvAbility.Enabled = False
                Me.ChkInvAbility.value = vbUnchecked
            End If
        End If

    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

    If KeyCode = vbKeyF9 Then
        CmdUndo_Click
    End If

    If KeyCode = vbKeyF11 Then
        If CmdType = False Then
            CmdEdit_Click
        End If
    End If

    If KeyCode = vbKeyF10 Then
        If CmdType = True Then
            CmdEdit_Click
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            CmdExit_Click
        End If
    End If

End Sub

Private Sub Form_Load()
    Dim StrSQL As String
    Dim Msg As String

    On Error GoTo ErrTrap
    CenterForm Me

    FormPostion Me, GetPostion
    StrSQL = "SELECT * From TblUsers where UserID<>1"
    fill_combo DCboUserName, StrSQL
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DCboUserName
    ShowData
    Msg = "«·√„ﬂ«‰Ì«  «· «·Ì… Â«„… Ãœ« ÕÌÀ «‰Â«  ÿ»ﬁ ›Ï ›« Ê—… «·»Ì⁄"
    lbl(0).Caption = Msg
    'CmdEdit.ButtonImage = ImgLstTree.ListImages("Edit").Picture
    TabMain.CurrTab = 0
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set cSearchDcbo = Nothing
End Sub

Private Sub ShowData()
    Dim IntColName As Integer
    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    With Me.FG
        .Redraw = flexRDNone
        .GridLines = flexGridNone
        .Rows = 1
        IntColName = 1
        .AddItem "‘«‘«  «·»—‰«„Ã"
        .Rowdata(.Rows - 1) = "1G"
        .IsSubtotal(.Rows - 1) = True
        .Cell(flexcpFontBold, .Rows - 1, 1) = True
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarComplete
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        .NodeClosedPicture = mdifrmmain.ImgLstTree.ListImages("Closed_Node").Picture
        .NodeOpenPicture = mdifrmmain.ImgLstTree.ListImages("Open_Node").Picture
        .RowHeightMin = 300
        .ScrollTrack = False
        .ScrollTips = True
        .SheetBorder = vbWhite
    
        .AddItem "«·»Ì«‰«  «·√”«”Ì…"
        .Rowdata(.Rows - 1) = "1G"
        .IsSubtotal(.Rows - 1) = True
        .RowOutlineLevel(.Rows - 1) = 2
    
        .AddItem "»Ì«‰«  «·„ÊŸ›Ì‰"
        .Rowdata(.Rows - 1) = "FrmEmployee"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Supplier").Picture
    
        .AddItem "»Ì«‰«  «·⁄„·«¡"
        .Rowdata(.Rows - 1) = "FrmCustemers"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Clients").Picture
    
        .AddItem "»Ì«‰«  «·„Ê—œÌ‰"
        .Rowdata(.Rows - 1) = "FrmCompany"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
    
        .AddItem "»Ì«‰«  «·„Ã„Ê⁄« "
        .Rowdata(.Rows - 1) = "FrmGroups"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Groups").Picture
    
        .AddItem "»Ì«‰«  «·√’‰«›"
        .Rowdata(.Rows - 1) = "FrmItems"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Item2").Picture

        .AddItem "»Ì«‰«  «·„Œ«“‰"
        .Rowdata(.Rows - 1) = "FrmStoreData"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Closed_Node").Picture
        
        .AddItem "»Ì«‰«  «·»‰Êﬂ"
        .Rowdata(.Rows - 1) = "FrmBanksData"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
       
        .AddItem "»Ì«‰«  «·Œ“‰"
        .Rowdata(.Rows - 1) = "FrmBoxesData"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
    
        .AddItem "«·„⁄«„·«  «· Ã«—Ì…"
        .Rowdata(.Rows - 1) = "Group"
        .IsSubtotal(.Rows - 1) = True
        .RowOutlineLevel(.Rows - 1) = 2
    
        .AddItem "ﬁ«∆„… «·√”⁄«—"
        .Rowdata(.Rows - 1) = "FrmMainPriceList"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("LIST").Picture
    
        .AddItem "⁄—÷ √”⁄«—"
        .Rowdata(.Rows - 1) = "FrmShowPrice"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Wizared").Picture
    
        .AddItem "⁄—Ê÷ Ã«Â“…"
        .Rowdata(.Rows - 1) = "FrmTemplate"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Favorites").Picture
    
        .AddItem "›« Ê—… «·»Ì⁄"
        .Rowdata(.Rows - 1) = "FrmSaleBill"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Purchase").Picture

        .AddItem "›« Ê—… «·‘—«¡"
        .Rowdata(.Rows - 1) = "FrmBillBuy"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
    
        .AddItem "„— Ã⁄ «·„»Ì⁄« "
        .Rowdata(.Rows - 1) = "FrmReturnSalling"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Request").Picture
    
        .AddItem "„— Ã⁄ «·„‘ —Ì« "
        .Rowdata(.Rows - 1) = "FrmReturnpurchases"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("ReturnBack").Picture
    
        .AddItem "«·’Ì«‰…"
        .Rowdata(.Rows - 1) = "FrmMaintenence"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Maintenance").Picture
    
        .AddItem "«· ·›Ì« "
        .Rowdata(.Rows - 1) = "FrmDestruction"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Delete").Picture

        .AddItem "«·—’Ìœ «·«›  «ÕÌ"
        .Rowdata(.Rows - 1) = "FrmOpeningBalance"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Balance").Picture
    
        .AddItem "Ã—œ «·„Œ«“‰"
        .Rowdata(.Rows - 1) = "FrmGard"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("qty").Picture
    
        .AddItem " ÕÊÌ· «·»÷«⁄… „‰ «·„Œ“‰"
        .Rowdata(.Rows - 1) = "FrmMoving"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
    
        .AddItem "«·√ÃÂ“… «·„ «Õ…"
        .Rowdata(.Rows - 1) = "FrmSearchSerial"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Items").Picture

        .AddItem "«·«” ⁄·«„ ⁄‰ ”Ì—Ì«·"
        .Rowdata(.Rows - 1) = "FrmSerialData"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    
        .AddItem "«·√’‰«› «·„ÿ·Ê»…"
        .Rowdata(.Rows - 1) = "FrmRequest"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Required").Picture
    
        .AddItem "«·„⁄«„·«  «·„«·Ì…"
        .Rowdata(.Rows - 1) = "Currency"
        .IsSubtotal(.Rows - 1) = True
        .RowOutlineLevel(.Rows - 1) = 2
    
        .AddItem " √‰Ê«⁄ «·„’—Ê›« "
        .Rowdata(.Rows - 1) = "FrmExpensesType"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("COPY").Picture
    
        .AddItem " «·„’—Ê›« "
        .Rowdata(.Rows - 1) = "FrmExpenses"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
    
        .AddItem " «·„œ›Ê⁄« "
        .Rowdata(.Rows - 1) = "FrmPayments"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    
        .AddItem " «·„ﬁ»Ê÷« "
        .Rowdata(.Rows - 1) = "FrmCashing"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Currency").Picture
    
        .AddItem "  ‰»ÌÂ«  «·√Ê—«ﬁ «·„«·Ì… «·„” Õﬁ…"
        .Rowdata(.Rows - 1) = "FrmPaymentTime"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("qty").Picture
    
        .AddItem " Õ’Ì· «·√ﬁ”«ÿ"
        .Rowdata(.Rows - 1) = "FrmReceiptPart"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("receipt").Picture
    
        .AddItem "«·√ﬁ”«ÿ «·„ÿ·Ê»…"
        .Rowdata(.Rows - 1) = "FrmInstallmentMustPay"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
     
        '-----------------------------
        .AddItem "≈Ìœ«⁄ ›Ï «·Œ“‰…"
        .Rowdata(.Rows - 1) = "FrmBoxDeposit"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
    
        .AddItem "”Õ» „‰ «·Œ“‰…"
        .Rowdata(.Rows - 1) = "FrmBoxDrawing"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
    
        .AddItem "—’Ìœ «·Œ“‰… «·√‰"
        .Rowdata(.Rows - 1) = "FrmBoxesAccounts"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
    
        '------------------------------
        '=================================
        .AddItem "‘∆Ê‰ «·„ÊŸ›Ì‰"
        .Rowdata(.Rows - 1) = "Employees"
        .IsSubtotal(.Rows - 1) = True
        .RowOutlineLevel(.Rows - 1) = 2
    
        .AddItem "√‰Ê«⁄ «·√Ã«“« "
        .Rowdata(.Rows - 1) = "FrmVacancy"
        .RowOutlineLevel(.Rows - 1) = 3
        '.Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = MDIFrmMain.ImgLstTree.ListImages("Report").Picture
    
        .AddItem "≈⁄œ«œ „Ê«⁄Ìœ «·Õ÷Ê— “«·≈‰’—«›"
        .Rowdata(.Rows - 1) = "FrmTimeSetting"
        .RowOutlineLevel(.Rows - 1) = 3
        '.Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = MDIFrmMain.ImgLstTree.ListImages("Report").Picture
    
        .AddItem "«·„ﬂ«›«¯ "
        .Rowdata(.Rows - 1) = "FrmMkafea"
        .RowOutlineLevel(.Rows - 1) = 3
        '.Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = MDIFrmMain.ImgLstTree.ListImages("Report").Picture
    
        .AddItem "«·Œ’Ê„« "
        .Rowdata(.Rows - 1) = "FrmKhsm"
        .RowOutlineLevel(.Rows - 1) = 3
        '.Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = MDIFrmMain.ImgLstTree.ListImages("Report").Picture
    
        .AddItem "„Ê«⁄Ìœ «·Õ÷Ê—"
        .Rowdata(.Rows - 1) = "FrmPresentTime"
        .RowOutlineLevel(.Rows - 1) = 3
        '.Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = MDIFrmMain.ImgLstTree.ListImages("Report").Picture
    
        .AddItem "„Ê«⁄Ìœ «·≈‰’—«›"
        .Rowdata(.Rows - 1) = "FrmGoTime"
        .RowOutlineLevel(.Rows - 1) = 3
        '.Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = MDIFrmMain.ImgLstTree.ListImages("Report").Picture
    
        .AddItem " ”ÃÌ· «·€Ì«»"
        .Rowdata(.Rows - 1) = "FrmAbsent"
        .RowOutlineLevel(.Rows - 1) = 3
        '.Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = MDIFrmMain.ImgLstTree.ListImages("Report").Picture
    
        .AddItem "—Ê« » «·„ÊŸ›Ì‰"
        .Rowdata(.Rows - 1) = "FrmEmpSalary"
        .RowOutlineLevel(.Rows - 1) = 3
        '.Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = MDIFrmMain.ImgLstTree.ListImages("Report").Picture
    
        .AddItem "”·› «·„ÊŸ›Ì‰"
        .Rowdata(.Rows - 1) = "FrmEmpsAdvance"
        .RowOutlineLevel(.Rows - 1) = 3
        '.Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = MDIFrmMain.ImgLstTree.ListImages("Report").Picture
    
        .AddItem "—œ ”·› «·„ÊŸ›Ì‰"
        .Rowdata(.Rows - 1) = "FrmEmpsAdvancePayed"
        .RowOutlineLevel(.Rows - 1) = 3
        '.Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = MDIFrmMain.ImgLstTree.ListImages("Report").Picture
    
        '=================================
        .AddItem "«· ﬁ«—Ì—"
        .Rowdata(.Rows - 1) = "Reports"
        .IsSubtotal(.Rows - 1) = True
        .RowOutlineLevel(.Rows - 1) = 2
    
        .AddItem "«· ﬁ«—Ì— «·⁄«„…"
        .Rowdata(.Rows - 1) = "FrmReports"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Report").Picture
    
        .AddItem "«· ﬁ—Ì— «·ÌÊ„Ì"
        .Rowdata(.Rows - 1) = "FrmDailtyReport"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("DReport").Picture
 
        .AddItem "√œÊ« "
        .Rowdata(.Rows - 1) = "Group"
        .IsSubtotal(.Rows - 1) = True
        .RowOutlineLevel(.Rows - 1) = 2
    
        .AddItem " «·»«—ﬂÊœ"
        .Rowdata(.Rows - 1) = "FrmBarcode"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("barcode").Picture
    
        .AddItem "ŒÌ«—« "
        .Rowdata(.Rows - 1) = "FrmOptions"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Maintenance").Picture
    
        .AddItem "≈÷«›… „” Œœ„ ÃœÌœ"
        .Rowdata(.Rows - 1) = "FrmAddUser"
        .RowOutlineLevel(.Rows - 1) = 3
        Dim x As IPictureDisp
        Dim i As Long
        i = mdifrmmain.ilsIcons.ItemIndex("AddUser")
        Set x = mdifrmmain.ilsIcons.ItemPicture(mdifrmmain.ilsIcons.ItemKey(i - 1))
    
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = x
    
        .AddItem "Õ–› „” Œœ„"
        .Rowdata(.Rows - 1) = "FrmDelUser"
        .RowOutlineLevel(.Rows - 1) = 3
        i = mdifrmmain.ilsIcons.ItemIndex("DelUser")
        Set x = mdifrmmain.ilsIcons.ItemPicture(mdifrmmain.ilsIcons.ItemKey(i - 1))
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = x
       
        .AddItem "’·«ÕÌ«  «·„” Œœ„Ì‰"
        .Rowdata(.Rows - 1) = "FrmUserAbility"
        .RowOutlineLevel(.Rows - 1) = 3
    
        i = mdifrmmain.ilsIcons.ItemIndex("Pass")
        Set x = mdifrmmain.ilsIcons.ItemPicture(mdifrmmain.ilsIcons.ItemKey(i - 1))
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = x
       
        .AddItem " „’„„ «· ﬁ«—Ì—" ' 6x
        .Rowdata(.Rows - 1) = "FrmReportDesigner"
        .RowOutlineLevel(.Rows - 1) = 3
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("ScreenName")) = mdifrmmain.ImgLstTree.ListImages("Design").Picture
    
        .Cell(flexcpPictureAlignment, .FixedRows, .ColIndex("ScreenName"), .Rows - 1, .ColIndex("ScreenName")) = flexAlignRightCenter
        .AutoSize 0, .Cols - 1, False
        .Redraw = True
        .Outline 1
    End With

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub Opt_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim RowNum As Integer

    Select Case Index

        Case 0
            Opt(0).value = True
            Opt(1).value = False
            Opt(2).value = False

            For RowNum = 1 To FG.Rows - 1

                If left(FG.Rowdata(RowNum), 3) = "Frm" Then
                    FG.TextMatrix(RowNum, FG.ColIndex("Apility")) = True
                End If

            Next RowNum

        Case 1
            Opt(1).value = True
            Opt(0).value = False
            Opt(2).value = False

            For RowNum = 1 To FG.Rows - 1

                If left(FG.Rowdata(RowNum), 3) = "Frm" Then
                    FG.TextMatrix(RowNum, FG.ColIndex("Apility")) = False
                End If

            Next RowNum

        Case 2
            Opt(2).value = True
            Opt(1).value = False
            Opt(0).value = False
    End Select

    Exit Sub
ErrTrap:
End Sub
