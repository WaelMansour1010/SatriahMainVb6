VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmAqarListOfOwner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄ﬁ«—«  «·„·«ﬂ"
   ClientHeight    =   7965
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   8670
   HelpContextID   =   440
   Icon            =   "FrmAqarListOfOwner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   8670
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8655
      _cx             =   15266
      _cy             =   15690
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
      Align           =   0
      AutoSizeChildren=   7
      BorderWidth     =   2
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
      Begin VB.Frame Frame3 
         Height          =   6615
         Left            =   4800
         TabIndex        =   15
         Top             =   960
         Width           =   3855
         Begin VB.Label lblCompanyname 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·”« —Ì…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   5295
            Left            =   120
            TabIndex        =   16
            Top             =   2520
            Width           =   3495
         End
         Begin VB.Image Image3 
            Height          =   2310
            Left            =   120
            Picture         =   "FrmAqarListOfOwner.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3660
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2175
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   7530
         Width           =   8565
         _cx             =   15108
         _cy             =   3836
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
         BorderWidth     =   2
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
         Begin VB.TextBox txtId 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   2400
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   2520
            Visible         =   0   'False
            Width           =   5205
         End
         Begin VB.TextBox txtMessage 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   2415
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   420
            Visible         =   0   'False
            Width           =   5205
         End
         Begin VB.CheckBox ChkShow 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·«  ŸÂ— Â–Â «·‰«›–… ⁄‰œ  ‘€Ì· «·»—‰«„Ã"
            ForeColor       =   &H000000FF&
            Height          =   1050
            Left            =   3465
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   2835
            Width           =   4875
         End
         Begin ImpulseButton.ISButton CmdExit 
            Cancel          =   -1  'True
            Height          =   810
            Left            =   120
            TabIndex        =   5
            Top             =   -120
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   1429
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
            ButtonImage     =   "FrmAqarListOfOwner.frx":10A48
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
         Begin ImpulseButton.ISButton CmdPrint 
            Height          =   825
            Left            =   1950
            TabIndex        =   6
            Top             =   2100
            Visible         =   0   'False
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   1455
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
            ButtonImage     =   "FrmAqarListOfOwner.frx":10DE2
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton SendMessage 
            Height          =   585
            Left            =   960
            TabIndex        =   8
            Top             =   2625
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1032
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«—”«·"
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
            Index           =   9
            Left            =   960
            TabIndex        =   14
            Top             =   840
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â"
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
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»ÕÀ »«·ÂÊÌ…"
            Height          =   510
            Left            =   7485
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   2520
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»ÕÀ »«·«”„"
            Height          =   510
            Left            =   7740
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   420
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ì „  ÕœÌœ Â–Â «·»Ì«‰«  »‰«¡« ⁄·Ï «· «—ÌŒ «·Õ«·Ì ›Ì «·ÃÂ«“"
            ForeColor       =   &H000000FF&
            Height          =   675
            Left            =   4035
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   3405
            Visible         =   0   'False
            Width           =   4425
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   6480
         Left            =   0
         TabIndex        =   2
         Top             =   960
         Width           =   4755
         _cx             =   8387
         _cy             =   11430
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
         Rows            =   50
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmAqarListOfOwner.frx":1117C
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
      Begin VB.CheckBox Check17 
         Alignment       =   1  'Right Justify
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   255
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   765
         Width           =   1425
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   2520
         Picture         =   "FrmAqarListOfOwner.frx":11347
         Stretch         =   -1  'True
         Top             =   240
         Width           =   525
      End
      Begin VB.Image Image2 
         Height          =   615
         Left            =   120
         Picture         =   "FrmAqarListOfOwner.frx":14FAF
         Stretch         =   -1  'True
         Top             =   120
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "FrmAqarListOfOwner.frx":15E2B
         Top             =   30
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label LblCaption 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         Caption         =   "⁄ﬁ«—«  «·„·«ﬂ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   990
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   30
         Width           =   8580
      End
   End
End
Attribute VB_Name = "FrmAqarListOfOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Askinterval As String
Dim Askcount As Integer
Dim sql As String

'Private Sub Check17_Click()
'    Dim i As Integer
'
'    If Check17.value = vbChecked Then
'
'        With Me.Fg
'
'            For i = 1 To .Rows - 2
'
'                .TextMatrix(i, .ColIndex("Send")) = True
'            Next i
'
'        End With
'
'    Else
'
'        With Me.Fg
'
'            For i = 1 To .Rows - 2
'
'                .TextMatrix(i, .ColIndex("Send")) = False
'            Next i

'        End With

'    End If

'End Sub


'Private Sub Cmd_Click(Index As Integer)
'print_report sql
'End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

'Private Sub CmdPrint_Click()
'
'    If DoPremis(Do_Print, Me.name, True) = False Then
'        Exit Sub
'    End If
'
'    On Error GoTo ErrTrap
'    Dim Reports As ClsRepoerts
'    Dim StrSQL As String
'
'    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_InstallmentMustPayed", True)
'    Askcount = GetSetting(StrAppRegPath, "Setting", "count_InstallmentMustPayed", True)
'
'    'StrSQL = "select * From QestNotReceipted where  DueDate<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
'    ' StrSQL = StrSQL + " order by CusName,Transaction_ID,QeqtNum"
'
'    StrSQL = "SELECT     TOP 100 PERCENT dbo.QryCust_Qest.QestID, dbo.QryCust_Qest.NoteID, dbo.QryCust_Qest.QeqtNum, dbo.QryCust_Qest.PartID, dbo.QryCust_Qest.[Value], "
'    StrSQL = StrSQL + " dbo.QryCust_Qest.DueDate, dbo.QryCust_Qest.Receipt, dbo.QryCust_Qest.Summition, dbo.QryCust_Qest.CustID, dbo.QryCust_Qest.CusName,"
'    StrSQL = StrSQL + "  dbo.QryCust_Qest.Transaction_ID , dbo.QryCust_Qest.Transaction_Date, dbo.Transactions.NoteSerial1"
'    StrSQL = StrSQL + " FROM         dbo.QryCust_Qest LEFT OUTER JOIN"
'    StrSQL = StrSQL + "  dbo.Transactions ON dbo.QryCust_Qest.Transaction_ID = dbo.Transactions.Transaction_ID"
'    StrSQL = StrSQL + " WHERE     (dbo.QryCust_Qest.QestID NOT IN"
'    StrSQL = StrSQL + " (SELECT     QestID"
'    StrSQL = StrSQL + "  from InstallmentDet_Junc_Receipt"
'    StrSQL = StrSQL + " WHERE     Status <> 1))"
'    StrSQL = StrSQL + "  and DueDate <" & SQLDate(Date, True) & "'"
'    StrSQL = StrSQL + "  order by CusName,QryCust_Qest.Transaction_ID,QeqtNum"
'
 '   Set Reports = New ClsRepoerts
 '   Reports.QestMustPayed StrSQL, , LblCaption.Caption
 '   Exit Sub
'ErrTrap:
'End Sub

'Private Sub Fg_BeforeEdit(ByVal Row As Long, _
'                          ByVal Col As Long, _
'                          Cancel As Boolean)

''    If Col <> Fg.ColIndex("Send") Then
 '       Cancel = True
 '   End If

'End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim RowNum As Integer
    Dim ReCount As Integer

 
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
 
    LoadIcons
'loadgrid
 
    
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Exit Sub
ErrTrap:
End Sub
Function loadgrid(Optional owner As Integer = 0)
   Dim BGround As New ClsBackGroundPic
    Dim BolShowRequest As Boolean
If owner <> 0 Then
    Dim RsTemp As New ADODB.Recordset
        Dim StrSQL As String
        StrSQL = " SELECT     *"
        StrSQL = StrSQL + " FROM         dbo.TblAqar where ownerid= " & owner & ""
If txtMessage.Text <> "" Then
                     StrSQL = StrSQL & " and  aqarname like'%" & txtMessage.Text & "%'"
       
  
End If

If txtId.Text <> "" Then
 
                StrSQL = StrSQL & " and  aqarNo =" & txtId.Text
      
End If

       ' If SystemOptions.UserInterface = ArabicInterface Then
            StrSQL = StrSQL + " Order by aqarname "
       ' Else
       '     StrSQL = StrSQL + " Order by aqarname "
       ' End If

     With Fg
            .Rows = .FixedRows
 
                
        End With
        
Dim ReCount As Integer
Dim RowNum As Integer
sql = StrSQL
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        With Fg
            .Rows = .FixedRows

            For ReCount = 1 To RsTemp.RecordCount
                .Rows = .Rows + 1
                RowNum = .Rows - 1
                    .TextMatrix(RowNum, .ColIndex("ser")) = RowNum
                ', dbo.QryCust_Qest.CustID
               ' If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("aqarname").value), "", RsTemp("aqarname").value)
               ' Else
               '     .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("CusNamee").value), "", RsTemp("CusNamee").value)
'
'                End If
 '.TextMatrix(RowNum, .ColIndex("remark2")) = IIf(IsNull(RsTemp("remark2").value), "", RsTemp("remark2").value)
 '               .TextMatrix(RowNum, .ColIndex("Numbers")) = IIf(IsNull(RsTemp("Cus_mobile").value), "", RsTemp("Cus_mobile").value)
            
                RsTemp.MoveNext
            Next ReCount

            .AutoSize 0, .Cols - 1, False
        End With

    End If

    Fg.WallPaper = BGround.Picture
    BolShowRequest = GetSetting(StrAppRegPath, "View_Type", "InstallmentMustPayed", True)

    If BolShowRequest = True Then
        ChkShow.value = Unchecked
    Else
        ChkShow.value = Checked
    End If
End If
End Function
Private Sub ChangeLang()
    Me.Caption = "Installment Must Pay"
    LblCaption.Caption = Me.Caption
    ChkShow.Caption = "Dont Show at Start"
    Label1.Caption = "Data Based in your System Date"
    Me.CmdExit.Caption = "Exit"
    Me.CmdPrint.Caption = "Print"

    With Me.Fg
        .TextMatrix(0, .ColIndex("Name")) = "Customer Name"
        .TextMatrix(0, .ColIndex("BillIID")) = "BillI ID"
        .TextMatrix(0, .ColIndex("TransDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("QestNum")) = "installm. #"
        .TextMatrix(0, .ColIndex("DueDate")) = "DueDate"
        .TextMatrix(0, .ColIndex("value")) = "value"

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If ChkShow.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "InstallmentMustPayed", False
    Else
        SaveSetting StrAppRegPath, "View_Type", "InstallmentMustPayed", True
    End If

    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub

Private Sub LoadIcons()
    On Error GoTo ErrTrap

    With Fg
        .Cell(flexcpPicture, 0, .ColIndex("Name")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillIID")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("TransDate")) = mdifrmmain.ImgLstTree.ListImages("qty").Picture
        .Cell(flexcpPicture, 0, .ColIndex("QestNum")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("DueDate")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Value")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    Exit Sub
ErrTrap:
End Sub



Private Sub ImgFavorites_Click()
AddTofaforites Me.name, Me.Caption, Me.Caption
End Sub

Private Sub LblCaption_Click()
    On Error GoTo ErrTrap

    If Me.WindowState = vbNormal Then
        Me.WindowState = vbMaximized
    Else
        Me.WindowState = vbNormal
    End If

    Exit Sub
ErrTrap:
End Sub

'Function GetNumbers()

'End Function

'Private Sub SendMessage_Click()
'    Dim Numbers As String
'    Dim RowNum As Integer
'    Dim Opt As Integer
'    Dim CurrentMessage As String
'    Numbers = ""
'
'    With Fg
'
'        For RowNum = .FixedRows To .Rows - 1
'
'            If .Cell(flexcpChecked, RowNum, .ColIndex("Send")) = flexChecked Then
'
'                '  MsgBox (.TextMatrix(RowNum, .ColIndex("Numbers")))
'                If (.TextMatrix(RowNum, .ColIndex("Numbers"))) <> "" Then
'                    If Numbers = "" Then
'                        Numbers = (.TextMatrix(RowNum, .ColIndex("Numbers")))
'                    Else
'                        Numbers = Numbers & "," & (.TextMatrix(RowNum, .ColIndex("Numbers")))
'                    End If
'
'                End If
'            End If
'
'        Next RowNum
'
'        CurrentMessage = txtMessage.text

'        If Numbers = "" Then Exit Sub
'        SMSSeTTings.SendMessage CurrentMessage, Numbers
'        SMSSeTTings.Hide
'
'    End With

'End Sub

'Private Sub txtId_Change()
'loadgrid
'End Sub

'Private Sub txtMessage_Change()
'loadgrid
'End Sub
