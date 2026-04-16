VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmSearchItemShow 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ⁄—Ê÷ «·«’‰«›"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13920
   Icon            =   "FrmSearchItemShow.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   13920
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
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   645
      Index           =   1
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   3360
      Width           =   10635
      Begin VB.TextBox TxtGroup 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   240
         Width           =   4275
      End
      Begin VB.TextBox txtitem 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   4275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„Ã„Ê⁄…"
         Height          =   195
         Index           =   9
         Left            =   9615
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·’‰›"
         Height          =   195
         Index           =   7
         Left            =   4335
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   1035
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   2400
      Width           =   2775
      Begin VB.ComboBox DcbTypePoliceyDit 
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox DcbtypPolicep 
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Œ’„ ‘«„·"
         Height          =   195
         Index           =   15
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   330
         Width           =   900
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Œ’„ «·„Œ’’"
         Height          =   195
         Index           =   14
         Left            =   1455
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   660
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ ‰Â«Ì…«·⁄—÷"
      Height          =   1035
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   2400
      Width           =   3855
      Begin MSComCtl2.DTPicker DTEndFPicker 
         Height          =   330
         Left            =   90
         TabIndex        =   26
         Top             =   270
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DTEndTPicker 
         Height          =   330
         Left            =   90
         TabIndex        =   27
         Top             =   630
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   13
         Left            =   3255
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   660
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   12
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ »œ«Ì… «·⁄—÷"
      Height          =   1035
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2400
      Width           =   3855
      Begin MSComCtl2.DTPicker DTStarFPicker 
         Height          =   330
         Left            =   90
         TabIndex        =   21
         Top             =   270
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DTStarTPicker 
         Height          =   330
         Left            =   90
         TabIndex        =   22
         Top             =   630
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   11
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   8
         Left            =   3255
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   660
         Width           =   480
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   645
      Index           =   0
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1740
      Width           =   6675
      Begin VB.TextBox TxtNameShow 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   5355
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·⁄—÷"
         Height          =   195
         Index           =   0
         Left            =   5535
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·⁄—÷"
      Height          =   1035
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   3255
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   630
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   2175
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   660
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄—÷"
      Height          =   645
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1740
      Width           =   3795
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   2775
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   1020
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   525
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   10
      Top             =   3600
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
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
      TabIndex        =   11
      Top             =   3600
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
      TabIndex        =   12
      Top             =   3600
      Width           =   735
      _ExtentX        =   1296
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
   Begin VSFlex8Ctl.VSFlexGrid FgItemPloice 
      Height          =   1755
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   13905
      _cx             =   24527
      _cy             =   3096
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
      Rows            =   1
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmSearchItemShow.frx":038A
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3060
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3060
      Width           =   1785
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2700
      Width           =   2775
   End
End
Attribute VB_Name = "FrmSearchItemShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0

 GetData
           
        Case 1
         
            Me.FgItemPloice.Clear flexClearScrollable, flexClearEverything
            FgItemPloice.Rows = 1
            Me.FgItemPloice.Enabled = True
            clear_all Me
Me.DtpDateFrom.value = ""
Me.DtpDateTo.value = ""
Me.DTEndFPicker.value = ""
Me.DTEndTPicker.value = ""
Me.DTStarFPicker.value = ""
Me.DTStarTPicker.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub


Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
    ' Dcombos.GetClientName Me.DCEmp_Name
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
     If SystemOptions.UserInterface = EnglishInterface Then
     DcbtypPolicep.AddItem "Dis Value"
     DcbtypPolicep.AddItem "Dis Rate"
     DcbtypPolicep.AddItem "Dis Same Item"
     DcbtypPolicep.AddItem "Dis Another Item"
     
     DcbTypePoliceyDit.AddItem "Dis Value"
     DcbTypePoliceyDit.AddItem "Dis Rate"
     DcbTypePoliceyDit.AddItem "Dis Same Item"
     DcbTypePoliceyDit.AddItem "Dis Another Item"
   Else
   
 Me.DcbtypPolicep.AddItem "Œ’„ ﬁÌ„…"
 Me.DcbtypPolicep.AddItem "Œ’„ ‰”»…"
 Me.DcbtypPolicep.AddItem "Œ’„ ﬂ„Ì…„‰ ‰›” «·’‰›"
 Me.DcbtypPolicep.AddItem "Œ’„ ﬂ„Ì… „‰ ’‰› «Œ—"
 Me.DcbTypePoliceyDit.AddItem "Œ’„ ﬁÌ„…"
 Me.DcbTypePoliceyDit.AddItem "Œ’„ ‰”»…"
 Me.DcbTypePoliceyDit.AddItem "Œ’„ ﬂ„Ì…„‰ ‰›” «·’‰›"
 Me.DcbTypePoliceyDit.AddItem "Œ’„ ﬂ„ÌÂ „‰ ’‰› «Œ—"
 End If
    Set DCboSearch = New clsDCboSearch
    'Set DCboSearch.Client = Me.DCEmp_Name
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  '  CenterForm Me
'GetData
'    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

   ' With Me.FgItemPloice
   '     Set .WallPaper = GrdBack.Picture
   '     .AutoSize 0, .Cols - 1, False
   ' End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo
SetDtpickerDate Me.DTEndFPicker
SetDtpickerDate Me.DTEndTPicker
SetDtpickerDate Me.DTStarFPicker
SetDtpickerDate Me.DTStarTPicker
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

StrSQL = " SELECT     dbo.TblItemShow.RecordDate, dbo.TblItemShow.NameShow, dbo.TblItemShow.StartSDate, dbo.TblItemShow.ID, dbo.TblItemShow.EndDate, "
StrSQL = StrSQL & "                      TblItems_1.ItemID AS itemidd, TblItems_1.ItemName AS ItemNameD, TblItems_1.ItemNamee AS ItemNameD, dbo.Groups.GroupID, dbo.Groups.GroupName,"
StrSQL = StrSQL & "                      dbo.TblItemShow.TypePoliceD, dbo.TblItemShow.TypePoliceP, dbo.TblItemShow.AllPolice, dbo.TblItemShow.PrivatePolice, dbo.TblItems.ItemNamee,"
StrSQL = StrSQL & "                      dbo.TblItems.itemname"
StrSQL = StrSQL & " FROM         dbo.TblItemShow LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItems TblItems_1 ON dbo.TblItemShow.ItemIDD = TblItems_1.ItemID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Groups ON dbo.TblItemShow.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItems ON dbo.TblItemShow.ItemID = dbo.TblItems.ItemID"
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblItemShow.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblItemShow.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    ' If val(FrmCarAuthontication.TxtOrder.text) <> 0 Then
    '    If BolBegine = True Then
    '        StrWhere = StrWhere & " dbo.TblCardAuthorizationReform.ID =" & val(FrmCarAuthontication.TxtOrder.text) & ""
    '    Else
    '        BolBegine = True
    '        StrWhere = " Where dbo.TblCardAuthorizationReform.ID =" & val(FrmCarAuthontication.TxtOrder.text) & ""
    '    End If
    'End If

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblItemShow.ID <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblItemShow.ID <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If
    '///////////////////
      If Me.txtitem.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblItems.itemname like '%" & Me.txtitem.text & "%'"
            StrWhere = StrWhere & " or  TblItems_1.ItemName  like '%" & Me.txtitem.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblItems.itemname like '%" & Me.txtitem.text & "%' or  TblItems_1.ItemName  like '%" & Me.txtitem.text & "%'"
           ' StrWhere = " Where dbo.TblItems1.ItemNameD like '%" & Me.txtitem.text & "%'"
        End If
    End If
     If Me.TxtGroup.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.Groups.GroupName like '%" & Me.TxtGroup.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.Groups.GroupName like '%" & Me.TxtGroup.text & "%'"
        End If
    End If
     If TxtNameShow.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblItemShow.NameShow like '%" & Me.TxtNameShow.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblItemShow.NameShow like '%" & Me.TxtNameShow.text & "%'"
        End If
    End If

   If Me.DcbtypPolicep.ListIndex <> -1 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblItemShow.TypePoliceP=" & Me.DcbtypPolicep.ListIndex & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblItemShow.TypePoliceP=" & Me.DcbtypPolicep.ListIndex & ""
        End If
    End If
If Me.DcbTypePoliceyDit.ListIndex <> -1 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblItemShow.TypePoliceD=" & Me.DcbTypePoliceyDit.ListIndex & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblItemShow.TypePoliceD=" & Me.DcbTypePoliceyDit.ListIndex & ""
        End If
    End If
   ' If Me.DCUser.BoundText <> "" Then
   ''     If BolBegine = True Then
    ''        StrWhere = StrWhere & " AND    dbo.TblCardAuthorizationReform.UserID=" & Me.DCUser.BoundText & ""
     ''   Else
      ''      BolBegine = True
       '     StrWhere = " Where    dbo.TblCardAuthorizationReform.UserID=" & Me.DCUser.BoundText & ""
       ' End If
    'End If

    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblItemShow.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblItemShow.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblItemShow.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblItemShow.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------
  If Not IsNull(Me.DTEndFPicker.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblItemShow.EndDate >=" & SQLDate(Me.DTEndFPicker.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblItemShow.EndDate >=" & SQLDate(Me.DTEndFPicker.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTEndTPicker.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblItemShow.EndDate <=" & SQLDate(Me.DTEndTPicker.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblItemShow.EndDate <=" & SQLDate(Me.DTEndTPicker.value, True) & ""
        End If
    End If
    
    
      If Not IsNull(Me.DTStarFPicker.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblItemShow.StartSDate >=" & SQLDate(DTStarFPicker.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblItemShow.StartSDate >=" & SQLDate(DTStarFPicker.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTStarTPicker.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblItemShow.StartSDate <=" & SQLDate(Me.DTStarTPicker.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblItemShow.StartSDate <=" & SQLDate(Me.DTStarTPicker.value, True) & ""
        End If
    End If
    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblItemShow.ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else

        With Me.FgItemPloice
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
            
             If Not (IsNull(rs("StartSDate").value)) Then
                    .TextMatrix(i, .ColIndex("start")) = Format(rs("StartSDate").value, "yyyy/M/d")
                End If
                 If Not (IsNull(rs("EndDate").value)) Then
                    .TextMatrix(i, .ColIndex("enddate")) = Format(rs("EndDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("nameshow")) = IIf(IsNull(rs("NameShow").value), "", rs("NameShow").value)
                Me.DcbTypePoliceyDit.ListIndex = IIf(IsNull(rs("TypePoliceD").value), -1, rs("TypePoliceD").value)
                Me.DcbtypPolicep.ListIndex = IIf(IsNull(rs("TypePoliceP").value), -1, rs("TypePoliceP").value)
                .TextMatrix(i, .ColIndex("typeprivte")) = DcbTypePoliceyDit.text
                .TextMatrix(i, .ColIndex("trypeall")) = DcbtypPolicep.text
               If IsNull(rs("ItemName").value) Then
               .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("ItemNameD").value), "", rs("ItemNameD").value)
               Else
               
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                End If
                .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
                rs.MoveNext
            Next i

           ' .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub

Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
  Me.Caption = "Search Item Show"
lbreg.Caption = "Date Show"
Frame1.Caption = "Start Show"
Frame2.Caption = "End Show"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(4).Caption = "From"
lbl(3).Caption = "To"
lbl(11).Caption = "From"
lbl(8).Caption = "To"
lbl(9).Caption = "Group"
lbl(7).Caption = "Item"
lbl(12).Caption = "From"
lbl(3).Caption = "To"
lbl(0).Caption = "Name Show"
lbl(8).Caption = "To"
lbl(13).Caption = "To"
lbl(2).Caption = "Total"
'Me.lbreg.Caption = "Date Registration"
Me.lbprocess.Caption = "Show No"
lbl(15).Caption = "Type Dis All"
lbl(14).Caption = "Type Dis Private"
     With Me.FgItemPloice
        .TextMatrix(0, .ColIndex("Ser")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "Code"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("nameshow")) = "Name Show"
        .TextMatrix(0, .ColIndex("start")) = "StartShow"
       .TextMatrix(0, .ColIndex("enddate")) = "EndDate"
         .TextMatrix(0, .ColIndex("trypeall")) = "Type Dis All"
       .TextMatrix(0, .ColIndex("typeprivte")) = "Type Dis Private"
       .TextMatrix(0, .ColIndex("name")) = "Item"
       .TextMatrix(0, .ColIndex("group")) = "Group"
    End With
  '
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

