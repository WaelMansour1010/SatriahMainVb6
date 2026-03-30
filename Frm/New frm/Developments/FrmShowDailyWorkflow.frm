VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmShowDailyWorkflow 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÊÞÇÑíÑ ÇáÇíÑÇÏ æÇáãÕÑÉÝ"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12030
   Icon            =   "FrmShowDailyWorkflow.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnClear 
      Caption         =   "ãÓÍ"
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   8880
      Width           =   1095
      _cx             =   1931
      _cy             =   873
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
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   6645
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   12075
      Begin VB.CommandButton Command1 
         Caption         =   "ÚÑÖ"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7230
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   825
      End
      Begin VB.Frame Frame3 
         Height          =   2655
         Left            =   9000
         TabIndex        =   1
         Top             =   600
         Width           =   3015
         Begin VB.Image Image1 
            Height          =   2475
            Left            =   0
            Picture         =   "FrmShowDailyWorkflow.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2955
         End
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Bindings        =   "FrmShowDailyWorkflow.frx":28E2
         Height          =   315
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
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
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   5595
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   8745
         _cx             =   15425
         _cy             =   9869
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
         Rows            =   50
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmShowDailyWorkflow.frx":28F7
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
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÇáãæÙÝ"
         Height          =   285
         Index           =   10
         Left            =   7680
         TabIndex        =   10
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ßæÏ ÇáÕäÝ"
         Height          =   255
         Index           =   31
         Left            =   6525
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   8160
         Width           =   2925
      End
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   1320
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "ÚÑÖ ÇáÊÞÑíÑ"
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
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   825
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   14205
      _cx             =   25056
      _cy             =   1455
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   24
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
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   6
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
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ãÑÇÌÚå æäÞííã ÓíÑ ÇáÚãá"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   6360
         TabIndex        =   14
         Top             =   120
         Width           =   4980
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   4920
         Picture         =   "FrmShowDailyWorkflow.frx":2988
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "ØÈÞÇ áãÓÊÃÌÑ ãÍÏÏ"
      Height          =   195
      Index           =   5
      Left            =   5400
      TabIndex        =   3
      Top             =   2040
      Width           =   1290
   End
End
Attribute VB_Name = "FrmShowDailyWorkflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic


Private Sub ChangeLang()
On Error GoTo ErrTrap
    Me.Caption = "Workflow daily report"
   Label5.Caption = Me.Caption
  ' lbl(25).Caption = Label5.Caption

lbl(10).Caption = "Employee"
With Grid
.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("NoteDate")) = "Date"
.TextMatrix(0, .ColIndex("show")) = "Show"
End With
Command1.Caption = "Show"
ErrTrap:
End Sub

Private Sub Command1_Click()
DcboEmpName_Click (0)
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hwnd
End Sub
Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
  
  
    Set Dcombos = New ClsDataCombos
        Dcombos.GetEmployees DcboEmpName
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    Resize_Form Me

End Sub
Sub GetEmployee(Optional EmpID As Integer)
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim i As Integer
Dim sql As String
Grid.Clear flexClearScrollable, flexClearEverything
sql = "select * from TblDailyWorkflow where PerformType=5  and EmpID=" & EmpID & " "
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
With Grid

.Rows = Rs5.RecordCount + 1
Rs5.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs5("TransDate").value), "", Rs5("TransDate").value)
.TextMatrix(i, .ColIndex("OrderNo")) = IIf(IsNull(Rs5("id").value), 0, Rs5("id").value)
Rs5.MoveNext
Next i
End With
End If
End Sub

Private Sub Grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With Me.Grid
Select Case .ColKey(Col)
Case "show"
  If checkApility("FrmDailyWorkflow") = False Then
                Exit Sub
            End If
 Unload FrmDailyWorkflow
Load FrmDailyWorkflow
FrmDailyWorkflow.show
   FrmDailyWorkflow.FindRec val(.TextMatrix(Row, .ColIndex("OrderNo")))

End Select
End With
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.Grid
Select Case .ColKey(Col)
 Case "show"
            .ColComboList(.ColIndex("show")) = "..."
     End Select
    End With
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If

End Sub
Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
 If val(DcboEmpName.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , val(DcboEmpName.BoundText), EmpCode
    TxtSearchCode.Text = EmpCode
    GetEmployee val(DcboEmpName.BoundText)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    'Set DCboSearch = Nothing
End Sub

