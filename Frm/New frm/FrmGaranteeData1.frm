VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FRMGranteeData1 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ”ÃÌ· »Ì«‰«  «·’Ì«‰… «·œÊ—Ì…"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5010
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmGaranteeData1.frx":0000
      Left            =   1320
      List            =   "FrmGaranteeData1.frx":000D
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtvlaue 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Text            =   "12"
      Top             =   1320
      Width           =   855
   End
   Begin VB.OptionButton GranteeTypeopt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "»œÊ‰ «·ﬁÿ⁄"
      Height          =   195
      Index           =   0
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   960
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton GranteeTypeopt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„⁄ «·ﬁÿ⁄"
      Height          =   195
      Index           =   1
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   8520
      Visible         =   0   'False
      Width           =   8775
   End
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   405
      Left            =   1020
      TabIndex        =   7
      Top             =   6570
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "Õ›Ÿ"
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
   Begin VB.TextBox TxtComment 
      Alignment       =   1  'Right Justify
      Height          =   975
      Left            =   30
      MaxLength       =   255
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   8910
      Width           =   4425
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   405
      Left            =   60
      TabIndex        =   8
      Top             =   6570
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "«·€«¡"
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
   Begin MSComCtl2.DTPicker GranteeStartDate 
      Height          =   330
      Left            =   1440
      TabIndex        =   9
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   393216
      Format          =   96468993
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker GranteeEndDate 
      Height          =   330
      Left            =   1440
      TabIndex        =   14
      Top             =   1680
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   393216
      Format          =   96468993
      CurrentDate     =   38784
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2355
      Left            =   360
      TabIndex        =   21
      Top             =   3840
      Width           =   4095
      _cx             =   7223
      _cy             =   4154
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   22
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmGaranteeData1.frx":003D
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
   Begin MSComCtl2.DTPicker DTRegMaintDate 
      Height          =   330
      Left            =   2280
      TabIndex        =   22
      Top             =   3120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   393216
      Format          =   96468993
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   390
      Index           =   20
      Left            =   1320
      TabIndex        =   24
      Top             =   3120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈÷«›…"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmGaranteeData1.frx":02EC
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   390
      Index           =   21
      Left            =   3120
      TabIndex        =   25
      Top             =   6480
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " Õ–› ”ÿ—"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmGaranteeData1.frx":0686
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·’Ì«‰…"
      Height          =   375
      Index           =   15
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3240
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·’Ì«‰…"
      Height          =   375
      Index           =   14
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  «·’Ì«‰…"
      Height          =   375
      Index           =   13
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‘Â—"
      Height          =   255
      Index           =   12
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "› —… «·÷„«‰"
      Height          =   255
      Index           =   11
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1320
      Width           =   2115
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·÷„«‰"
      Height          =   255
      Index           =   10
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   10155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ ‰Â«Ì… «·÷„«‰"
      Height          =   255
      Index           =   9
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1800
      Width           =   2115
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   8
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   7
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ »œ«Ì… «·÷„«‰"
      Height          =   255
      Index           =   6
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   960
      Width           =   2115
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4665
      X2              =   240
      Y1              =   2280
      Y2              =   2295
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   5
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   -1860
      Width           =   3645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·’Ì«‰…"
      Height          =   255
      Index           =   2
      Left            =   6810
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   660
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   4
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   -360
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «··ÊÕ…"
      Height          =   255
      Index           =   1
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   2115
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   3
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   300
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì·"
      Height          =   255
      Index           =   0
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   300
      Width           =   2115
   End
End
Attribute VB_Name = "FRMGranteeData1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Fg As VSFlex8UCtl.vsFlexGrid

Public LngRow As Long

Public LngCol As Long

Public AllDate As String

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 20
            addrow

        Case 21
            RemoveGridRow
    End Select

End Sub

Public Sub FillGridWithData()

    On Error GoTo ErrTrap
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    'strInputString = seriallist
    strFilterText = ","
 
    astrSplitItems = Split(Me.AllDate, strFilterText)
    Grid.Rows = UBound(astrSplitItems) + 1

    For intX = 0 To UBound(astrSplitItems)
        Grid.TextMatrix(intX + 1, Grid.ColIndex("MaDate")) = Format$(astrSplitItems(intX), "dd/mm/yyyy")
    Next
     
ErrTrap:
End Sub

Private Sub RemoveGridRow()

    With Me.Grid

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
    Dim Msg As String
    Dim ExpiryDate As Date
    Dim Askinterval As String

    If Not Fg Is Nothing Then
 
        If Me.Fg.ColIndex("GranteeType") <> -1 Then
 
            Fg.TextMatrix(LngRow, Fg.ColIndex("GranteeType")) = IIf(GranteeTypeopt(0).value = True, 0, 1)
        End If

        If Me.Fg.ColIndex("GranteeStartDate") <> -1 Then
 
            Fg.TextMatrix(LngRow, Fg.ColIndex("GranteeStartDate")) = GranteeStartDate.value
        End If

        If Me.Fg.ColIndex("GranteeEndDate") <> -1 Then
 
            Fg.TextMatrix(LngRow, Fg.ColIndex("GranteeEndDate")) = GranteeEndDate.value
        End If

        If Me.Fg.ColIndex("RegularMaintenancedates") <> -1 Then
 
            Fg.TextMatrix(LngRow, Fg.ColIndex("RegularMaintenancedates")) = AllDate
        End If

        If Me.Fg.ColIndex("guaranteeTime") <> -1 Then
 
            Fg.TextMatrix(LngRow, Fg.ColIndex("guaranteeTime")) = val(txtvlaue.text)
  
        End If

        Unload Me
    End If

End Sub

Sub addrow()
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
    Dim des As String
 
    Me.Grid.Rows = Me.Grid.Rows + 1
    LngRow = Me.Grid.Rows - 1

    With Me.Grid
 
        .TextMatrix(LngRow, .ColIndex("MaDate")) = (DTRegMaintDate.value)
        .AutoSize 0, .Cols - 1, False
    End With
  
    ReLineGrid
 
End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
    AllDate = ""

    With Me.Grid

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("MaDate")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                AllDate = AllDate & .TextMatrix(i, .ColIndex("MaDate")) & ","
         
            End If

        Next i
   
    End With

End Sub

Private Sub Form_Load()
    CenterForm Me

    FormPostion Me, GetPostion

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        cahngelang
    End If

    Me.CmdOk.ButtonStyle = impActive
    Set CmdOk.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Save").Picture
    CmdOk.ButtonPositionImage = impRightOfText

    Me.CMDCancel.ButtonStyle = impActive
    Set CMDCancel.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").Picture
    CMDCancel.ButtonPositionImage = impRightOfText
    'GranteeStartDate.value = Date
    GranteeEndDate.value = Date
    DTRegMaintDate.value = Date

End Sub

Function cahngelang()
    Me.Caption = "Guarantee Data"

    lbl(1).Caption = "ItemCode"
    lbl(2).Caption = "Item Name"
    lbl(10).Caption = "G. Type"
    GranteeTypeopt(0).Caption = "WithOut Part"
    GranteeTypeopt(1).Caption = "With Part"
    lbl(6).Caption = "Guarantee  Start Date"
    lbl(9).Caption = "Guarantee  Emd Date"
    lbl(11).Caption = "Guarantee Period"
    lbl(12).Caption = "Month"
    lbl(13).Caption = "preventive maintenance Dates"
    Cmd(20).Caption = "ADD"
    Cmd(21).Caption = "Delete Row"
    CmdOk.Caption = "Save"
    CMDCancel.Caption = "Cancel"

    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("MaDate")) = "preventive maintenance Dates"
    End With

End Function

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub
 
Public Sub txtvlaue_Change()
    Me.GranteeEndDate.value = DateAdd("M", val(Me.txtvlaue), Me.GranteeStartDate.value)
End Sub

Private Sub txtvlaue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, txtvlaue.text, 0)
End Sub
