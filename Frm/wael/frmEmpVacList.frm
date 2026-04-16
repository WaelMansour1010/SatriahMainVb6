VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Begin VB.Form frmEmpVacList 
   Caption         =   "«·ÿ·»« "
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.CommandButton cmdOK 
      Caption         =   "ok"
      Height          =   420
      Left            =   2655
      TabIndex        =   1
      Top             =   6735
      Width           =   1530
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   5130
      TabIndex        =   0
      Top             =   6720
      Width           =   1530
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   6540
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8655
      _cx             =   15266
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEmpVacList.frx":0000
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
      Begin VSFlex8UCtl.VSFlexGrid Fg2 
         Height          =   6540
         Left            =   0
         TabIndex        =   3
         Top             =   -30
         Visible         =   0   'False
         Width           =   8655
         _cx             =   15266
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   3
         GridLinesFixed  =   2
         GridLineWidth   =   2
         Rows            =   50
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   600
         RowHeightMax    =   600
         ColWidthMin     =   50
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEmpVacList.frx":00FA
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
         TextStyle       =   1
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
Attribute VB_Name = "frmEmpVacList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public code As String
Public FromDate As String
Public ToDate As String
Public salType As Boolean
Public notes As String
Public mIndex As Integer
Public mDeta As String
Public projectId As String
Public ProjectName As String


Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdOk_Click()
    If mIndex = 0 Then
    ' frmEmpVacList.Fg.TextMatrix(row, frmEmpVacList.Fg.ColIndex("Id")) = empDic("id")
        code = FG.TextMatrix(FG.row, FG.ColIndex("Code")) ' = empDic("employeeCode")
        '   fromdate            frmEmpVacList.Fg.TextMatrix(row, frmEmpVacList.Fg.ColIndex("Name")) = empDic("employeeName")
        FromDate = FG.TextMatrix(FG.row, FG.ColIndex("from")) '= Replace(empDic("startDate"), "T00:00:00", "")
        ToDate = FG.TextMatrix(FG.row, FG.ColIndex("to")) '= Replace(empDic("endDate"), "T00:00:00", "")
        notes = FG.TextMatrix(FG.row, FG.ColIndex("notes")) '= empDic("notes")
        salType = FG.TextMatrix(FG.row, FG.ColIndex("Sal")) '= empDic("chkSallary")
    ElseIf mIndex = 1 Then
    
          '   code = FG.TextMatrix(FG.row, FG.ColIndex("Code")) ' = empDic("employeeCode")
        '   fromdate            frmEmpVacList.Fg.TextMatrix(row, frmEmpVacList.Fg.ColIndex("Name")) = empDic("employeeName")
        FromDate = Fg2.TextMatrix(Fg2.row, Fg2.ColIndex("Transaction_Date")) '= Replace(empDic("startDate"), "T00:00:00", "")
        code = Fg2.TextMatrix(Fg2.row, Fg2.ColIndex("NoteSerial1")) '= Replace(empDic("endDate"), "T00:00:00", "")
        mDeta = Fg2.TextMatrix(Fg2.row, Fg2.ColIndex("Deta")) '= empDic("notes")
       ProjectName = Fg2.TextMatrix(Fg2.row, Fg2.ColIndex("ProjectName")) '= empDic("notes")
       projectId = Fg2.TextMatrix(Fg2.row, Fg2.ColIndex("projectId")) '= empDic("notes")

    End If
    Unload Me
End Sub

