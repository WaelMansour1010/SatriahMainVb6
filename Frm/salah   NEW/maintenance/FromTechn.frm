VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FromTechnicians 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "«Œ Ì«— «·›‰ÌÌ‰"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8790
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3675
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
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
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   405
      Left            =   1020
      TabIndex        =   1
      Top             =   3210
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
      TabIndex        =   0
      Top             =   4950
      Width           =   4425
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   405
      Left            =   60
      TabIndex        =   2
      Top             =   3210
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
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3060
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8715
      _cx             =   15372
      _cy             =   5397
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
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FromTechn.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   270
      Index           =   21
      Left            =   8040
      TabIndex        =   6
      Top             =   3000
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   476
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ–›"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FromTechn.frx":0106
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   8
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4425
      X2              =   0
      Y1              =   3120
      Y2              =   3135
   End
End
Attribute VB_Name = "FromTechnicians"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Click(Index As Integer)

    With Me.fg

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

'Public fg As VSFlex8UCtl.vsFlexGrid

'Public LngRow As Long

'Public LngCol As Long

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
Dim i As Integer
Dim auth, auth1 As String
  '  If Not FrmCarAuthontication.fg Is Nothing Then
   ' FrmCarAuthontication.fg.TextMatrix(FrmCarAuthontication.LngRow, FrmCarAuthontication.LngCol) = XPDtbBill.value    'Trim$(Me.TxtComment.text)
   auth = ""
   auth1 = ""
   For i = 1 To fg.Rows - 2
   auth = auth + ";" + fg.TextMatrix(i, fg.ColIndex("EmpID1"))
   auth1 = auth1 + "^" + fg.TextMatrix(i, fg.ColIndex("remark"))
   Next i
FrmEmpDepartments1.fg.TextMatrix(FrmEmpDepartments1.LngRow, FrmEmpDepartments1.fg.ColIndex("audStr")) = auth
FrmEmpDepartments1.fg.TextMatrix(FrmEmpDepartments1.LngRow, FrmEmpDepartments1.fg.ColIndex("remark")) = auth1
        Unload Me
   ' End If

End Sub
Public Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter  As Integer
       
    IntCounter = 0

    With fg

        For i = .FixedRows To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("EmpID1"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("serial")) = IntCounter
   
            End If

        Next i
 
    End With
    
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
 Dim StrAccountCode As String
Dim StrAccountCode1 As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim StrComboList As String

    With fg
               
    

        Select Case .ColKey(Col)
 
            Case "Emp_name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("EmpID1"), False, True)
                .TextMatrix(Row, .ColIndex("EmpID1")) = StrAccountCode
                'StrAccountCodepu = StrAccountCode
               ' StrSQL = "select * from TblExtraExpeneses where Id=" & val(StrAccountCode)
               ' rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
               '  If rs.RecordCount > 0 Then
               '     .TextMatrix(Row, .ColIndex("typeexpen")) = IIf(IsNull(rs("TypeExtrExpen").value), 0, rs("TypeExtrExpen").value)
              '  Else
               '     .TextMatrix(Row, .ColIndex("typeexpen")) = ""
               ' End If
                   End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub

Private Sub Fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With fg

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
        Select Case .ColKey(Col)
            
            Case "EmpID"
               Cancel = True
          
               '  Case "comp"
               ' fg.ComboList = ""
               '  Case "bill"
               ' fg.ComboList = ""
        End Select

    End With

    fg.ComboList = ""
End Sub
Private Sub Fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With fg

        Select Case .ColKey(Col)

            Case "Emp_name"
          ' StrSQL = " SELECT     dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.Supervisors.Emp_ID, dbo.Technicians.Emp_ID1"
          '      StrSQL = StrSQL & " FROM         dbo.TblEmployee INNER JOIN"
          '     StrSQL = StrSQL & "       dbo.Technicians ON dbo.TblEmployee.Emp_ID = dbo.Technicians.Emp_ID1 INNER JOIN"
          '    StrSQL = StrSQL & "          dbo.Supervisors ON dbo.Technicians.Emp_ID = dbo.Supervisors.Emp_ID"
  'SQL = StrSQL & " Where (dbo.Supervisors.Emp_id =" & val(FrmEmpDepartments1.StrAccountCodepu) & ")"
 StrSQL = " select * from  TblEmployee "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = fg.BuildComboList(rs, "Emp_Name", "Emp_ID")
                Else
                    StrComboList = fg.BuildComboList(rs, "Emp_Namee", "Emp_ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                 Case "audStr"
                .ColComboList(.ColIndex("audStr")) = "..."
        End Select

    End With

End Sub

Private Sub ChangeLang()
    CmdCancel.Caption = "Cancel"
CmdOk.Caption = "Save"
Cmd(21).Caption = "Delete"
Me.Caption = "Select Tchnicians"
     With Me.fg
        .TextMatrix(0, .ColIndex("serial")) = "Serial"
        .TextMatrix(0, .ColIndex("EmpID")) = " No.Tchnicians"
        .TextMatrix(0, .ColIndex("Emp_name")) = "Tchnicians Name "
.TextMatrix(0, .ColIndex("remark")) = "Remarks"
    End With
 

    
End Sub
  Public Sub EditRec(StrTable As String, _
                   RecID As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    

End Sub
Private Sub Form_Load()

If FrmEmpDepartments1.TxtModFlg.text = "R" And FrmEmpDepartments1.LngRow <> 0 Then
FrmEmpDepartments1.FiLLTXT (FrmEmpDepartments1.fg.TextMatrix(FrmEmpDepartments1.LngRow, FrmEmpDepartments1.fg.ColIndex("EmpID1")))
Else
fg.Rows = 2
fg.Enabled = True
End If
    CenterForm Me

fg.Enabled = True
    FormPostion Me, GetPostion

    Me.CmdOk.ButtonStyle = impActive
    Set CmdOk.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Save").Picture
    CmdOk.ButtonPositionImage = impRightOfText

    Me.CmdCancel.ButtonStyle = impActive
    Set CmdCancel.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").Picture
    CmdCancel.ButtonPositionImage = impRightOfText

If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

