VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmEmpOper 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "‘«‘… «”„«¡ «·⁄„·Ì‰ ›Ì «·„‘—Ê⁄"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16485
   Icon            =   "FrmEmpOper.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   16485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame10 
      Caption         =   "«”„«¡ «·⁄«„·Ì‰ ›Ì «·„‘—Ê⁄"
      Height          =   4335
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   480
      Width           =   16455
      Begin VB.TextBox txt_employee_count 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9960
         TabIndex        =   17
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox txt_emp_salary 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   6360
         TabIndex        =   16
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   16335
         Begin VB.CommandButton Command2 
            BackColor       =   &H80000007&
            Caption         =   "«œ—«Ã"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox TxtEmpcount 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6840
            TabIndex        =   3
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox TxtCount 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4680
            TabIndex        =   4
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton Option5 
            Alignment       =   1  'Right Justify
            Caption         =   " Œ’Ì’ ›⁄·Ì"
            Height          =   255
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Alignment       =   1  'Right Justify
            Caption         =   " ﬁœÌ—Ì"
            Height          =   255
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   120
            Value           =   -1  'True
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo dcJobTypeName 
            Height          =   315
            Left            =   11520
            TabIndex        =   0
            Top             =   120
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   2040
            TabIndex        =   5
            Top             =   120
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   92667905
            CurrentDate     =   38784
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "⁄œœ «·«Ì«„"
            Height          =   255
            Left            =   5520
            TabIndex        =   15
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   " «—ÌŒ «· Œ’Ì’"
            Height          =   255
            Left            =   3480
            TabIndex        =   14
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "⁄œœ «·⁄„«·"
            Height          =   255
            Left            =   8160
            TabIndex        =   13
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "«Œ — «·„Â‰… «·„ÿ·Ê»…"
            Height          =   255
            Left            =   14640
            TabIndex        =   12
            Top             =   120
            Width           =   1575
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   2700
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   16200
         _cx             =   28575
         _cy             =   4762
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
         Rows            =   3
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmEmpOper.frx":038A
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
         Index           =   8
         Left            =   14640
         TabIndex        =   21
         Top             =   3840
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
         ButtonImage     =   "FrmEmpOper.frx":055E
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label Label27 
         Caption         =   "«Ã„«·Ì ⁄œœ «·⁄„·"
         Height          =   255
         Left            =   11640
         TabIndex        =   20
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label29 
         Caption         =   "ﬁÌ„… «ÃÊ— «·⁄„«·"
         Height          =   255
         Left            =   8040
         TabIndex        =   19
         Top             =   3840
         Width           =   1815
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   6
      Top             =   4920
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "Õ›Ÿ"
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
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
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
      TabIndex        =   8
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   ""
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "«”„«¡ «·⁄«„·Ì‰ ›Ì «·„‘—Ê⁄"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   15
      TabIndex        =   9
      Top             =   0
      Width           =   16410
   End
End
Attribute VB_Name = "FrmEmpOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Dim currentterms As String
Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer
    Dim StartWeek As Double
    Dim EndWeek As Double
    Dim EarlyStartWeek As Double
    Dim EarlyEndWeek As Double
    Dim rs As ADODB.Recordset

IntCounter = 0
  txt_employee_count.Text = 0
  txt_emp_salary.Text = 0
        Set rs = New ADODB.Recordset
             

    With VSFlexGrid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                 .TextMatrix(i, .ColIndex("FullCode")) = currentterms & "-" & .TextMatrix(i, .ColIndex("LineNo"))
               ' .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                '.TextMatrix(i, .ColIndex("fullcode")) = current_terms & "-" & IntCounter
               .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("Count"))) * val(.TextMatrix(i, .ColIndex("daysalary")))
                txt_employee_count.Text = IntCounter
                txt_emp_salary.Text = val(txt_emp_salary.Text) + val(.TextMatrix(i, .ColIndex("total")))
            End If
        
        Next i

  
    End With



End Sub


Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
    save
    Unload Me
' GetData
           
      '  Case 1
           ' clear_all Me
'Me.DtpDateFrom.value = ""
'Me.DtpDateTo.value = ""
      '      If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
      '      Else
               ' Me.lbl(0).Caption = "Search Results"
      '      End If

      '  Case 2
      '      Unload Me
       Case 24
     '  AddNewFgRowother
       Case 8
            DeleteFgRowAther
    End Select

End Sub
Sub save()
Dim str As String
Dim i As Integer
str = ""

With Me.VSFlexGrid1
For i = 1 To .Rows - 1
 If .TextMatrix(i, .ColIndex("name")) <> "" Then
 str = str & Trim(.TextMatrix(i, .ColIndex("id"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("jobid"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("daysalary"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("Count"))) & "#"
 str = str & Trim("@")
  str = str & Chr(13)
  str = Trim(str)
 End If
Next

Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("emps")) = str
Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("total_salary")) = val(txt_emp_salary.Text)




End With
End Sub

Private Sub DeleteFgRowAther()

    With Me.VSFlexGrid1

        If .Row = -1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        .RemoveItem .Row
        '.AutoSize 0, .Cols - 1, False
     ReLineGrid
    End With

End Sub
'Private Sub AddNewFgRowother()
'
'    Dim Msg As String
'    Dim LngFindRow As Long
'    Dim LngNewRow As Long
'
'    If Me.DcbAccount.BoundText = "" Then
''        Msg = "  ÌÃ»  ÕœÌœ «”„ «·Õ”«»"
  '      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 '       Me.DcbAccount.SetFocus
 '
''        End If
'''
' '   End With
' LngNewRow = ModFgLib.SetFgForNewRow(FG, FG.ColIndex("Account_Code"))
'
'    With Me.FG
'
'    .TextMatrix(LngNewRow, .ColIndex("Serial")) = LngNewRow
'    .TextMatrix(LngNewRow, .ColIndex("Account_code1")) = Me.TxtAccountCode.text
'        .TextMatrix(LngNewRow, .ColIndex("Account_Code")) = Trim(Me.DcbAccount.BoundText)
'        .TextMatrix(LngNewRow, .ColIndex("Account_Name")) = Me.DcbAccount.text
'
'        .TextMatrix(LngNewRow, .ColIndex("TypeValue")) = Me.DcbTypevalue.ListIndex
'        .TextMatrix(LngNewRow, .ColIndex("TypeValuename")) = Me.DcbTypevalue.text
'        .TextMatrix(LngNewRow, .ColIndex("Vlue")) = val(Me.TxtValue.text)
'        .TextMatrix(LngNewRow, .ColIndex("Remark")) = Me.TxtRemark.text
'
        '.AutoSize 0, .Cols - 1, False'    If val(Me.TxtValue.text) = 0 Then
'        Msg = " ÌÃ» «œŒ«· «·ﬁÌ„Â «Ê «·‰”»Â "
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        Me.TxtValue.SetFocus
'        Exit Sub
'    End If
'
'
'
'    If val(Me.DcbTypevalue.ListIndex) = -1 Then
'        Msg = " ÌÃ»  ÕœÌœ  ‰Ê⁄ «·ﬁÌ„Â"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        Me.DcbTypevalue.SetFocus
'        Exit Sub
'    End If
'
'   ' With Me.Fg
'   '     LngFindRow = .FindRow(val(Me.Dcbiteem.BoundText), .FixedRows, .ColIndex("ItemID"), False, True)
''
''        If LngFindRow <> -1 Then
''            Msg = "Â–« «·’‰› „ÊÃÊœ ›⁄·«"
''            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
''            .SetFocus
''            Exit Sub
''        End If
'''
' '   End With
' LngNewRow = ModFgLib.SetFgForNewRow(FG, FG.ColIndex("Account_Code"))
'
'    With Me.FG
'
'    .TextMatrix(LngNewRow, .ColIndex("Serial")) = LngNewRow
'    .TextMatrix(LngNewRow, .ColIndex("Account_code1")) = Me.TxtAccountCode.text
'        .TextMatrix(LngNewRow, .ColIndex("Account_Code")) = Trim(Me.DcbAccount.BoundText)
'        .TextMatrix(LngNewRow, .ColIndex("Account_Name")) = Me.DcbAccount.text
'
'        .TextMatrix(LngNewRow, .ColIndex("TypeValue")) = Me.DcbTypevalue.ListIndex
'        .TextMatrix(LngNewRow, .ColIndex("TypeValuename")) = Me.DcbTypevalue.text
'        .TextMatrix(LngNewRow, .ColIndex("Vlue")) = val(Me.TxtValue.text)
'        .TextMatrix(LngNewRow, .ColIndex("Remark")) = Me.TxtRemark.text
'
        '.AutoSize 0, .Cols - 1, False
'    End With

Sub RetriveEmpInfo(Optional EmpID As Double = 0, Optional ByRef EmpName As String, Optional ByRef fullcode As String)
Dim rs1 As ADODB.Recordset
Dim Sql As String
Set rs1 = New ADODB.Recordset
Sql = " select * from TblEmployee where Emp_ID =" & EmpID & ""
rs1.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If rs1.RecordCount > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
EmpName = IIf(IsNull(rs1("Emp_Name").value), "", rs1("Emp_Name").value)
Else
EmpName = IIf(IsNull(rs1("Emp_Namee").value), "", rs1("Emp_Namee").value)
End If
fullcode = IIf(IsNull(rs1("Fullcode").value), "", rs1("Fullcode").value)

End If
End Sub
Sub RetriveJob(Optional JobID As Double = 0, Optional ByRef jobname As String)
Dim rs1 As ADODB.Recordset
Dim Sql As String
Set rs1 = New ADODB.Recordset
Sql = " select * from TblEmpJobsTypes where JobTypeID =" & JobID & ""
rs1.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If rs1.RecordCount > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
jobname = IIf(IsNull(rs1("JobTypeName").value), "", rs1("JobTypeName").value)
Else
jobname = IIf(IsNull(rs1("JobTypeNamee").value), "", rs1("JobTypeNamee").value)
End If
End If
End Sub

Private Sub Retrive(Optional project_id As Integer = 0, Optional Pand As Integer = 0, Optional Oper As Integer = 0)
    Dim jobname As String
    Dim fullcode As String
    Dim EmpName As String
    Dim i As Integer
    Dim astrSplit2tems2() As String
    Dim astrSplitItems() As String
    Dim ItemName As String
    Dim j As Integer
    Dim st As String
    Dim nElements As Integer
    'On Error GoTo ErrTrap
   ' VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
   ' VSFlexGrid1.Rows = 2
   ' VSFlexGrid1.Enabled = True
   
          
   ' StrSQL = "SELECT     dbo.TblEmpOper.ID, dbo.TblEmpOper.ProjectID, dbo.TblEmpOper.Pand, dbo.TblEmpOper.Opr, dbo.TblEmpOper.daysalary, dbo.TblEmpOper.[Count], "
   ' StrSQL = StrSQL & "                  dbo.TblEmpOper.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
   ' StrSQL = StrSQL & "                    dbo.TblEmpOper.JobID , dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee"
   ' StrSQL = StrSQL & "    FROM         dbo.TblEmpOper LEFT OUTER JOIN"
   ' StrSQL = StrSQL & "                    dbo.TblEmpJobsTypes ON dbo.TblEmpOper.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
   'StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.TblEmpOper.EmpID = dbo.TblEmployee.Emp_ID"
'StrSQL = StrSQL & "   Where (dbo.TblEmpOper.Projectid =" & project_id & ") And (dbo.TblEmpOper.Pand = " & Pand & ") And (dbo.TblEmpOper.OPR =" & Oper & ")"
'    Set RsDev = New ADODB.Recordset
'    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

'    If Not (RsDev.BOF Or RsDev.EOF) Then
'        RsDev.MoveFirst
  
        With Me.VSFlexGrid1
        If Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("emps")) <> "" Then
            st = Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("emps"))
            st = Trim(st)
            astrSplitItems = Split(st, "@")
            nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
           .Rows = .FixedRows + nElements
           For j = 0 To nElements - 1
           astrSplit2tems2 = Split(astrSplitItems(j), "#")
          i = j + 1
                
                .TextMatrix(i, .ColIndex("id")) = val(astrSplit2tems2(0))
                 RetriveEmpInfo val(astrSplit2tems2(0)), EmpName, fullcode
                 RetriveJob val(astrSplit2tems2(1)), jobname
                .TextMatrix(i, .ColIndex("daysalary")) = val(astrSplit2tems2(2))
                .TextMatrix(i, .ColIndex("Count")) = val(astrSplit2tems2(3))
                .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("Count"))) * val(.TextMatrix(i, .ColIndex("daysalary")))
                .TextMatrix(i, .ColIndex("code")) = fullcode
                .TextMatrix(i, .ColIndex("jobid")) = val(astrSplit2tems2(1))
              
               .TextMatrix(i, .ColIndex("name")) = EmpName
               .TextMatrix(i, .ColIndex("jobname")) = jobname
                Next j

           ' Me.txt_opr_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        End If
        End With

    
          
    ReLineGrid

End Sub










Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub
Function calcnets()

    With Me.VSFlexGrid1
        txt_employee_count = .Rows - 2
        Me.txt_emp_salary.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
    End With
 
End Function
Private Sub Command2_Click()

 Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String

    If Me.dcJobTypeName.BoundText <> "" Then
        If Not IsNumeric(TxtCount.Text) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ  ⁄œœ «·«Ì«„   "
            Else
                Msg = " SPecify No of Days  "
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtCount.SetFocus
            Exit Sub
        End If
        
        If Not IsNumeric(TxtEmpcount.Text) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ  ⁄œœ «·„ÿ·Ê»Ì‰ „‰ Â–… «·„Â‰…  "
            Else
                Msg = "Specify No oF labors From this Job  "
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtCount.SetFocus
            Exit Sub
        End If

        If Option4.value = True Then ' ﬁœÌ— ›ﬁÿ
            StrSQL = "SELECT     ROUND((ISNULL(dbo.TblEmployee.Emp_Salary, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_sakn, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_bus, 0) " & "       + ISNULL(dbo.TblEmployee.Emp_Salary_food, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_others, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_mob, 0) " & "       + ISNULL(dbo.TblEmployee.Emp_Salary_mang, 0)) / 30, 2) AS daysalary, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, " & "       dbo.TblEmployee.JobTypeID , dbo.TblEmployee.project_id, dbo.TblEmpJobsTypes.JobTypeName " & "  , dbo.TblEmpJobsTypes.JobTypeNamee,dbo.TblEmployee.Emp_Namee , dbo.TblEmployee.fullcode     FROM         dbo.TblEmployee INNER JOIN" & "       dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID " & "  WHERE      dbo.TblEmployee.JobTypeID =" & val(Me.dcJobTypeName.BoundText)
            ' StrSQL = "SELECT  round((isnull(Emp_Salary,0)+isnull(Emp_Salary_sakn,0) +isnull(Emp_Salary_bus,0)   +isnull(Emp_Salary_food,0)  +isnull(Emp_Salary_others,0)  +isnull(Emp_Salary_mob,0)  +isnull(Emp_Salary_mang,0))/30,2) as daysalary,* from TblEmployee Where  JobTypeID= " & Val(Me.dcJobTypeName.BoundText)
        ElseIf Option5.value = True Then
            ' StrSQL = "SELECT  round((isnull(Emp_Salary,0)+isnull(Emp_Salary_sakn,0) +isnull(Emp_Salary_bus,0)   +isnull(Emp_Salary_food,0)  +isnull(Emp_Salary_others,0)  +isnull(Emp_Salary_mob,0)  +isnull(Emp_Salary_mang,0))/30,2) as daysalary,* from TblEmployee Where  project_id=0 and JobTypeID= " & Val(Me.dcJobTypeName.BoundText) '   Œ’Ì’ ›⁄·Ï
            StrSQL = "SELECT     ROUND((ISNULL(dbo.TblEmployee.Emp_Salary, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_sakn, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_bus, 0) " & "       + ISNULL(dbo.TblEmployee.Emp_Salary_food, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_others, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_mob, 0) " & "       + ISNULL(dbo.TblEmployee.Emp_Salary_mang, 0)) / 30, 2) AS daysalary, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, " & "       dbo.TblEmployee.JobTypeID , dbo.TblEmployee.project_id, dbo.TblEmpJobsTypes.JobTypeName " & "    , dbo.TblEmpJobsTypes.JobTypeNamee,dbo.TblEmployee.Emp_Namee , dbo.TblEmployee.fullcode   FROM         dbo.TblEmployee INNER JOIN " & "       dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID " & " WHERE      dbo.TblEmployee.JobTypeID =" & val(Me.dcJobTypeName.BoundText) & " and ( dbo.TblEmployee.project_id =0 OR  dbo.TblEmployee.project_id IS NULL)"
                
        End If
      StrSQL = StrSQL & " and dbo.TblEmployee.jopstatusid=1"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Dim lastrow As Integer
        Dim X As Integer

        If rs.RecordCount > 0 Then
           ' If Option5.value = True Then
                If rs.RecordCount < val(TxtEmpcount.Text) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "«·⁄œœ «·„ÿ·Ê» „‰ «·⁄„«· €Ì— „ Ê›— Â·  —Ìœ  «· ﬂ„·… »«·⁄œœ «·„ÊÃÊœ" & Chr(13)
                        Msg = Msg & "  ‰⁄„  ﬂ„·…"
                        Msg = Msg & "  ·«  «·€«¡" & Chr(13)
                    Else
                        Msg = "No Of Labors not exist Now,continue with avilable " & Chr(13)
                        Msg = Msg & "  Yes -continue  "
                        Msg = Msg & " No - cancel" & Chr(13)
                    End If

                    X = MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

                    If X = vbNo Then
                        Exit Sub
                    End If
                
                End If
           ' End If

            rs.MoveFirst
    Dim i As Integer
            With Me.VSFlexGrid1
                lastrow = .Rows - 1
               If rs.RecordCount > val(TxtEmpcount.Text) Then
               .Rows = .Rows + val(TxtEmpcount.Text)
               Else
                .Rows = .Rows + rs.RecordCount
End If
                For i = lastrow To .Rows - 2
                
                    .TextMatrix(i, .ColIndex("LineNo")) = i
                    .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                    .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs("fullcode").value), "", rs("fullcode").value)
                    
                   If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                    .TextMatrix(i, .ColIndex("jobname")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                    Else
                     .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                    .TextMatrix(i, .ColIndex("jobname")) = IIf(IsNull(rs("JobTypeNamee").value), "", rs("JobTypeNamee").value)
                    
                   ', dbo.TblEmpJobsTypes.JobTypeNamee,dbo.TblEmployee.Emp_Namee , dbo.TblEmployee.fullcode
                    
                    End If

                    .TextMatrix(i, .ColIndex("jobid")) = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)
                    .TextMatrix(i, .ColIndex("daysalary")) = GetEmployeeSalaryAccordingToComponent(val(IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)), "")
                    .TextMatrix(i, .ColIndex("daysalary")) = Round(val(.TextMatrix(i, .ColIndex("daysalary")) / 30), 2)
                    .TextMatrix(i, .ColIndex("Count")) = val(Me.TxtCount.Text)
                    .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("daysalary"))) * val(.TextMatrix(i, .ColIndex("Count")))
                    rs.MoveNext
                Next

            End With

            calcnets
        Else

            If Option4.value = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "€Ì— „ Ê›— ⁄„«· »Â–… «·„Â‰…  "
                Else
                    Msg = "No Labors assigned to this job  "
                End If

            ElseIf Option5.value = True Then

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "€Ì— „ Ê›— ⁄„«· »Â–… «·„Â‰… «Ê «‰ ﬂ· «·⁄„«· „Œ’’Ì‰ ·„‘«—Ì⁄ «Ê ⁄„·Ì«  «Œ—Ï  "
                Else
                    Msg = "No Labors assigned to this job Or all Labors Allocated to another Project Process  "
                End If
            End If
                     
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Õœœ «·„Â‰… «Ê·« «·„ÿ·Ê»… «Ê·« "
        Else
            Msg = "Specify Job Firstly "
        End If

        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        dcJobTypeName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim EmployeeSalary As Double
    With VSFlexGrid1

        Select Case .ColKey(Col)
 
            Case "name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
         
              StrSQL = "  SELECT     dbo.TblEmployee.Fullcode, dbo.TblEmployee.JobTypeID, ROUND((ISNULL(dbo.TblEmployee.Emp_Salary, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_sakn, 0)"
              StrSQL = StrSQL & "        + ISNULL(dbo.TblEmployee.Emp_Salary_bus, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_food, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_others, 0)"
              StrSQL = StrSQL & "         + ISNULL(dbo.TblEmployee.Emp_Salary_mob, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_mang, 0)) / 30, 2) AS daysalary, dbo.TblEmpJobsTypes.JobTypeName,"
              StrSQL = StrSQL & "         dbo.TblEmpJobsTypes.JobTypeNamee , dbo.TblEmployee.Emp_id"
              StrSQL = StrSQL & "        FROM         dbo.TblEmployee LEFT OUTER JOIN"
              StrSQL = StrSQL & "         dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
              StrSQL = StrSQL & "  Where (dbo.TblEmployee.Emp_id = " & val(StrAccountCode) & ")"
                Set rs = Nothing
            EmployeeSalary = GetEmployeeSalaryAccordingToComponent(val(StrAccountCode), "")
            EmployeeSalary = Round(EmployeeSalary / 30)
                If StrAccountCode <> "" Then
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                    .TextMatrix(Row, .ColIndex("JobID")) = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)
                        .TextMatrix(Row, .ColIndex("code")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
                        .TextMatrix(Row, .ColIndex("daysalary")) = EmployeeSalary
                        
                                        
                    .TextMatrix(Row, .ColIndex("total")) = val(.TextMatrix(Row, .ColIndex("daysalary"))) * val(.TextMatrix(Row, .ColIndex("Count")))


                        If SystemOptions.UserInterface = ArabicInterface Then
                         .TextMatrix(Row, .ColIndex("jobname")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                        Else
                         .TextMatrix(Row, .ColIndex("jobname")) = IIf(IsNull(rs("JobTypeNamee").value), "", rs("JobTypeNamee").value)
                        End If
                    End If
                End If
            
                '.TextMatrix(Row, .ColIndex("id")) = get_Expenses_id(StrAccountCode)
        
            Case "code"
                  
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If
StrSQL = " SELECT     dbo.TblEmployee.JobTypeID, ROUND((ISNULL(dbo.TblEmployee.Emp_Salary, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_sakn, 0)"
StrSQL = StrSQL & "                      + ISNULL(dbo.TblEmployee.Emp_Salary_bus, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_food, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_others, 0)"
StrSQL = StrSQL & "                      + ISNULL(dbo.TblEmployee.Emp_Salary_mob, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_mang, 0)) / 30, 2) AS daysalary, dbo.TblEmpJobsTypes.JobTypeName,"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes.JobTypeNamee , dbo.TblEmployee.Emp_id, dbo.TblEmployee.emp_name, dbo.TblEmployee.Emp_Namee"
StrSQL = StrSQL & " FROM         dbo.TblEmployee LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
StrSQL = StrSQL & " WHERE     (dbo.TblEmployee.Fullcode = " & .TextMatrix(Row, Col) & ")"
              
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
          .TextMatrix(Row, .ColIndex("JobID")) = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)
                    .TextMatrix(Row, .ColIndex("id")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                    .TextMatrix(Row, .ColIndex("daysalary")) = IIf(IsNull(rs("daysalary").value), "", rs("daysalary").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Row, .ColIndex("name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                          EmployeeSalary = GetEmployeeSalaryAccordingToComponent(val(.TextMatrix(Row, .ColIndex("id"))), "")
            EmployeeSalary = Round(EmployeeSalary / 30)


                                   .TextMatrix(Row, .ColIndex("daysalary")) = EmployeeSalary
                        
                                        
                    .TextMatrix(Row, .ColIndex("total")) = val(.TextMatrix(Row, .ColIndex("daysalary"))) * val(.TextMatrix(Row, .ColIndex("Count")))


                    .TextMatrix(Row, .ColIndex("jobname")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                    Else
                    .TextMatrix(Row, .ColIndex("name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                     .TextMatrix(Row, .ColIndex("jobname")) = IIf(IsNull(rs("JobTypeNamee").value), "", rs("JobTypeNamee").value)
                    End If
                Else
                    .TextMatrix(Row, .ColIndex("id")) = ""
                    .TextMatrix(Row, .ColIndex("name")) = ""
                End If

        End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
        txt_employee_count = .Rows - 2
        Me.txt_emp_salary.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
   
    End With

    ReLineGrid
End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid1

        Select Case .ColKey(Col)

            Case "name"
                StrSQL = "select * from TblEmployee where jopstatusid=1"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = VSFlexGrid1.BuildComboList(rs, "Emp_Name", "Emp_ID")
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub
Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid1

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
        Select Case .ColKey(Col)
            
            Case "name"
                Exit Sub
        End Select

    End With

    VSFlexGrid1.ComboList = ""
End Sub





Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
Dim My_SQL As String
Dim rwOp As Integer
Dim Xpid As Integer
Dim rwpand As Integer
    Set Dcombos = New ClsDataCombos
 
   'Dcombos.GetAccountingCodes Me.DcbAccount
Frame10.Enabled = True

    Set DCboSearch = New clsDCboSearch
   
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    Set GrdBack = New ClsBackGroundPic
    If SystemOptions.UserInterface = ArabicInterface Then
 My_SQL = "   SELECT DISTINCT TOP 100 PERCENT dbo.TblEmployee.JobTypeID, dbo.TblEmpJobsTypes.JobTypeName "
 Else
  My_SQL = "   SELECT DISTINCT TOP 100 PERCENT dbo.TblEmployee.JobTypeID, dbo.TblEmpJobsTypes.JobTypeNamee"
 End If
My_SQL = My_SQL & " FROM         dbo.TblEmployee LEFT OUTER JOIN"
My_SQL = My_SQL & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
My_SQL = My_SQL & " ORDER BY dbo.TblEmployee.JobTypeID"

' My_SQL = "  select JobTypeID,JobTypeName from TblEmpJobsTypes  order by JobTypeName  "
    fill_combo dcJobTypeName, My_SQL
'    With Me.Fg
'        Set .WallPaper = GrdBack.Picture
'        .AutoSize 0, .Cols - 1, False
'    End With
If Projects.TxtModFlg.Text = "N" Then
VSFlexGrid1.Rows = 1
Cmd(0).Enabled = True

ElseIf Projects.TxtModFlg.Text = "E" Then
Cmd(0).Enabled = True
VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
Else
Cmd(0).Enabled = False

End If
currentterms = Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("fullcode"))
    If SystemOptions.UserInterface = ArabicInterface Then
                    Frame10.Caption = " „ÊŸ›Ì «·⁄„·ÌÂ —ﬁ„ : " & currentterms
                Else
                    Frame10.Caption = "Employees For Process No: " & currentterms
                End If
            
   Xpid = val(Projects.txt_project_id.Text)
    rwOp = val(Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("id")))
    
    rwpand = val(Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("ProjectDes_ID")))

Retrive Xpid, rwpand, rwOp
If Projects.TxtModFlg.Text <> "R" Then
 
VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
Else
 

End If

 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
'

Private Sub ChangeLang()
 
 FrmEmpOper.Caption = "Employee Worked in Poroject"
 Label5.Caption = Me.Caption
 Frame10.Caption = Me.Caption
 Label3.Caption = "Select Job"
 Option4.Caption = "Exp."
 Option5.Caption = "Act."
 Label4.Caption = "Lab. Count"
 Label11.Caption = "Days. Count"
 Label10.Caption = "Allocation Date"
 Cmd(8).Caption = "Delete"
 Label27.Caption = "Labors Total Count"
  Label29.Caption = "Totals"
  Cmd(0).Caption = "save"
  Cmd(1).Caption = "Clear"
  Command2.Caption = "Add"
  
  
      With Me.VSFlexGrid1
     
      .TextMatrix(0, .ColIndex("LineNo")) = "I"
        .TextMatrix(0, .ColIndex("code")) = "Code"
        .TextMatrix(0, .ColIndex("name")) = "Name"
        .TextMatrix(0, .ColIndex("jobname")) = "jobnamey"
        .TextMatrix(0, .ColIndex("daysalary")) = "Day Salary"
        .TextMatrix(0, .ColIndex("count")) = "Count"
        .TextMatrix(0, .ColIndex("total")) = "total"
        
     
 
    End With
    

  '
End Sub









