VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmEvaluationReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6840
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   10404
   Icon            =   "FrmEvaluationReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10404
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
   Begin VB.CommandButton btnClear 
      BackColor       =   &H00E2E9E9&
      Caption         =   "ãÓÍ"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   5328
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   10395
      Begin VB.Frame Frame3 
         Height          =   4575
         Left            =   6000
         TabIndex        =   4
         Top             =   120
         Width           =   4332
         Begin VB.Image Image1 
            Height          =   3672
            Left            =   0
            Picture         =   "FrmEvaluationReport.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   4272
         End
         Begin VB.Label lblCompanyname 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ÇáÓÇÊÑíÉ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.6
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   1095
            Left            =   240
            TabIndex        =   5
            Top             =   3840
            Width           =   2895
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   3612
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   5868
         _cx             =   10350
         _cy             =   6371
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   9.6
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   12648447
         ForeColor       =   128
         FrontTabColor   =   14871017
         BackTabColor    =   8454143
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "ÊÞÑíÑ ÇáÊÞííã"
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3264
            Index           =   2
            Left            =   24
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   24
            Width           =   5808
            _cx             =   10245
            _cy             =   5757
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
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
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÇáÝÊÑå"
               Height          =   735
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   1920
               Width           =   5775
               Begin MSComCtl2.DTPicker DtpDateFrom 
                  Height          =   330
                  Left            =   2640
                  TabIndex        =   20
                  Top             =   270
                  Width           =   1575
                  _ExtentX        =   2773
                  _ExtentY        =   572
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   85327875
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker DtpDateTo 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   21
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2773
                  _ExtentY        =   572
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   85327875
                  CurrentDate     =   41640
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Åáì"
                  Height          =   195
                  Index           =   3
                  Left            =   1830
                  RightToLeft     =   -1  'True
                  TabIndex        =   23
                  Top             =   240
                  Width           =   480
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ãä"
                  Height          =   195
                  Index           =   4
                  Left            =   5010
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   240
                  Width           =   540
               End
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00E2E9E9&
               Height          =   2055
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   -120
               Width           =   5775
               Begin VB.TextBox oneEmp_Code 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   3468
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   1080
                  Width           =   840
               End
               Begin MSDataListLib.DataCombo DcbBranch 
                  Height          =   312
                  Left            =   120
                  TabIndex        =   14
                  Top             =   360
                  Width           =   4176
                  _ExtentX        =   7366
                  _ExtentY        =   508
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo dcS 
                  Height          =   312
                  Left            =   120
                  TabIndex        =   15
                  Top             =   720
                  Width           =   4176
                  _ExtentX        =   7366
                  _ExtentY        =   508
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo OneEmployee 
                  Height          =   288
                  Left            =   120
                  TabIndex        =   25
                  Top             =   1080
                  Width           =   3324
                  _ExtentX        =   5863
                  _ExtentY        =   508
                  _Version        =   393216
                  Style           =   2
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ãÚíÇÑ"
                  Height          =   288
                  Index           =   0
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   720
                  Width           =   1092
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ãæÙÝ"
                  Height          =   288
                  Index           =   38
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   1080
                  Width           =   1092
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ÝÑÚ"
                  Height          =   288
                  Index           =   37
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   360
                  Width           =   1092
               End
            End
         End
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "ÔÇÔÉ ÊÞÇÑíÑ ÇáÊÞííã "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1020
         Index           =   25
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   4080
         Width           =   5772
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1092
         Left            =   120
         Top             =   4080
         Width           =   5772
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   492
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   6240
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   868
      ButtonPositionImage=   1
      Caption         =   "ÎÑæÌ"
      BackColor       =   14871017
      FontSize        =   7.8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8700
      _ExtentY        =   508
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   492
      Index           =   0
      Left            =   3000
      TabIndex        =   10
      Top             =   6240
      Width           =   1128
      _ExtentX        =   1990
      _ExtentY        =   868
      ButtonPositionImage=   1
      Caption         =   "ÚÑÖ ÇáÊÞÑíÑ"
      BackColor       =   14871017
      FontSize        =   7.8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "ØÈÞÇ áãÓÊÃÌÑ ãÍÏÏ"
      Height          =   195
      Index           =   5
      Left            =   5400
      TabIndex        =   8
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÔÇÔÉ ÊÞÇÑíÑ ÇáÊÞííã "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10464
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmEvaluationReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecID As String
Dim ii As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim DCboSearch As clsDCboSearch

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

Label5.Caption = "Reports of Evaluations"
lbl(25).Caption = Label5.Caption
lbl(37).Caption = "Branch"
lbl(0).Caption = "Standered"
lbl(38).Caption = "Employee"
Frame1.Caption = "Duration"
lbl(3).Caption = "To"
lbl(4).Caption = "From"
btnClear.Caption = "Clear"
Cmd(0).Caption = "Show Report"
Cmd(2).Caption = "Exit"
lblCompanyname.Caption = "AL SATTARYAH"
C1Tab1.Caption = " Evaluation Report"
 
End Sub
Private Sub btnClear_Click()
clear_all Me
'DcbTypeMain.Enabled = False
'TxtSearchCode.Enabled = False
'DcbEmp.Enabled = False
'DcbDept.Enabled = False
DtpDateFrom.value = ""
DtpDateTo.value = ""
End Sub

Private Sub Chorder_Click()

End Sub

Private Sub CHReq_Click()

End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0

'GetData
print_ProfitDist
          
        Case 1
          
        Case 2
            Unload Me
            Case 3
'print_report
    End Select

End Sub




Private Sub DcbEmp_Change()
DcbEmp_Click (0)
End Sub

Private Sub DcbEmp_Click(Area As Integer)


End Sub

Private Sub dcEmp_Change()
Dim val1, val2, str As String
If OneEmployee.BoundText = "" Then Exit Sub
oneEmp_Code.text = ""

    str = " select * From TblEmployee where  Emp_ID = " & val(OneEmployee.BoundText)
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        val1 = IIf(IsNull(Rs_Temp("FullCode").value), "", Rs_Temp("FullCode").value)
     Else
        val1 = ""
    End If
    
    oneEmp_Code.text = val1
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub



Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)


End Sub
Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    Dim str As String

    
    
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches DcbBranch
    fill_combo dcS, " Select ID,EName From TblEvaluationStandered  "
    
    Dcombos.GetEmployees OneEmployee
    
    DtpDateFrom.value = ""
    DtpDateTo.value = ""

    Resize_Form Me
    If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
    End If
    
    DtpDateFrom.value = Now
    DtpDateTo.value = Now
    
    DtpDateFrom.value = Null
    DtpDateTo.value = Null

End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Private Function Selection_Query() As String

Dim str As String
Set Rs_Temp = New ADODB.Recordset

str = str & "     SELECT dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Fullcode, dbo.TblEvaluationStandered.EName, dbo.TblEvaluationStandered.ENameE, dbo.TblEmployee.Emp_Name,"
str = str & "     dbo.TblEmployee.Emp_Code, dbo.TblEvaluation_Details.ID, dbo.TblEvaluation_Details.HID, dbo.TblEvaluation_Details.StanderedID,"
str = str & "     dbo.TblEvaluation_Details.Emp_ID AS Expr1, dbo.TblEvaluation_Details.PreDegree, dbo.TblEvaluation_Details.MaxDgree, dbo.TblEvaluation_Details.Curr_Dynamic,"
str = str & "     dbo.TblEvaluation_Details.sum_Degrees, dbo.TblEvaluation_Details.Manual_Degree, dbo.TblEvaluation_Details.Final_Evaluation, dbo.TblEvaluation_Details.Remarks,"
str = str & "     dbo.TblEvaluation_Details.EvalTitle"
 str = str & "    FROM     dbo.TblEvaluation_Details INNER JOIN"
str = str & "     dbo.TblEmpEvaluation ON dbo.TblEvaluation_Details.HID = dbo.TblEmpEvaluation.ID LEFT OUTER JOIN"
str = str & "     dbo.TblEmployee ON dbo.TblEvaluation_Details.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
str = str & "     dbo.TblEvaluationStandered ON dbo.TblEvaluation_Details.StanderedID = dbo.TblEvaluationStandered.ID"

str = str & " where 1 = 1 "

If OneEmployee.BoundText <> "" Then
        str = str & " and   TblEvaluation_Details.Emp_ID =  " & val(OneEmployee.BoundText)
ElseIf DcbBranch.BoundText <> "" Then
        str = str & "  and  dbo.TblEmpEvaluation.Branch  =  " & val(DcbBranch.BoundText)

ElseIf dcS.BoundText <> "" Then
        str = str & " and dbo.TblEvaluation_Details.StanderedID =  " & val(dcS.BoundText)
End If


If Not IsNull(DtpDateFrom.value) Then
         str = str & "  and TblEvaluationStandered.SDate >= " & SQLDate(DtpDateFrom.value, True) & ""
End If

If Not IsNull(DtpDateTo.value) Then
         str = str & "  and TblEvaluationStandered.SDate <= " & SQLDate(DtpDateTo.value, True) & ""
End If

str = str & "  order by Emp_Name  "

Selection_Query = str
End Function


Public Sub GetData()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
 
    StrSQL = Selection_Query

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Msg = "áÇÊæÌÏ ÈíÇäÇÊ ááÚÑÖ ÊæÇÝÞ ÔÑæØ ÇáÊÞÑíÑ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
        rs.MoveFirst
        print_report StrSQL
    End If


End Sub
Function print_report(Optional NoteSerial As String)
     


End Function


Private Sub oneEmp_Code_Change()


Dim val1, val2, str As String
If oneEmp_Code.text = "" Then Exit Sub
'EmployeeID.BoundText = ""

    str = " select * From TblEmployee where  fullcode = '" & oneEmp_Code.text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        val1 = IIf(IsNull(Rs_Temp("Emp_ID").value), "", Rs_Temp("Emp_ID").value)
     Else
        val1 = ""
    End If
    OneEmployee.BoundText = val1

End Sub

Private Sub OneEmployee_Click(Area As Integer)
Dim val1, val2, str As String
If OneEmployee.BoundText = "" Then Exit Sub
oneEmp_Code.text = ""

    str = " select * From TblEmployee where  Emp_ID = " & val(OneEmployee.BoundText)
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        val1 = IIf(IsNull(Rs_Temp("FullCode").value), "", Rs_Temp("FullCode").value)
     Else
        val1 = ""
    End If
    
    oneEmp_Code.text = val1

End Sub


Function print_ProfitDist(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
Dim str As String

str = str & "     SELECT dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Fullcode, dbo.TblEvaluationStandered.EName, dbo.TblEvaluationStandered.ENameE, dbo.TblEmployee.Emp_Name,"
str = str & "     dbo.TblEmployee.Emp_Code, dbo.TblEvaluation_Details.ID, dbo.TblEvaluation_Details.HID, dbo.TblEvaluation_Details.StanderedID,"
str = str & "     dbo.TblEvaluation_Details.Emp_ID AS Expr1, dbo.TblEvaluation_Details.PreDegree, dbo.TblEvaluation_Details.MaxDgree, dbo.TblEvaluation_Details.Curr_Dynamic,"
str = str & "     dbo.TblEvaluation_Details.sum_Degrees, dbo.TblEvaluation_Details.Manual_Degree, dbo.TblEvaluation_Details.Final_Evaluation, dbo.TblEvaluation_Details.Remarks,"
str = str & "     dbo.TblEvaluation_Details.EvalTitle"
str = str & "    FROM     dbo.TblEvaluation_Details INNER JOIN"
str = str & "     dbo.TblEmpEvaluation ON dbo.TblEvaluation_Details.HID = dbo.TblEmpEvaluation.ID LEFT OUTER JOIN"
str = str & "     dbo.TblEmployee ON dbo.TblEvaluation_Details.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
str = str & "     dbo.TblEvaluationStandered ON dbo.TblEvaluation_Details.StanderedID = dbo.TblEvaluationStandered.ID"

str = str & " where 1 = 1 "

If OneEmployee.BoundText <> "" Then
        str = str & " and   TblEvaluation_Details.Emp_ID =  " & val(OneEmployee.BoundText)
ElseIf DcbBranch.BoundText <> "" Then
        str = str & "  and  dbo.TblEmpEvaluation.Branch  =  " & val(DcbBranch.BoundText)

ElseIf dcS.BoundText <> "" Then
        str = str & " and dbo.TblEvaluation_Details.StanderedID =  " & val(dcS.BoundText)
End If


If Not IsNull(DtpDateFrom.value) Then
         str = str & "  and TblEvaluationStandered.SDate >= " & SQLDate(DtpDateFrom.value, True) & ""
End If

If Not IsNull(DtpDateTo.value) Then
         str = str & "  and TblEvaluationStandered.SDate <= " & SQLDate(DtpDateTo.value, True) & ""
End If

str = str & "  order by Emp_Name  "






 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_EvaluationReport.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_EvaluationReport.rpt"
       End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open str, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "áÇÊæÌÏ ÈíÇäÇÊ ááÚÑÖ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
End Function


