VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAncestorEmployeeReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ﬁ—Ì— ”·› «·„ÊŸ›Ì‰"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10890
   Icon            =   "FrmAncestorEmployeeReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   10890
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   4695
      Left            =   6480
      TabIndex        =   22
      Top             =   720
      Width           =   4335
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
         Height          =   1095
         Left            =   480
         TabIndex        =   23
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   3675
         Index           =   1
         Left            =   120
         Picture         =   "FrmAncestorEmployeeReport.frx":6852
         Stretch         =   -1  'True
         Top             =   120
         Width           =   4395
      End
   End
   Begin VB.Frame lblLW 
      BackColor       =   &H00E2E9E9&
      Height          =   4695
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   720
      Width           =   6465
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”·›  Õ·Ì·Ï"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -30
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   240
         Width           =   3255
      End
      Begin VB.CheckBox ChkStatus 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈ŸÂ«— ﬂ· «·„ÊŸ›Ì‰ „⁄ «·„‰ ÂÌ… Œœ„« Â„"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Frame lbprocess 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1440
         Width           =   6195
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Height          =   735
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   0
            Width           =   3315
            Begin VB.OptionButton Opr 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "="
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   315
               Index           =   0
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton Opr 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "= >"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   315
               Index           =   4
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton Opr 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "= <"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   315
               Index           =   2
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton Opr 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   3
               Left            =   840
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton Opr 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   1
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.TextBox TxtIDFrom 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·”·›…"
            Height          =   195
            Index           =   14
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   975
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   6255
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo DcboEmpName 
            Height          =   315
            Left            =   90
            TabIndex        =   16
            Top             =   600
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   556
            _Version        =   393216
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„ÊŸ›"
            Height          =   285
            Index           =   4
            Left            =   4800
            TabIndex        =   18
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   285
            Index           =   7
            Left            =   4680
            TabIndex        =   17
            Top             =   600
            Width           =   1485
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Height          =   1065
         Left            =   0
         TabIndex        =   9
         Top             =   2130
         Width           =   6255
         Begin VB.OptionButton optTypeDate 
            Alignment       =   1  'Right Justify
            Caption         =   "» «—ÌŒ «·œ›⁄…"
            Height          =   375
            Index           =   1
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optTypeDate 
            Alignment       =   1  'Right Justify
            Caption         =   "» «—ÌŒ «·”·›…"
            Height          =   345
            Index           =   0
            Left            =   3780
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker DtpDateFrom 
            Height          =   330
            Left            =   3120
            TabIndex        =   10
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   143654915
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DtpDateTo 
            Height          =   330
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   143654915
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰  «—ÌŒ"
            Height          =   195
            Index           =   9
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   0
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   720
            Width           =   1080
         End
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1395
         Left            =   120
         Top             =   3270
         Width           =   6255
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "›Ì Õ«·… ⁄œ„  ÕœÌœ «Ì Õﬁ· ”Ê› ÌﬂÊ‰ «· ﬁ—Ì— ≈Ã„«·Ì"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1290
         Index           =   25
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   3330
         Width           =   6255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   5400
      Width           =   10815
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   5
         Top             =   240
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   661
         Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
         ButtonImage     =   "FrmAncestorEmployeeReport.frx":8DAA
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
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   661
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
         ButtonImage     =   "FrmAncestorEmployeeReport.frx":F60C
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
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
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
         ButtonImage     =   "FrmAncestorEmployeeReport.frx":15E6E
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
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   10905
      Begin VB.Image Image1 
         Height          =   615
         Index           =   0
         Left            =   8880
         Picture         =   "FrmAncestorEmployeeReport.frx":3FA90
         Stretch         =   -1  'True
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ﬁ—Ì— ”·› «·„ÊŸ›Ì‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   120
         Width           =   3360
      End
   End
   Begin VB.ComboBox DcbMoth 
      Height          =   315
      ItemData        =   "FrmAncestorEmployeeReport.frx":4103E
      Left            =   15360
      List            =   "FrmAncestorEmployeeReport.frx":41040
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Dcbyear 
      Height          =   315
      ItemData        =   "FrmAncestorEmployeeReport.frx":41042
      Left            =   15360
      List            =   "FrmAncestorEmployeeReport.frx":41044
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "FrmAncestorEmployeeReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public Sub GetData()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
StrSQL = " SELECT     dbo.TblEmpAdvance.AdvanceDate, dbo.TblEmpAdvance.AdvanceID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, "
StrSQL = StrSQL & "                       dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblEmpAdvance.Emp_ID, dbo.TblEmpAdvance.AdvanceValue,"
StrSQL = StrSQL & "                      dbo.TblEmpAdvance.PaymentCounts, dbo.TblEmpAdvance.FirstMonthPayment, dbo.TblEmpAdvance.FirstYearPayment,"
StrSQL = StrSQL & "                      dbo.GetReminingAdvancValue(dbo.TblEmpAdvance.AdvanceID, dbo.TblEmpAdvance.Emp_ID) AS ReminValue ,"
StrSQL = StrSQL & " dbo.GetReminingAdvancQst(dbo.TblEmpAdvance.AdvanceID, dbo.TblEmpAdvance.Emp_ID) AS ReminQst"
StrSQL = StrSQL & " FROM         dbo.TblEmpAdvance LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblEmpAdvance.Emp_ID = dbo.TblEmployee.Emp_ID"
  StrWhere = " Where 1 = 1 "
   '///////////////////
   
          If (Me.DcboEmpName.Text <> "") And (val(DcboEmpName.BoundText) <> 0) Then
             StrWhere = StrWhere & " AND dbo.TblEmpAdvance.Emp_ID =" & Me.DcboEmpName.BoundText & ""
          End If

     If Not IsNull(Me.DtpDateFrom.value) Then
        StrWhere = StrWhere & " AND  dbo.TblEmpAdvance.AdvanceDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
        If Not IsNull(Me.DtpDateTo.value) Then
         StrWhere = StrWhere & " AND  dbo.TblEmpAdvance.AdvanceDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If
If (TxtIDFrom.Text) <> "" Then
If Opr(0).value = True Then
StrWhere = StrWhere & " AND dbo.TblEmpAdvance.AdvanceValue =" & val(TxtIDFrom.Text) & ""
ElseIf Opr(1).value = True Then
StrWhere = StrWhere & " AND dbo.TblEmpAdvance.AdvanceValue > " & val(TxtIDFrom.Text) & ""
ElseIf Opr(2).value = True Then
StrWhere = StrWhere & " AND dbo.TblEmpAdvance.AdvanceValue >=" & val(TxtIDFrom.Text) & ""
ElseIf Opr(3).value = True Then
StrWhere = StrWhere & " AND dbo.TblEmpAdvance.AdvanceValue < " & val(TxtIDFrom.Text) & ""
ElseIf Opr(4).value = True Then
StrWhere = StrWhere & " AND dbo.TblEmpAdvance.AdvanceValue <=" & val(TxtIDFrom.Text) & ""
End If
End If

    StrSQL = StrSQL & StrWhere

 StrSQL = StrSQL & "  ORDER BY  dbo.TblEmpAdvance.AdvanceID"
 

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"sdfsf
        End If
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_report StrSQL

'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

           
 

    End If

End Sub

Public Sub GetDataDetails()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
StrSQL = " SELECT     dbo.TblEmpAdvance.AdvanceID, dbo.TblEmpAdvance.Emp_ID, dbo.TblEmpAdvance.AdvanceValue, dbo.TblEmpAdvance.AdvanceDate,"
StrSQL = StrSQL & "                       dbo.TblEmpAdvanceDetails.PartNO, dbo.TblEmpAdvanceDetails.PartValue, dbo.TblEmpAdvanceDetails.Payed, dbo.TblEmpAdvanceDetails.Payed1,"
StrSQL = StrSQL & "                       dbo.TblEmpAdvanceDetails.Remark, dbo.TblEmpAdvance.PaymentCounts, dbo.TblEmpAdvance.AdvanceType, dbo.TblEmployee.Emp_Code,"
StrSQL = StrSQL & "                       dbo.TblEmployee.emp_name , dbo.TblEmployee.Emp_Namee, dbo.TblEmpAdvanceDetails.PartDate"
StrSQL = StrSQL & " FROM         dbo.TblEmpAdvance INNER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmpAdvanceDetails ON dbo.TblEmpAdvance.AdvanceID = dbo.TblEmpAdvanceDetails.AdvanceID INNER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee ON dbo.TblEmpAdvance.Emp_ID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & " Where 1 = 1  "

   '///////////////////
   
          If (Me.DcboEmpName.Text <> "") And (val(DcboEmpName.BoundText) <> 0) Then
             StrWhere = StrWhere & " AND dbo.TblEmpAdvance.Emp_ID =" & Me.DcboEmpName.BoundText & ""
          End If

     If Not IsNull(Me.DtpDateFrom.value) Then
        If optTypeDate(0) Then
            StrWhere = StrWhere & " AND  dbo.TblEmpAdvanceDetails.PartDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            StrWhere = StrWhere & " AND  dbo.TblEmpAdvance.AdvanceDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
      End If
        If Not IsNull(Me.DtpDateTo.value) Then
            If optTypeDate(0) Then
                StrWhere = StrWhere & " AND  dbo.TblEmpAdvanceDetails.PartDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
            Else
                StrWhere = StrWhere & " AND  dbo.TblEmpAdvance.AdvanceDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
            End If
      End If
If (TxtIDFrom.Text) <> "" Then
If Opr(0).value = True Then
StrWhere = StrWhere & " AND dbo.TblEmpAdvance.AdvanceValue =" & val(TxtIDFrom.Text) & ""
ElseIf Opr(1).value = True Then
StrWhere = StrWhere & " AND dbo.TblEmpAdvance.AdvanceValue > " & val(TxtIDFrom.Text) & ""
ElseIf Opr(2).value = True Then
StrWhere = StrWhere & " AND dbo.TblEmpAdvance.AdvanceValue >=" & val(TxtIDFrom.Text) & ""
ElseIf Opr(3).value = True Then
StrWhere = StrWhere & " AND dbo.TblEmpAdvance.AdvanceValue < " & val(TxtIDFrom.Text) & ""
ElseIf Opr(4).value = True Then
StrWhere = StrWhere & " AND dbo.TblEmpAdvance.AdvanceValue <=" & val(TxtIDFrom.Text) & ""
End If
End If

    StrSQL = StrSQL & StrWhere

 StrSQL = StrSQL & "  ORDER BY  dbo.TblEmpAdvance.AdvanceID"
 

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"sdfsf
        End If
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_report StrSQL

'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

           
 

    End If

End Sub
'Public Sub GetData()
'    Dim StrSQL As String
'      Dim StrWhere As String
'    Dim BolBegine As Boolean
'    Dim rs As ADODB.Recordset
'    Dim Msg As String
'    Dim i As Integer
'StrSQL = "SELECT     dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.Note_Value, dbo.Notes.CashingType, dbo.Notes.salary_or_advance, "
'StrSQL = StrSQL & "                      dbo.Notes.EmpAccountCode, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
'StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1,"
''StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee2 , dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4 , dbo.TblEmployee.jopstatusid"
'StrSQL = StrSQL & " FROM         dbo.Notes LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Notes.EmpAccountCode = dbo.TblEmployee.Account_code"
'StrSQL = StrSQL & " Where (dbo.Notes.NoteType = 5) And (dbo.Notes.CashingType <= 5) And (dbo.Notes.salary_or_advance = 1)"
'
'   If ChkStatus.value = vbUnchecked Then
'' StrSQL = StrSQL & " and dbo.TblEmployee.jopstatusid <> 2"
 'StrSQL = StrSQL & " and dbo.TblEmployee.jopstatusid <> 5"
 'StrSQL = StrSQL & " and dbo.TblEmployee.jopstatusid <> 6"
 'End If
'
'  StrWhere = ""
'   '///////////////////
'
'          If (Me.DcboEmpName.Text <> "") And (val(DcboEmpName.BoundText) <> 0) Then
'             StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_ID =" & Me.DcboEmpName.BoundText & ""
'          End If
'
'     If Not IsNull(Me.DtpDateFrom.value) Then
'        StrWhere = StrWhere & " AND dbo.Notes.NoteDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
''      End If
 '       If Not IsNull(Me.DtpDateTo.value) Then
 '        StrWhere = StrWhere & " AND dbo.Notes.NoteDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
 '     End If
'If (TxtIDFrom.Text) <> "" Then
'If Opr(0).value = True Then
'StrWhere = StrWhere & " AND dbo.Notes.Note_Value =" & val(TxtIDFrom.Text) & ""
'ElseIf Opr(1).value = True Then
'StrWhere = StrWhere & " AND dbo.Notes.Note_Value > " & val(TxtIDFrom.Text) & ""
'ElseIf Opr(2).value = True Then
'StrWhere = StrWhere & " AND dbo.Notes.Note_Value >=" & val(TxtIDFrom.Text) & ""
'ElseIf Opr(3).value = True Then
'StrWhere = StrWhere & " AND dbo.Notes.Note_Value < " & val(TxtIDFrom.Text) & ""
'ElseIf Opr(4).value = True Then
'StrWhere = StrWhere & " AND dbo.Notes.Note_Value <=" & val(TxtIDFrom.Text) & ""
'End If
'End If
'
'    StrSQL = StrSQL & StrWhere
'
' StrSQL = StrSQL & "  ORDER BY  dbo.Notes.NoteID"
'
'
'    Set rs = New ADODB.Recordset
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.BOF Or rs.EOF Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
'        ElseIf SystemOptions.UserInterface = EnglishInterface Then
'          '  Me.lbl(10).Caption = "Search Results=0"
'        End If
'
'        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        Exit Sub
'    Else
' rs.MoveFirst
' print_report StrSQL
'
'
'            If SystemOptions.UserInterface = ArabicInterface Then
'             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
'            ElseIf SystemOptions.UserInterface = EnglishInterface Then
'               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
'            End If
'
'
'
'
'    End If
'
'End Sub
Function print_report(Optional NoteSerial As String)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   
        If Check1.value = False Then
             If SystemOptions.UserInterface = ArabicInterface Then
             
                 StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAdvanceEmployeeReports.rpt"
                 Else
                 StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAdvanceEmployeeReportsE.rpt"
                 
            End If
        Else
                If SystemOptions.UserInterface = ArabicInterface Then
             
                    StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAdvanceEmployeeReports2.rpt"
                 Else
                    StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAdvanceEmployeeReportsE.rpt"
                 
            End If
        End If
  


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "Not Found Data to Show"
        End If
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
       'If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then xReport.ParameterFields(3).AddCurrentValue Format(Me.XPDtbFrom.value, "yyyy/M/d")
       ' If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(4).AddCurrentValue Format(Me.XPDtpTo.value, "yyyy/M/d")
       ' If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then xReport.ParameterFields(5).AddCurrentValue Me.Txt_DateHigriFrom.value
       ' If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(6).AddCurrentValue Me.Txt_DateHigriTO.value
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
  
    End If

   
  If DtpDateFrom.value <> "" And DtpDateTo.value <> "" Then
   xReport.ParameterFields(8).AddCurrentValue DtpDateFrom.value

    xReport.ParameterFields(10).AddCurrentValue DtpDateTo.value
  '  xReport.ParameterFields(11).AddCurrentValue DtpDateToH.value
    End If

  Dim Total As String
  Dim totl As Double


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
Private Sub ChangeLang()
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Show"
    Cmd(2).Caption = "Exit"
     Me.Caption = "advances Employee Report"
     Me.Label1(2).Caption = Me.Caption
    Me.lbl(4).Caption = "Employee Code"
    Me.lbl(7).Caption = "Employee Name"
    Me.lbl(14).Caption = "Value"
    Me.lbl(9).Caption = "From Date"
    Me.lbl(0).Caption = "To"
    lblCompanyname.Caption = "El-Sattaryh"
   ChkStatus.Caption = "All Employees With End Service"
   lbl(25).Caption = "If you do not select any field report will be Total"
  End Sub

Private Sub Cmd_Click(Index As Integer)
    Select Case Index

        Case 0
        If Check1.value = vbChecked Then
            GetDataDetails
        Else
            GetData
        End If
          
        Case 1
            clear_all Me

DtpDateFrom.value = ""
DtpDateTo.value = ""

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
            Case 3
End Select
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
  If val(DcboEmpName.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
End Sub

Private Sub Form_Activate()
 PutFormOnTop Me.hwnd
End Sub

Private Sub Form_Load()
Dim Dcombos As ClsDataCombos
Set Dcombos = New ClsDataCombos
Dcombos.GetEmployees Me.DcboEmpName
DtpDateFrom.value = ""
DtpDateTo.value = ""
    Resize_Form Me
    If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtIDFrom.Text, 1)
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub
