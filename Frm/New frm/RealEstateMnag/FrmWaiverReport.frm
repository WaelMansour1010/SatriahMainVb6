VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmWaiverReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
   Icon            =   "FrmWaiverReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnClear 
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ﬁ«—Ì— «· ’›Ì« "
      Height          =   4485
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   10395
      Begin VB.CheckBox chkIsLegalAffairs 
         Alignment       =   1  'Right Justify
         Caption         =   "«ŸÂ«— «· ’›Ì«  «·„Õ«·… ··‘∆Ê‰ «·ﬁ«‰Ê‰Ì… ›ﬁÿ"
         Height          =   255
         Left            =   1500
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   180
         Width           =   3705
      End
      Begin VB.Frame Frame1 
         Caption         =   " ÕœÌœ —ﬁ„ «· ’›ÌÂ"
         Height          =   585
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   2640
         Width           =   3060
         Begin VB.TextBox TxtTo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox TXtFrom 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   39
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "„‰"
            Height          =   255
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "«·Ï"
            Height          =   255
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CheckBox chknotllpayed 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·„  ”œœ »«·ﬂ«„·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CheckBox chknotllCollected 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·„  Õ’· »«·ﬂ«„·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CheckBox chkoutflow 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Â—Ê»"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CheckBox chkoutCondition 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘—Êÿ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtUnitNo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   30
         Top             =   2280
         Width           =   4935
      End
      Begin VB.TextBox txtContNo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   4935
      End
      Begin VB.Frame XPPnlTime 
         BackColor       =   &H00E2E9E9&
         Caption         =   "›Ï «·› —…"
         Height          =   1185
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   3240
         Width           =   2415
         Begin MSComCtl2.DTPicker XPDtbFrom 
            Height          =   345
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   183894017
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker XPDtpTo 
            Height          =   345
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   183894017
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   24
            Top             =   360
            Width           =   465
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   23
            Top             =   720
            Width           =   465
         End
      End
      Begin VB.Frame FrameDateH 
         Caption         =   " ÕœÌœ «· «—ÌŒ «·ÂÃ—Ì"
         Height          =   1185
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   3240
         Width           =   2220
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigriFrom 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigriTO 
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "«·Ï"
            Height          =   255
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "„‰"
            Height          =   255
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3855
         Left            =   6960
         TabIndex        =   12
         Top             =   120
         Width           =   3375
         Begin VB.Image Image1 
            Height          =   2190
            Left            =   240
            Picture         =   "FrmWaiverReport.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2940
         End
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
            Left            =   240
            TabIndex        =   13
            Top             =   2400
            Width           =   2895
         End
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcsupplier 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcbAqarType 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCAkarUnit 
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·—ﬁ„ «·ÊÕœ…"
         Height          =   195
         Index           =   7
         Left            =   5400
         TabIndex        =   31
         Top             =   2280
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·—ﬁ„ «·⁄ﬁœ"
         Height          =   195
         Index           =   6
         Left            =   5400
         TabIndex        =   28
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·„” √Ã— „Õœœ"
         Height          =   195
         Index           =   3
         Left            =   5400
         TabIndex        =   25
         Top             =   840
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·‰Ê⁄ «·ÊÕœ…"
         Height          =   195
         Index           =   9
         Left            =   5400
         TabIndex        =   10
         Top             =   1920
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ«·⁄ﬁ«— „⁄Ì‰"
         Height          =   195
         Index           =   1
         Left            =   5400
         TabIndex        =   9
         Top             =   1560
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·›—⁄ „⁄Ì‰"
         Height          =   195
         Index           =   0
         Left            =   5400
         TabIndex        =   5
         Top             =   480
         Width           =   1020
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   5280
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   873
      ButtonPositionImage=   1
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
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   5280
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   873
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   240
      TabIndex        =   26
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   240
      Picture         =   "FrmWaiverReport.frx":10A48
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ»ﬁ« ·„” √Ã— „Õœœ"
      Height          =   195
      Index           =   5
      Left            =   5400
      TabIndex        =   27
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ«—Ì— «· ’›Ì« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   -105
      TabIndex        =   6
      Top             =   0
      Width           =   10440
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmWaiverReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim ii As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Private Sub btnClear_Click()
clear_all Me
XPDtpTo.value = ""
    xpdtbfrom.value = ""
End Sub

Private Sub Cmd_Click(index As Integer)

    Select Case index

        Case 0
       

 GetData
            
        Case 1
            clear_all Me
'DtpDateFrom.value = ""
'DtpDateTo.value = ""
'Me.DtStart.value = ""
'Me.DtEnd.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
            Case 3
'print_report
    End Select

End Sub






Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    

   XPDtpTo.value = ""
    xpdtbfrom.value = ""
    Set Dcombos = New ClsDataCombos
    
    Dcombos.GetIqar dcbAqarType
    Dcombos.GetCustomersSuppliers 56, Me.dcsupplier
    Dcombos.getAkarUnit Me.DCAkarUnit
    Dcombos.GetBranches DcbBranch
    Set cSearch = New clsDCboSearch
    My_SQL = "TblContract"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Resize_Form Me
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
    'gr = 9
    'Order = 9

'StrSQL = "SELECT     dbo.TblFiterWaiver.FilterDateH, dbo.TblFiterWaiver.FilterDate, dbo.TblAqar.Aqarid, dbo.TblCustemers.CusName, dbo.TblAkarUnit.name, dbo.TblAqar.aqarname,  dbo.TblAqar.aqarNo, dbo.TblFiterWaiver.BranchID, dbo.TblAqarDetai.unitno, dbo.TblBranchesData.branch_name, dbo.TblContract.ContNo, dbo.TblCustemers.CusID, dbo.TblAkarUnit.id FROM         dbo.TblContract INNER JOIN  dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid RIGHT OUTER JOIN dbo.TblFiterWaiver LEFT OUTER JOIN  dbo.TblAqarDetai ON dbo.TblFiterWaiver.ApartmentID = dbo.TblAqarDetai.Id LEFT OUTER JOIN  dbo.TblBranchesData ON dbo.TblFiterWaiver.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN   dbo.TblAkarUnit ON dbo.TblFiterWaiver.unittype = dbo.TblAkarUnit.id LEFT OUTER JOIN  dbo.TblCustemers ON dbo.TblFiterWaiver.RenterID = dbo.TblCustemers.CusID ON dbo.TblAqar.Aqarid = dbo.TblFiterWaiver.BulidID"
StrSQL = " SELECT     ISNULL(dbo.Getfitterdata(dbo.TblFiterWaiver.ID, 5, '02-Mar-2015', '02-Mar-2015'), 0) AS totalpayed, ISNULL(dbo.Getfitterdata(dbo.TblFiterWaiver.ID, 4, "
 StrSQL = StrSQL & "                     '02-Mar-2015', '02-Mar-2015'), 0) AS totalcollected, ISNULL(dbo.TblFiterWaiver.net, 0) AS net, dbo.TblFiterWaiver.ID, dbo.TblFiterWaiver.FilterDateH,"
 StrSQL = StrSQL & "                     dbo.TblFiterWaiver.FilterDate, dbo.TblAqar.Aqarid, dbo.TblCustemers.CusName, dbo.TblAkarUnit.name, dbo.TblAqar.aqarname, dbo.TblAqar.aqarNo,"
 StrSQL = StrSQL & "                     dbo.TblFiterWaiver.BranchID, dbo.TblAqarDetai.unitno, dbo.TblBranchesData.branch_name, dbo.TblCustemers.CusID, dbo.TblFiterWaiver.ContNo,"
StrSQL = StrSQL & "                      dbo.TblFiterWaiver.ContractNo,TblFiterWaiver.IsLegalAffairs,TblFiterWaiver.LegalAffairs"
StrSQL = StrSQL & " FROM         dbo.TblFiterWaiver LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqarDetai ON dbo.TblFiterWaiver.ApartmentID = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblFiterWaiver.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAkarUnit ON dbo.TblFiterWaiver.unittype = dbo.TblAkarUnit.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblFiterWaiver.RenterID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblFiterWaiver.BulidID = dbo.TblAqar.Aqarid"
 



StrSQL = StrSQL & " Where  (1 = 1)"
    BolBegine = False
    StrWhere = ""
        If val(Me.txtFrom.text) <> 0 Then
        StrSQL = StrSQL & " and  dbo.TblFiterWaiver.ID >=" & val(Me.txtFrom.text) & ""
     
    End If
If chkIsLegalAffairs.value = vbChecked Then
        StrSQL = StrSQL & " and  dbo.TblFiterWaiver.IsLegalAffairs = 1 "
End If
    If val(Me.txtto.text) <> 0 Then
    
       StrSQL = StrSQL & " AND dbo.TblFiterWaiver.ID <=" & val(Me.txtto.text) & ""
      End If
 If chknotllCollected.value = vbChecked Then
 StrSQL = StrSQL & "AND   isnull(dbo.TblFiterWaiver.net,0)  >0  and   isnull(dbo.TblFiterWaiver.net,0) >isnull( dbo.Getfitterdata(dbo.TblFiterWaiver.ID,4,'" & SQLDate(xpdtbfrom.value) & "' ,'" & SQLDate(XPDtpTo.value) & "'),0) "
 
 End If
 
    
 If chknotllpayed.value = vbChecked Then
 StrSQL = StrSQL & " AND  isnull(dbo.TblFiterWaiver.net,0)  <0  and   abs(isnull(dbo.TblFiterWaiver.net,0))  >isnull( dbo.Getfitterdata(dbo.TblFiterWaiver.ID,5,'" & SQLDate(xpdtbfrom.value) & "' ,'" & SQLDate(XPDtpTo.value) & "'),0) "
 
 End If
     
     


        
If Me.DcbBranch.BoundText <> "" Then
StrWhere = StrWhere & " AND tblFiterWaiver.BranchID = " & val(Me.DcbBranch.BoundText)
'gr = 0
End If


If Me.dcbAqarType.BoundText <> "" Then
'gr = 1
StrWhere = StrWhere & " AND dbo.TblAqar.Aqarid = " & val(Me.dcbAqarType.BoundText)
'gr = 1
End If


If Me.dcsupplier.BoundText <> "" Then
StrWhere = StrWhere & " AND tblCustemers.CusID = " & val(dcsupplier.BoundText)
'gr = 2
End If



If Me.DCAkarUnit.BoundText <> "" Then
StrWhere = StrWhere & " AND  TblAkarUnit.id = " & val(DCAkarUnit.BoundText)
'gr = 2
End If

If Me.TxtContNo.text <> "" Then
StrWhere = StrWhere & " AND  dbo.TblFiterWaiver.ContractNo ='" & (TxtContNo.text) & "'"
'gr = 2
End If

If Me.txtUnitNo.text <> "" Then
StrWhere = StrWhere & " AND  TblAqarDetai.unitno = " & val(txtUnitNo.text)
'gr = 2
End If

 If Me.xpdtbfrom <> Empty Or Me.xpdtbfrom <> Null Then
        StrWhere = StrWhere + " and (TblFiterWaiver.FilterDate >=" & SQLDate(Me.xpdtbfrom.value, True) & ")"
    End If

    If Me.XPDtpTo <> Empty Or Me.XPDtpTo <> Null Then
        StrWhere = StrWhere + " and (TblFiterWaiver.FilterDate <=" & SQLDate(XPDtpTo.value, True) & ")"
    End If


    '-----------------------------------

If chkoutflow.value = vbChecked Then
 StrWhere = StrWhere + " and outflow=1 "

End If

If chkoutCondition.value = vbChecked Then
 StrWhere = StrWhere + " and outCondition=1 "

End If

    StrSQL = StrSQL & StrWhere
 
  
  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
   '  Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
 rs.MoveFirst
' MsgBox rs("Granty").value
 print_report StrSQL
'print_report StrSQL
       ' With Me.Fg
       '     .Clear flexClearScrollable, flexClearEverything
       '     .Rows = .FixedRows
       '     .Rows = rs.RecordCount + .FixedRows
'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

           
 

    End If

End Sub
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



        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Waiver.rpt"
            
       End If
           
        
            
    ' If Me.RdDept.value = True Then
           ' StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byDept.rpt"
     '       Else
      '      If Me.RdSuper.value = True Then
       '     StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1bySuper.rpt"
        '    Else
         '   If Me.RdFitter.value = True Then
           ' StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byFitter.rpt"
          ' Else
             
            '        If Me.RdAll2.value = True Then
         '   StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1all.rpt"
          '  Else
           '  StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1.rpt"
            
     '      End If
      '      End If
       '     End If
        '     End If
         '   End If
          '  End If
        '    End If
           ' End If
          '  End If
       '      End If
           '
      '  End If



    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        If xpdtbfrom.value <> Null Or xpdtbfrom.value <> "" Then xReport.ParameterFields(3).AddCurrentValue Format(Me.xpdtbfrom.value, "yyyy/M/d")
        If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(4).AddCurrentValue Format(Me.XPDtpTo.value, "yyyy/M/d")
        If xpdtbfrom.value <> Null Or xpdtbfrom.value <> "" Then xReport.ParameterFields(5).AddCurrentValue Me.Txt_DateHigriFrom.value
        If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(6).AddCurrentValue Me.Txt_DateHigriTO.value
    
    If chknotllCollected.value = vbChecked Then xReport.ParameterFields(7).AddCurrentValue True
     If chknotllpayed.value = vbChecked Then xReport.ParameterFields(8).AddCurrentValue True
    
  
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
        xReport.ParameterFields(3).AddCurrentValue Format(Me.xpdtbfrom.value, "yyyy/M/d")
        xReport.ParameterFields(4).AddCurrentValue Format(Me.XPDtpTo.value, "yyyy/M/d")
       xReport.ParameterFields(5).AddCurrentValue Me.Txt_DateHigriFrom.value
        xReport.ParameterFields(6).AddCurrentValue Me.Txt_DateHigriTO.value
    End If

   ' xReport.ParameterFields(3).AddCurrentValue user_name
   ' xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
        'xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
'Dim gr, order As Integer
' xReport.ParameterFields(14).AddCurrentValue Order
 'xReport.ParameterFields(15).AddCurrentValue gr
 ' xReport.ParameterFields(15).AddCurrentValue gr
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
  Dim total As String
  Dim totl As Double
 ' totl = val(LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
 ' total = totl
 '  xReport.ParameterFields(12).AddCurrentValue Me.lbTotalMente.Caption
 '     xReport.ParameterFields(13).AddCurrentValue LbToTalExtra.Caption
 '       xReport.ParameterFields(14).AddCurrentValue total
   ' xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , NoteSerial

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function


Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub Txt_DateHigriFrom_LostFocus()
 VBA.Calendar = vbCalGreg
            xpdtbfrom.value = ToGregorianDate(Txt_DateHigriFrom.value)
End Sub

Private Sub Txt_DateHigriTO_LostFocus()
 VBA.Calendar = vbCalGreg
            XPDtpTo.value = ToGregorianDate(Txt_DateHigriTO.value)
End Sub






'Public Function GetBranchIDFromCode(Optional brancHcode As String, _
'Optional ByRef Emp_id As Integer) ' As Integer
'
'    Dim sql As String
'    Dim rs As New ADODB.Recordset
'    Dim Balance As Double
'    Dim id As Integer
'
'
'
'    sql = "select * from TblBranchesData where branch_code= '" & brancHcode & "'"
'
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        id = IIf(IsNull(rs("branch_Id").value), 0, rs("branch_Id").value)
'    Else
'        id = 0
'    End If
'
'    rs.Close
'    Emp_id = id
'    'GetBranchIDFromCode = id

'End Function





Private Sub TXtFrom_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.txtFrom.text, 1)
End Sub

Private Sub txtto_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.txtto.text, 1)
End Sub

Private Sub xpdtbfrom_Change()
If Not IsNull(xpdtbfrom.value) Then
 Txt_DateHigriFrom.value = ToHijriDate(xpdtbfrom.value)
 End If
End Sub

Private Sub XPDtpTo_Change()
If Not IsNull(XPDtpTo.value) Then
 Txt_DateHigriTO.value = ToHijriDate(XPDtpTo.value)
 End If
End Sub
