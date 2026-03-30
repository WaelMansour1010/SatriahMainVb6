VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmAqarReport1 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   Icon            =   "FrmAqarReport1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnClear 
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   2520
      TabIndex        =   33
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   4605
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   10395
      Begin VB.Frame FrameDateH 
         Caption         =   " ÕœÌœ «· «—ÌŒ «·ÂÃ—Ì"
         Height          =   1185
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   3120
         Width           =   2220
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigriFrom 
            Height          =   315
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigriTO 
            Height          =   315
            Left            =   120
            TabIndex        =   48
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "„‰"
            Height          =   255
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "«·Ï"
            Height          =   255
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   720
            Width           =   495
         End
      End
      Begin VB.Frame XPPnlTime 
         BackColor       =   &H00E2E9E9&
         Caption         =   "›Ï «·› —…"
         Height          =   1185
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   3120
         Width           =   2415
         Begin MSComCtl2.DTPicker XPDtbFrom 
            Height          =   345
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   108003329
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker XPDtpTo 
            Height          =   345
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   108003329
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   45
            Top             =   720
            Width           =   465
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   44
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.OptionButton opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… „œÌ—Ì «·«›—⁄"
         Height          =   255
         Index           =   3
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   2400
         Width           =   1815
      End
      Begin VB.OptionButton opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «· Õ’Ì·"
         Height          =   255
         Index           =   2
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   3600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «· ”ÊÌﬁ  Ê «· Õ’Ì·"
         Height          =   255
         Index           =   1
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «·≈œ«—…"
         Height          =   255
         Index           =   0
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   2040
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.ComboBox cbopaymentType 
         Height          =   315
         ItemData        =   "FrmAqarReport1.frx":038A
         Left            =   3480
         List            =   "FrmAqarReport1.frx":0397
         TabIndex        =   36
         Text            =   "«·ﬂ·"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtStreet 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   14160
         TabIndex        =   34
         Top             =   3720
         Width           =   4935
      End
      Begin VB.Frame Frame3 
         Height          =   4455
         Left            =   6960
         TabIndex        =   31
         Top             =   120
         Width           =   3375
         Begin VB.Image Image1 
            Height          =   2310
            Left            =   120
            Picture         =   "FrmAqarReport1.frx":03AF
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3300
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
            Height          =   5295
            Left            =   120
            TabIndex        =   32
            Top             =   2520
            Width           =   2895
         End
      End
      Begin VB.TextBox txtCodeBranch 
         Height          =   285
         Left            =   6360
         TabIndex        =   30
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtCodeCustomer 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   16440
         TabIndex        =   29
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtCodeOwner 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   16440
         TabIndex        =   28
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtCodeSalesRep 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         TabIndex        =   27
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5400
         TabIndex        =   26
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtRoomCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   14160
         TabIndex        =   25
         Top             =   4080
         Width           =   4935
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
         Left            =   12360
         TabIndex        =   7
         Top             =   1200
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcbAqarType 
         Height          =   315
         Left            =   12360
         TabIndex        =   8
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dbcClient 
         Height          =   315
         Left            =   12360
         TabIndex        =   11
         Top             =   1560
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcbSalesSpec 
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dbcAqarStatus 
         Height          =   315
         Left            =   14160
         TabIndex        =   13
         Top             =   2280
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
         Left            =   14160
         TabIndex        =   22
         Top             =   3360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcbCityId2 
         Height          =   315
         Left            =   14160
         TabIndex        =   23
         Top             =   3000
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcmCityID 
         Height          =   315
         Left            =   14160
         TabIndex        =   24
         Top             =   2640
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
         Caption         =   "‰Ê⁄ «·⁄„·Ì…"
         Height          =   195
         Index           =   5
         Left            =   5850
         TabIndex        =   35
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·⁄œœ «·€—›"
         Height          =   195
         Index           =   11
         Left            =   13560
         TabIndex        =   21
         Top             =   4080
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·‘«—⁄ „⁄Ì‰"
         Height          =   195
         Index           =   10
         Left            =   11880
         TabIndex        =   20
         Top             =   3720
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·‰Ê⁄ «·ÊÕœ…"
         Height          =   195
         Index           =   9
         Left            =   11880
         TabIndex        =   19
         Top             =   3360
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·„œÌ‰… „⁄Ì‰…"
         Height          =   195
         Index           =   8
         Left            =   11880
         TabIndex        =   18
         Top             =   3000
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·ÕÏ „⁄Ì‰"
         Height          =   195
         Index           =   7
         Left            =   13560
         TabIndex        =   17
         Top             =   2640
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·Õ«·… «·⁄ﬁ«—"
         Height          =   195
         Index           =   6
         Left            =   13560
         TabIndex        =   16
         Top             =   2280
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·„‰œÊ» „Õœœ"
         Height          =   195
         Index           =   4
         Left            =   5400
         TabIndex        =   15
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·„” √Ã— „Õœœ"
         Height          =   195
         Index           =   3
         Left            =   14520
         TabIndex        =   14
         Top             =   1560
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·„«·ﬂ „Õœœ"
         Height          =   195
         Index           =   2
         Left            =   14520
         TabIndex        =   10
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ«·⁄ﬁ«— „⁄Ì‰"
         Height          =   195
         Index           =   1
         Left            =   14520
         TabIndex        =   9
         Top             =   840
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
      Left            =   1200
      TabIndex        =   0
      Top             =   5400
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
      Left            =   0
      TabIndex        =   1
      Top             =   5400
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
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   2520
      Picture         =   "FrmAqarReport1.frx":10A6D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ«—Ì— «·⁄„Ê·«‰"
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
      Left            =   -30
      TabIndex        =   6
      Top             =   0
      Width           =   10380
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
Attribute VB_Name = "FrmAqarReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecID As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Private Sub Check2_Click()

End Sub

Private Sub btnClear_Click()
clear_all Me
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

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


Private Sub ChangeLang()
 



  '
End Sub

Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    
    Set Dcombos = New ClsDataCombos
    
    Dcombos.GetIqar dcbAqarType
    
    Dcombos.GetCountriesGovernCities dcmCityID
    
    Dcombos.getCountriesGovernments dcbCityId2
    
    Dcombos.GetCustomersSuppliers 2, Me.dcsupplier
    
    Dcombos.getAkarUnit Me.DCAkarUnit
    
    Dcombos.GetSalesRepData Me.dcbSalesSpec
    
    Dcombos.GetCustomersSuppliers 1, Me.dbcClient
    
    Dcombos.GetBranches DcbBranch
    
   Dcombos.GetRentStatus dbcAqarStatus
    
    
    
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

'StrSQL = "SELECT     dbo.TblAqarDetai.unitdesc, dbo.TblAqarDetai.unitno, dbo.TblAqarDetai.ACCount, dbo.TblAqarDetai.Status, dbo.TblAqarDetai.RentValue, dbo.TblAqar.ownerid,  TblCustemers_1.CusID, TblCustemers_1.CusName AS ownername, TblCustemers_1.CusNamee AS ownernamee, dbo.TblAqarDetai.customerid,     TblCustemers_1.CusName, TblCustemers_1.CusNamee, dbo.TblAqar.Aqarid, dbo.TblAqar.aqarname, dbo.branches.branch_id, TblCustemers_2.CusName AS Expr1,    dbo.TblAqar.CountryID, dbo.TblAqar.cityid, dbo.TblAqar.streetname, dbo.TblAqar.BranchId, dbo.TblAkarUnit.name AS UnitName, dbo.TblAkarUnit.namee AS UnitNamee,  dbo.TblRentStatus.name AS RentStatusName, dbo.TblRentStatus.namee AS RentStatusNamee  FROM  dbo.TblAqarDetai INNER JOIN  dbo.TblAqar ON dbo.TblAqarDetai.Aqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN dbo.TblRentStatus ON dbo.TblAqarDetai.Status = dbo.TblRentStatus.id LEFT OUTER JOIN  dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id RIGHT OUTER JOIN "
'StrSQL = StrSQL + " dbo.branches ON dbo.TblAqar.BranchId = dbo.branches.branch_id LEFT OUTER JOIN  dbo.TblCustemers TblCustemers_1 ON dbo.TblAqarDetai.customerid = TblCustemers_1.CusID LEFT OUTER JOIN  dbo.TblCustemers TblCustemers_2 ON dbo.TblAqar.ownerid = TblCustemers_2.CusID"

 If opt(0).value = True Then

StrSQL = " SELECT     SUM(dbo.Notes.Note_Value) AS branchRevenue, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
StrSQL = StrSQL + "  FROM         dbo.Notes LEFT OUTER JOIN"
StrSQL = StrSQL + "  dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL + "  WHERE     (dbo.Notes.NoteType = 4) AND (dbo.Notes.CashingType = 8)"

 
            If Me.DcbBranch.BoundText <> "" Then
            StrWhere = StrWhere & " AND Notes.branch_no = " & val(Me.DcbBranch.BoundText)
            
            End If
            
            
            
                If Me.XPDtbFrom <> Empty Or Me.XPDtbFrom <> Null Then
        StrWhere = StrWhere + " and (NoteDate >=" & SQLDate(Me.XPDtbFrom.value, True) & ""
    End If

    If Me.XPDtpTo <> Empty Or Me.XPDtpTo <> Null Then
        StrWhere = StrWhere + " and NoteDate <=" & SQLDate(XPDtpTo.value, True) & ")"
    End If
    

    StrSQL = StrSQL & StrWhere
StrSQL = StrSQL + "  GROUP BY dbo.Notes.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"


ElseIf opt(1).value = True Then

StrSQL = " SELECT     dbo.ContracttBillInstallmentsDone.RecordDate, dbo.ContracttBillInstallmentsDone.RecordDateH, dbo.ContracttBillInstallmentsDone.CommissionsPayed, "
StrSQL = StrSQL + "  dbo.ContracttBillInstallmentsDone.RentValuePayed, dbo.ContracttBillInstallmentsDone.paymentType, dbo.ContracttBillInstallmentsDone.istallid,"
StrSQL = StrSQL + " dbo.TblContractInstallments.InstallNo, dbo.TblContractInstallments.Installdate, dbo.TblContractInstallments.InstalldateH, dbo.TblContract.NoteSerial1,"
StrSQL = StrSQL + " dbo.TblAqarDetai.unitno, dbo.ContracttBillInstallmentsDone.WaterPayed, dbo.ContracttBillInstallmentsDone.ElectricPayed,"
StrSQL = StrSQL + " dbo.ContracttBillInstallmentsDone.TelandNetPayed, dbo.Notes.empid, TblEmployee_1.Emp_Code, TblEmployee_1.Emp_Name,"
StrSQL = StrSQL + " TblEmployee_1.Emp_Namee, dbo.TblAqar.aqarname, dbo.TblAkarUnit.name AS unittypename, dbo.TblAkarUnit.namee AS unittypenamee, dbo.TblContract.CusID,"
StrSQL = StrSQL + " dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblContract.OutContract, dbo.ContracttBillInstallmentsDone.NoteID,"
StrSQL = StrSQL + " dbo.Notes.NoteSerial1 AS CashingSerialNo, dbo.Notes.branch_no , dbo.ContracttBillInstallmentsDone.CommisionValue"
StrSQL = StrSQL + " FROM         dbo.TblAqarDetai LEFT OUTER JOIN"
StrSQL = StrSQL + " dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id RIGHT OUTER JOIN"
StrSQL = StrSQL + "  dbo.TblContract INNER JOIN"
StrSQL = StrSQL + " dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo LEFT OUTER JOIN"
StrSQL = StrSQL + " dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL + " dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid ON dbo.TblAqarDetai.Id = dbo.TblContract.UnitNo RIGHT OUTER JOIN"
StrSQL = StrSQL + " dbo.Notes RIGHT OUTER JOIN"
StrSQL = StrSQL + " dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
StrSQL = StrSQL + " dbo.ContracttBillInstallmentsDone ON TblEmployee_1.Emp_ID = dbo.ContracttBillInstallmentsDone.empid ON"
StrSQL = StrSQL + " dbo.Notes.NoteID = dbo.ContracttBillInstallmentsDone.NoteID ON dbo.TblContractInstallments.id = dbo.ContracttBillInstallmentsDone.istallid"
 

StrSQL = StrSQL + "  where 1=1 "
'Where (dbo.ContracttBillInstallmentsDone.PaymentType = 1)

 
            If Me.DcbBranch.BoundText <> "" Then
            StrWhere = StrWhere & " AND Notes.branch_no = " & val(Me.DcbBranch.BoundText)
            
            End If
            
            
            
                If Me.XPDtbFrom <> Empty Or Me.XPDtbFrom <> Null Then
        StrWhere = StrWhere + " and (NoteDate >=" & SQLDate(Me.XPDtbFrom.value, True) & ""
    End If

    If Me.XPDtpTo <> Empty Or Me.XPDtpTo <> Null Then
        StrWhere = StrWhere + " and NoteDate <=" & SQLDate(XPDtpTo.value, True) & ")"
    End If
    

           If cbopaymentType.ListIndex <> -1 And cbopaymentType.ListIndex <> 0 Then
            StrWhere = StrWhere & " AND ContracttBillInstallmentsDone.paymentType= " & val(cbopaymentType.ListIndex)
            
            End If
            
                If Me.dcbSalesSpec.BoundText <> "" Then
            StrWhere = StrWhere & " AND   Notes.empid = " & val(Me.dcbSalesSpec.BoundText)
            
            End If
            
    StrSQL = StrSQL & StrWhere
 

ElseIf opt(3).value = True Then

StrSQL = " SELECT     dbo.ContracttBillInstallmentsDone.RecordDate, dbo.ContracttBillInstallmentsDone.RecordDateH, dbo.ContracttBillInstallmentsDone.CommissionsPayed, "
StrSQL = StrSQL + "  dbo.ContracttBillInstallmentsDone.RentValuePayed, dbo.ContracttBillInstallmentsDone.paymentType, dbo.ContracttBillInstallmentsDone.istallid,"
StrSQL = StrSQL + " dbo.TblContractInstallments.InstallNo, dbo.TblContractInstallments.Installdate, dbo.TblContractInstallments.InstalldateH, dbo.TblContract.NoteSerial1,"
StrSQL = StrSQL + " dbo.TblAqarDetai.unitno, dbo.ContracttBillInstallmentsDone.WaterPayed, dbo.ContracttBillInstallmentsDone.ElectricPayed,"
StrSQL = StrSQL + " dbo.ContracttBillInstallmentsDone.TelandNetPayed, dbo.ContracttBillInstallmentsDone.empid, TblEmployee_1.Emp_Code, TblEmployee_1.Emp_Name,"
StrSQL = StrSQL + " TblEmployee_1.Emp_Namee, dbo.TblAqar.aqarname, dbo.TblAkarUnit.name AS unittypename, dbo.TblAkarUnit.namee AS unittypenamee, dbo.TblContract.CusID,"
StrSQL = StrSQL + " dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblContract.OutContract, dbo.ContracttBillInstallmentsDone.NoteID,"
StrSQL = StrSQL + " dbo.Notes.NoteSerial1 AS CashingSerialNo, dbo.Notes.branch_no , dbo.ContracttBillInstallmentsDone.CommisionValue"
StrSQL = StrSQL + " FROM         dbo.TblAqarDetai LEFT OUTER JOIN"
StrSQL = StrSQL + " dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id RIGHT OUTER JOIN"
StrSQL = StrSQL + "  dbo.TblContract INNER JOIN"
StrSQL = StrSQL + " dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo LEFT OUTER JOIN"
StrSQL = StrSQL + " dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL + " dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid ON dbo.TblAqarDetai.Id = dbo.TblContract.UnitNo RIGHT OUTER JOIN"
StrSQL = StrSQL + " dbo.Notes RIGHT OUTER JOIN"
StrSQL = StrSQL + " dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
StrSQL = StrSQL + " dbo.ContracttBillInstallmentsDone ON TblEmployee_1.Emp_ID = dbo.ContracttBillInstallmentsDone.empid ON"
StrSQL = StrSQL + " dbo.Notes.NoteID = dbo.ContracttBillInstallmentsDone.NoteID ON dbo.TblContractInstallments.id = dbo.ContracttBillInstallmentsDone.istallid"
 

StrSQL = StrSQL + "  where 1=1 "
'Where (dbo.ContracttBillInstallmentsDone.PaymentType = 1)

 
            If Me.DcbBranch.BoundText <> "" Then
            StrWhere = StrWhere & " AND Notes.branch_no = " & val(Me.DcbBranch.BoundText)
            
            End If
            
            
            
                If Me.XPDtbFrom <> Empty Or Me.XPDtbFrom <> Null Then
        StrWhere = StrWhere + " and (NoteDate >=" & SQLDate(Me.XPDtbFrom.value, True) & ""
    End If

    If Me.XPDtpTo <> Empty Or Me.XPDtpTo <> Null Then
        StrWhere = StrWhere + " and NoteDate <=" & SQLDate(XPDtpTo.value, True) & ")"
    End If
    

      
            
         '       If Me.dcbSalesSpec.BoundText <> "" Then
         '   StrWhere = StrWhere & " AND   empid = " & val(Me.dcbSalesSpec.BoundText)
         '
     '       End If
            
    StrSQL = StrSQL & StrWhere
 


End If





   '-----------------------------------

 
  
  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
   '  Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
 rs.MoveFirst
' MsgBox rs("Granty").value
If opt(0).value = True Then
 print_report StrSQL, 0
 ElseIf opt(1).value = True Then
 print_report StrSQL, 1
ElseIf opt(3).value = True Then
 print_report StrSQL, 3
 End If
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
Function print_report(Optional StrSQL As String, Optional reportno As Integer)
     
  '  Set rs = New ADODB.Recordset
  '  rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

If reportno = 0 Then

        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Aqar1.rpt"
            
       End If
           
        
ElseIf reportno = 1 Then

        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_AqarCommissions.rpt"
            
       End If
       
ElseIf reportno = 3 Then

        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_AqarCommissionsMangements.rpt"
            
       End If
       
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
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
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
StrReportTitle = ""
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
If opt(1).value = True Then
            If cbopaymentType.ListIndex = 0 Or cbopaymentType.ListIndex = -1 Then
                    StrReportTitle = "⁄„Ê·«  «· ”ÊÌﬁ Ê«· Õ’Ì·" & Chr(13) '& StrAccountName
             ElseIf cbopaymentType.ListIndex = 1 Then
             StrReportTitle = "⁄„Ê·«  «· ”ÊÌﬁ  " & Chr(13) '& StrAccountName
              ElseIf cbopaymentType.ListIndex = 2 Then
             StrReportTitle = "⁄„Ê·«   «· Õ’Ì·" & Chr(13) '& StrAccountName
             End If
 End If
 
If opt(3).value = True Then
 StrReportTitle = "⁄„Ê·«   «·«œ«—…" & Chr(13)
End If

    If Me.XPDtbFrom.value <> Empty Or Me.XPDtbFrom.value <> Null Then
          StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.XPDtbFrom.value, "yyyy/M/d") & "  "
        End If
        If Me.XPDtpTo.value <> Empty Or Me.XPDtpTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.XPDtpTo.value, "yyyy/M/d") & " "
        End If
        
        xReport.ParameterFields(2).AddCurrentValue StrReportTitle
        
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
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function


Private Sub Text1_Change()

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.name, Me.Caption, Me.Caption
End Sub

Private Sub Txt_DateHigriFrom_LostFocus()
  VBA.Calendar = vbCalGreg
            XPDtbFrom.value = ToGregorianDate(Txt_DateHigriFrom.value)
               
End Sub

Private Sub Txt_DateHigriTO_LostFocus()
  VBA.Calendar = vbCalGreg
            XPDtpTo.value = ToGregorianDate(Txt_DateHigriTO.value)
           
End Sub

Private Sub txtCodeBranch_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
       GetBranchIDFromCode txtCodeBranch.text, EmpID
       DcbBranch.BoundText = EmpID
    End If
End Sub

Private Sub txtCodeCustomer_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode txtCodeCustomer.text, EmpID
        dbcClient.BoundText = EmpID
    End If
End Sub


Private Sub txtCodeOwner_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode txtCodeCustomer.text, EmpID
        dbcClient.BoundText = EmpID
    End If
End Sub


Private Sub txtCodeSalesRep_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode txtCodeSalesRep.text, EmpID
        dcbSalesSpec.BoundText = EmpID
    End If

End Sub


Private Sub txtRoomCount_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, txtRoomCount)
End Sub

Public Function GetBranchIDFromCode(Optional brancHcode As String, _
Optional ByRef Emp_id As Integer) ' As Integer
            
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim id As Integer
    

    
    sql = "select * from TblBranchesData where branch_code= '" & brancHcode & "'"
   
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        id = IIf(IsNull(rs("branch_Id").value), 0, rs("branch_Id").value)
    Else
        id = 0
    End If

    rs.Close
    Emp_id = id
    'GetBranchIDFromCode = id

End Function




 






Private Sub xpdtbfrom_Change()
 
     On Error Resume Next
         Txt_DateHigriFrom.value = ToHijriDate(XPDtbFrom.value)
       
 
End Sub

Private Sub XPDtpTo_Change()
On Error Resume Next
         Txt_DateHigriTO.value = ToHijriDate(XPDtpTo.value)

End Sub
