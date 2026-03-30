VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmApproveInvoice 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4905
   ClientLeft      =   1410
   ClientTop       =   2970
   ClientWidth     =   6450
   Icon            =   "FrmApproveInvoice.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   6450
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   4905
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   6435
      _cx             =   11351
      _cy             =   8652
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
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   960
         Top             =   2520
      End
      Begin VB.TextBox TxtRemarks 
         Alignment       =   1  'Right Justify
         Height          =   795
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   3600
         Width           =   6225
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   2820
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   480
         Width           =   6465
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÇáÚãíá ÊÎØì ÍÏ ÇáÇÆÊãÇä"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   2535
            Index           =   0
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   120
            Width           =   6120
         End
      End
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   0
         Width           =   7545
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÇáÚãíá ÊÎØì ÍÏ ÇáÇÆÊãÇä"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   2
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   0
            Width           =   4800
         End
      End
      Begin ALLButtonS.ALLButton cmdAdd 
         Height          =   345
         Left            =   3360
         TabIndex        =   16
         Tag             =   "Delete Row"
         Top             =   4440
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "ÍÝÙ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmApproveInvoice.frx":6852
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton1 
         Height          =   345
         Left            =   4920
         TabIndex        =   17
         Tag             =   "Delete Row"
         Top             =   4440
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "ÇÑÓÇá ááÇÚÊãÇÏ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmApproveInvoice.frx":686E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton2 
         Height          =   345
         Left            =   1800
         TabIndex        =   18
         Tag             =   "Delete Row"
         Top             =   4440
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "ÊÍÏíË ÇáÈíÇäÇÊ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmApproveInvoice.frx":688A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton3 
         Height          =   345
         Left            =   120
         TabIndex        =   19
         Tag             =   "Delete Row"
         Top             =   4440
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "ÎÑæÌ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmApproveInvoice.frx":68A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   " ãáÇÍÙÇÊ"
         Height          =   285
         Index           =   6
         Left            =   1800
         TabIndex        =   21
         Top             =   3360
         Width           =   2550
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmApproveInvoice.frx":68C2
      Left            =   18360
      List            =   "FrmApproveInvoice.frx":68D2
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   18600
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   18720
      TabIndex        =   5
      Tag             =   "ãä ÝÖáß ÃÏÎá ÑÞã ÇáÞÖíÉ"
      Top             =   960
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   -2147483624
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   18360
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   18480
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveInvoice.frx":68EB
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveInvoice.frx":6C85
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveInvoice.frx":701F
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveInvoice.frx":73B9
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveInvoice.frx":7753
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveInvoice.frx":7AED
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveInvoice.frx":7E87
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmApproveInvoice.frx":8421
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   18480
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "ÊÍÏíË ÞÇÚÏÉ ÇáÈíÇäÇÊ"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÊÍÏíË"
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
      ButtonImage     =   "FrmApproveInvoice.frx":87BB
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "ØÈÇÚÉ ÇáÈíÇäÇÊ "
      Top             =   120
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ØÈÇÚÉ "
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
      ButtonImage     =   "FrmApproveInvoice.frx":F01D
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   19800
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ááÈÍË ÅÖÛØ åÐÇ ÇáãÝÊÇÍ Ãæ ÅÖÛØ "
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÈÍË"
      BackColor       =   14871017
      FontSize        =   9.75
      FontName        =   "Arial"
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmApproveInvoice.frx":1587F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "ÇáãÓÊÎÏã"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   18360
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmApproveInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub ALLButton1_Click()
save
ALLButton1.Enabled = False
End Sub

Private Sub ALLButton2_Click()
If CheckApprove() = True Then
CmdAdd.Enabled = True
End If
End Sub
Private Sub ChangeLang()
Label1(2).Caption = "Credit Limit Of Customer"
lbl(6).Caption = "Remarks"
CmdAdd.Caption = "Save"
ALLButton2.Caption = "Update"
ALLButton3.Caption = "Exit"
ALLButton1.Caption = "Send To Approve"
End Sub
Private Sub ALLButton3_Click()
Unload Me
End Sub
Function CheckApprove() As Boolean
Dim sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset

sql = " SELECT     ISNULL(FlagApproved, 0)"
sql = sql & " From dbo.TblAproveInvoice"
sql = sql & " Where (EmpID = " & val(frmsalebill.DcboEmp.BoundText) & " ) And (CusID = " & val(frmsalebill.DBCboClientName.BoundText) & ") And (IsNull(FlagApproved, 0) = 1)and TransDate=" & SQLDate(frmsalebill.XPDtbBill.value, True) & ""
sql = sql & " and BillValue=" & val(frmsalebill.LblTotal.Caption) + val(frmsalebill.lblInstComm.Caption) + val(frmsalebill.TxtValueAdded.Text) & ""
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
CheckApprove = True
Else
CheckApprove = False
End If
End Function

Private Sub cmdAdd_Click()
frmsalebill.FlgAproved = 1
Unload Me
End Sub

Private Sub Form_Load()
CmdAdd.Enabled = False
ALLButton1.Enabled = True
frmsalebill.FlgAproved = 0
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
End Sub
Sub save()
Dim i As Integer
    Dim sql As String
    Dim Rs3 As ADODB.Recordset
    Dim SngCreditLiimt As Single
    Dim SngCreditLimitCredit As Single
    Dim SngCusAccount As Single
    Dim Msg As String
    Dim StrTemp As String
    Dim IntRes As Integer
    Dim DepitInterval As Integer
    Dim DepitIntervalID As Integer
    Dim Rs4 As ADODB.Recordset
    sql = "Select Account_Code,CreditLimit,CreditLimitCredit,DepitInterval,DepitIntervalID From TblCustemers Where CusID=" & val(frmsalebill.DBCboClientName.BoundText) & ""
    Set Rs3 = New ADODB.Recordset
    Rs3.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Rs3.RecordCount > 0 Then
        SngCreditLiimt = IIf(IsNull(Rs3("CreditLimit").value), 0, Rs3("CreditLimit").value)
        SngCreditLimitCredit = IIf(IsNull(Rs3("CreditLimitCredit").value), 0, Rs3("CreditLimitCredit").value)
        DepitInterval = IIf(IsNull(Rs3("DepitInterval").value), 0, Rs3("DepitInterval").value)
        DepitIntervalID = IIf(IsNull(Rs3("DepitIntervalID").value), 0, Rs3("DepitIntervalID").value)
    End If
If DepitIntervalID = 1 Then
DepitInterval = DepitInterval * 30
ElseIf DepitIntervalID = 2 Then
DepitInterval = DepitInterval * 365
End If


    Dim Account_Code As String
     Dim FirstPeriod As Date
   getFirstPeriodDateInthisYear FirstPeriod
        
  Account_Code = GetMyAccountCode("TblCustemers", "CusID", val(frmsalebill.DBCboClientName.BoundText))  '
  SngCusAccount = GetActualAccountBalance(Account_Code, 0, FirstPeriod, Date)
  SngCusAccount = SngCusAccount - GetSumOfGeForOneAccount(Account_Code, Transaction_ID, 0)

        
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset

Set Rs4 = New ADODB.Recordset
sql = " SELECT     UserID"
sql = sql & " From dbo.TblUsers"
sql = sql & " Where (AllowAprovedSalesBill = 1)"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
Rs4.MoveFirst
sql = "Select * from TblAproveInvoice where 1=-1"
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
For i = 1 To Rs4.RecordCount
Rs2.AddNew
Rs2("CusID").value = val(frmsalebill.DBCboClientName.BoundText)
Rs2("EmpID").value = val(frmsalebill.DcboEmp.BoundText)
Rs2("IssueDate").value = frmsalebill.DtpDelayDate.value
Rs2("TransDate").value = frmsalebill.XPDtbBill.value
Rs2("BillValue").value = val(frmsalebill.LblTotal.Caption) + val(frmsalebill.lblInstComm.Caption) + val(frmsalebill.TxtValueAdded.Text)
Rs2("NoDay").value = DateDiff("d", frmsalebill.DtpDelayDate.value, frmsalebill.XPDtbBill.value)
Rs2("SkipNoDay").value = DepitInterval
Rs2("Value").value = SngCusAccount
Rs2("Remarks").value = Me.TxtRemarks.Text
Rs2("SkipValue").value = SngCreditLiimt
Rs2("UserID").value = IIf(IsNull(Rs4("UserID").value), 0, Rs4("UserID").value)
Rs2.update
Rs4.MoveNext
Next i
End If
End Sub

Private Sub Timer1_Timer()
If CmdAdd.Enabled = True Then Exit Sub
ALLButton2_Click
End Sub
