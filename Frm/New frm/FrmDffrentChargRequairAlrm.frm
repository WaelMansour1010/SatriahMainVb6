VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmDffrentChargRequairAlrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "‘«‘…  ‰»ÌÂ«  «·›—ﬁ »Ì‰ «·ﬂ„Ì«  «·„ÿ·Ê»Â Ê«·ﬂ„Ì«  «·„‘ÕÊ‰Â"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16380
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   16380
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
   Begin VB.Frame Frame4 
      Caption         =   "»ÕÀ »Õ”»"
      Height          =   735
      Left            =   10920
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1440
      Width           =   5415
      Begin VB.TextBox TxtItemCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3165
         TabIndex        =   25
         Top             =   240
         Width           =   1650
      End
      Begin MSDataListLib.DataCombo DcboItems 
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·’‰›"
         Height          =   435
         Index           =   3
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1800
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   16365
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9600
         TabIndex        =   29
         Top             =   600
         Width           =   810
      End
      Begin VB.Frame Frame3 
         Caption         =   "»ÕÀ »Õ”»"
         Height          =   735
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   120
         Width           =   9375
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Height          =   315
            Left            =   4680
            TabIndex        =   23
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Œ“‰"
            Height          =   435
            Index           =   1
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   315
            Index           =   2
            Left            =   8280
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "«·› —Â"
         Height          =   735
         Left            =   10920
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   5415
         Begin MSComCtl2.DTPicker Fromdate 
            Height          =   330
            Left            =   2760
            TabIndex        =   12
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   71041025
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker todate 
            Height          =   330
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   71041025
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "≈«·Ï"
            Height          =   435
            Index           =   14
            Left            =   1860
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   315
            Index           =   0
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   585
         End
      End
      Begin ImpulseButton.ISButton Cmd1 
         Height          =   585
         Index           =   20
         Left            =   9600
         TabIndex        =   16
         Top             =   1080
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1032
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "⁄—÷"
         BackColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmDffrentChargRequairAlrm.frx":0000
         ColorButton     =   8438015
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÕœÌÀ ﬂ·"
         Height          =   435
         Index           =   4
         Left            =   9600
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   780
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   855
         Left            =   0
         Top             =   840
         Width           =   9375
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Â–Â «·‘«‘…  ﬁÊ„ »«ŸÂ«— «·›—ﬁ »Ì‰ «·ﬂ„Ì«  «·„ÿ·Ê»Â Ê«·ﬂ„Ì«  «·„‘ÕÊ‰Â"
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
         Height          =   780
         Index           =   25
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   840
         Width           =   9375
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   16395
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   6360
         Top             =   240
      End
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   510
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Text            =   "modflag"
         Top             =   120
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Frame Frmo2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   2
            Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
            Top             =   15
            Width           =   2340
            _ExtentX        =   4128
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "«·„” Œœ„"
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
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   45
            Width           =   855
         End
      End
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   2760
         Top             =   0
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
               Picture         =   "FrmDffrentChargRequairAlrm.frx":039A
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDffrentChargRequairAlrm.frx":03F8
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDffrentChargRequairAlrm.frx":0456
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDffrentChargRequairAlrm.frx":04B4
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDffrentChargRequairAlrm.frx":0512
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDffrentChargRequairAlrm.frx":0570
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDffrentChargRequairAlrm.frx":05CE
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDffrentChargRequairAlrm.frx":062C
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‘«‘…  ‰»ÌÂ«  «·›—ﬁ »Ì‰ «·ﬂ„Ì«  «·„ÿ·Ê»Â Ê«·ﬂ„Ì«  «·„‘ÕÊ‰Â"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   2
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   10080
      End
   End
   Begin ImpulseButton.ISButton btnCancel 
      Height          =   330
      Left            =   120
      TabIndex        =   7
      Top             =   8040
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   582
      ButtonStyle     =   1
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
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton CmdPrint 
      Height          =   390
      Left            =   1080
      TabIndex        =   8
      Top             =   8040
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄…"
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
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
      Height          =   5580
      Left            =   0
      TabIndex        =   10
      Top             =   2400
      Width           =   16335
      _cx             =   28813
      _cy             =   9842
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
      Rows            =   12
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmDffrentChargRequairAlrm.frx":068A
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
   Begin ImpulseButton.ISButton Cmd1 
      Height          =   390
      Index           =   0
      Left            =   3120
      TabIndex        =   17
      Top             =   8040
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " ÕœÌÀ"
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
      ButtonImage     =   "FrmDffrentChargRequairAlrm.frx":084E
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   390
      Left            =   2040
      TabIndex        =   28
      Top             =   8040
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   688
      ButtonStyle     =   1
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
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmDffrentChargRequairAlrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

     Dim My_SQL As String
Private Sub BtnCancel_Click()
    Me.Hide
End Sub


Public Sub fillgrid(Optional Str As String)

  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
 My_SQL = " SELECT     dbo.TblItems.HaveSerial, dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_Date AS Transaction_DateH, dbo.Transactions.Transaction_Type, "
 My_SQL = My_SQL + "                      dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
 My_SQL = My_SQL + "                       dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Transactions.CashCustomerName, dbo.TblCustemers.Cus_Phone,"
 My_SQL = My_SQL + "                       dbo.TblCustemers.Cus_mobile, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
 My_SQL = My_SQL + "                       dbo.TblBranchesData.branch_id AS branch_idH, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
 My_SQL = My_SQL + "                      dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.*"
 My_SQL = My_SQL + "  FROM         dbo.TblEmployee RIGHT OUTER JOIN"
 My_SQL = My_SQL + "                      dbo.Transactions ON dbo.TblEmployee.Emp_ID = dbo.Transactions.Emp_ID LEFT OUTER JOIN"
 My_SQL = My_SQL + "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
 My_SQL = My_SQL + "                      dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 My_SQL = My_SQL + "                      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
 My_SQL = My_SQL + "                      dbo.TblItems RIGHT OUTER JOIN"
 My_SQL = My_SQL + "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID LEFT OUTER JOIN"
 My_SQL = My_SQL + "                     dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
 My_SQL = My_SQL + " where  dbo.Transactions.Transaction_Type=55"
 If Not (IsNull(Me.Fromdate.value)) Then
 My_SQL = My_SQL + " and (dbo.Transactions.Transaction_Date >='" & SQLDate(Fromdate.value) & "')"
 End If
 If Not (IsNull(Me.todate.value)) Then
 My_SQL = My_SQL + " and (dbo.Transactions.Transaction_Date <='" & SQLDate(todate.value) & "')"
 End If
If Me.DcboItems.text <> "" And val(Me.DcboItems.BoundText) <> 0 Then
My_SQL = My_SQL + "and Transaction_Details.Item_ID =" & val(Me.DcboItems.BoundText) & ""
End If
If Me.Dcbranch.text <> "" And val(Me.Dcbranch.BoundText) <> 0 Then
My_SQL = My_SQL + "and TblBranchesData.branch_id =" & val(Me.Dcbranch.BoundText) & ""
End If
If Me.DCboStoreName.text <> "" And val(Me.DCboStoreName.BoundText) <> 0 Then
My_SQL = My_SQL + "and Transactions.StoreID =" & val(Me.DCboStoreName.BoundText) & ""
End If
 

My_SQL = My_SQL + "   order by  dbo.Transactions.Transaction_Serial "


 
 
         


   
Dim ActualTotal As Double
'rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
              .TextMatrix(i, .ColIndex("Ser")) = i
              .TextMatrix(i, .ColIndex("Transaction_Serial")) = (IIf(IsNull(rs.Fields("Transaction_Serial").value), 0, rs.Fields("Transaction_Serial").value))
              .TextMatrix(i, .ColIndex("Transaction_ID")) = (IIf(IsNull(rs.Fields("Transaction_ID").value), "", rs.Fields("Transaction_ID").value))
              .TextMatrix(i, .ColIndex("Transaction_Date")) = (IIf(IsNull(rs.Fields("Transaction_DateH").value), "", rs.Fields("Transaction_DateH").value))
              .TextMatrix(i, .ColIndex("ShipedQty")) = (IIf(IsNull(rs.Fields("ShipedQty").value), 0, rs.Fields("ShipedQty").value))
              .TextMatrix(i, .ColIndex("showqty")) = IIf(IsNull(rs.Fields("showqty").value), 0, rs.Fields("showqty").value)
              .TextMatrix(i, .ColIndex("diff")) = val(.TextMatrix(i, .ColIndex("ShipedQty"))) - val(.TextMatrix(i, .ColIndex("showqty")))
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Cus")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
              .TextMatrix(i, .ColIndex("ItemName")) = (IIf(IsNull(rs.Fields("ItemName").value), "", rs.Fields("ItemName").value))
              .TextMatrix(i, .ColIndex("branch_name")) = (IIf(IsNull(rs.Fields("branch_name").value), "", rs.Fields("branch_name").value))
              .TextMatrix(i, .ColIndex("StoreName")) = (IIf(IsNull(rs.Fields("StoreName").value), "", rs.Fields("StoreName").value))
            Else
            .TextMatrix(i, .ColIndex("Cus")) = (IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value))
              .TextMatrix(i, .ColIndex("ItemName")) = (IIf(IsNull(rs.Fields("ItemNamee").value), "", rs.Fields("ItemNamee").value))
              .TextMatrix(i, .ColIndex("branch_name")) = (IIf(IsNull(rs.Fields("branch_name").value), "", rs.Fields("branch_namee").value))
              .TextMatrix(i, .ColIndex("StoreName")) = (IIf(IsNull(rs.Fields("StoreNamee").value), "", rs.Fields("StoreNamee").value))
            End If
            If rs.Fields("CusID").value = 2 Then
              .TextMatrix(i, .ColIndex("Cus")) = (IIf(IsNull(rs.Fields("CashCustomerName").value), "", rs.Fields("CashCustomerName").value))
              End If

        rs.MoveNext
            Next i
 
            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With

End Sub










Private Sub Cmd1_Click(Index As Integer)
fillgrid
End Sub

Private Sub CmdPrint_Click()
    print_report My_SQL
End Sub
Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDiffChargRequairedAlarm.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDiffChargRequairedAlarm.rpt"
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

    xReport.ParameterFields(3).AddCurrentValue user_name
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       If Not (IsNull(Me.Fromdate.value)) Then
        xReport.ParameterFields(6).AddCurrentValue Fromdate.value
       End If
      If Not (IsNull(Me.todate.value)) Then
        xReport.ParameterFields(7).AddCurrentValue todate.value
      End If
        ' xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
'   xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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


Private Sub DcboItems_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
       Reload
End If
End Sub
Sub Reload()
 Dim Dcombos As New ClsDataCombos
      Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsNames Me.DcboItems
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetBranches Me.Dcbranch

End Sub





Private Sub DCboStoreName_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
       Reload
End If
End Sub



Private Sub dcBranch_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
       Reload
End If
End Sub

Private Sub Form_Load()
 
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
   
      

    If SystemOptions.UserInterface = EnglishInterface Then

        SetInterface Me
        cahngelang
    End If
    Reload
Fromdate.value = Date
todate.value = Date
fillgrid
Dim Askcount As Integer
Askcount = GetSetting(StrAppRegPath, "Setting", "CountAlarmMinutes", 5)
        
    Timer1.interval = val(Askcount) * 1000
     
     
End Sub

Function cahngelang()
    Label1(2).Caption = "The Difference Between The Required Quantities and Charged"
    Me.Caption = Label1(2).Caption
    lbl(25).Caption = Label1(2).Caption
    lbl(0).Caption = "From"
    lbl(14).Caption = "To"
    Frame2.Caption = "Period"
    Frame3.Caption = "Search By"
    Frame4.Caption = "Search By"
    lbl(2).Caption = "Branch"
    lbl(1).Caption = "Store"
    lbl(3).Caption = "Item"
    Cmd1(20).Caption = "Add"
    Cmd1(0).Caption = "Update"
    CmdPrint.Caption = "Print"
    ISButton1.Caption = "Clear"
    btnCancel.Caption = "Exit"
    With GridInstallments
    .TextMatrix(0, .ColIndex("Ser")) = "Serial"
    .TextMatrix(0, .ColIndex("Transaction_Serial")) = "Tran_Serial"
    .TextMatrix(0, .ColIndex("Transaction_Date")) = "Date"
    .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
    .TextMatrix(0, .ColIndex("StoreName")) = "Store"
    .TextMatrix(0, .ColIndex("ItemName")) = "Item"
    .TextMatrix(0, .ColIndex("ShipedQty")) = "Qty Required "
    .TextMatrix(0, .ColIndex("showqty")) = "Qty Charged "
    .TextMatrix(0, .ColIndex("diff")) = "Difference"
   
    End With
    


End Function



Private Sub DcboItems_Click(Area As Integer)
    DcboItems_Change
End Sub

Private Sub DcboItems_Change()
       Me.TxtItemCode.text = GetItemCode(val(Me.DcboItems.BoundText))
  End Sub

Private Sub ISButton1_Click()
            clear_all Me
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
todate.value = Date
Fromdate.value = Date
End Sub

Private Sub Timer1_Timer()
Cmd1_Click (0)
End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        If TxtItemCode.text = "" Then
            Me.DcboItems.BoundText = ""
        Else
            Me.DcboItems.BoundText = GetItemID(Trim$(Me.TxtItemCode.text))
        End If
    End If

End Sub


 
