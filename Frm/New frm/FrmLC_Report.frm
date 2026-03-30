VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmLC_Report 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   6732
      Left            =   -120
      TabIndex        =   2
      Top             =   600
      Width           =   10152
      Begin VB.Frame Frame3 
         Height          =   5892
         Left            =   5760
         TabIndex        =   47
         Top             =   120
         Width           =   4332
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
            Height          =   1092
            Left            =   720
            TabIndex        =   48
            Top             =   5160
            Width           =   2892
         End
         Begin VB.Image Image1 
            Height          =   4932
            Left            =   120
            Picture         =   "FrmLC_Report.frx":0000
            Stretch         =   -1  'True
            Top             =   120
            Width           =   4140
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÿ»«⁄…"
         Height          =   492
         Left            =   5760
         TabIndex        =   44
         Top             =   6120
         Width           =   4332
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·«‰ Â«¡ "
         Height          =   732
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   5160
         Width           =   5412
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   348
            Left            =   2280
            TabIndex        =   40
            Top             =   240
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102891521
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   348
            Left            =   360
            TabIndex        =   41
            Top             =   240
            Width           =   1452
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102891521
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   288
            Index           =   6
            Left            =   1800
            TabIndex        =   43
            Top             =   240
            Width           =   468
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   288
            Index           =   5
            Left            =   3240
            TabIndex        =   42
            Top             =   240
            Width           =   1068
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «Œ— ‘Õ‰"
         Height          =   732
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   5880
         Width           =   5412
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   348
            Left            =   2280
            TabIndex        =   33
            Top             =   240
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102891521
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker DTPicker6 
            Height          =   348
            Left            =   360
            TabIndex        =   34
            Top             =   240
            Width           =   1452
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102891521
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   288
            Index           =   8
            Left            =   1800
            TabIndex        =   36
            Top             =   240
            Width           =   468
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   288
            Index           =   7
            Left            =   3240
            TabIndex        =   35
            Top             =   240
            Width           =   1068
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·«€·«ﬁ"
         Height          =   612
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   4560
         Width           =   5412
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   348
            Left            =   2280
            TabIndex        =   29
            Top             =   120
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102891521
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   348
            Left            =   360
            TabIndex        =   30
            Top             =   120
            Width           =   1452
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102891521
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   288
            Index           =   4
            Left            =   1800
            TabIndex        =   32
            Top             =   120
            Width           =   468
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   288
            Index           =   3
            Left            =   3240
            TabIndex        =   31
            Top             =   120
            Width           =   1068
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   3852
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   5412
         Begin VB.ComboBox cbValue 
            DataSource      =   "Adodc1"
            Height          =   288
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   1560
            Width           =   972
         End
         Begin VB.TextBox txtValue 
            Height          =   288
            Left            =   360
            TabIndex        =   45
            Top             =   1560
            Width           =   2652
         End
         Begin VB.TextBox TXtPrimaryInvoiceNo 
            Height          =   372
            Left            =   360
            TabIndex        =   37
            Top             =   3360
            Width           =   3612
         End
         Begin VB.TextBox txtName 
            Height          =   372
            Left            =   360
            TabIndex        =   23
            Top             =   1080
            Width           =   3612
         End
         Begin VB.TextBox TXTLCNO 
            Height          =   372
            Left            =   360
            TabIndex        =   21
            Top             =   240
            Width           =   3612
         End
         Begin MSDataListLib.DataCombo DCLC 
            Height          =   288
            Left            =   360
            TabIndex        =   12
            Top             =   720
            Width           =   3612
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCCUrrency 
            Height          =   288
            Left            =   360
            TabIndex        =   15
            Top             =   1920
            Width           =   3612
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCBank 
            Height          =   288
            Left            =   360
            TabIndex        =   17
            Top             =   2280
            Width           =   3612
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCCountry 
            Height          =   288
            Left            =   360
            TabIndex        =   19
            Top             =   2640
            Width           =   3612
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   288
            Left            =   360
            TabIndex        =   25
            Top             =   3000
            Width           =   3612
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·›« Ê—… «·„»œ∆Ì…"
            Height          =   252
            Left            =   3960
            TabIndex        =   38
            Top             =   3480
            Width           =   1212
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê—œ"
            Height          =   228
            Index           =   0
            Left            =   4284
            TabIndex        =   26
            Top             =   3000
            Width           =   888
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„"
            Height          =   252
            Left            =   4320
            TabIndex        =   24
            Top             =   1200
            Width           =   852
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„"
            Height          =   252
            Left            =   4320
            TabIndex        =   22
            Top             =   360
            Width           =   852
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·œÊ·…"
            Height          =   228
            Index           =   8
            Left            =   4284
            TabIndex        =   20
            Top             =   2640
            Width           =   888
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·»‰ﬂ"
            Height          =   228
            Index           =   27
            Left            =   4260
            TabIndex        =   18
            Top             =   2280
            Width           =   912
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄„·…"
            Height          =   228
            Index           =   59
            Left            =   4164
            TabIndex        =   16
            Top             =   1920
            Width           =   1008
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„…"
            Height          =   228
            Index           =   7
            Left            =   4284
            TabIndex        =   14
            Top             =   1560
            Width           =   888
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·‰Ê⁄"
            Height          =   228
            Index           =   52
            Left            =   4260
            TabIndex        =   13
            Top             =   720
            Width           =   912
         End
      End
      Begin VB.Frame XPPnlTime 
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·› Õ"
         Height          =   588
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   3960
         Width           =   5412
         Begin MSComCtl2.DTPicker XPDtbFrom 
            Height          =   348
            Left            =   2280
            TabIndex        =   5
            Top             =   120
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102891521
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker XPDtpTo 
            Height          =   348
            Left            =   360
            TabIndex        =   6
            Top             =   120
            Width           =   1452
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102891521
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   288
            Index           =   2
            Left            =   3240
            TabIndex        =   8
            Top             =   120
            Width           =   1068
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   288
            Index           =   0
            Left            =   1800
            TabIndex        =   7
            Top             =   120
            Width           =   468
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   492
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   1128
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
      TabIndex        =   9
      Top             =   2040
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
      Caption         =   "ÿ»ﬁ« ·„” √Ã— „Õœœ"
      Height          =   195
      Index           =   5
      Left            =   5400
      TabIndex        =   10
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "       «· ﬁ«—Ì— «·„ Œ’’… ··«⁄ „«œ«       "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
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
      Width           =   10092
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
Attribute VB_Name = "FrmLC_Report"
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
Dim DCboSearch As clsDCboSearch


Private Sub btnClear_Click()
clear_all Me
End Sub








Private Sub Command1_Click()

Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String



'MySQL = "  SELECT dbo.BanksData.BankName, dbo.currency.name AS CurrencyName, dbo.LCTypes.name AS TypeName, dbo.LCTypes.namee AS TypeNameE, dbo.TblLC.CountryId,"
'MySQL = MySQL & "                  dbo.TblLC.BankId, dbo.TblLC.LCTyperId, dbo.TblCountriesData.CountryName, dbo.TblLC.Value, dbo.TblLC.LCNO, dbo.TblLC.Todate, dbo.TblLC.Name, dbo.TblLC.FromDate,"
'MySQL = MySQL & "                          dbo.TblLC.CloseDate , dbo.TblLC.LastParcilDate, dbo.TblLC.VendorID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee"
'MySQL = MySQL & "        FROM     dbo.LCTypes RIGHT OUTER JOIN"
'    MySQL = MySQL & "                      dbo.currency RIGHT OUTER JOIN"
'   MySQL = MySQL & "                       dbo.TblCountriesData RIGHT OUTER JOIN"
'    MySQL = MySQL & "                      dbo.TblLC LEFT OUTER JOIN"
'      MySQL = MySQL & "                    dbo.TblCustemers ON dbo.TblLC.VendorId = dbo.TblCustemers.CusID ON dbo.TblCountriesData.CountryID = dbo.TblLC.CountryId ON"
'        MySQL = MySQL & "                  dbo.currency.id = dbo.TblLC.CurrencyId ON dbo.LCTypes.id = dbo.TblLC.LCTyperId LEFT OUTER JOIN"
''     MySQL = MySQL & "                     dbo.BanksData ON dbo.TblLC.BankId = dbo.BanksData.BankID"

MySQL = "  SELECT BanksData_1.BankName, dbo.currency.name AS CurrencyName, dbo.LCTypes.name AS TypeName, dbo.LCTypes.namee AS TypeNameE, dbo.TblLC.CountryId,"
MySQL = MySQL & "                         dbo.TblLC.BankId, dbo.TblLC.LCTyperId, dbo.TblCountriesData.CountryName, dbo.TblLC.Value, dbo.TblLC.LCNO, dbo.TblLC.Todate, dbo.TblLC.Name, dbo.TblLC.FromDate,"
 MySQL = MySQL & "                        dbo.TblLC.CloseDate, dbo.TblLC.LastParcilDate, dbo.TblLC.VendorId, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblLC.Namee, dbo.TblLC.OpenValue,"
MySQL = MySQL & "                         dbo.TblLC.Remarks, dbo.TblLC.NoOfParcil, dbo.TblLC.PaymentTypeID, dbo.TblLC.ChequeNumber, dbo.TblLC.ChequeDueDate, dbo.TblBoxesData.BoxName,"
MySQL = MySQL & "                         dbo.BanksData.BankName AS BankName2, dbo.currency.nameE AS CurrencyNameE, dbo.BanksData.BankNamee AS BankNameE2,"
MySQL = MySQL & "                         BanksData_1.BankNamee AS BankNameE, dbo.TblBoxesData.BoxNameE"
MySQL = MySQL & "       FROM     dbo.TblCountriesData RIGHT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblCustemers RIGHT OUTER JOIN"
 MySQL = MySQL & "                        dbo.TblBoxesData RIGHT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblLC LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.BanksData ON dbo.TblLC.BankID2 = dbo.BanksData.BankID ON dbo.TblBoxesData.BoxID = dbo.TblLC.BoxID ON dbo.TblCustemers.CusID = dbo.TblLC.VendorId ON"
MySQL = MySQL & "                         dbo.TblCountriesData.CountryID = dbo.TblLC.CountryId LEFT OUTER JOIN"
 MySQL = MySQL & "                        dbo.currency ON dbo.TblLC.CurrencyId = dbo.currency.id LEFT OUTER JOIN"
 MySQL = MySQL & "                        dbo.LCTypes ON dbo.TblLC.LCTyperId = dbo.LCTypes.id LEFT OUTER JOIN"
 MySQL = MySQL & "                        dbo.BanksData AS BanksData_1 ON dbo.TblLC.BankId = BanksData_1.BankID"



    MySQL = MySQL & "         where  1 =1   "
    
    
   If cbValue.text <> "" Then
             MySQL = MySQL & " and  value  " & cbValue.text & "  " & val(txtValue.text)
   End If
    
    If DCLC.BoundText <> "" Then
         MySQL = MySQL & " and  LCTyperID = " & val(DCLC.BoundText)
    End If

    If DCCUrrency.BoundText <> "" Then
        MySQL = MySQL & " and  CUrrencyID  = " & val(DCCUrrency.BoundText)
    End If

    If DCBank.BoundText <> "" Then
        MySQL = MySQL & " and  tbllc.BankID  = " & val(DCBank.BoundText)
    End If
    
    If DCCountry.BoundText <> "" Then
        MySQL = MySQL & " and  tbllc.CountryID = " & DCCountry.BoundText
    End If

    If DBCboClientName.BoundText <> "" Then
        MySQL = MySQL & " AND   tbllc.VendorID =  " & val(DBCboClientName.BoundText)
    End If
    
'    If DCBranch.BoundText <> "" Then
'        MySQL = MySQL & " AND   tbllc.BranchID =  " & val(DCBranch.BoundText)
'    End If
    
    
    
'//////////
    If Not IsNull(XPDtbFrom.value) Then
            MySQL = MySQL & " and  FromDate >=  " & SQLDate(Me.XPDtbFrom.value, True) & ""
    End If
    
    If Not IsNull(XPDtpTo.value) Then
            MySQL = MySQL & "  and   FromDate <=  " & SQLDate(Me.XPDtpTo.value, True) & ""
    End If
    
    
   '///////////////////
    If Not IsNull(DTPicker1.value) Then
            MySQL = MySQL & " and  ToDate >=  " & SQLDate(Me.DTPicker1.value, True) & ""
    End If
    
    If Not IsNull(DTPicker2.value) Then
            MySQL = MySQL & "  and   ToDate <=  " & SQLDate(Me.DTPicker2.value, True) & ""
    End If
    
    
       '///////////////////
    If Not IsNull(DTPicker3.value) Then
            MySQL = MySQL & " and  CloseDate >=  " & SQLDate(Me.DTPicker3.value, True) & ""
    End If
    
    If Not IsNull(DTPicker4.value) Then
            MySQL = MySQL & "  and   CloseDate <=  " & SQLDate(Me.DTPicker4.value, True) & ""
    End If
    
           '///////////////////
    If Not IsNull(DTPicker5.value) Then
            MySQL = MySQL & " and  LastParcilDate >=  " & SQLDate(Me.DTPicker5.value, True) & ""
    End If
    
    If Not IsNull(DTPicker6.value) Then
            MySQL = MySQL & "  and   CloseDate <=  " & SQLDate(Me.DTPicker6.value, True) & ""
    End If
    
    
    
    
    If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW" & "\rpt_LC.rpt"
    Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW" & "\rpt_LC_E.rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo





    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
       
        
       If Not IsNull(XPDtbFrom.value) Then
    xReport.ParameterFields(2).AddCurrentValue (SQLDate(Me.XPDtbFrom.value, True))
    End If
    
    If Not IsNull(XPDtpTo.value) Then
  xReport.ParameterFields(4).AddCurrentValue (SQLDate(Me.XPDtpTo.value, True))
    End If
    
     Dim ss As Integer
RsData.MoveLast
ss = RsData.RecordCount
    Dim dd As String
    dd = "" & ss & ""
  xReport.ParameterFields(5).AddCurrentValue dd
  
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng
           
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



End Sub


Private Sub ChangeLang()
'XPLbl(1).Caption = "Branch"
Label5.Caption = "Special Reports For LC"
Me.Caption = Label5.Caption
Label2.Caption = "No."
XPLbl(52).Caption = "52"
Label3.Caption = "Name"
XPLbl(7).Caption = "Value"
XPLbl(59).Caption = "Currency"
XPLbl(27).Caption = "Bank"
XPLbl(8).Caption = "Country"
XPLbl(0).Caption = "Vendor"
Label4.Caption = "Beginning Inv. No."
XPPnlTime.Caption = "Open Date"
Frame1.Caption = "Close Date"
Frame5.Caption = "End Date"
Frame6.Caption = "last shipment date"
lbl(2).Caption = "From"
lbl(3).Caption = "From"
lbl(5).Caption = "From"
lbl(7).Caption = "From"
lbl(0).Caption = "To"
lbl(4).Caption = "To"
lbl(6).Caption = "To"
lbl(8).Caption = "To"
Command1.Caption = "Print"
lblCompanyname.Caption = "El-Sattaryh"
End Sub

Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    
   If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
   End If
    
   XPDtbFrom.value = Date
   XPDtpTo.value = Date
    
    Set Dcombos = New ClsDataCombos
    
   Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
   Dcombos.GetLCTypesName Me.DCLC
   Dcombos.GetCUrrencyNames Me.DCCUrrency
   Dcombos.GetBanks Me.DCBank
   Dcombos.GetCountriesNames Me.DCCountry
      
   
With Me.cbValue
.AddItem ">"
.AddItem ">="
.AddItem "<"
.AddItem "<="
.AddItem "="
End With
    
    
  

   
   
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

Private Sub Txt_DateHigriFrom_LostFocus()
 VBA.Calendar = vbCalGreg
            'XPDtbFrom.value = ToGregorianDate(Txt_DateHigriFrom.value)
End Sub

Private Sub Txt_DateHigriTO_LostFocus()
 VBA.Calendar = vbCalGreg
            'XPDtpTo.value = ToGregorianDate(Txt_DateHigriTO.value)
End Sub


Private Sub xpdtbfrom_Change()
If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then
' Txt_DateHigriFrom.value = ToHijriDate(XPDtbFrom.value)
 End If
End Sub

Private Sub XPDtpTo_Change()
If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then
' Txt_DateHigriTO.value = ToHijriDate(XPDtpTo.value)
 End If
End Sub
