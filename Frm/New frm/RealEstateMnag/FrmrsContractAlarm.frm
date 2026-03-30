VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmRsContractAlarm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ‰»ÌÂ«  «·⁄ﬁÊœ «·„‰ ÂÌÂ"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17790
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9540
   ScaleWidth      =   17790
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Õœœ «·› —…"
      Height          =   1440
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   600
      Width           =   13725
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   27
         Top             =   1050
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "≈ŸÂ«— «·ﬂ·"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   26
         Top             =   480
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«Œ›«¡ «·⁄ﬁÊœ «·„Ãœœ…"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   25
         Top             =   240
         Width           =   2535
         _Version        =   786432
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "≈ŸÂ«— «·⁄ﬁÊœ «·„‰ ÂÌ… «·„Ãœœ… ›ﬁÿ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.CheckBox chek 
         Alignment       =   1  'Right Justify
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   195
         Left            =   12480
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Fromdate 
         Height          =   375
         Left            =   10815
         TabIndex        =   18
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   211746817
         CurrentDate     =   41640
      End
      Begin MSComCtl2.DTPicker todate 
         Height          =   375
         Left            =   6600
         TabIndex        =   19
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   211746817
         CurrentDate     =   41640
      End
      Begin Dynamic_Byte.NourHijriCal Fromdate√H 
         Height          =   375
         Left            =   9000
         TabIndex        =   20
         Top             =   240
         Width           =   1815
         _extentx        =   3201
         _extenty        =   661
      End
      Begin Dynamic_Byte.NourHijriCal todateH 
         Height          =   375
         Left            =   4800
         TabIndex        =   21
         Top             =   240
         Width           =   1815
         _extentx        =   3201
         _extenty        =   661
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   390
         Index           =   9
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   688
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "≈÷«›…"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmrsContractAlarm.frx":0000
         DrawFocusRectangle=   0   'False
      End
      Begin MSDataListLib.DataCombo dcBranch 
         Height          =   315
         Left            =   4800
         TabIndex        =   23
         Top             =   720
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   255
         Index           =   3
         Left            =   2070
         TabIndex        =   28
         Top             =   750
         Width           =   2505
         _Version        =   786432
         _ExtentX        =   4419
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«ŸÂ«— «·⁄ﬁÊœ  «·€Ì— «·„ÊÀﬁ… ›ﬁÿ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   195
         Index           =   32
         Left            =   8040
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —… „‰"
         Height          =   315
         Index           =   0
         Left            =   12600
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "≈«·Ï"
         Height          =   435
         Index           =   14
         Left            =   8220
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   8520
      Width           =   885
      Begin VB.CommandButton menue 
         BackColor       =   &H8000000D&
         Caption         =   " ÕœÌÀ"
         Height          =   435
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«—”«· "
      Height          =   495
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   18945
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
               Picture         =   "FrmrsContractAlarm.frx":039A
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmrsContractAlarm.frx":03F8
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmrsContractAlarm.frx":0456
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmrsContractAlarm.frx":04B4
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmrsContractAlarm.frx":0512
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmrsContractAlarm.frx":0570
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmrsContractAlarm.frx":05CE
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmrsContractAlarm.frx":062C
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   8040
         Picture         =   "FrmrsContractAlarm.frx":068A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ‰»ÌÂ«  «·⁄ﬁÊœ «·„‰ ÂÌÂ"
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
         Left            =   10440
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   5880
      End
   End
   Begin ImpulseButton.ISButton btnCancel 
      Height          =   330
      Left            =   120
      TabIndex        =   7
      Top             =   8640
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
      Top             =   8640
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
      Height          =   6120
      Left            =   0
      TabIndex        =   17
      Top             =   2370
      Width           =   17790
      _cx             =   31380
      _cy             =   10795
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
      Cols            =   46
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmrsContractAlarm.frx":42F2
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
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   0
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Â–… «·‘«‘…  ﬁÊ„ »≈ŸÂ«— «·⁄ﬁÊœ «·„‰ ÂÌÂ ÿ»ﬁ« ·Â–« «·ÌÊ„"
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
      TabIndex        =   10
      Top             =   720
      Width           =   4095
   End
End
Attribute VB_Name = "FrmRsContractAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

     Dim My_SQL As String
Private Sub BtnCancel_Click()
    Me.Hide
End Sub


Public Sub FillGrid(Optional str As String)

  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
  
My_SQL = " SELECT     dbo.TblCustemers.CusName, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.CusNamee, dbo.TblContract.NoteSerial1, dbo.TblContract.CusID, "
My_SQL = My_SQL + "                      dbo.TblContract.TodateH, dbo.TblContract.FromdateH, dbo.TblContract.StrDate, dbo.TblContract.EndDate, dbo.TblContract.Iqar, dbo.TblAqar.aqarNo,"
 My_SQL = My_SQL + "                     dbo.TblAqar.aqarname, dbo.TblContract.UnitType, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblContract.UnitNo, dbo.TblAqarDetai.unitno AS unitnoName,"
 My_SQL = My_SQL + "                     dbo.TblContract.TotalContract , dbo.TblContract.Water, dbo.TblContract.Electricity, dbo.TblContract.CommiValue, dbo.TblContract.phone , dbo.TblContract.ContNo , dbo.TblContract.Branch_NO "
My_SQL = My_SQL + " FROM         dbo.TblContract LEFT OUTER JOIN"
 My_SQL = My_SQL + "                     dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
 My_SQL = My_SQL + "                     dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
 My_SQL = My_SQL + "                     dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid"


'If Check1.value = 1 Then
'My_SQL = My_SQL + " WHERE     (dbo.TblContract.EndDate >=" & SQLDate(Me.Fromdate.value, True) & ")and (EndDate <= " & SQLDate(todate.value, True) & ")"
'Else
'My_SQL = My_SQL + " WHERE     (dbo.TblContract.EndDate =" & SQLDate(Date, True) & ")"
'End If
If Rd(1).value = True Then
My_SQL = My_SQL + " WHERE     (dbo.TblContract.Renew = 0 or dbo.TblContract.Renew is null)"
End If

If Rd(3).value = True Then
My_SQL = My_SQL + " WHERE     (dbo.TblContract.Accredit = 0 or dbo.TblContract.Accredit is null)"
End If


If Rd(0).value = True Then
My_SQL = My_SQL + " WHERE     (dbo.TblContract.Renew = 1)"
End If
If Rd(2).value = True Then
My_SQL = My_SQL + " WHERE     (1 = 1)"
End If
 If (Me.Fromdate <> Empty Or Me.Fromdate <> Null) Then
    My_SQL = My_SQL + "  and    (dbo.TblContract.EndDate >=" & SQLDate(Me.Fromdate.value, True) & ")"
    End If
If (Me.ToDate <> Empty Or Me.ToDate <> Null) Then
    My_SQL = My_SQL + " and     (dbo.TblContract.EndDate <=" & SQLDate(Me.ToDate.value, True) & ")"
' Else
'       My_SQL = My_SQL + " and     (dbo.TblContract.EndDate =" & SQLDate(Date, True) & ")"
End If
If SystemOptions.usertype = UserAdminAll Then
If val(dcBranch.BoundText) <> 0 Then
 My_SQL = My_SQL + "   AND (dbo.TblContract.Branch_NO = " & val(dcBranch.BoundText) & ")"
 End If
 Else
  My_SQL = My_SQL + "   AND (dbo.TblContract.Branch_NO = " & Current_branch & ")"
 End If
    'My_SQL = My_SQL + " and (dbo.TblContract.StrDate >='" & SQLDate(Fromdate.value) & "')"
    'My_SQL = My_SQL + " and (dbo.TblContract.StrDate <='" & SQLDate(todate.value) & "')"
     
My_SQL = My_SQL + " and (dbo.TblContract.EndContract IS NULL) "
 

My_SQL = My_SQL + "   order by  dbo.TblContract.ContNo "
   
Dim ActualTotal As Double
'rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .rows - 1
              .TextMatrix(i, .ColIndex("NoteSerial1")) = (IIf(IsNull(rs.Fields("NoteSerial1").value), 0, rs.Fields("NoteSerial1").value))
               .TextMatrix(i, .ColIndex("aqarname")) = (IIf(IsNull(rs.Fields("aqarname").value), "", rs.Fields("aqarname").value))

.TextMatrix(i, .ColIndex("Cus_mobile")) = (IIf(IsNull(rs.Fields("Cus_mobile").value), "", rs.Fields("Cus_mobile").value))
  .TextMatrix(i, .ColIndex("unitnoName")) = (IIf(IsNull(rs.Fields("unitnoName").value), "", rs.Fields("unitnoName").value))
                'TblCustemers.Cus_mobile
 .TextMatrix(i, .ColIndex("FromdateH")) = (IIf(IsNull(rs.Fields("FromdateH").value), ToHijriDate(Date), rs.Fields("FromdateH").value))
  .TextMatrix(i, .ColIndex("StrDate")) = IIf(IsNull(rs.Fields("StrDate").value), Date, rs.Fields("StrDate").value)
   .TextMatrix(i, .ColIndex("TodateH")) = (IIf(IsNull(rs.Fields("TodateH").value), ToHijriDate(Date), rs.Fields("TodateH").value))
  .TextMatrix(i, .ColIndex("EndDate")) = IIf(IsNull(rs.Fields("EndDate").value), Date, rs.Fields("EndDate").value)
  
  
     
                      '    ActualTotal = getinsttPayedTocontract(val(rs.Fields("id").value))
 '.TextMatrix(i, .ColIndex("payed")) = ActualTotal
  '.TextMatrix(i, .ColIndex("Remains")) = val(.TextMatrix(i, .ColIndex("Value"))) - ActualTotal

'If ActualTotal = 0 Then
'          .Cell(flexcpBackColor, i, 1, i, 37) = vbRed
'Else
'          .Cell(flexcpBackColor, i, 1, i, 37) = vbYellow
'End If
       .TextMatrix(i, .ColIndex("Iqar")) = (IIf(IsNull(rs.Fields("Iqar").value), "", rs.Fields("Iqar").value))
       .TextMatrix(i, .ColIndex("UnitType")) = (IIf(IsNull(rs.Fields("UnitType").value), "", rs.Fields("UnitType").value))
         .TextMatrix(i, .ColIndex("UnitNo")) = (IIf(IsNull(rs.Fields("UnitNo").value), "", rs.Fields("UnitNo").value))
     .TextMatrix(i, .ColIndex("CusID")) = (IIf(IsNull(rs.Fields("CusID").value), "", rs.Fields("CusID").value))
   
   If SystemOptions.UserInterface = ArabicInterface Then
   .TextMatrix(i, .ColIndex("name")) = (IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value))
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
   Else
   .TextMatrix(i, .ColIndex("name")) = (IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value))
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value))
   End If

    .TextMatrix(i, .ColIndex("TotalContract")) = (IIf(IsNull(rs.Fields("TotalContract").value), 0, rs.Fields("TotalContract").value))
    .TextMatrix(i, .ColIndex("Water")) = (IIf(IsNull(rs.Fields("Water").value), 0, rs.Fields("Water").value))
    .TextMatrix(i, .ColIndex("CommiValue")) = (IIf(IsNull(rs.Fields("CommiValue").value), 0, rs.Fields("CommiValue").value))
    .TextMatrix(i, .ColIndex("Water")) = (IIf(IsNull(rs.Fields("Water").value), 0, rs.Fields("Water").value))
    .TextMatrix(i, .ColIndex("Electricity")) = (IIf(IsNull(rs.Fields("Electricity").value), 0, rs.Fields("Electricity").value))
    .TextMatrix(i, .ColIndex("Phone")) = (IIf(IsNull(rs.Fields("Phone").value), 0, rs.Fields("Phone").value))
 
       .TextMatrix(i, .ColIndex("net")) = val(.TextMatrix(i, .ColIndex("TotalContract"))) + val(.TextMatrix(i, .ColIndex("Water"))) + val(.TextMatrix(i, .ColIndex("Electricity"))) + val(.TextMatrix(i, .ColIndex("Phone")))
    .TextMatrix(i, .ColIndex("ContNo")) = (IIf(IsNull(rs.Fields("ContNo").value), 0, rs.Fields("ContNo").value))
'.TextMatrix(i, .ColIndex("Doneofall")) = (IIf(IsNull(rs.Fields("Doneofall").value), 0, rs.Fields("Doneofall").value))

        rs.MoveNext
            Next i
 
            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With
ReLineGrid
End Sub



Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
 
  '  Dim Percenrage As Double
 
 
   ' IntCounter = 0
  'Me.TxtTotalContract.text = 0
  'Me.TxtCommiValue.text = 0
  '  Me.TxtInsuranceValue.text = 0
  '    Me.TxtWater.text = 0
  '    Me.TxtElectricity.text = 0
  '      Me.TxtPhone.text = 0
     
    With Me.GridInstallments

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("CusName")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                
             If chek.value = vbChecked Then
                .TextMatrix(i, .ColIndex("Send")) = -1
                Else
                .TextMatrix(i, .ColIndex("Send")) = 0
      End If
  'Me.TxtTotalContract.text = val(Me.TxtTotalContract.text) + .TextMatrix(i, .ColIndex("RentValue"))
  'Me.TxtCommiValue.text = val(Me.TxtCommiValue.text) + .TextMatrix(i, .ColIndex("Commissions"))
  'Me.TxtInsuranceValue.text = val(Me.TxtInsuranceValue.text) + .TextMatrix(i, .ColIndex("Insurance"))
  'Me.TxtWater.text = val(Me.TxtWater.text) + .TextMatrix(i, .ColIndex("Water"))
  'Me.TxtElectricity.text = val(Me.TxtElectricity.text) + .TextMatrix(i, .ColIndex("Electric"))
  'Me.TxtPhone.text = val(Me.TxtPhone.text) + .TextMatrix(i, .ColIndex("TelandNet"))
  
 ' End If
  
     
         
            End If

        Next i
   
    End With

End Sub


Private Sub chek_Click()
ReLineGrid
End Sub



Private Sub Cmd_Click(index As Integer)
FillGrid
End Sub

Private Sub CmdPrint_Click()
    On Error Resume Next
    'Dim GrdBack As ClsBackGroundPic
    print_report My_SQL
    'Grid.ExtendLastCol = True
 '   Grid.WallPaper = Nothing
    'Grid.AutoSize  0, Grid.Cols - 1, False
'    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
    'Printer.RightToLeft = True
    'Printer.Print ("Employee Salary Report")

'    Me.Grid.PrintGrid " ‰»Ì…    „” Œ·’«  ·„  ”œœ »«·ﬂ«„·", True, 2, 1, 1500
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
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepContractAlarm.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepContractAlarm.rpt"
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

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    If Not IsNull(Fromdate.value) Then
 xReport.ParameterFields(14).AddCurrentValue Fromdate.value
 xReport.ParameterFields(16).AddCurrentValue Fromdate√H.value
 End If
 If Not IsNull(ToDate.value) Then
   xReport.ParameterFields(15).AddCurrentValue ToDate.value
    xReport.ParameterFields(17).AddCurrentValue TodateH.value
 End If
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function
Private Sub Command1_Click()
    Dim Numbers As String
    Dim RowNum As Integer
    Dim Opt As Integer
    Dim CurrentMessage As String
    Numbers = ""

    With GridInstallments

        For RowNum = .FixedRows To .rows - 1
    
            If .Cell(flexcpChecked, RowNum, .ColIndex("Send")) = flexChecked Then

                '  MsgBox (.TextMatrix(RowNum, .ColIndex("Numbers")))
                If (.TextMatrix(RowNum, .ColIndex("Cus_mobile"))) <> "" Then
                    If Numbers = "" Then
                        Numbers = (.TextMatrix(RowNum, .ColIndex("Cus_mobile")))
                    Else
                        Numbers = Numbers & "," & (.TextMatrix(RowNum, .ColIndex("Cus_mobile")))
                    End If
             
                End If
            End If
          
        Next RowNum
      
        CurrentMessage = ComposMessage(Me.Name)  ', 0, "", Me.TXTMessageDES.text, Opt)

        If Numbers = "" Then Exit Sub
        SMSSeTTings.SendMessage CurrentMessage, Numbers
        SMSSeTTings.Hide
                                    
    End With

End Sub

Private Sub Form_Load()
 
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
     
      ToDate.value = Date
      
  Dim Dcombos As ClsDataCombos
 Set Dcombos = New ClsDataCombos
Dcombos.GetBranches dcBranch
Rd(1).value = True
If SystemOptions.UserInterface = EnglishInterface Then

        SetInterface Me
        cahngelang
    End If
'Dtday.value = Date
FillGrid
End Sub

Function cahngelang()
    Label1(2).Caption = "Project Invoices Not Payed"
    Me.Caption = Label1(2).Caption
    Frame1.Caption = "Color Map"
    Label3.Caption = "Fully"
    Label5.Caption = "Partial"

    Me.Caption = Label1(2).Caption
    CmdPrint.Caption = "Print"
    btnCancel.Caption = "Cancel"



End Function

Private Sub FromDate_Change()
If Not (IsNull(Fromdate.value)) Then
 Fromdate√H.value = ToHijriDate(Fromdate.value)
 End If
End Sub

Private Sub Fromdate√H_LostFocus()
 VBA.Calendar = vbCalGreg
 Fromdate.value = ToGregorianDate(Fromdate√H.value)
End Sub

Private Sub GridInstallments_CellButtonClick(ByVal row As Long, ByVal Col As Long)
With Me.GridInstallments
Select Case .ColKey(Col)
Case "Sho"
Load RSContract
RSContract.show
'RSContract.RereivID = val(.TextMatrix(.Row, .ColIndex("ContNo")))
RSContract.BtnLast_Click
RSContract.RereivID = val(.TextMatrix(.row, .ColIndex("ContNo")))
RSContract.FindRec val(.TextMatrix(.row, .ColIndex("ContNo")))
End Select
End With
End Sub

Private Sub GridInstallments_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.GridInstallments
Select Case .ColKey(Col)
Case "Sho"
'.ColComboList(.ColIndex("timEnter")) = "..."
.ColComboList(.ColIndex("Sho")) = "..."
End Select
End With
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub menue_Click()
FillGrid
End Sub



Private Sub Rd_Click(index As Integer)
FillGrid
End Sub

Private Sub ToDate_Change()
If Not (IsNull(ToDate.value)) Then
 TodateH.value = ToHijriDate(ToDate.value)
 End If
End Sub

Private Sub ToDateH_LostFocus()
 VBA.Calendar = vbCalGreg
 ToDate.value = ToGregorianDate(TodateH.value)
End Sub
