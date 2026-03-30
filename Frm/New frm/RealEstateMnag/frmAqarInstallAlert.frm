VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAqarInstallAlert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ‰»ÌÂ«  œ›⁄«  «·⁄ﬁ«—"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17715
   Icon            =   "frmAqarInstallAlert.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   17715
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Õœœ «·› —…"
      Height          =   960
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   480
      Width           =   12045
      Begin VB.CheckBox chek 
         Alignment       =   1  'Right Justify
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   195
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker Fromdate 
         Height          =   375
         Left            =   9015
         TabIndex        =   14
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93388801
         CurrentDate     =   41640
      End
      Begin MSComCtl2.DTPicker todate 
         Height          =   375
         Left            =   4680
         TabIndex        =   15
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93388801
         CurrentDate     =   41640
      End
      Begin Dynamic_Byte.NourHijriCal Fromdate√H 
         Height          =   375
         Left            =   7200
         TabIndex        =   16
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
      End
      Begin Dynamic_Byte.NourHijriCal todateH 
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   510
         Index           =   9
         Left            =   1800
         TabIndex        =   20
         Top             =   120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   900
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
         ButtonImage     =   "frmAqarInstallAlert.frx":058A
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "≈«·Ï"
         Height          =   435
         Index           =   14
         Left            =   6420
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —… „‰"
         Height          =   315
         Index           =   0
         Left            =   10800
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«—”«· "
      Height          =   495
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   17745
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
               Picture         =   "frmAqarInstallAlert.frx":0924
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAqarInstallAlert.frx":0CBE
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAqarInstallAlert.frx":1058
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAqarInstallAlert.frx":13F2
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAqarInstallAlert.frx":178C
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAqarInstallAlert.frx":1B26
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAqarInstallAlert.frx":1EC0
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAqarInstallAlert.frx":245A
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   720
         Picture         =   "frmAqarInstallAlert.frx":27F4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ‰»ÌÂ«  œ›⁄«  «·⁄ﬁ«—"
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
         Left            =   11640
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   0
         Width           =   5880
      End
   End
   Begin ImpulseButton.ISButton btnCancel 
      Height          =   330
      Left            =   240
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
      ButtonImage     =   "frmAqarInstallAlert.frx":645C
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton CmdPrint 
      Height          =   390
      Left            =   1320
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
      ButtonImage     =   "frmAqarInstallAlert.frx":67F6
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
      Height          =   6885
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   17685
      _cx             =   31194
      _cy             =   12144
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
      Rows            =   50
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAqarInstallAlert.frx":6B90
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
      Editable        =   0
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
      Height          =   735
      Left            =   0
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Â–… ‘«‘…  ‰»Ì«  œ›⁄«  «·⁄ﬁ«—"
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
      Height          =   660
      Index           =   25
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   600
      Width           =   5655
   End
End
Attribute VB_Name = "frmAqarInstallAlert"
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
    
   ' My_SQL = "SELECT    dbo.TblAqar.aqarname, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblContract.ContNo, dbo.TblContractInstallments.InstallNo, dbo.TblContractInstallments.Installdate, dbo.TblContractInstallments.InstalldateH, dbo.TblAqar.Aqarid,  dbo.TblAqrOwin.AllowDateH , dbo.TblAqrOwin.AllowDate, dbo.TblContractInstallments.installValue FROM  dbo.TblAqar INNER JOIN  dbo.TblCustemers ON dbo.TblAqar.ownerid = dbo.TblCustemers.CusID LEFT OUTER JOIN  dbo.TblAqrOwin ON dbo.TblAqar.Aqarid = dbo.TblAqrOwin.AqrID LEFT OUTER JOIN  dbo.TblContract ON dbo.TblAqar.Aqarid = dbo.TblContract.Iqar LEFT OUTER JOIN dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo"
    My_SQL = " SELECT     dbo.TblAqar.aqarname, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblAqar.Aqarid, dbo.TblAqrOwin.AllowDateH, dbo.TblAqrOwin.AllowDate, "
    My_SQL = My_SQL & "                  dbo.TblAqrOwin.[value], dbo.TblAqrOwin.RecDate, dbo.TblAqrOwin.RecDateH, dbo.TblAqrOwin.DMY, dbo.TblAqrOwin.Cont, dbo.TblAqrOwin.PaymentNo,"
    My_SQL = My_SQL & "                  dbo.TblAqrOwin.TotalPayed, dbo.GetOwnerPayment(dbo.TblAqrOwin.ID) AS PayedValue, dbo.TblAqrOwin.ID"
    My_SQL = My_SQL & "    FROM         dbo.TblAqar LEFT OUTER JOIN"
    My_SQL = My_SQL & "                  dbo.TblCustemers ON dbo.TblAqar.ownerid = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                  dbo.TblAqrOwin ON dbo.TblAqar.Aqarid = dbo.TblAqrOwin.AqrID"
    My_SQL = My_SQL & " WHERE     ((dbo.TblAqrOwin.TotalPayed = 0) OR"
    My_SQL = My_SQL & "                  (dbo.TblAqrOwin.TotalPayed IS NULL)) and ( dbo.TblAqrOwin.[value]<>0) "

     If (Me.FromDate <> Empty Or Me.FromDate <> Null) Then
    My_SQL = My_SQL + "  and    (AllowDate >=" & SQLDate(Me.FromDate.value, True) & ")"
    End If
If (Me.ToDate <> Empty Or Me.ToDate <> Null) Then
    My_SQL = My_SQL + " and     (AllowDate <=" & SQLDate(Me.ToDate.value, True) & ")"
' Else
'       My_SQL = My_SQL + " and     (dbo.TblContract.EndDate =" & SQLDate(Date, True) & ")"
End If




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
             .TextMatrix(i, .ColIndex("aqarname")) = (IIf(IsNull(rs.Fields("aqarname").value), 0, rs.Fields("aqarname").value))
             .TextMatrix(i, .ColIndex("ownername")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
             .TextMatrix(i, .ColIndex("InstallNo")) = (IIf(IsNull(rs.Fields("PaymentNo").value), "", rs.Fields("PaymentNo").value))
             .TextMatrix(i, .ColIndex("date")) = (IIf(IsNull(rs.Fields("RecDate").value), "", rs.Fields("RecDate").value))
             .TextMatrix(i, .ColIndex("dateh")) = (IIf(IsNull(rs.Fields("RecDateH").value), "", rs.Fields("RecDateH").value))
             .TextMatrix(i, .ColIndex("InstallValue")) = IIf(IsNull(rs.Fields("value").value), 0, rs.Fields("value").value)
             .TextMatrix(i, .ColIndex("AllowDateH")) = (IIf(IsNull(rs.Fields("AllowDate").value), "", rs.Fields("AllowDate").value))
             .TextMatrix(i, .ColIndex("AllowDate")) = IIf(IsNull(rs.Fields("AllowDateH").value), "", rs.Fields("AllowDateH").value)
             .TextMatrix(i, .ColIndex("PayedValue")) = IIf(IsNull(rs.Fields("PayedValue").value), 0, rs.Fields("PayedValue").value)
             .TextMatrix(i, .ColIndex("ReValue")) = val(.TextMatrix(i, .ColIndex("InstallValue"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
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

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("ownerName")) <> "" Then
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





Private Sub Cmd_Click(Index As Integer)
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
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_IstallmentAlert.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_IstallmentAlert.rpt"
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
        If FromDate.value <> Null Or FromDate.value <> "" Then xReport.ParameterFields(14).AddCurrentValue Format(Me.FromDate.value, "yyyy/M/d")
        If ToDate.value <> Null Or ToDate.value <> "" Then xReport.ParameterFields(15).AddCurrentValue Format(Me.ToDate.value, "yyyy/M/d")
        If FromDate.value <> Null Or FromDate.value <> "" Then xReport.ParameterFields(16).AddCurrentValue Me.Fromdate√H.value
        If ToDate.value <> Null Or ToDate.value <> "" Then xReport.ParameterFields(17).AddCurrentValue Me.todateH.value
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
         If FromDate.value <> Null Or FromDate.value <> "" Then xReport.ParameterFields(14).AddCurrentValue Format(Me.FromDate.value, "yyyy/M/d")
        If ToDate.value <> Null Or ToDate.value <> "" Then xReport.ParameterFields(15).AddCurrentValue Format(Me.ToDate.value, "yyyy/M/d")
        If FromDate.value <> Null Or FromDate.value <> "" Then xReport.ParameterFields(16).AddCurrentValue Me.Fromdate√H.value
        If ToDate.value <> Null Or ToDate.value <> "" Then xReport.ParameterFields(17).AddCurrentValue Me.todateH.value
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
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
Private Sub Command1_Click()
    Dim Numbers As String
    Dim RowNum As Integer
    Dim Opt As Integer
    Dim CurrentMessage As String
    Numbers = ""

    With GridInstallments

        For RowNum = .FixedRows To .Rows - 1
    
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
 
 
FromDate.value = Date
ToDate.value = Date

 
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
    ' Fromdate.value = Date
    '  todate.value = rentInstallmentdate
      

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

Private Sub menue_Click()
FillGrid
End Sub

Private Sub FromDate_Change()
 If (FromDate.value <> Null Or FromDate.value <> "") Then
     Fromdate√H.value = ToHijriDate(FromDate.value)
 End If
End Sub

Private Sub Fromdate√H_LostFocus()
 VBA.Calendar = vbCalGreg
 FromDate.value = ToGregorianDate(Fromdate√H.value)
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ToDate_Change()
If ToDate.value <> Null Or ToDate.value <> "" Then
 todateH.value = ToHijriDate(ToDate.value)
 End If
End Sub

Private Sub ToDateH_LostFocus()
 VBA.Calendar = vbCalGreg
 ToDate.value = ToGregorianDate(todateH.value)
End Sub
