VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmOldInstallments 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÇáÏÝÚÇÊ ÞÈá ÇáÊÚÏíá"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11940
   Icon            =   "frmOldInstallments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   90
      TabIndex        =   0
      Top             =   2850
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÎÑæÌ"
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
      Height          =   2685
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   11805
      _cx             =   20823
      _cy             =   4736
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
      Cols            =   61
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmOldInstallments.frx":038A
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
   Begin VB.Label LblTotalQasts 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "ÇÌãÇáí ÇáÏÝÚÇÊ"
      Height          =   285
      Index           =   34
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2880
      Width           =   1890
   End
   Begin VB.Label LblNotPayed 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "ÛíÑ ãÓÏÏ"
      Height          =   285
      Index           =   36
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2880
      Width           =   1410
   End
End
Attribute VB_Name = "FrmOldInstallments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Public ContNo As String
Private Sub Cmd_Click(Index As Integer)
    Dim Msg As String
    On Error GoTo ErrTrap

    Select Case Index

        Case 0



        Case 2
            Unload Me
    End Select

    Exit Sub

End Sub



End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
  
    Set rs = New ADODB.Recordset
   
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    CenterForm Me

    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    Retrive
    Exit Sub
ErrTrap:
End Sub
Private Sub ChangeLang()
 
 
    Me.Caption = "Search Moving Vouchers"
 
    XPLbl(0).Caption = "VCHR#"
 
    XPLbl(1).Caption = "From"
    XPLbl(2).Caption = "User"
    XPLbl(3).Caption = "To"
    XPLbl(4).Caption = "From Store"
    XPLbl(4).Caption = "To Store"
 
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Delete"
    Cmd(2).Caption = "Exit"
 
    With FG
    
    .TextMatrix(0, .ColIndex("count")) = "I"
        .TextMatrix(0, .ColIndex("Serial")) = "Vchr#"
        .TextMatrix(0, .ColIndex("BillDate")) = " Date"
        .TextMatrix(0, .ColIndex("ClientNmae")) = "From Store "
        .TextMatrix(0, .ColIndex("StorName")) = "To Store"
 
  
    End With

End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    GridInstallments.Clear flexClearScrollable, flexClearEverything
    Dim My_SQL  As String
    Dim i As Long

Dim notpayed As Double
notpayed = 0
 
My_SQL = " SELECT  * from TblContractInstallmentsOld"
My_SQL = My_SQL & " WHERE     (ContNo =" & val(ContNo) & ")  order by InstallNo"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'    rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
           .TextMatrix(i, .ColIndex("DevID")) = (IIf(IsNull(rs.Fields("DevID").value), 0, rs.Fields("DevID").value))
          .TextMatrix(i, .ColIndex("Installid")) = (IIf(IsNull(rs.Fields("id").value), 0, rs.Fields("id").value))
          .TextMatrix(i, .ColIndex("TempInstal")) = (IIf(IsNull(rs.Fields("TempInstal").value), 0, rs.Fields("TempInstal").value))
          .TextMatrix(i, .ColIndex("InstallNo")) = (IIf(IsNull(rs.Fields("InstallNo").value), 0, rs.Fields("InstallNo").value))
          .TextMatrix(i, .ColIndex("hijri")) = (IIf(IsNull(rs.Fields("hijri").value), 1, rs.Fields("hijri").value))
          .TextMatrix(i, .ColIndex("DES")) = (IIf(IsNull(rs.Fields("DES").value), "", rs.Fields("DES").value))
          .TextMatrix(i, .ColIndex("Due_DateH")) = Format((IIf(IsNull(rs.Fields("Installdateh").value), ToHijriDate(Date), rs.Fields("Installdateh").value)), "yyyy/MM/dd")
           .TextMatrix(i, .ColIndex("Due_Date")) = Format(IIf(IsNull(rs.Fields("Installdate").value), Date, rs.Fields("Installdate").value), "yyyy/MM/dd")
        'yyyy/MM/dd
       .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(rs.Fields("installValue").value), 0, rs.Fields("installValue").value))
       .TextMatrix(i, .ColIndex("ServiceArbon")) = (IIf(IsNull(rs.Fields("ServiceArbon").value), 0, rs.Fields("ServiceArbon").value))
       .TextMatrix(i, .ColIndex("NoteSerial")) = (IIf(IsNull(rs.Fields("NoteSerial").value), "", rs.Fields("NoteSerial").value))
       .TextMatrix(i, .ColIndex("NoteSerial1")) = (IIf(IsNull(rs.Fields("NoteSerial1").value), "", rs.Fields("NoteSerial1").value))
       .TextMatrix(i, .ColIndex("NoteId")) = (IIf(IsNull(rs.Fields("NoteId").value), "", rs.Fields("NoteId").value))
If Not IsNull(rs.Fields("Status").value) Then
             If rs.Fields("Status").value = 0 Then
                    .Cell(flexcpChecked, i, .ColIndex("Status")) = flexUnchecked
            Else
                     .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked
                       notpayed = notpayed + val(.TextMatrix(i, .ColIndex("Value")))
            End If

End If

    .TextMatrix(i, .ColIndex("RentValue")) = (IIf(IsNull(rs.Fields("RentValue").value), 0, rs.Fields("RentValue").value))
    .TextMatrix(i, .ColIndex("VATPayed")) = (IIf(IsNull(rs.Fields("VATPayed").value), 0, rs.Fields("VATPayed").value))
    .TextMatrix(i, .ColIndex("VATValue")) = (IIf(IsNull(rs.Fields("VATValue").value), 0, rs.Fields("VATValue").value))
    .TextMatrix(i, .ColIndex("Commissions")) = (IIf(IsNull(rs.Fields("Commissions").value), 0, rs.Fields("Commissions").value))
    .TextMatrix(i, .ColIndex("Insurance")) = (IIf(IsNull(rs.Fields("Insurance").value), 0, rs.Fields("Insurance").value))
    .TextMatrix(i, .ColIndex("Water")) = (IIf(IsNull(rs.Fields("Water").value), 0, rs.Fields("Water").value))
    .TextMatrix(i, .ColIndex("Electric")) = (IIf(IsNull(rs.Fields("Electric").value), 0, rs.Fields("Electric").value))
    .TextMatrix(i, .ColIndex("TelandNet")) = (IIf(IsNull(rs.Fields("TelandNet").value), 0, rs.Fields("TelandNet").value))
 .TextMatrix(i, .ColIndex("NpayedValue")) = (IIf(IsNull(rs.Fields("NpayedValue").value), 0, rs.Fields("NpayedValue").value))
        
    .TextMatrix(i, .ColIndex("OldValue")) = (IIf(IsNull(rs.Fields("OldValue").value), 0, rs.Fields("OldValue").value))
'    .TextMatrix(i, .ColIndex("Remains")) = (IIf(IsNull(rs.Fields("Remains").value), 0, rs.Fields("Remains").value))
    
    
    .TextMatrix(i, .ColIndex("RentValuePayed")) = (IIf(IsNull(rs.Fields("RentValuePayed").value), 0, rs.Fields("RentValuePayed").value))
    .TextMatrix(i, .ColIndex("CommissionsPayed")) = (IIf(IsNull(rs.Fields("CommissionsPayed").value), 0, rs.Fields("CommissionsPayed").value))
    .TextMatrix(i, .ColIndex("InsurancePayed")) = (IIf(IsNull(rs.Fields("InsurancePayed").value), 0, rs.Fields("InsurancePayed").value))
    .TextMatrix(i, .ColIndex("WaterPayed")) = (IIf(IsNull(rs.Fields("WaterPayed").value), 0, rs.Fields("WaterPayed").value))
    .TextMatrix(i, .ColIndex("ElectricPayed")) = (IIf(IsNull(rs.Fields("ElectricPayed").value), 0, rs.Fields("ElectricPayed").value))
    .TextMatrix(i, .ColIndex("TelandNetPayed")) = (IIf(IsNull(rs.Fields("TelandNetPayed").value), 0, rs.Fields("TelandNetPayed").value))
'   .TextMatrix(i, .ColIndex("Payed")) = (IIf(IsNull(rs.Fields("Payed").value), 0, rs.Fields("Payed").value))
  '.TextMatrix(i, .ColIndex("Remains")) = val(.TextMatrix(i, .ColIndex("Value"))) - val(.TextMatrix(i, .ColIndex("Payed")))
  ''// 19 08 2015
  .TextMatrix(i, .ColIndex("Rent1")) = (IIf(IsNull(rs.Fields("Rent1").value), 0, rs.Fields("Rent1").value))
  .TextMatrix(i, .ColIndex("RentArbon")) = (IIf(IsNull(rs.Fields("RentArbon").value), 0, rs.Fields("RentArbon").value))
  .TextMatrix(i, .ColIndex("NetRent")) = (IIf(IsNull(rs.Fields("NetRent").value), 0, rs.Fields("NetRent").value))
  .TextMatrix(i, .ColIndex("Commissions1")) = (IIf(IsNull(rs.Fields("Commissions1").value), 0, rs.Fields("Commissions1").value))
  .TextMatrix(i, .ColIndex("CommissionsArbon")) = (IIf(IsNull(rs.Fields("CommissionsArbon").value), 0, rs.Fields("CommissionsArbon").value))
  .TextMatrix(i, .ColIndex("NetCommissions")) = (IIf(IsNull(rs.Fields("NetCommissions").value), 0, rs.Fields("NetCommissions").value))
  .TextMatrix(i, .ColIndex("Insurance1")) = (IIf(IsNull(rs.Fields("Insurance1").value), 0, rs.Fields("Insurance1").value))
  .TextMatrix(i, .ColIndex("InsuranceArbon")) = (IIf(IsNull(rs.Fields("InsuranceArbon").value), 0, rs.Fields("InsuranceArbon").value))
  .TextMatrix(i, .ColIndex("NetInsurance")) = (IIf(IsNull(rs.Fields("NetInsurance").value), 0, rs.Fields("NetInsurance").value))
  .TextMatrix(i, .ColIndex("Water1")) = (IIf(IsNull(rs.Fields("Water1").value), 0, rs.Fields("Water1").value))
  .TextMatrix(i, .ColIndex("WaterArbon")) = (IIf(IsNull(rs.Fields("WaterArbon").value), 0, rs.Fields("WaterArbon").value))
  
  .TextMatrix(i, .ColIndex("Electric1")) = (IIf(IsNull(rs.Fields("Electric1").value), 0, rs.Fields("Electric1").value))
  .TextMatrix(i, .ColIndex("ElectricArbon")) = (IIf(IsNull(rs.Fields("ElectricArbon").value), 0, rs.Fields("ElectricArbon").value))
  .TextMatrix(i, .ColIndex("NetElectric")) = (IIf(IsNull(rs.Fields("Electric").value), 0, rs.Fields("Electric").value))
  .TextMatrix(i, .ColIndex("NetWater")) = (IIf(IsNull(rs.Fields("Water").value), 0, rs.Fields("Water").value))
  '.TextMatrix(i, .ColIndex("NetElectric")) = (IIf(IsNull(rs.Fields("NetElectric").value), 0, rs.Fields("NetElectric").value))
  '.TextMatrix(i, .ColIndex("NetWater")) = (IIf(IsNull(rs.Fields("NetWater").value), 0, rs.Fields("NetWater").value))
  
  ''//
  Dim X As String
  Dim RentValuePayed   As Double
  Dim CommissionsPayed  As Double
  Dim InsurancePayed    As Double
  Dim WaterPayed   As Double
  Dim ElectricPayed   As Double
  Dim TelandNetPayed  As Double
  Dim payed As Double
  Dim VATPayed As Double
'   getinsttPayedTocontract(val(rs.Fields("id").value), RentValuePayed, CommissionsPayed, InsurancePayed, WaterPayed, ElectricPayed, TelandNetPayed)
            payed = getinsttPayedTocontract(val(rs.Fields("id").value), RentValuePayed, CommissionsPayed, InsurancePayed, WaterPayed, ElectricPayed, TelandNetPayed, , , , VATPayed)

.TextMatrix(i, .ColIndex("RentValuePayed")) = RentValuePayed
.TextMatrix(i, .ColIndex("CommissionsPayed")) = CommissionsPayed
.TextMatrix(i, .ColIndex("InsurancePayed")) = InsurancePayed
.TextMatrix(i, .ColIndex("WaterPayed")) = WaterPayed
.TextMatrix(i, .ColIndex("ElectricPayed")) = ElectricPayed
.TextMatrix(i, .ColIndex("TelandNetPayed")) = TelandNetPayed
.TextMatrix(i, .ColIndex("VATPayed")) = VATPayed
     
    '      payed = payed + (IIf(IsNull(rs.Fields("RentValuePayed").value), 0, rs.Fields("RentValuePayed").value)) 'val(rs("RentValuePayed").value)
     
  '        payed = payed + (IIf(IsNull(rs.Fields("CommissionsPayed").value), 0, rs.Fields("CommissionsPayed").value))  ' val(rs("CommissionsPayed").value)
     
  '        payed = payed + (IIf(IsNull(rs.Fields("InsurancePayed").value), 0, rs.Fields("InsurancePayed").value))  '   val(rs("InsurancePayed").value)
  '
  '        payed = payed + (IIf(IsNull(rs.Fields("WaterPayed").value), 0, rs.Fields("WaterPayed").value))  ' val(rs("WaterPayed").value)
  '
  '        payed = payed + (IIf(IsNull(rs.Fields("ElectricPayed").value), 0, rs.Fields("ElectricPayed").value))     'val(rs("ElectricPayed").value)
  '
  '      payed = payed + (IIf(IsNull(rs.Fields("TelandNetPayed").value), 0, rs.Fields("TelandNetPayed").value)) ' val(rs("TelandNetPayed").value)
  '
        .TextMatrix(i, .ColIndex("Payed")) = payed
              
  .TextMatrix(i, .ColIndex("Remains")) = val(.TextMatrix(i, .ColIndex("Value"))) - val(.TextMatrix(i, .ColIndex("Payed")))
   ' .TextMatrix(i, .ColIndex("payedPayed")) = (IIf(IsNull(rs.Fields("payedPayed").value), 0, rs.Fields("payedPayed").value))
   ' .TextMatrix(i, .ColIndex("RemainsPayed")) = (IIf(IsNull(rs.Fields("RemainsPayed").value), 0, rs.Fields("RemainsPayed").value))
    
       .TextMatrix(i, .ColIndex("lastPayedDate")) = Format((IIf(IsNull(rs.Fields("lastPayedDate").value), Format(Date, "yyyy/MM/dd"), rs.Fields("lastPayedDate").value)), "yyyy/MM/dd")
 .TextMatrix(i, .ColIndex("lastPayedDateH")) = Format((IIf(IsNull(rs.Fields("lastPayedDateH").value), Format(ToHijriDate(Date), "yyyy/MM/dd"), rs.Fields("lastPayedDateH").value)), "yyyy/MM/dd")
     .TextMatrix(i, .ColIndex("allocations")) = (IIf(IsNull(rs.Fields("allocations").value), 0, rs.Fields("allocations").value))
.TextMatrix(i, .ColIndex("Countsofall")) = (IIf(IsNull(rs.Fields("Countsofall").value), 0, rs.Fields("Countsofall").value))
.TextMatrix(i, .ColIndex("Doneofall")) = (IIf(IsNull(rs.Fields("Doneofall").value), 0, rs.Fields("Doneofall").value))

        rs.MoveNext
            Next i
If rs.RecordCount > 0 Then
  Me.LblTotalQasts.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
Else
Me.LblTotalQasts.Caption = 0
End If
            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False
        'Me.LblTotalQasts.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        .RowHeight(-1) = 300
    End With


    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub


