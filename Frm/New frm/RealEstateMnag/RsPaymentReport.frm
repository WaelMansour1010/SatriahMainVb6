VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RsPayemntReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ﬁ—Ì—  Õ’Ì· «·«ÌÃ«—« "
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15180
   Icon            =   "RsPaymentReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   15180
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   -15
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   15105
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
         Left            =   120
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
         Left            =   480
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
               Picture         =   "RsPaymentReport.frx":058A
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RsPaymentReport.frx":0924
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RsPaymentReport.frx":0CBE
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RsPaymentReport.frx":1058
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RsPaymentReport.frx":13F2
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RsPaymentReport.frx":178C
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RsPaymentReport.frx":1B26
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RsPaymentReport.frx":20C0
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   6120
         Picture         =   "RsPaymentReport.frx":245A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ﬁ—Ì—  Õ’Ì· «·«ÌÃ«—« "
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
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   5040
      End
   End
   Begin ImpulseButton.ISButton btnCancel 
      Height          =   330
      Left            =   120
      TabIndex        =   7
      Top             =   5160
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
      ButtonImage     =   "RsPaymentReport.frx":60C2
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton CmdPrint 
      Height          =   390
      Left            =   1080
      TabIndex        =   8
      Top             =   5160
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
      ButtonImage     =   "RsPaymentReport.frx":645C
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4035
      Left            =   0
      TabIndex        =   9
      Top             =   6240
      Width           =   13995
      _cx             =   24686
      _cy             =   7117
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
      Rows            =   2
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"RsPaymentReport.frx":67F6
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
      ExplorerBar     =   0
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
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   3660
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   15075
      _cx             =   26591
      _cy             =   6456
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   24
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"RsPaymentReport.frx":6A2F
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
      ExplorerBar     =   0
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
   Begin MSDataListLib.DataCombo DcboGovernmentID 
      Height          =   315
      Left            =   10440
      TabIndex        =   11
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   720
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   270
      Left            =   7680
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   476
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   10383715
      CustomFormat    =   "yyyy/M/d"
      Format          =   95223811
      CurrentDate     =   37140
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   270
      Left            =   5520
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   476
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   10383715
      CustomFormat    =   "yyyy/M/d"
      Format          =   95223811
      CurrentDate     =   37140
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·„«·ﬂ"
      Height          =   285
      Index           =   1
      Left            =   13290
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   720
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·Ï"
      Height          =   405
      Index           =   23
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   720
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰"
      Height          =   285
      Index           =   22
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   720
      Width           =   810
   End
End
Attribute VB_Name = "RsPayemntReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub FillGridWithData()

    'On Error GoTo ErrTrap
 
    Dim i As Integer
    Dim X As Integer
    Dim rs As ADODB.Recordset
 
    Dim ActualTotal As Double
    Dim result As Double
    Dim resultpercentage As Double
    Dim sql As String
    sql = "SELECT  * FROM     project_billl  "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
 
        Exit Sub
    End If
 
    With Me.Grid
        .Rows = 1
        .Clear flexClearScrollable
 
        rs.MoveFirst
        
        For X = 1 To rs.RecordCount
       
            ActualTotal = getBillPayedToproject(val(rs.Fields("id").value))
            result = val(rs.Fields("total").value) - ActualTotal
            Dim Total As Double
            Total = IIf(val(rs.Fields("total").value) = 0, 1, val(rs.Fields("total").value))
            resultpercentage = Round(ActualTotal / Total * 100, 2)
 
            If val(rs.Fields("total").value) > ActualTotal Then
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
            
                .TextMatrix(i, .ColIndex("bill_date")) = IIf(IsNull(rs.Fields("bill_date").value), "", rs.Fields("bill_date").value)
                .TextMatrix(i, .ColIndex("project_no")) = IIf(IsNull(rs.Fields("project_no").value), "", rs.Fields("project_no").value)
                .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(rs.Fields("project_name").value), "", rs.Fields("project_name").value)
            
                .TextMatrix(i, .ColIndex("End_user_name")) = IIf(IsNull(rs.Fields("End_user_name").value), "", rs.Fields("End_user_name").value)
            
                .TextMatrix(i, .ColIndex("Sub_user_name")) = IIf(IsNull(rs.Fields("Sub_user_name").value), "", rs.Fields("Sub_user_name").value)
            
                .TextMatrix(i, .ColIndex("bill_to")) = IIf(IsNull(rs.Fields("bill_to").value), "", rs.Fields("bill_to").value)
 
                .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(rs.Fields("total").value), "", rs.Fields("total").value)
            
                .TextMatrix(i, .ColIndex("ActualTotal")) = ActualTotal
                .TextMatrix(i, .ColIndex("ResultPercentage")) = resultpercentage
                .TextMatrix(i, .ColIndex("Result")) = result

                If result = val(.TextMatrix(i, .ColIndex("total"))) Then
                    .Cell(flexcpBackColor, i, 12, i, 12) = vbRed
                Else
                    .Cell(flexcpBackColor, i, 12, i, 12) = vbYellow
                End If

            End If

            rs.MoveNext
        Next

        rs.Close
 
        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

Private Sub BtnCancel_Click()
    Me.Hide
End Sub

Private Sub CmdPrint_Click()
    On Error Resume Next
    Dim GrdBack As ClsBackGroundPic
    'Grid.ExtendLastCol = True
    Grid.WallPaper = Nothing
    'Grid.AutoSize  0, Grid.Cols - 1, False
    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
    'Printer.RightToLeft = True
    'Printer.Print ("Employee Salary Report")

    Me.Grid.PrintGrid " ‰»Ì…    „” Œ·’«  ·„  ”œœ »«·ﬂ«„·", True, 2, 1, 1500
End Sub

Private Sub Form_Load()
 
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    FillGridWithData

    If SystemOptions.UserInterface = EnglishInterface Then

        SetInterface Me
        cahngelang
    End If

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

    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("id")) = " Bill ID"
        .TextMatrix(0, .ColIndex("bill_date")) = "Bill Date  "
        .TextMatrix(0, .ColIndex("Project_name")) = "Project Name"
        .TextMatrix(0, .ColIndex("End_user_name")) = "End_user_name"
        .TextMatrix(0, .ColIndex("Sub_user_name")) = "Sub_user_name"
        .TextMatrix(0, .ColIndex("total")) = "Bill Total"
        .TextMatrix(0, .ColIndex("ActualTotal")) = "Payed"
        .TextMatrix(0, .ColIndex("result")) = "Variance"
        .TextMatrix(0, .ColIndex("resultpercentage")) = "Variance%"

    End With

End Function

Private Sub ImgFavorites_Click()
AddTofaforites Me.name, Me.Caption, Me.Caption


End Sub
