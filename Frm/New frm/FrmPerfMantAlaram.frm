VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPerfMantAlaram 
   Caption         =   " ‰»ÌÂ«  «·’Ì«‰… «·Êﬁ«∆Ì… ·Â–« «·ÌÊ„"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   20250
   Icon            =   "FrmPerfMantAlaram.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.CheckBox Check17 
      Alignment       =   1  'Right Justify
      Caption         =   " ÕœÌœ «·ﬂ·"
      Height          =   375
      Left            =   18840
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   600
      Width           =   1185
   End
   Begin VB.TextBox txtMessage 
      Alignment       =   1  'Right Justify
      Height          =   750
      Left            =   3615
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Text            =   "FrmPerfMantAlaram.frx":058A
      Top             =   8520
      Width           =   5205
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   13440
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   720
      Width           =   1575
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   -15
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   30
      Width           =   20265
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   2
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   3
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
            TabIndex        =   4
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
               Picture         =   "FrmPerfMantAlaram.frx":05B6
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPerfMantAlaram.frx":0950
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPerfMantAlaram.frx":0CEA
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPerfMantAlaram.frx":1084
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPerfMantAlaram.frx":141E
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPerfMantAlaram.frx":17B8
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPerfMantAlaram.frx":1B52
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPerfMantAlaram.frx":20EC
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ‰»ÌÂ«  «·’Ì«‰… «·Êﬁ«∆Ì… ·Â–« «·ÌÊ„"
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
         Left            =   15360
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   120
         Width           =   4320
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   7035
      Left            =   0
      TabIndex        =   0
      Top             =   1170
      Width           =   20235
      _cx             =   35692
      _cy             =   12409
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
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmPerfMantAlaram.frx":2486
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
   Begin ImpulseButton.ISButton btnCancel 
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   4680
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
      ButtonImage     =   "FrmPerfMantAlaram.frx":26C5
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton CmdPrint 
      Height          =   390
      Left            =   120
      TabIndex        =   9
      Top             =   8760
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
      ButtonImage     =   "FrmPerfMantAlaram.frx":2A5F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin MSComCtl2.DTPicker dbDate 
      Height          =   345
      Left            =   16440
      TabIndex        =   10
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      Format          =   94699521
      CurrentDate     =   41275
   End
   Begin ImpulseButton.ISButton SendMessage 
      Height          =   585
      Left            =   2160
      TabIndex        =   15
      Top             =   8700
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1032
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "«—”«·"
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
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—”«·…"
      Height          =   510
      Left            =   9060
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   6960
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·”Ì—Ì«·"
      Height          =   195
      Index           =   0
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«· «—ÌŒ"
      Height          =   195
      Index           =   3
      Left            =   17880
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   720
      Width           =   990
   End
End
Attribute VB_Name = "FrmPerfMantAlaram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub FillGridWithData(Optional SerialNO As String = "")

    'On Error GoTo ErrTrap

    Dim i As Integer
    Dim X As Integer
    Dim rs As ADODB.Recordset

    With Me.Grid
        .Rows = 1
    End With

    Dim ActualTotal As Double
    Dim result As Double
    Dim resultpercentage As Double
    Dim sql As String
    On Error Resume Next
    sql = "SELECT   dbo.TblCustemers.Cus_mobile,   dbo.TBLRegularMaint.DateOfRegularMaint, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.Transactions.Transaction_Date, "
    sql = sql & "  dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TBLRegularMaint.GranteeType, dbo.TBLRegularMaint.GranteeStartDate,"
    sql = sql & " dbo.TBLRegularMaint.GranteeEndDate, dbo.TBLRegularMaint.Done, dbo.TBLRegularMaint.DoneDate, dbo.TBLRegularMaint.ItemSerial,"
    sql = sql & " dbo.Transactions.NoteSerial1"
    sql = sql & " FROM         dbo.TBLRegularMaint INNER JOIN"
    sql = sql & " dbo.TblItems ON dbo.TBLRegularMaint.itemid = dbo.TblItems.ItemID INNER JOIN"
    sql = sql & "  dbo.Transactions ON dbo.TBLRegularMaint.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
    sql = sql & "     dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"

    If SerialNO = "" Then
        sql = sql & "  WHERE     (dbo.TBLRegularMaint.DateOfRegularMaint =" & SQLDate(dbDate.value, True) & " )"
    Else
        sql = sql & "  WHERE   ItemSerial='" & SerialNO & "'"
    End If
 
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
 
        Exit Sub
    End If

    i = 0

    With Me.Grid
        .Rows = 1
        .Clear flexClearScrollable
      
        rs.MoveFirst

        For X = 1 To rs.RecordCount
 
            If val(rs.Fields("total").value) < ActualTotal Then
                i = i + 1
                .Rows = .Rows + 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs.Fields("NoteSerial1").value), "", rs.Fields("NoteSerial1").value)
            
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(rs.Fields("Transaction_Date").value), "", rs.Fields("Transaction_Date").value)
            
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs.Fields("ItemCode").value), "", rs.Fields("ItemCode").value)
                .TextMatrix(i, .ColIndex("Cus_mobile")) = IIf(IsNull(rs.Fields("Cus_mobile").value), "", rs.Fields("Cus_mobile").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs.Fields("ItemName").value), "", rs.Fields("ItemName").value)
                Else
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs.Fields("ItemNamee").value), "", rs.Fields("ItemNamee").value)
                End If
            
                .TextMatrix(i, .ColIndex("ItemSerial")) = IIf(IsNull(rs.Fields("ItemSerial").value), "", rs.Fields("ItemSerial").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value)
                Else
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value)
            
                End If
                     
                '       .Cell(flexcpBackColor, I, 8, I, 8) = vbRed
                Dim GranteeType As String

                If SystemOptions.UserInterface = ArabicInterface Then
            
                    If val(.TextMatrix(i, .ColIndex("GranteeType"))) = 0 Then
                        GranteeType = "»œÊ‰ ﬁÿ⁄ "
                    Else
                        GranteeType = "„⁄ ﬁÿ⁄"
                    End If
            
                Else
                        
                    If val(.TextMatrix(i, .ColIndex("GranteeType"))) = 0 Then
                        GranteeType = "Without Spare Parts"
                    Else
                        GranteeType = "With Spare Parts"
                    End If
            
                End If
                  
                .TextMatrix(i, .ColIndex("GranteeType")) = GranteeType
             
                .TextMatrix(i, .ColIndex("GranteeStartDate")) = IIf(IsNull(rs.Fields("GranteeStartDate").value), "", rs.Fields("GranteeStartDate").value)
            
                .TextMatrix(i, .ColIndex("GranteeEndDate")) = IIf(IsNull(rs.Fields("GranteeEndDate").value), "", rs.Fields("GranteeEndDate").value)
            
                .TextMatrix(i, .ColIndex("DateOfRegularMaint")) = IIf(IsNull(rs.Fields("DateOfRegularMaint").value), "", rs.Fields("DateOfRegularMaint").value)

            End If

            rs.MoveNext
        Next

        rs.Close
 
        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

Private Sub BtnCancel_Click()
    Unload Me
End Sub

 
Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.Grid
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Check")) = True
            Next i

        End With

    Else

        With Me.Grid

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Check")) = False
            Next i

        End With

    End If

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

    Me.Grid.PrintGrid " ‰»Ì… »«·»‰Êœ «· Ì  ⁄œ  «·ﬁÌ„… «·„Œ’’… ·Â«", True, 2, 1, 1500
End Sub

Function cahngelang()
    Label1(3).Caption = "Date"
    Label1(0).Caption = "Serial No"

    Label1(2).Caption = "Preventive maintenance alerts for this day"
    Me.Caption = Label1(2).Caption
    CmdPrint.Caption = "Print"
    btnCancel.Caption = "Cancel"

    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "Transaction#"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Transaction Date"
        .TextMatrix(0, .ColIndex("ItemCode")) = "ItemCode"
        .TextMatrix(0, .ColIndex("ItemName")) = "ItemName"
        .TextMatrix(0, .ColIndex("ItemSerial")) = "Item Serial"
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
        .TextMatrix(0, .ColIndex("GranteeType")) = " Guarantee Type"
        .TextMatrix(0, .ColIndex("GranteeStartDate")) = "Guarantee Start Date"
        .TextMatrix(0, .ColIndex("GranteeEndDate")) = "Guarantee End Date"
        .TextMatrix(0, .ColIndex("DateOfRegularMaint")) = "Date Of Preventive maintenance"
    End With

End Function

Private Sub DbDAte_Change()
    FillGridWithData
End Sub

Private Sub Form_Load()
    dbDate.value = Date
    FillGridWithData
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        cahngelang
    End If

End Sub

Private Sub SendMessage_Click()
 Dim Numbers As String
    Dim RowNum As Integer
    Dim Opt As Integer
    Dim CurrentMessage As String
    Numbers = ""

    With Grid

        For RowNum = .FixedRows To .Rows - 1
    
            If .Cell(flexcpChecked, RowNum, .ColIndex("Check")) = flexChecked Then

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
      
        CurrentMessage = txtMessage.text

        If Numbers = "" Then Exit Sub
        SMSSeTTings.SendMessage CurrentMessage, Numbers
        SMSSeTTings.Hide
                                    
    End With

End Sub

Private Sub Text1_Change()
    FillGridWithData (Text1.text)
End Sub

