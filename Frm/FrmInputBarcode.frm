VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInputBarcode 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2805
   ScaleWidth      =   3750
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
   Begin VB.TextBox txtNo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Text            =   "1"
      Top             =   600
      Width           =   2385
   End
   Begin VB.TextBox TxtName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Text            =   "barcode1"
      Top             =   240
      Width           =   2385
   End
   Begin ImpulseButton.ISButton CmdOk 
      Default         =   -1  'True
      Height          =   405
      Left            =   1020
      TabIndex        =   4
      Top             =   2250
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "ÿ»«⁄Â"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtComment 
      Alignment       =   1  'Right Justify
      Height          =   975
      Left            =   30
      MaxLength       =   255
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3630
      Width           =   4425
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   405
      Left            =   60
      TabIndex        =   5
      Top             =   2250
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "«·€«¡"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DcbUnit 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker ProductionDate 
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   104857601
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker ExpiryDate 
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   104857601
      CurrentDate     =   38784
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·«‰ Â«¡"
      Height          =   255
      Index           =   9
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·«‰ «Ã"
      Height          =   255
      Index           =   2
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ÊÕœ…"
      Height          =   255
      Index           =   1
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label lblindex 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄œœ «·«” Ìﬂ—« "
      Height          =   255
      Index           =   0
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   8
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   7
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "k‰„Ê–Ã «·»«—ﬂÊœ"
      Height          =   255
      Index           =   6
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   3720
      X2              =   0
      Y1              =   2160
      Y2              =   2175
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   5
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   660
      Width           =   3645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   4
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   3
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   300
      Width           =   1545
   End
End
Attribute VB_Name = "FrmInputBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public fg As VSFlex8UCtl.vsFlexGrid

'Public LngRow As Long

'Public LngCol As Long

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
Dim LongRow As Double
Dim Price As Double
Dim VatYou As Double
Dim Vat As Double
Dim Total As Double
If TxtName.Text = "" Then
MsgBox "Ì—ÃÏ «œŒ«· «”„ «·»«—ﬂÊœ"
Exit Sub
End If
If Me.txtNo.Text = "" Then
MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ «·«” Ìﬂ—« "
Exit Sub
End If
Price = GetUnitSalesPrice(val(FrmItems.XPTxtID.Text), val(Me.DcbUnit.BoundText))
VatYou = PercentgValueAddedBarcode(Date, val(FrmItems.XPTxtID.Text), 21)
If VatYou <> 0 Then
Vat = Price * VatYou / 100
Else
Vat = 0
End If
Total = Price + Vat

If Me.lblindex = 1 Then
FrmItems.PrintBarCode val(txtNo.Text), TxtName.Text, FrmItems.TxtbarCodeNO.Text, Price, , lblindex, VatYou, Vat, Total, ProductionDate.value, ExpiryDate.value
Else
LongRow = FrmItems.LngRow
If FrmItems.GridItemsDetails2.TextMatrix(LongRow, FrmItems.GridItemsDetails2.ColIndex("ParrtNoCode")) <> "" Then
FrmItems.PrintBarCode val(txtNo.Text), TxtName.Text, FrmItems.GridItemsDetails2.TextMatrix(LongRow, FrmItems.GridItemsDetails2.ColIndex("ParrtNoCode")), Price, FrmItems.GridItemsDetails2.TextMatrix(LongRow, FrmItems.GridItemsDetails2.ColIndex("ItemDetailedCode")), lblindex
Else
MsgBox "·«Ì„ﬂ‰ «·ÿ»«⁄Â ·⁄œ„ ÊÃÊœ «·»«—ﬂÊœ"
Exit Sub
End If
End If
        Unload Me
'    End If

End Sub
Private Sub ChangeLang()
    CmdCancel.Caption = "Cancel"
CmdOk.Caption = "Save"
lbl(6).Caption = "DateEnter"
lbl(0).Caption = "Number of sticker"
lbl(1).Caption = "Unit"
lbl(2).Caption = "Production Date"
lbl(9).Caption = "Expiry date"
CmdOk.Caption = "Print"
CmdCancel.Caption = "Cancel"

Me.Caption = "Register BarCode "
End Sub

Private Sub Form_Load()
    CenterForm Me
    Dim My_SQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "  SELECT     dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName"
    Else
    My_SQL = "  SELECT     dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitNamee"
        ChangeLang
    End If
    My_SQL = My_SQL & "   FROM         dbo.TblItemsUnits LEFT OUTER JOIN"
    My_SQL = My_SQL & "                  dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
     My_SQL = My_SQL & "    WHERE     dbo.TblItemsUnits.ItemID = " & val(FrmItems.XPTxtID.Text)
    fill_combo DcbUnit, My_SQL
    Me.DcbUnit.BoundText = GetDefultUnit(val(FrmItems.XPTxtID.Text))
    
    FormPostion Me, GetPostion
    ProductionDate.value = Date
ProductionDate_Change
End Sub
Function GetDefultUnit(Optional ItemID As Double) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     UnitID"
sql = sql & " From dbo.TblItemsUnits"
sql = sql & " Where (ItemID = " & ItemID & ") And (DefaultUnit = 1)"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetDefultUnit = IIf(IsNull(Rs3("UnitID").value), 0, Rs3("UnitID").value)
Else
GetDefultUnit = 0
End If
End Function
Function GetUnitSalesPrice(Optional ItemID As Double, Optional UnitID As Double) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     UnitSalesPrice"
sql = sql & " From dbo.TblItemsUnits"
sql = sql & " Where (ItemID = " & ItemID & ") And (UnitID = " & UnitID & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetUnitSalesPrice = IIf(IsNull(Rs3("UnitSalesPrice").value), 0, Rs3("UnitSalesPrice").value)
Else
GetUnitSalesPrice = 0
End If
End Function

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub ProductionDate_Change()
 Dim Item_ID As Long
    Dim GroupID As Double
    Dim ExpiryValue As Integer
    Dim ExpiryType As Integer
    Dim Askinterval As String
Item_ID = val(FrmItems.XPTxtID.Text)

GetItemIDExpiry Item_ID, ExpiryType, ExpiryValue

      If ExpiryType = 0 Then
            Askinterval = "D"
        ElseIf ExpiryType = 1 Then
            Askinterval = "M"
        ElseIf ExpiryType = 2 Then
            Askinterval = "YYYY"
        End If
'
            ExpiryDate.value = DateAdd(Askinterval, val(ExpiryValue), ProductionDate)


  
End Sub
