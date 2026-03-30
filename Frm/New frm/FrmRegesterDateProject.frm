VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRegesterDateProject 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ”ÃÌ·  «—ÌŒ  "
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3885
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2040
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   408
      Left            =   1020
      TabIndex        =   4
      Top             =   1536
      Width           =   948
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "Õ›Ÿ"
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
      Top             =   2430
      Width           =   4425
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   408
      Left            =   60
      TabIndex        =   5
      Top             =   1536
      Width           =   948
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
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   330
      Left            =   720
      TabIndex        =   6
      Top             =   240
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   582
      _Version        =   393216
      Format          =   94765057
      CurrentDate     =   38784
   End
   Begin Dynamic_Byte.NourHijriCal XPDtbBillH 
      Height          =   312
      Left            =   720
      TabIndex        =   10
      Top             =   600
      Width           =   1836
      _ExtentX        =   3228
      _ExtentY        =   556
   End
   Begin MSComCtl2.DTPicker dtpTime 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "H:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   336
      Left            =   720
      TabIndex        =   12
      Top             =   960
      Width           =   1872
      _ExtentX        =   3307
      _ExtentY        =   582
      _Version        =   393216
      Format          =   94765058
      CurrentDate     =   38784
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”«⁄…"
      Height          =   252
      Index           =   1
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   960
      Width           =   1152
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·œŒÊ· Â‹ "
      Height          =   252
      Index           =   0
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   600
      Width           =   1152
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   8
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·œŒÊ·"
      Height          =   255
      Index           =   6
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   240
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4425
      X2              =   0
      Y1              =   1320
      Y2              =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   252
      Index           =   5
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1260
      Width           =   3648
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   252
      Index           =   4
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1548
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   3
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   1545
   End
End
Attribute VB_Name = "FrmRegesterDateProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SendForm As String
'Public fg As VSFlex8UCtl.vsFlexGrid
'Public LngRow As Long
'Public LngCol As Long
Dim Rs_Temp5 As ADODB.Recordset

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
    Dim Msg As String
    Dim dateenter As Date
    Dim timEnter As Date
    Dim Askinterval As String
    
    If SendForm = "" Then
            If Not Projects.GridSub Is Nothing Then
                    ' FrmCarAuthontication.fg.TextMatrix(FrmCarAuthontication.LngRow, FrmCarAuthontication.LngCol) = XPDtbBill.value    'Trim$(Me.TxtComment.text)
                    Projects.GridSub.TextMatrix(Projects.LngRow, Projects.GridSub.ColIndex("subdate")) = XPDtbBill.value
                    Askinterval = "dd/mm/yyyy"
                    '  End If
                    
                    If Projects.GridSub.ColIndex("subdate") <> -1 Then
                           
                            Projects.GridSub.TextMatrix(Projects.LngRow, Projects.GridSub.ColIndex("subdate")) = XPDtbBill.value ' dateenter
                    End If
                    Unload Me
            End If
    ElseIf SendForm = "AttributionContract" Then
            With FrmAttributionContract.VSFlexGrid1
                    If GetDurationStart <= XPDtbBillH.value Then
                            FrmAttributionContract.VSFlexGrid1.TextMatrix(.Row, .ColIndex("Embark")) = XPDtbBill.value
                            FrmAttributionContract.VSFlexGrid1.TextMatrix(.Row, .ColIndex("EmbarkH")) = XPDtbBillH.value
                        '   TblBookingRequest2 FrmAttributionContract.MDate
                        '   TblBookingRequest2 FrmAttributionContract.Cal_AcutualWorkDays
                            Unload Me
                    Else
                            MsgBox ("·«Ì„ﬂ‰ «‰ ÌﬂÊ‰  «—ÌŒ «·„»«‘—… «ﬁ· „‰  «—ÌŒ »œ«Ì… «·”‰… «·œ—«”Ì…")
                    End If
            End With
    ElseIf SendForm = "BookingRequest" Then
            FrmBookingRequest.Grid.TextMatrix(FrmBookingRequest.Grid.Row, FrmBookingRequest.Grid.ColIndex("Date")) = XPDtbBill.value
            FrmBookingRequest.Grid.TextMatrix(FrmBookingRequest.Grid.Row, FrmBookingRequest.Grid.ColIndex("Time")) = Format(dtpTime.value, "HH:mm")
            Unload Me
        ElseIf SendForm = "BookingRequest2" Then
            FrmBookingRequest2.Grid.TextMatrix(FrmBookingRequest2.Grid.Row, FrmBookingRequest2.Grid.ColIndex("Date")) = XPDtbBill.value
            FrmBookingRequest2.Grid.TextMatrix(FrmBookingRequest2.Grid.Row, FrmBookingRequest2.Grid.ColIndex("Time")) = Format(dtpTime.value, "HH:mm")
            Unload Me
    ElseIf SendForm = "VehicleOperatorOrder" Then
         '   FrmVehicleOperatorOrder.Grid.TextMatrix(FrmVehicleOperatorOrder.Grid.Row, FrmVehicleOperatorOrder.Grid.ColIndex("Date")) = XPDtbBill.value
         '   FrmVehicleOperatorOrder.Grid.TextMatrix(FrmVehicleOperatorOrder.Grid.Row, FrmVehicleOperatorOrder.Grid.ColIndex("Time")) = Format(dtpTime.value, "HH:mm")
            Unload Me
    End If

End Sub
Private Sub ChangeLang()
    CmdCancel.Caption = "Cancel"
CmdOk.Caption = "Save"
lbl(6).Caption = "DateEnter"

Me.Caption = "Register Date "
End Sub

Private Sub Form_Load()
    CenterForm Me

    FormPostion Me, GetPostion

    Me.CmdOk.ButtonStyle = impActive
    Set CmdOk.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Save").Picture
    CmdOk.ButtonPositionImage = impRightOfText

    Me.CmdCancel.ButtonStyle = impActive
    Set CmdCancel.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").Picture
    CmdCancel.ButtonPositionImage = impRightOfText
    XPDtbBill.value = Date

If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub XPDtbBill_Change()
  XPDtbBillH.value = ToHijriDate(XPDtbBill.value)
End Sub


Private Sub XPDtbBillH_LostFocus()
   VBA.Calendar = vbCalGreg
            XPDtbBill.value = ToGregorianDate(XPDtbBillH.value)
End Sub



Public Function GetDurationStart() As String
    
    Dim i  As Integer, str As String
    i = val(FrmAttributionContract.dcDuration.BoundText)
    str = " select * from tbldurations where id =  " & i
    Set Rs_Temp5 = New ADODB.Recordset
   Rs_Temp5.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
   If Rs_Temp5.RecordCount > 0 Then
                GetDurationStart = IIf(IsNull(Rs_Temp5("fromdateh").value), "", Rs_Temp5("fromdateh").value)
   End If
       
    
End Function


