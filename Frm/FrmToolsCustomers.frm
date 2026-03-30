VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmToolsCustomers 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ÿ»ÿ ›Ê« Ì— «·⁄„·«¡"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   Icon            =   "FrmToolsCustomers.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   6705
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ–› Â–« «·⁄„Ì· »⁄œ ⁄„· «· ÕœÌÀ"
      Height          =   315
      Left            =   1020
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2970
      Width           =   2835
   End
   Begin MSDataListLib.DataCombo DcboCustomer 
      Height          =   315
      Left            =   990
      TabIndex        =   1
      Top             =   1350
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboCustomer2 
      Height          =   315
      Left            =   990
      TabIndex        =   2
      Top             =   2580
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton ISBXPBtnOK 
      Height          =   375
      Left            =   990
      TabIndex        =   5
      Top             =   4080
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
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
      BackStyle       =   0
      ButtonImage     =   "FrmToolsCustomers.frx":038A
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
   Begin ImpulseButton.ISButton ISBXPBtnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   4080
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈·€«¡"
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
      ButtonImage     =   "FrmToolsCustomers.frx":0724
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
   Begin ImpulseButton.ISButton CmdCusSearch 
      Height          =   315
      Index           =   0
      Left            =   540
      TabIndex        =   8
      Top             =   1350
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   556
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "..."
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
      ButtonImage     =   "FrmToolsCustomers.frx":0ABE
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton CmdCusSearch 
      Height          =   315
      Index           =   1
      Left            =   540
      TabIndex        =   9
      Top             =   2610
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   556
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "..."
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
      ButtonImage     =   "FrmToolsCustomers.frx":0E58
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Image Img 
      Height          =   480
      Left            =   5700
      Picture         =   "FrmToolsCustomers.frx":11F2
      Top             =   360
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6735
      X2              =   60
      Y1              =   3840
      Y2              =   3855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·⁄„Ì· «·„ﬂ—— √Ê «·„‘«»Â…"
      Height          =   405
      Index           =   2
      Left            =   5220
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2580
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·⁄„Ì· «·√’·Ï"
      Height          =   315
      Index           =   1
      Left            =   5220
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1350
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Index           =   0
      Left            =   1140
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   4485
   End
End
Attribute VB_Name = "FrmToolsCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TTPD As clstooltipdemand

Private Sub CmdCusSearch_Click(Index As Integer)

    Select Case Index

        Case 0
            Load FrmCustemerSearch
            FrmCustemerSearch.SearchType = 1
            FrmCustemerSearch.RetrunType = 1
            Set FrmCustemerSearch.DcboCustomers = Me.DcboCustomer
            FrmCustemerSearch.Show vbModal

        Case 1
            Load FrmCustemerSearch
            FrmCustemerSearch.SearchType = 1
            FrmCustemerSearch.RetrunType = 1
            Set FrmCustemerSearch.DcboCustomers = Me.DcboCustomer2
            FrmCustemerSearch.Show vbModal
    End Select

End Sub

Private Sub Form_Load()
    Dim Msg As String
    Dim Dcombos As ClsDataCombos
    Set TTPD = New clstooltipdemand
    Msg = "ÌÃ» ≈” Œœ«„ Â–Â «·√œ«… »⁄‰«Ì… ÊÕ—’ .. ÕÌÀ «‰Â« „„ﬂ‰ «‰  ﬂÊ‰ Œÿ—… Ãœ« "
    Msg = Msg & " Ê«‰  ÌÃ» «‰  ” Œœ„Â« ›Ï Õ«·… ÊÃÊœ  ﬂ—«— ›Ï «”„ «Ï ⁄„Ì· «Ê „Ê—œ "
    Msg = Msg & " ÊÊÃÊœ ›Ê« Ì— «Ê Õ—ﬂ«  „”Ã·… „⁄ «ﬂÀ— „‰ ⁄„Ì· ÊÂÊ ›Ï «·√’· ⁄„Ì· Ê«Õœ "
    Me.lbl(0).Caption = Msg
    CenterForm Me

    FormPostion Me, GetPostion
    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DcboCustomer, False
    Dcombos.GetCustomersSuppliers 0, Me.DcboCustomer2, False

    Set TTPD.m_From = Me
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub ISBXPBtnCancel_Click()
    Unload Me
End Sub
 
Private Sub ISBXPBtnOK_Click()
    Dim Msg As String
    Dim StrSQL As String
    Dim IntRes As VbMsgBoxResult
    Dim rs As ADODB.Recordset
    Dim StrTemp As String
    Dim i As Integer

    If Me.DcboCustomer.BoundText = "" Then
        Msg = "ÌÃ» ≈Œ Ì«— «·⁄„Ì· «·√’·Ï ..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.DcboCustomer.SetFocus
        ShowToolTips 0
        Exit Sub
    End If

    If Me.DcboCustomer2.BoundText = "" Then
        Msg = "ÌÃ» ≈Œ Ì«— «”„ «·⁄„Ì· «·„‘«»Â…  ..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.DcboCustomer2.SetFocus
        ShowToolTips 1
        Exit Sub
    End If

    If Me.DcboCustomer.BoundText = Me.DcboCustomer2.BoundText Then
        Msg = "ÌÃ» ≈Œ Ì«— ⁄„Ì· €Ì— «·⁄„Ì· «·√’·Ï"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.DcboCustomer.SetFocus
        Exit Sub
    End If

    Msg = "ÌÃ» „·«ÕŸ… «‰ Â–Â «·⁄„·Ì… ·«Ì„ﬂ‰ «· —«Ã⁄ ›ÌÂ«.."
    Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«—"
    IntRes = MsgBox(Msg, vbQuestion + vbDefaultButton2 + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo, App.Title)

    If IntRes = vbNo Then
        Exit Sub
    End If

    StrSQL = "SELECT COUNT(Transactions.Transaction_ID) AS TransCount, TransactionTypes.TransactionTypeName"
    StrSQL = StrSQL + " FROM Transactions INNER JOIN TransactionTypes ON Transactions.Transaction_Type = "
    StrSQL = StrSQL + " TransactionTypes.Transaction_Type "
    StrSQL = StrSQL + " Where Transactions.CusID=" & Me.DcboCustomer.BoundText & ""
    StrSQL = StrSQL + " GROUP BY TransactionTypes.TransactionTypeName"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        StrTemp = ""
        StrTemp = "‰Ê⁄ «·Õ—ﬂ…:-" & vbTab & vbTab & vbTab & "⁄œœÂ«:-"

        For i = 1 To rs.RecordCount
            StrTemp = StrTemp & Chr(13) & Chr(10)
            StrTemp = StrTemp & rs("TransactionTypeName").value & vbTab & rs("TransCount").value
            rs.MoveNext
        Next i

        Msg = "Â–Â «·Õ—ﬂ«  ÂÏ «· Ï ”Ê› Ì „  ÕœÌÀÂ« «Ê ‰ﬁ·Â«  "
        Msg = Msg & Chr(13) & " „‰ «·⁄„Ì· :-" & Me.DcboCustomer.text
        Msg = Msg & Chr(13) & "≈·Ï «·⁄„Ì· :-" & Me.DcboCustomer2.text
        Msg = Msg & Chr(13) & StrTemp
        Msg = Msg & Chr(13) & ""
        Msg = Msg & Chr(13) & ""
        Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ›Ï ⁄·„Ì… «· ÕœÌÀ ..øø"
        IntRes = MsgBox(Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbQuestion + vbYesNo, App.Title)

        If IntRes = vbNo Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If

        StrSQL = "Update Transactions Set CusID=" & Me.DcboCustomer2.BoundText & ""
        StrSQL = StrSQL + " Where CusID=" & Me.DcboCustomer.BoundText
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Update NOTES Set CusID=" & Me.DcboCustomer2.BoundText & ""
        StrSQL = StrSQL + " Where CusID=" & Me.DcboCustomer.BoundText
        Cn.Execute StrSQL, , adExecuteNoRecords
        Msg = " „  ⁄„·Ì… «· ÕœÌÀ »‰Ã«Œ...!!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Else
        Msg = "·« ÊÃœ «Ï ›Ê« Ì— «Ê Õ—ﬂ«   Ã«—Ì… „— »ÿ… »Â–« «·⁄„Ì· "
        Msg = Msg & Chr(13) & Me.DcboCustomer.text
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

End Sub

Private Sub ShowToolTips(intIndex As Integer)
    TTPD.Destroy
    TTPD.Style = TTBalloon
    TTPD.Icon = TTIconWarning
    TTPD.Centered = True
    TTPD.RightToLeft = True
    TTPD.Title = "≈Œ Ì«— «”„ «·⁄„Ì· «·√’·Ï"
    TTPD.TipText = "ÌÃ» ≈ŒÌ «— «”„ «·⁄„Ì· «Ê «·„Ê—œ «·√’·Ï..." & Chr(13) & "«·–Ï ”Ê› Ì „ —»ÿ ﬂ· «·›Ê« Ì— «Ê «·Õ—ﬂ«  »Â"
    TTPD.PopupOnDemand = True
    TTPD.VisibleTime = 5000

    Select Case intIndex

        Case 0
            TTPD.Title = "≈Œ Ì«— «”„ «·⁄„Ì· «·√’·Ï"
            TTPD.TipText = "ÌÃ» ≈ŒÌ «— «”„ «·⁄„Ì· «Ê «·„Ê—œ «·√’·Ï..." & Chr(13) & "«·–Ï ”Ê› Ì „ —»ÿ ﬂ· «·›Ê« Ì— «Ê «·Õ—ﬂ«  »Â"
            TTPD.CreateToolTip Me.DcboCustomer.hWnd
            TTPD.Show (DcboCustomer.Width / Screen.TwipsPerPixelY), (DcboCustomer.Height / Screen.TwipsPerPixelX - 1)

        Case 1
            TTPD.Title = "≈Œ Ì«— «”„ «·⁄„Ì· «·„‘«»Â…"
            TTPD.TipText = "ÌÃ» ≈ŒÌ «— «”„ «·⁄„Ì· «Ê «·„Ê—œ «·„‘«»Â…..." & Chr(13) & "«·–Ï ”Ê› Ì „ ‰ﬁ· «Ê  ⁄œÌ·  ﬂ· «·›Ê« Ì— «Ê «·Õ—ﬂ«  «·Œ«’… »Â"

            TTPD.CreateToolTip Me.DcboCustomer2.hWnd
            TTPD.Show (DcboCustomer2.Width / Screen.TwipsPerPixelY), (DcboCustomer2.Height / Screen.TwipsPerPixelX - 1)
    End Select

End Sub
