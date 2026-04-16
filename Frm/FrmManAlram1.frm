VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmManAlram 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "≈—”«·  ‰»ÌÂ » Ã„Ì⁄ ÃÂ«“ ≈·Ï ﬁ”„ «·’Ì«‰…"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6510
   Icon            =   "FrmManAlram.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox InvSerial 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   4110
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   1230
      Width           =   1365
   End
   Begin VB.TextBox TxtInvID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3570
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSDataListLib.DataCombo DcboUsers 
      Height          =   315
      Left            =   3000
      TabIndex        =   12
      Top             =   4170
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.ComboBox CboPriority 
      Height          =   315
      Left            =   4080
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2010
      Width           =   1395
   End
   Begin VB.TextBox TxtMsg 
      Alignment       =   1  'Right Justify
      Height          =   1155
      Left            =   90
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2820
      Width           =   6345
   End
   Begin MSComCtl2.DTPicker Dtp 
      Height          =   345
      Left            =   4110
      TabIndex        =   5
      Top             =   1620
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      _Version        =   393216
      Format          =   63438849
      CurrentDate     =   39362
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„⁄·Ê„«  ⁄‰ ›« Ê—… «·»Ì⁄"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1425
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   810
      Width           =   3495
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Index           =   11
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   930
         Width           =   2415
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·›« Ê—…"
         Height          =   285
         Index           =   10
         Left            =   2490
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   930
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Index           =   9
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   630
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Index           =   8
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   330
         Width           =   2415
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·⁄„Ì·:"
         Height          =   285
         Index           =   7
         Left            =   2490
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   630
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·⁄„Ì·:"
         Height          =   285
         Index           =   6
         Left            =   2490
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   330
         Width           =   975
      End
   End
   Begin VB.TextBox TxtID 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   4110
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   1365
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   675
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6510
      _cx             =   11483
      _cy             =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "≈—”«·  ‰»ÌÂ » Ã„Ì⁄ ÃÂ«“ ≈·Ï ﬁ”„ «·’Ì«‰…"
      Align           =   1
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   7
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin ImpulseButton.ISButton XPBtnOK 
      Default         =   -1  'True
      Height          =   345
      Left            =   990
      TabIndex        =   13
      Top             =   4740
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„Ê«›ﬁ"
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
      ButtonImage     =   "FrmManAlram.frx":038A
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton XPBtnCancel 
      Height          =   345
      Left            =   90
      TabIndex        =   14
      Top             =   4740
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   609
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
      ButtonImage     =   "FrmManAlram.frx":0724
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
   Begin ImpulseButton.ISButton CmdSearchTrans 
      Height          =   315
      Left            =   3570
      TabIndex        =   15
      Top             =   1230
      Width           =   495
      _ExtentX        =   873
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
      ButtonImage     =   "FrmManAlram.frx":0ABE
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton XpNew 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   2040
      TabIndex        =   24
      Top             =   4800
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÃœÌœ"
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
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„” Œœ„ «·„Õ—— ·· ‰»ÌÂ"
      Height          =   255
      Index           =   5
      Left            =   4830
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   4200
      Width           =   1605
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   30
      X2              =   6450
      Y1              =   4590
      Y2              =   4590
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—”«·… ≈·Ï ﬁ”„ «·’Ì«‰…"
      Height          =   255
      Index           =   4
      Left            =   4710
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "œ—Ã… «· ‰»ÌÂ"
      Height          =   315
      Index           =   3
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2010
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   345
      Index           =   2
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1620
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·›« Ê—…"
      Height          =   345
      Index           =   1
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1230
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «· ‰»ÌÂ"
      Height          =   345
      Index           =   0
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   945
   End
End
Attribute VB_Name = "FrmManAlram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSearchTrans_Click()
Load FrmBuySearch
FrmBuySearch.DealingForm = InvoiceTransaction
Set FrmBuySearch.ExtraRetrunObject = Me.TxtInvID
FrmBuySearch.CboPayMentType.Enabled = True
FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄"
FrmBuySearch.Show 1
End Sub

Private Sub Form_Load()
Dim Dcombos As ClsDataCombos
CenterForm Me
Set Dcombos = New ClsDataCombos
Dcombos.GetUsers Me.DcboUsers
Me.DcboUsers.BoundText = User_ID
SetDtpickerDate Me.Dtp
With Me.CboPriority
    .Clear
    .AddItem "Â«„ Ãœ«"
    .ItemData(.NewIndex) = 3
    .AddItem "⁄«œÏ"
    .ItemData(.NewIndex) = 2
    .AddItem "÷⁄Ì›"
    .ItemData(.NewIndex) = 1
    .ListIndex = 1
End With
End Sub

Private Sub InvSerial_KeyDown(KeyCode As Integer, Shift As Integer)
''Dim Rs As ADODB.Recordset
''Dim StrSQL As String
'If KeyCode = vbKeyReturn Then
'    MsgBox ""
'    If Trim$(Me.InvSerial.text) = "" Then
'        Exit Sub
'    Else
'        PutTrans
'    End If
'End If
End Sub

Private Sub InvSerial_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    If Trim$(Me.InvSerial.text) = "" Then
        Exit Sub
    Else
        PutTrans
    End If
End If
End Sub
Private Function PutTrans() As Boolean
Dim StrTemp As String
Dim Msg As String
If Trim(Me.InvSerial.text) = "" Then
    Me.TxtInvID.text = ""
Else
    StrTemp = GetTransIDSerial(0, , Trim(Me.InvSerial.text), 2)
    If StrTemp = "" Then
        Msg = "·« ÊÃœ ›« Ê—… »Â–« «·—ﬁ„ ... ≈” Œœ„ ‘«‘… «·»ÕÀ.!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        PutTrans = False
    Else
        If Trim$(Me.TxtInvID.text) <> StrTemp Then
            Me.TxtInvID.text = StrTemp
        End If
        PutTrans = True
    End If
End If
End Function

Private Sub TxtInvID_Change()
Dim Rs As ADODB.Recordset
Dim StrSQL As String

If Trim$(Me.TxtInvID.text) = "" Then
    Me.InvSerial.text = ""
    Me.lbl(8).Caption = ""
    Me.lbl(11).Caption = ""
Else
    Set Rs = New ADODB.Recordset
    StrSQL = "SELECT  dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial," & _
    "dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusName,dbo.QryOneTransactionTotal(" & _
    "dbo.Transactions.Transaction_ID) AS TotalX "
    StrSQL = StrSQL + " FROM dbo.Transactions INNER JOIN dbo.TblCustemers ON " & _
    "dbo.Transactions.CusID = dbo.TblCustemers.CusID "
    StrSQL = StrSQL + " Where dbo.Transactions.Transaction_ID=" & Val(Me.TxtInvID.text)
    Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (Rs.BOF Or Rs.EOF) Then
        Me.InvSerial.text = IIf(IsNull(Rs("Transaction_Serial").Value), "", Rs("Transaction_Serial").Value)
        Me.lbl(8).Caption = IIf(IsNull(Rs("CusName").Value), "", Rs("CusName").Value)
        Me.lbl(11).Caption = IIf(IsNull(Rs("TotalX").Value), "", Rs("TotalX").Value)
    Else
        Me.InvSerial.text = ""
        Me.lbl(8).Caption = ""
        Me.lbl(11).Caption = ""
    End If
    Rs.Close
    Set Rs = Nothing
    
End If
End Sub

Private Sub XPBtnCancel_Click()
Unload Me


End Sub
Private Sub XPBtnOK_Click()
Dim Rs As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim StrSQL As String
Dim Msg As String

StrSQL = "SELECT dbo.TblManAlram.TableID, dbo.TblManAlram.TransID, dbo.TblManAlram.AlramDate," & _
"dbo.TblManAlram.AlramPriority, dbo.TblManAlram.AlramMsg,dbo.TblManAlram.State," & _
"dbo.Transactions.Transaction_Serial, dbo.TblUsers.UserName "
StrSQL = StrSQL + " FROM dbo.TblManAlram INNER JOIN dbo.Transactions ON dbo.TblManAlram.TransID" & _
"= dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUsers ON dbo.TblManAlram.UserID =" & _
"dbo.TblUsers.UserID "
StrSQL = StrSQL + " Where  dbo.TblManAlram.TransID=" & Val(Me.TxtInvID.text) & ""
Set RsTemp = New ADODB.Recordset
RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Not (RsTemp.BOF Or RsTemp.EOF) Then
    Msg = "⁄›Ê« ... "
    Msg = Msg & Chr(13) & "«·›« Ê—… —ﬁ„ " & RsTemp("Transaction_Serial").Value
    Msg = Msg & Chr(13) & "”Ã· ·Â«  ‰»ÌÂ ›Ï ÌÊ„ " & DisplayDate(RsTemp("AlramDate").Value)
    Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  ﬂ—«—  ”ÃÌ· «· ‰»ÌÂ....!!!"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub
End If
Set Rs = New ADODB.Recordset
Rs.Open "TblManAlram", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
Rs.AddNew
    Rs("TransID").Value = Val(Me.TxtInvID.text)
    Rs("AlramDate").Value = Me.Dtp.Value
    Rs("AlramPriority").Value = Me.CboPriority.ItemData(Me.CboPriority.ListIndex)
    Rs("AlramMsg").Value = Trim$(Me.TxtMsg.text)
    Rs("UserID").Value = Me.DcboUsers.BoundText
    Rs("State").Value = 0
Rs.update
Msg = " „ ≈—”«· «· ‰»ÌÂ.."
MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

End Sub

Private Sub XpNew_Click()
 If DoPremis(Do_New, Me.name, True) = False Then
            Exit Sub
        End If
'        TxtModFlg.text = "N"
        clear_all Me
        TxtID.text = CStr(new_id("TblManAlram", "TableID", "", True))
        Me.DcboUsers.BoundText = User_ID
        'TxtPaymentCounts.text = 1
        TxtID.SetFocus

End Sub


