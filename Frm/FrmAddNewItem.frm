VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmAddNewItem 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈÷«›… ’‰› ÃœÌœ"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   Icon            =   "FrmAddNewItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   4425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.ComboBox CboItemType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1890
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1860
      Width           =   1200
   End
   Begin VB.TextBox TxtGuarValue 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   1470
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3780
      Width           =   570
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "√”⁄«— «·’‰›"
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
      Height          =   1275
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2220
      Width           =   4275
      Begin VB.TextBox TxtDealerPrice 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   900
         Width           =   1005
      End
      Begin VB.TextBox TxtCusPrice 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   540
         Width           =   735
      End
      Begin VB.TextBox TxtPurPrice 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   210
         Width           =   2925
      End
      Begin VB.TextBox TxtSalesPrice 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”⁄— «·»Ì⁄(„” Â·ﬂ)"
         Height          =   405
         Index           =   9
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”⁄— «·»Ì⁄(⁄„Ì·)"
         Height          =   315
         Index           =   8
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   570
         Width           =   1155
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”⁄— «·‘—«¡"
         Height          =   315
         Index           =   5
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”⁄— «·»Ì⁄"
         Height          =   255
         Index           =   6
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   930
         Width           =   1185
      End
   End
   Begin VB.CheckBox ChkGuar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·Â ÷„«‰"
      Height          =   285
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3840
      Width           =   1125
   End
   Begin VB.TextBox TxtReQty 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1860
      Width           =   615
   End
   Begin MSDataListLib.DataCombo DcboGroup 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   450
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.TextBox XPTxtCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1155
      Width           =   2955
   End
   Begin VB.TextBox XPTxtName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1515
      Width           =   2955
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   2925
   End
   Begin VB.CheckBox XPChkSerial 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·Â ”Ì—Ì«·"
      Height          =   285
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3525
      Width           =   1125
   End
   Begin ImpulseButton.ISButton XPButton301 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   30
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4950
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
      BackStyle       =   0
      ButtonImage     =   "FrmAddNewItem.frx":038A
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton XPBtnsave 
      Height          =   345
      Left            =   900
      TabIndex        =   13
      Top             =   4950
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ›Ÿ"
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
      ButtonImage     =   "FrmAddNewItem.frx":0724
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   345
      Index           =   0
      Left            =   30
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1410
      _cx             =   2487
      _cy             =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   5
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
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
      Begin VB.OptionButton OptGaurType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Â—"
         Height          =   225
         Index           =   0
         Left            =   690
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   60
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.OptionButton OptGaurType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÌÊ„"
         Height          =   225
         Index           =   1
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   60
         Width           =   600
      End
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·’‰›"
      Height          =   315
      Index           =   11
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   1860
      Width           =   1275
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ÿ«„ «·÷„«‰"
      Height          =   255
      Index           =   10
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   3840
      Width           =   1185
   End
   Begin VB.Image Img 
      Height          =   240
      Index           =   1
      Left            =   4170
      Picture         =   "FrmAddNewItem.frx":0ABE
      Top             =   4290
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4350
      X2              =   0
      Y1              =   4860
      Y2              =   4845
   End
   Begin VB.Label LblInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
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
      Height          =   615
      Index           =   1
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   4170
      Width           =   4095
   End
   Begin VB.Image Img 
      Height          =   240
      Index           =   0
      Left            =   3540
      Picture         =   "FrmAddNewItem.frx":1048
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ÿ«„ «·”Ì—Ì«·"
      Height          =   255
      Index           =   7
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3540
      Width           =   1185
   End
   Begin VB.Label LblInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   810
      Width           =   3405
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õœ ≈⁄«œ… «·ÿ·»"
      Height          =   315
      Index           =   4
      Left            =   780
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1860
      Width           =   1065
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·’‰›"
      Height          =   315
      Index           =   3
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1155
      Width           =   1305
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·’‰›"
      Height          =   315
      Index           =   2
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1485
      Width           =   1305
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Ã„Ê⁄…"
      Height          =   315
      Index           =   0
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   450
      Width           =   1305
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·’‰›"
      Height          =   315
      Index           =   1
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   90
      Width           =   1305
   End
End
Attribute VB_Name = "FrmAddNewItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Private m_DealingForm As GridTransType
Dim cDcboSearch As clsDCboSearch

Private Sub ChkGuar_Click()

    If ChkGuar.value = vbChecked Then
        Me.TxtGuarValue.Enabled = True
        Me.OptGaurType(0).Enabled = True
        Me.OptGaurType(1).Enabled = True
    Else
        Me.TxtGuarValue.Enabled = False
        Me.OptGaurType(0).Enabled = False
        Me.OptGaurType(1).Enabled = False
    End If

End Sub

Private Sub DcboGroup_Change()
    WriteItemCode
End Sub

Private Sub DcboGroup_Click(Area As Integer)
    WriteItemCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap
    'If Shift = 2 Then
    '    If KeyCode = vbKeyX Then
    '        XPButton301_Click
    '    End If
    'End If
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    Dim StrSQL As String
    Dim Dcombos As ClsDataCombos
    Dim Msg As String

    On Error GoTo ErrTrap

    With Me.CboItemType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "”·⁄…"
            .AddItem "Œœ„…"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            .AddItem "Goods"
            .AddItem "Services"
        End If

        .ListIndex = 0
    End With

    Msg = "Ì›÷· «‰  ﬁÊ„ √Ê·« »≈Œ Ì«— «·„Ã„Ê⁄… «· Ï  —Ìœ √‰  ÷Ì› ·Â« «·’›"
    Msg = Msg & " Ê–·ﬂ Õ Ï Ì⁄—÷ ·ﬂ ﬂÊœ √Œ— ’‰› ≈÷Ì› ≈·Ï «·„Ã„Ê⁄… Õ Ï  ” ÿÌ⁄ «‰  Õœœ ﬂÊœ «·’‰› «·ÃœÌœ "
    lblInfo(1).Caption = Msg
    CenterForm Me

    FormPostion Me, GetPostion
    StrSQL = "select * From Groups where GroupID<>1"

    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemSGroups Me.DCboGroup
    Set cDcboSearch = New clsDCboSearch
    Set cDcboSearch.Client = Me.DCboGroup

    Set rs = New ADODB.Recordset
    rs.Open "[TblItems]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    XPTxtID.text = CStr(new_id("TblItems", "ItemID", "", True))
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        If rs.EditMode <> adEditNone Then
            rs.CancelUpdate
        End If

        rs.Close
        Set rs = Nothing
    End If

    FormPostion Me, SavePostion
    Set cDcboSearch = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtCusPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtCusPrice.text, 0)
End Sub

Private Sub TxtDealerPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtDealerPrice.text)
End Sub

Private Sub TxtPurPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPurPrice.text)
End Sub

Private Sub TxtReQty_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtReQty.text, 1)
End Sub

Private Sub TxtSalesPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtSalesPrice.text)
End Sub

Private Sub XPBtnsave_Click()
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim BeginTrans As Boolean
    On Error GoTo ErrTrap

    If XPTxtName.text = "" Then
        Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·’‰›"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtName.SetFocus
        Exit Sub
    End If

    If XPTxtCode.text = "" Then
        Msg = "„‰ ›÷·ﬂ √œŒ· ﬂÊœ «·’‰›...!!"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtCode.SetFocus
        Exit Sub
    End If

    StrSQL = "select * From TblItems where ItemName='" & Trim(XPTxtName.text) & "'"
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsTemp.RecordCount > 0 Then
        Msg = "ÌÊÃœ ’‰› „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·’‰›"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    StrSQL = "select * From TblItems where ItemCode='" & Trim(XPTxtCode.text) & "'"
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsTemp.RecordCount > 0 Then
        Msg = "ÌÊÃœ ’‰› „”Ã· „”»ﬁ« »Â–« «·ﬂÊœ" & Chr(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·ﬂÊœ «·’ÕÌÕ " & Chr(13)
        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ ﬂÊœ «·’‰›"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If Me.DCboGroup.BoundText = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·„Ã„Ê⁄…" & Chr(13)
        Msg = Msg + "«· Ì Ì‰ „Ì «·ÌÂ« Â–« «·’‰›"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboGroup.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    Cn.BeginTrans
    BeginTrans = True
    rs.AddNew
    rs("ItemID").value = IIf(XPTxtID.text = "", "", val(XPTxtID.text))
    rs("ItemCode").value = IIf(XPTxtCode.text = "", "", Trim(XPTxtCode.text))
    rs("ItemName").value = IIf(XPTxtName.text = "", "", Trim(XPTxtName.text))
    rs("HaveSerial").value = XPChkSerial.value

    If Me.DCboGroup.BoundText = "" Then
        rs("GroupID").value = Null
    Else
        rs("GroupID").value = val(Me.DCboGroup.BoundText)
    End If

    rs("PurchasePrice").value = val(Me.TxtPurPrice.text)
    rs("SallingPrice").value = val(Me.TxtSalesPrice.text)
    rs("RequestLimit").value = val(Me.TxtReQty.text)
    rs("CustomerPrice").value = val(Me.TxtCusPrice.text)
    rs("DealerPrice").value = val(Me.TxtDealerPrice.text)

    If Me.ChkGuar.value = vbChecked Then
        rs("HaveGuarantee").value = Me.ChkGuar.value
        rs("GuaranteeValue").value = val(Me.TxtGuarValue.text)
        rs("GuaranteeType").value = IIf(OptGaurType(0).value = True, 0, 1)
    Else
        rs("HaveGuarantee").value = False
        rs("GuaranteeValue").value = 0
        rs("GuaranteeType").value = 0
    End If

    rs("IsArchive").value = 0

    If Me.CboItemType.ListIndex = 0 Then
        rs("ItemType").value = 0
    Else
        rs("ItemType").value = 1
    End If

    rs("AssbliedItem").value = False
    rs("RelatedItem").value = False
    rs.update
    Cn.CommitTrans
    BeginTrans = False
    DataPassing
    Msg = " „ Õ›Ÿ Â–« «·’‰› ... "
    Msg = Msg & Chr(13) & "Â·  —Ìœ ≈œŒ«· ’‰› «Œ— ...??"

    If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2 + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbNo Then
        Unload Me
    Else
        XPTxtID.text = CStr(new_id("TblItems", "ItemID", "", True))
        Me.XPTxtCode.text = ""
        Me.TxtPurPrice.text = ""
        Me.TxtSalesPrice.text = ""
        Me.TxtReQty.text = ""
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub XPButton301_Click()
    Unload Me
End Sub

Private Sub DataPassing()
    Dim StrSQL As String
    Dim StrList As String
    Dim RsNote As New ADODB.Recordset
    On Error GoTo ErrTrap
    StrSQL = "select * From TblItems"
    RsNote.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    Select Case Me.DealingForm

        Case PurchaseTransaction

            With mdifrmmain.ActiveForm
                StrList = .FG.BuildComboList(RsNote, "ItemName", "ItemID")

                If StrList <> "" Then
                    .FG.ColComboList(.FG.ColIndex("Name")) = "|" & StrList
                End If

                StrList = .FG.BuildComboList(RsNote, "ItemCode", "ItemID")

                If StrList <> "" Then
                    .FG.ColComboList(.FG.ColIndex("Code")) = "|" & StrList
                End If

                .FG.TextMatrix(.FG.Row, .FG.ColIndex("Code")) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))
                .FG.TextMatrix(.FG.Row, .FG.ColIndex("Name")) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))
            End With

        Case ShowPrice
            StrList = frmsalebill.FG.BuildComboList(RsNote, "ItemName", "ItemID")

            If StrList <> "" Then
                frmsalebill.FG.ColComboList(2) = "|" & StrList
            End If

            StrList = frmsalebill.FG.BuildComboList(RsNote, "ItemCode", "ItemID")

            If StrList <> "" Then
                frmsalebill.FG.ColComboList(1) = "|" & StrList
            End If

            frmsalebill.FG.TextMatrix(frmsalebill.FG.Row, 2) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))

        Case Maintenance

            With FrmMaintenence
                StrList = .FG.BuildComboList(RsNote, "ItemName", "ItemID")

                If StrList <> "" Then
                    .FG.ColComboList(.FG.ColIndex("Name")) = "|" & StrList
                End If

                StrList = .FG.BuildComboList(RsNote, "ItemCode", "ItemID")

                If StrList <> "" Then
                    .FG.ColComboList(.FG.ColIndex("Code")) = "|" & StrList
                End If

                .FG.TextMatrix(.FG.Row, .FG.ColIndex("Code")) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))
                .FG.TextMatrix(.FG.Row, .FG.ColIndex("Name")) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))
            End With

            '«·—’Ìœ «·«›  «ÕÌ
        Case OpeningBalance

            With FrmOpeningBalance
                StrList = .FG.BuildComboList(RsNote, "ItemName", "ItemID")

                If StrList <> "" Then
                    .FG.ColComboList(.FG.ColIndex("Name")) = "|" & StrList
                End If

                StrList = .FG.BuildComboList(RsNote, "ItemCode", "ItemID")

                If StrList <> "" Then
                    .FG.ColComboList(.FG.ColIndex("Code")) = "|" & StrList
                End If

                .FG.TextMatrix(.FG.Row, .FG.ColIndex("Code")) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))
                .FG.TextMatrix(.FG.Row, .FG.ColIndex("Name")) = IIf(IsNull(XPTxtID.text), "", Trim(XPTxtID.text))
            End With

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If

End Sub

Public Property Get DealingForm() As GridTransType
    DealingForm = m_DealingForm
End Property

Public Property Let DealingForm(ByVal vNewValue As GridTransType)
    'If vNewValue = OpeningBalance Or vNewValue = PurchaseTransaction Or vNewValue = InvoiceTransaction Then
    m_DealingForm = vNewValue
    'End If
End Property

Private Sub WriteItemCode()
    Dim Msg As String

    If Trim(Me.DCboGroup.BoundText) <> "" Then
        Msg = "ﬂÊœ √Œ— ’‰› ≈÷Ì› ≈·Ï «·„Ã„Ê⁄… "
        Msg = Msg & GetLastItemCode(val(Me.DCboGroup.BoundText))
    Else
        Msg = ""
    End If

    Me.lblInfo(0).Caption = Msg
End Sub
