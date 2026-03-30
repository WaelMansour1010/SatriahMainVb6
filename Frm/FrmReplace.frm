VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReplace 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÅÓÊÈÏÇá ÞØÚÉ ÊÈÚ ÇáÖãÇä"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   Icon            =   "FrmReplace.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3720
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtTransSerial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1290
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1020
      Width           =   1485
   End
   Begin VB.TextBox XPTxtMaintanenceID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1380
      Width           =   1485
   End
   Begin VB.TextBox TxtNewSerial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   210
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2820
      Width           =   2625
   End
   Begin VB.TextBox TxtItemSerial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   210
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2100
      Width           =   2625
   End
   Begin VB.TextBox TxtTransID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   60
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   660
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo DCboStoreName 
      Height          =   315
      Left            =   210
      TabIndex        =   6
      Top             =   2460
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "7"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboItemsName 
      Height          =   315
      Left            =   210
      TabIndex        =   4
      Top             =   1740
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   3210
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÅáÛÇÁ"
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
      ButtonImage     =   "FrmReplace.frx":038A
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton cmdok 
      Height          =   375
      Left            =   1110
      TabIndex        =   15
      Top             =   3210
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÍÝÙ"
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
      ButtonImage     =   "FrmReplace.frx":0724
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin MSComCtl2.DTPicker DtbReplaceDate 
      Height          =   315
      Left            =   1290
      TabIndex        =   1
      Top             =   660
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
      Format          =   96337921
      CurrentDate     =   38784
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÊÇÑíÎ ÇáÇÓÊÈÏÇá"
      Height          =   285
      Index           =   5
      Left            =   2820
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   675
      Width           =   1770
   End
   Begin VB.Label LblCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " ÅÓÊÈÏÇá ÞØÚÉ ÊÇáÝÉ ÊÈÚ ÇáÖãÇä"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   615
      Index           =   13
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   0
      Width           =   4845
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÑÞã ÚãáíÉ ÇáÕíÇäÉ"
      Height          =   285
      Index           =   3
      Left            =   2820
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1395
      Width           =   1770
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ãßÇä ÇáÞØÚÉ ÇáÌÏíÏÉ"
      Height          =   285
      Left            =   2820
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2475
      Width           =   1770
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÓíÑíÇá ÇáÞØÚÉ ÇáÌÏíÏÉ"
      Height          =   285
      Index           =   4
      Left            =   2820
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2835
      Width           =   1770
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÇÓã ÇáÕäÝ"
      Height          =   285
      Index           =   1
      Left            =   2820
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1755
      Width           =   1770
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÓíÑíÇá ÇáÞØÚÉ ÇáãÓÊÈÏáÉ"
      Height          =   285
      Index           =   0
      Left            =   2820
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2115
      Width           =   1770
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÑÞã ÝÇÊæÑÉ ÇáÈíÚ"
      Height          =   285
      Index           =   2
      Left            =   2820
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1035
      Width           =   1770
   End
End
Attribute VB_Name = "FrmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDcbo As clsDCboSearch

Private Sub CmdCancel_Click()
'    FrmMaintenence.Tag = ""
'    Unload Me
End Sub

Private Sub CmdOk_Click()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTest As New ADODB.Recordset
    Dim Msg As String
    StrSQL = "select * From QryGardComplete"
    StrSQL = StrSQL + " where ItemID=" & DCboItemsName.BoundText
    StrSQL = StrSQL + " AND Transaction_Details.ItemSerial='" & Trim(TxtNewSerial.text) & "'"
    StrSQL = StrSQL + " and StoreID=" & DCboStoreName.BoundText
    RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsTest.EOF Or RsTest.BOF Then
        Msg = "ÇáÞØÚÉ : " & DCboItemsName.text & Chr(13)
        Msg = Msg + " ÐÇÊ ÇáÓíÑíÇá : "
        Msg = Msg + Trim(TxtNewSerial.text) & Chr(13)
        Msg = Msg + "ÛíÑ ãæÌæÏÉ Ýí ÇáãÎÒä ÇáãÍÏÏ" & Chr(13)
        Msg = Msg + "åá ÊÑÛÈ  Ýí ãÚÑÝÉ ÓíÑíÇá ÇáÞØÚ ÇáãæÌæÏÉ" & Chr(13)
        Msg = Msg + "ãä åÐÇ ÇáäæÚ Ýí ÇáãÎÇÒä"

        If (MsgBox(Msg, vbYesNo + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)) = vbYes Then
            FrmSearchSerial.Tag = Me.Tag
            FrmSearchSerial.Txt.text = "Replace"
            FrmSearchSerial.show vbModal
        End If

        Exit Sub
    End If

    'FrmMaintenence.Tag = "XX"
    'Me.Hide
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    CenterForm Me

    FormPostion Me, GetPostion
    StrSQL = "SELECT * From TblStore"
    fill_combo Me.DCboStoreName, StrSQL
    StrSQL = "SELECT * From TblItems"
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsTemp.RecordCount > 0 Then

        With DCboItemsName
            Set .RowSource = RsTemp
            .BoundColumn = RsTemp(0).name
            .ListField = RsTemp(2).name
            .BoundText = ""
            .text = ""
        End With

    End If

    DtbReplaceDate.value = Date
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DCboItemsName
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub
