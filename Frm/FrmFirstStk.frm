VERSION 5.00
Begin VB.Form FrmFirstStk 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÊÍÏíÏ ÓÊíßÑ ÇáÈÏÇíÉ ..."
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   ControlBox      =   0   'False
   Icon            =   "FrmFirstStk.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÓÊíßÑ ÇáÈÏÇíÉ"
      Height          =   720
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   6300
      Width           =   5025
      Begin VB.TextBox TxtStkNum 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   0
         Top             =   225
         Width           =   480
      End
      Begin VB.CommandButton CmdEnd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ÎÑæÌ"
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   645
      End
      Begin VB.CommandButton CmdOk 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ãæÇÝÞ"
         Height          =   345
         Left            =   825
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÑÞã ÓÊíßÑ ÇáÈÏÇíÉ"
         Height          =   225
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   300
         Width           =   1350
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   5175
      X2              =   -75
      Y1              =   6285
      Y2              =   6285
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   5220
      X2              =   -30
      Y1              =   6300
      Y2              =   6300
   End
   Begin VB.Shape Border 
      BackColor       =   &H00C0C0FF&
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   315
      Left            =   45
      Top             =   45
      Width           =   510
   End
   Begin VB.Label LblSticker 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Visible         =   0   'False
      Width           =   510
   End
End
Attribute VB_Name = "FrmFirstStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Msg As String

Private Sub CmdEnd_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
    On Error GoTo ErrTrap

    Dim i As Integer
    Msg = "ÚÝæÇ åÐÇ ÇáÇÓÊíßÑ ÛíÑ ãæÌæÏ"

    If Trim(TxtStkNum.text) > LblSticker.count - 1 Then
        MsgBox Msg, vbOKOnly + vbMsgBoxRtlReading + vbMsgBoxRight, App.title
        TxtStkNum.text = 1
        Exit Sub
    ElseIf Trim(TxtStkNum.text) < 1 Then
        MsgBox Msg, vbOKOnly + vbMsgBoxRtlReading + vbMsgBoxRight, App.title
        TxtStkNum.text = 1
        Exit Sub
    End If

    FrmSetting.LblStkBegin.Caption = val(TxtStkNum.text)
    Unload Me
ErrTrap:
End Sub

Public Sub GridOfBarCode(ByVal StkCount As Integer, _
                         ByVal StkCols As Integer)
    On Error GoTo ErrTrap

    Dim BarWidth As Single
    Dim BarHeight As Single
    Dim IRow As Integer
    Dim StkRows As Integer
    Dim ICol As Integer
    Dim LblTop As Single
    Dim LblLeft As Single
    Dim sp As Single
    Dim IX As Integer
    Dim mm2Twp As Single
    StkRows = StkCount / StkCols
    mm2Twp = 56.7
    sp = 0.45 * mm2Twp
    BarWidth = ((88.4 * mm2Twp) - (StkCols * sp)) / StkCols
    BarHeight = ((109 * mm2Twp) - (StkRows * sp)) / StkRows
    IX = 1
    LblTop = LblSticker(0).top
    LblLeft = LblSticker(0).left

    For IRow = 1 To StkRows
        For ICol = 1 To StkCols
            Load LblSticker(LblSticker.count)
            LblSticker(LblSticker.count - 1).Move LblLeft, LblTop, BarWidth, BarHeight
            LblLeft = LblSticker(LblSticker.count - 1).left + sp + LblSticker(LblSticker.count - 1).Width
            LblSticker(LblSticker.count - 1).Caption = IX
            IX = IX + 1
            LblSticker(LblSticker.count - 1).Visible = True
        Next

        LblTop = LblSticker(LblSticker.count - 1).top + sp + LblSticker(LblSticker.count - 1).Height
        LblLeft = LblSticker(0).left
    Next

    BrdPos 1
    Border.ZOrder 0
    TxtStkNum.text = 1

    Me.Refresh
ErrTrap:
End Sub

Private Sub Form_Load()

End Sub

Private Sub LblSticker_MouseDown(Index As Integer, _
                                 Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 Y As Single)
    On Error GoTo ErrTrap

    RestBkColor
    BrdPos Index
    Border.ZOrder 0
    TxtStkNum.text = Index
    TxtStkNum.SetFocus
    LblSticker(Index).backcolor = &HFFFF00
ErrTrap:

End Sub

Public Sub BrdPos(ByVal Index As Integer)
    On Error GoTo ErrTrap

    With Border
        .Width = LblSticker(Index).Width '+ 0.48
        .Height = LblSticker(Index).Height ' + 0.48
        .left = LblSticker(Index).left '- 0.24
        .top = LblSticker(Index).top ' - 0.24
    End With

ErrTrap:

End Sub

Public Sub RestBkColor()
    Dim i As Integer

    For i = 1 To LblSticker.count - 1
        LblSticker(i).backcolor = vbWhite
    Next

End Sub

Private Sub TxtStkNum_KeyUp(KeyCode As Integer, _
                            Shift As Integer)
    On Error GoTo ErrTrap

    Dim i As Integer
    Msg = "ÚÝæÇ åÐÇ ÇáÇÓÊíßÑ ÛíÑ ãæÌæÏ"

    If Trim(TxtStkNum.text) > LblSticker.count - 1 Then
        MsgBox Msg, vbOKOnly + vbMsgBoxRtlReading + vbMsgBoxRight, App.title
        TxtStkNum.text = 1
        Exit Sub
    ElseIf Trim(TxtStkNum.text) < 1 Then
        MsgBox Msg, vbOKOnly + vbMsgBoxRtlReading + vbMsgBoxRight, App.title
        TxtStkNum.text = 1
        Exit Sub
    End If

    LblSticker_MouseDown val(TxtStkNum.text), 0, 0, 0, 0
    'TxtStkNum.SelStart = 0
    'TxtStkNum.SelLength = Len(TxtStkNum.Text)
ErrTrap:
End Sub
