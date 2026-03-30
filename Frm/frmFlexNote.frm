VERSION 5.00
Begin VB.Form frmFlexNote 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "Flex Cell Note"
   ClientHeight    =   1500
   ClientLeft      =   2940
   ClientTop       =   3390
   ClientWidth     =   2475
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   2475
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNote 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   855
      Left            =   240
      MaxLength       =   255
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2115
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Top             =   1110
      Width           =   540
   End
End
Attribute VB_Name = "frmFlexNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Ctl As Control

Private m_lRow As Long

Private m_lCol As Long

Private Sub Form_Deactivate()
    '' if we have a note, save it
    'If Len(TxtNote) Then
    '    If TxtNote.Text <> "íãßäß ßÊÇÈÉ ÊÚáíÞ åäÇ...º" & vbCrLf & "==================" Then
    '        m_Ctl.Cell(flexcpData, m_lRow, m_lCol) = TxtNote.Text
    '    Else
    '        m_Ctl.Cell(flexcpData, m_lRow, m_lCol) = ""
    '    End If
    ''no text? no note!
    'Else
    '    'Set m_Ctl.Cell(flexcpData, m_lRow, m_lCol) = Nothing
    '    'Set m_Ctl.Cell(flexcpPicture, m_lRow, m_lCol) = Nothing
    'End If
    ''Unload Me
End Sub

Private Sub Form_Load()
    lblNote.BackStyle = 1
    lblNote.BackColor = txtNote.BackColor

    If SystemOptions.UserInterface = ArabicInterface Then
        lblNote.RightToLeft = True
        txtNote.RightToLeft = True
        lblNote.Alignment = vbRightJustify
        txtNote.Alignment = vbRightJustify
    Else
        lblNote.RightToLeft = False
        txtNote.RightToLeft = False
        lblNote.Alignment = vbLeftJustify
        txtNote.Alignment = vbLeftJustify
    End If

End Sub

Private Sub Form_Resize()
    txtNote.Move 0, 0, ScaleWidth - 40, ScaleHeight - 40
    lblNote.Move txtNote.left, txtNote.top, txtNote.Width, txtNote.Height
End Sub

'Public Sub ShowNote(Ctl As Control, Row As Long, Col As Long)
'
'    ' save control info to update note later
'    Set m_Ctl = Ctl
'    m_lRow = Row
'    m_lCol = Col
'
'    ' copy font from parent control to look nice
'    Set txtNote.Font = Ctl.Font
'    Set lblNote.Font = Ctl.Font
'
'    ' calculate note position
'    Dim fLeft!, fTop!, fWid!, fHei!
'    With Ctl
'        fLeft = MDIFrmamin.left + .Parent.left + .left + .ColPos(Col) + .ColWidth(Col) + 200
'        fTop = MDIFrmamin.top + .Parent.top + .top + .RowPos(Row) + 300
'    End With
'
'    ' calculate note size
'    lblNote = txtNote
'    fWid = lblNote.Width + 300
'    fHei = lblNote.Height + 150
'
'    ' make sure note is not off the screen
'    If fLeft + fWid > MDIFrmamin.Width Then fLeft = fLeft - fWid - Ctl.ColWidth(Col) - 200
'    If fTop + fHei > MDIFrmamin.Height - 300 Then fTop = MDIFrmamin.Height - fHei - 300
'
'    ' show note (we stay up until deactivated)
'    Move fLeft, fTop, fWid, fHei
'    txtNote.SelStart = 32000
'    Visible = True
'End Sub
'
Private Sub txtNote_Change()
    ' resize note as the user types
    Dim fWid As Single, fHei As Single
    lblNote.Caption = txtNote.text
    fWid = lblNote.Width + 300
    fHei = lblNote.Height + 150

    If fWid < 2475 Then fWid = 2475
    If fHei < 1500 Then fHei = 1500
    Move left, top, fWid, fHei
End Sub

Public Sub EditComment()
    lblNote.Visible = False
    txtNote.Visible = True
    'TxtNote.SetFocus
End Sub
