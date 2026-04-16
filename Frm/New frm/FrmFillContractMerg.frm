VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmFillContractMerg 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "‘«‘…œ„Ã «·ÊÕœ«  «·⁄ﬁ«—ÌÂ"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3060
   Icon            =   "FrmFillContractMerg.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1545
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   0
      Top             =   1080
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   810
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "„”Õ"
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   30
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
      BackStyle       =   0
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "œ„Ã «·ÊÕœ«  «·⁄ﬁ«—ÌÂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   135
      TabIndex        =   3
      Top             =   0
      Width           =   2820
   End
End
Attribute VB_Name = "FrmFillContractMerg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Dim NewGrid As New ClsGrid
Dim currentterms As String


Sub GetContract()
Dim Rs1 As ADODB.Recordset
Dim sql As String
Set Rs1 = New ADODB.Recordset
Dim i As Integer
sql = " select ContNo from TblContract  "
Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

If Rs1.RecordCount > 0 Then
Rs1.MoveFirst
For i = 1 To Rs1.RecordCount

save val(IIf(IsNull(Rs1("ContNo").value), 0, Rs1("ContNo").value))
Rs1.MoveNext
Next i
End If
End Sub

Sub save(Optional id As Integer = 0)
Dim rs As ADODB.Recordset
Dim sql As String
Dim str As String
Dim i As Integer
Dim sg2 As String
str = ""
Set rs = New ADODB.Recordset

sql = " SELECT     dbo.TblIqrMerg.Cont, dbo.TblIqrMerg.TypeID, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblIqrMerg.UntID, dbo.TblAqarDetai.unitno"
sql = sql & "  FROM         dbo.TblAqarDetai RIGHT OUTER JOIN"
sql = sql & "                       dbo.TblIqrMerg ON dbo.TblAqarDetai.Id = dbo.TblIqrMerg.UntID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblAkarUnit ON dbo.TblIqrMerg.TypeID = dbo.TblAkarUnit.id"
sql = sql & "  Where (dbo.TblIqrMerg.cont = " & id & ")"

rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount > 0 Then


For i = 1 To rs.RecordCount
If SystemOptions.UserInterface = ArabicInterface Then
 str = str & IIf(IsNull(rs("name").value), "", rs("name").value)
 Else
 str = str & IIf(IsNull(rs("namee").value), "", rs("namee").value)
 End If
str = str & " "
 str = str & IIf(IsNull(rs("unitno").value), "", rs("unitno").value)
  str = str & ","
 str = str & Chr(13)
rs.MoveNext
Next i

End If
If str <> "" Then
sg2 = " update TblContract set StrMerg='" & str & "' where ContNo=" & id & ""
Cn.Execute sg2, , adExecuteNoRecords
End If

End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
    GetContract
    MsgBox " „ «· ÕœÌÀ"
    Unload Me

   
    End Select

End Sub





Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub


    
Private Sub Form_Load()
Dim rs As ADODB.Recordset
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos



    Set DCboSearch = New clsDCboSearch
   
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

                

    Set GrdBack = New ClsBackGroundPic




End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
'

