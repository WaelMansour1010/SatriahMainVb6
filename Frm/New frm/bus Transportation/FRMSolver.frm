VERSION 5.00
Begin VB.Form FRMSolver 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   10560
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
   Begin VB.CommandButton Command2 
      Caption         =   " ‰›Ì–"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox TxtMonthId 
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Text            =   "4"
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   " ‰›Ì–"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtContract 
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Text            =   "5"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox TxtREI 
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Text            =   "72"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox TxtReD 
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Text            =   "12"
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "—ﬁ„ «·‘Â—"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "«÷«›… ·ÿ·» «·’—› —ﬁ„"
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "«÷«›…  ⁄ﬁœ ··«” Õﬁ«ﬁ«  —ﬁ„"
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "—ﬁ„ «·⁄ﬁœ"
      Height          =   495
      Left            =   8160
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "FRMSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim StrSql As String
Dim AllIid As String

StrSql = "update TblMinistryContract_Installment set  VR_Paid=1 , VRID=" & val(TxtReD) & "   where type=2 and  idmc=" & val(txtContract) & " and monthid=" & TxtMonthId & ""
Cn.Execute StrSql
  AllIid = GetAllIDFromTblMinistrtyContract(val(TxtReD))
  MsgBox AllIid
Cn.Execute " update TblExchangeRequest2 set allid ='" & AllIid & "' Where ID =" & (val(TxtReD))

 
End Sub

Private Sub Command2_Click()
Dim StrSql As String
Dim AllIid As String


 
StrSql = "update TblAttributionInstallmentDivided  set DDEmbarkDateH=Null  , DDEmbarkDate =null" & " where DDEmbarkDateH='0'"
Cn.Execute StrSql


StrSql = "update TblAttributionInstallmentDivided  set RE_paid=1 ,REID=" & TxtREI & " where idac=" & txtContract & " and monthid=" & TxtMonthId & " "
Cn.Execute StrSql
  
   AllIid = GetAllIDFromTblAttributionInstallmentDivided(val(TxtREI))
  MsgBox AllIid
  
StrSql = "update   TblExchangeRequest    set allid ='" & AllIid & "'Where ID =" & (val(TxtREI)) & ""

 Cn.Execute StrSql
 
End Sub
