VERSION 5.00
Begin VB.Form frmCarDetails 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3630
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4815
   Icon            =   "frmCarDetails.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3612
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4812
      Begin VB.CommandButton Command1 
         Caption         =   "ãæÇÝÞ"
         Height          =   492
         Left            =   480
         TabIndex        =   10
         Top             =   2760
         Width           =   3012
      End
      Begin VB.TextBox txtVendor 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   480
         TabIndex        =   9
         Top             =   2160
         Width           =   3012
      End
      Begin VB.TextBox txtRecord 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   480
         TabIndex        =   7
         Top             =   1800
         Width           =   3012
      End
      Begin VB.TextBox txtCode 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   480
         TabIndex        =   5
         Top             =   1440
         Width           =   3012
      End
      Begin VB.TextBox txtBoard 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   3012
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáãæÑÏ"
         Height          =   372
         Left            =   3600
         TabIndex        =   8
         Top             =   2160
         Width           =   852
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "ÑÞã ÇáÓÌá"
         Height          =   372
         Left            =   3600
         TabIndex        =   6
         Top             =   1800
         Width           =   852
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "ßæÏ ÇáãæÑÏ"
         Height          =   372
         Left            =   3600
         TabIndex        =   4
         Top             =   1440
         Width           =   852
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ÑÞã ÇááæÍå"
         Height          =   372
         Left            =   3600
         TabIndex        =   2
         Top             =   1080
         Width           =   852
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "ÇáãÚÏå/ÇáÓíÇÑÉ ãÓÌáå ãä ÞÈá"
         Height          =   612
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   4332
      End
   End
End
Attribute VB_Name = "frmCarDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsTemp As ADODB.Recordset
Public board  As String
Public ven As Integer
Dim ddd As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
        
   Dim ss As String
   Set RsTemp = New ADODB.Recordset
   ss = "  select * From TblCustemers where Type=2  and cusid = " & ven
    RsTemp.Open ss, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If RsTemp.RecordCount > 0 Then
          
          txtBoard.Text = board
          txtRecord.Text = IIf(IsNull(RsTemp("recordno").value), "", RsTemp("recordno").value)
         TxtCode.Text = IIf(IsNull(RsTemp("fullcode").value), "", RsTemp("fullcode").value)
         txtVendor.Text = IIf(IsNull(RsTemp("cusname").value), "", RsTemp("cusname").value)
          
    End If

End Sub

