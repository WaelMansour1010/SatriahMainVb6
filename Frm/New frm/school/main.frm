VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ãÏÑÓÉ ÇáãäÇåÌ"
   ClientHeight    =   10650
   ClientLeft      =   2775
   ClientTop       =   1290
   ClientWidth     =   11280
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   Picture         =   "main.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   10650
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "user_id"
      DataSource      =   "Adodc7"
      Height          =   285
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "user_priviliges"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label user_id 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   255
      Left            =   13560
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   9000
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label TXTUSERNAME 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      Caption         =   "ÚÈÏ ÇáÓáÇã"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   13680
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   8400
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÇáãÓÊÎÏã ÇáÍÇáí"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   17280
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Menu file 
      Caption         =   "ãáÝ"
      Begin VB.Menu exit 
         Caption         =   "ÎÑæÌ"
      End
   End
   Begin VB.Menu r00t1 
      Caption         =   "ÇáãáÝÇÊ ÇáÇÓÇÓíÉ"
      Index           =   0
      NegotiatePosition=   3  'Right
      Begin VB.Menu m2 
         Caption         =   "ÇáÓäæÇÊ ÇáÏÑÇÓíÉ"
      End
      Begin VB.Menu m3 
         Caption         =   "ÇáÊÎÕÕÇÊ"
      End
      Begin VB.Menu XC20 
         Caption         =   "ÈíÇäÇÊ ÇáãÏÑÓíä"
      End
      Begin VB.Menu M4 
         Caption         =   "ÇäæÇÚ ÇáÇÞÓÇØ"
      End
      Begin VB.Menu m5 
         Caption         =   "ÇáÎÒíäÉ"
      End
      Begin VB.Menu m6 
         Caption         =   "ÇäæÇÚ ÇáÛÑÇãÇÊ"
      End
      Begin VB.Menu m7 
         Caption         =   "ÇäæÇÚ ÇáÃÔÊÑÇßÇÊ"
      End
      Begin VB.Menu m8 
         Caption         =   "ÇäæÇÚ ÇáÃäÔØÉ"
      End
      Begin VB.Menu m9 
         Caption         =   "ÃäæÇÚ   ÇäÐÇÑÇÊ ÇáØáÇÈ"
      End
      Begin VB.Menu m10 
         Caption         =   "ÇáÓäÉ ÇáÏÑÇÓíÉ ÇáÍÇáíÉ"
         Checked         =   -1  'True
      End
      Begin VB.Menu m11 
         Caption         =   "ãáÝ ÇáØÇáÈ"
      End
      Begin VB.Menu m12 
         Caption         =   "ÇäæÇÚ ÇáÊÍÕíáÇÊ"
      End
      Begin VB.Menu m13 
         Caption         =   "ÇäæÇÚ ÇáãÕÑæÝÇÊ"
      End
      Begin VB.Menu xc111 
         Caption         =   "ÈíÇäÇÊ ÇáÇÍíÇÁ"
      End
      Begin VB.Menu xc22 
         Caption         =   "ÈíÇäÇÊ ÇáÔæÇÑÚ"
      End
      Begin VB.Menu m14 
         Caption         =   "ÊÚÑíÝ ÇáÍÇÝáÉ"
      End
      Begin VB.Menu m15 
         Caption         =   "ÇÓã ãÏíÑ   ÇáãÏÑÓÉ ÇáÍÇáí"
      End
   End
   Begin VB.Menu z 
      Caption         =   "ÍÑßÉ ÇáÇáÊÍÇÞ"
      Begin VB.Menu m16 
         Caption         =   "ÇáÊÍÇÞ ÌÏíÏ"
      End
      Begin VB.Menu m17 
         Caption         =   "ÊÌÏíÏ ÇáÇáÊÍÇÞ"
      End
      Begin VB.Menu m18 
         Caption         =   "ÊÍÏíË ÚÖæíÉ"
         Visible         =   0   'False
      End
      Begin VB.Menu m19 
         Caption         =   "ÇÞÓÇØ"
      End
      Begin VB.Menu m20 
         Caption         =   "ÛÑÇãÇÊ"
      End
      Begin VB.Menu m21 
         Caption         =   "ÈÏá ÝÇÞÏ"
      End
      Begin VB.Menu ACTIVITY 
         Caption         =   "ÇáÇäÔØÉ"
         Begin VB.Menu m22 
            Caption         =   "ÇÖÇÝÉ äÔÇØ áØÇáÈ"
         End
         Begin VB.Menu m23 
            Caption         =   " áØÇáÈÊÌÏíÏ äÔÇØ"
         End
         Begin VB.Menu m24 
            Caption         =   "ÍÐÝ äÔÇØ ØÇáÈ"
         End
      End
      Begin VB.Menu m25 
         Caption         =   "ãæÇÝÞÉ ÇáæÒÇÑÉ"
      End
   End
   Begin VB.Menu t 
      Caption         =   "ÍÑßÉ ÇáÎÒíäÉ"
      Begin VB.Menu m26 
         Caption         =   "ÓÏÇÏ ÞíãÉ ÇáÇäÔØÉ  "
      End
      Begin VB.Menu m27 
         Caption         =   "ÓÏÇÏ ÑÓæã ÇáÇáÊÍÇÞ"
      End
      Begin VB.Menu m28 
         Caption         =   "ÑÓæã ÊÌÏíÏ   ÇáÇáÊÍÇÞ"
      End
      Begin VB.Menu m29 
         Caption         =   "ÏÝÚ ÑÓæã ÈÏá ÝÇÆÏ"
      End
      Begin VB.Menu m30 
         Caption         =   "ãÕÑæÝÇÊ"
      End
      Begin VB.Menu m31 
         Caption         =   "ÊÍÕíáÇÊ"
      End
   End
   Begin VB.Menu AA 
      Caption         =   "ÍÑßÉ ÇáßÇÑäíåÇÊ"
      Begin VB.Menu m32 
         Caption         =   "ÇáßÇÑäíåÇÊ ÇáÌÇåÒå ááØÈÇÚå"
      End
      Begin VB.Menu m33 
         Caption         =   "ÇáßÇÑäíåÇÊ ÇáãØÈæÚå æãÚÏå ááÊÓáíã"
      End
      Begin VB.Menu m34 
         Caption         =   "ÇÚÇÏÉ ØÈÇÚå ÇáßÇÑäíåÇÊ"
      End
   End
   Begin VB.Menu AF 
      Caption         =   "ãÊÇÈÚÉ ÇæáíÇÁ ÇáÇãæÑ"
      Begin VB.Menu x1 
         Caption         =   "Ïáíá ÇáÊáíÝæäÇÊ"
      End
      Begin VB.Menu x2 
         Caption         =   "ÇÏÇÑÉ ÇáÑÓÇÆá"
      End
      Begin VB.Menu x3 
         Caption         =   "ãÊÇÈÚÉ ÇáÛíÇÈ"
      End
      Begin VB.Menu x4 
         Caption         =   "ÇäÐÇÑÇÊ ÇáÝÕá"
      End
      Begin VB.Menu x5 
         Caption         =   "ÇÌÊãÇÚ ÇæáíÇÁ ÇáÇãæÑ"
      End
      Begin VB.Menu x6 
         Caption         =   "ÇáÊÍæíá ÇáØÈí ááØÇáÈ"
      End
   End
   Begin VB.Menu XCC 
      Caption         =   "ÇáÌÏæá ÇáÏÑÇÓí"
      Begin VB.Menu xc1 
         Caption         =   "ÇÚÏÇÏ ÇáÌÏæá ÇáÏÑÇÓí"
      End
      Begin VB.Menu xc2 
         Caption         =   "ØÈÇÚÉ ÇáÌÏæá ÇáÏÑÇÓí"
      End
      Begin VB.Menu XC3 
         Caption         =   "ØÈÇÚÉ ÌÏæá ßá ãÏÑÓ"
      End
   End
   Begin VB.Menu cx1 
      Caption         =   "ÇáßÊÈ ÇáÏÑÇÓíÉ"
      Begin VB.Menu cx2 
         Caption         =   "ÊÚÑíÝ ÇáßÊÈ"
      End
      Begin VB.Menu cx3 
         Caption         =   "ÍÑßÉ ÊÓáíã ÇáßÊÈ"
      End
   End
   Begin VB.Menu AG 
      Caption         =   "ÇáÊÞÇÑíÑ"
      Begin VB.Menu m35 
         Caption         =   "ÇáÌãÚíÉ ÇáÚãæãíÉ"
      End
      Begin VB.Menu m44 
         Caption         =   "ÊÞÑíÑ Úä ÇáÊÌÏíÏ Ýì ãÏå ãÍÏÏå "
      End
      Begin VB.Menu m43 
         Caption         =   "ÊÞÑíÑ Úä ÇáÌÏíÏ Ýì ãÏå ãÍÏÏå"
      End
      Begin VB.Menu dddd 
         Caption         =   "ÊÞÑíÑ ÇáÎÒíäÉ"
         Visible         =   0   'False
      End
      Begin VB.Menu m36 
         Caption         =   "ÊÞÑíÑ ÚÏÏ ÇáÅäÇË Çæ ÇáÐßæÑ (( ÊÌÏíÏ)) æ ((ÌÏíÏ)) Çæ ßá æÇÍÏ Úáì ÍÏå æáßä ãä ÎáÇá ÇáßÇÑäíåÇÊ"
         Visible         =   0   'False
      End
      Begin VB.Menu m37 
         Caption         =   "ÊÞÑíÑ ÇáÇÚÖÇÁ ÇáÊÇÈÚíä ÝæÞ 18 ÓäÉ"
         Visible         =   0   'False
      End
      Begin VB.Menu treport 
         Caption         =   "ÊÝÇÑíÑ ÇáÍÒíäÉ"
         Begin VB.Menu m38 
            Caption         =   "ÊÞÑíÑ ÍÇáÉ ÇáØÇáÈ"
         End
         Begin VB.Menu m39 
            Caption         =   " ÊÞÑíÑ ÇáÎÒíäÉ (( ÇáÍÇÝÙÉ ))"
         End
         Begin VB.Menu m40 
            Caption         =   "ÊÞÑíÑ ÇáÇäÔØÉ"
         End
      End
      Begin VB.Menu m41 
         Caption         =   " ÊÞÇÑíÑ ÇáßÇÑäíåÇÊ"
      End
   End
   Begin VB.Menu r7 
      Caption         =   "ÇáãÓÊÎÏãíä"
      Begin VB.Menu m42 
         Caption         =   "ÕáÇÍíÇÊ ÇáãÓÊÎÏãíä"
      End
      Begin VB.Menu M45 
         Caption         =   "ÊÍÏíÏ ãÓÇÑ ÇáÈÑäÇãÌ"
      End
      Begin VB.Menu ccc 
         Caption         =   ""
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim check As Integer



Private Sub Command1_Click()

End Sub

Private Sub cx2_Click()
books.Show
End Sub

Private Sub cx3_Click()
taslem.Show
End Sub

Private Sub Form_Activate()
Adodc7.CommandType = adCmdText
'Adodc7.RecordSource = " select * from user_priviliges  where [view]=0 and user_id=" & user_id.Caption

Adodc7.RecordSource = " select * from user_priviliges where user_id=" & user_id.Caption


Adodc7.Refresh



For i = 1 To Adodc7.Recordset.RecordCount
If Adodc7.Recordset.Fields!no = "m2" Then
m2.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m3" Then
m3.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m4" Then
M4.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m5" Then
m5.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m6" Then
m6.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m7" Then
m7.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m8" Then
m8.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m9" Then
m9.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m10" Then
m10.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m11" Then
m11.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m12" Then
m12.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m13" Then
m13.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m14" Then
m14.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m15" Then
m15.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m16" Then
m16.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m17" Then
m17.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m18" Then
m18.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m19" Then
m19.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m20" Then
m20.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m21" Then
m21.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m22" Then
m22.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m23" Then
m23.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m24" Then
m24.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m25" Then
m25.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m26" Then
m26.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m27" Then
m27.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m28" Then
m28.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m29" Then
m29.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m30" Then
m30.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m31" Then
m31.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m32" Then
m32.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m33" Then
m33.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m34" Then
m34.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m35" Then
m35.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m36" Then
m36.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m37" Then
m37.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m38" Then
m38.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m39" Then
m39.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m40" Then
m40.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m41" Then
m41.Enabled = Adodc7.Recordset.Fields![view]
Else
If Adodc7.Recordset.Fields!no = "m42" Then
m42.Enabled = Adodc7.Recordset.Fields![view]

Else
If Adodc7.Recordset.Fields!no = "m43" Then
m43.Enabled = Adodc7.Recordset.Fields![view]

Else
If Adodc7.Recordset.Fields!no = "m44" Then
m44.Enabled = Adodc7.Recordset.Fields![view]

Else
If Adodc7.Recordset.Fields!no = "m45" Then
M45.Enabled = Adodc7.Recordset.Fields![view]

 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
  End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
  End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
  

Adodc7.Recordset.MoveNext
Next i
End Sub

Private Sub Image1_Click()

End Sub

Private Sub m2_Click()
MEMBER_TYPES.Show
End Sub

Private Sub m32_Click()
READY_TO_PRINT.Show
End Sub

Private Sub m33_Click()
PRINTED_AND_READY.Show
End Sub

Private Sub m22_Click()
member_activity.Show
End Sub

Private Sub m3_Click()
DEPARTMENT.Show
End Sub

Private Sub m4_Click()
Installments_TYPES.Show
End Sub

Private Sub m14_Click()
Hafela.Show
'CARD_VALUE.Show
End Sub
 

Private Sub m41_Click()
REPORTSFRM.Show
REPORTSFRM.case_id = 14
REPORTSFRM.Command1.Caption = "ÊÞÑíÑ Úä ÇáßÇÑäíåÇÊ "

End Sub

Private Sub m35_Click()
REPORTSFRM.Show
REPORTSFRM.case_id = 7
REPORTSFRM.Frame2.Visible = True
REPORTSFRM.Frame3.Visible = False
REPORTSFRM.Command1.Caption = " ÊÞÑíÑ ÇáÌãÚíÉ ÇáÚãæãíÉ   "


End Sub

Private Sub m15_Click()
center_manger.Show
End Sub

Private Sub m42_Click()
user_priviliges.Show
End Sub

Private Sub m43_Click()
REPORTSFRM.Show
REPORTSFRM.case_id = 5
REPORTSFRM.Command1.Caption = "ÊÞÑíÑ Úä ÇáÌÏíÏ Ýì ãÏå ãÍÏÏå "

End Sub

Private Sub m44_Click()
REPORTSFRM.Show
REPORTSFRM.case_id = 4
REPORTSFRM.Command1.Caption = "ÊÞÑíÑ Úä ÇáÊÌÏíÏ Ýì ãÏå ãÍÏÏå "

End Sub

Private Sub M45_Click()
IMAGE_PATH_FRM.Show
End Sub

Private Sub m5_Click()
Treasury.Show
End Sub

Private Sub dddd_Click()
REPORTSFRM.Show
REPORTSFRM.case_id = 6
REPORTSFRM.Command1.Caption = "ÊÞÑíÑ ÇáÎÒíäÉ Úä ÊÌÏíÏ ÇáÚÖæíÉ"

End Sub

Private Sub m24_Click()
delete_member_activity.Show
End Sub

Private Sub m6_Click()
Fine_TYPES.Show
End Sub

Private Sub m21_Click()
losed_card.Show
End Sub

Private Sub EXIT_Click()
Dim X As Integer
X = MsgBox("åá ÊÑíÏ äÃßíÏ ÇáÎÑæÌ ãä ÇáÈÑäÇãÌ", vbYesNo + vbExclamation)

If X = vbYes Then
End

End If
End Sub

Private Sub m7_Click()
Subscription_TYPES.Show
End Sub

Private Sub Form_Load()
'IMAGE_PATH_FRM.Show
'On Error Resume Next
'Image1.Picture = LoadPicture(IMAGE_PATH_FRM.IMAGE_PATH & "\IMAGES\bg.jpg")
check = 0

'Image1.Top = 0
'Image1.Height = 0
'Image1.Width = Me.Width
'Image1.Height = Me.Height
End Sub

Private Sub m8_Click()
ACTIVITIES_type.Show
End Sub

Private Sub m10_Click()
this_year.Show
End Sub

Private Sub m11_Click()
MEMBERS.Show
End Sub

Private Sub m19_Click()
ADD_MEMBER_INSTALLMENTS.Show
End Sub

Private Sub m9_Click()
STOP_MEMBER_TYPE.Show
End Sub

Private Sub m17_Click()
update_member.Show

End Sub

Private Sub newwww_Click()
End Sub

Private Sub m36_Click()
REPORTSFRM.Show
REPORTSFRM.case_id = 10
REPORTSFRM.Command1.Caption = "ÊÞÑíÑ ÚÏÏ ÇáÅäÇË Çæ ÇáÐßæÑ (( ÊÌÏíÏ)) æ ((ÌÏíÏ)) Çæ ßá æÇÍÏ Úáì ÍÏå æáßä ãä ÎáÇá ÇáßÇÑäíåÇÊ"

End Sub

Private Sub m26_Click()
PAY_NEW_ACTIVITY.Show
End Sub

Private Sub PAY_ACTIVITY_Click()

End Sub

Private Sub m29_Click()
pay_losed_ard.Show
End Sub

Private Sub m18_Click()
alarm_frm.Show
alarm_frm.txtcheck.Caption = "1"
End Sub

Private Sub Q_Click()

End Sub

Private Sub m25_Click()
SECURITY_FORM.Show
End Sub

Private Sub pay_Click()

End Sub

Private Sub m37_Click()

Form3.case_id = 1
Form3.Show

End Sub

Private Sub m23_Click()
renew_member_activity.Show
End Sub

Private Sub m31_Click()
services.Show
End Sub

Private Sub m30_Click()
masrouf.Show
End Sub

 

Private Sub m40_Click()
REPORTSFRM.Show
REPORTSFRM.case_id = 9
REPORTSFRM.Command1.Caption = " ÊÞÑíÑ  ÇáÇäÔØÉ - ãÇáíÇÊ   "

End Sub

Private Sub m28_Click()
operation_from.Show
 operation_from.Caption = "ÔÇÔÉ ÊÌÏíÏ ÇáÚÖæíÉ  "
operation_from.Adodc1.CommandType = adCmdText
operation_from.Adodc1.RecordSource = "select *  FROM OPERATIONS where PAYED=0 and (operation_type= 'ÊÌÏíÏ ÚÖæíÉ' OR operation_type= 'ÊÌÏíÏ äÔÇØ')"
operation_from.Adodc1.Refresh


If operation_from.Adodc1.Recordset.RecordCount > 0 Then
operation_from.Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub m16_Click()
new_members.Show
End Sub

Private Sub U_Click()

End Sub

Private Sub m38_Click()
REPORTSFRM.Show
REPORTSFRM.case_id = 8
REPORTSFRM.Command1.Caption = " ÊÞÑíÑ ÍÇáÉ ÇáÚÖæíÉ   "

End Sub

Private Sub m20_Click()
ADD_MEMBER_FINES.Show
End Sub

Private Sub m1_Click()
this_year.Show
End Sub



 

Private Sub m39_Click()
REPORTSFRM.Show
REPORTSFRM.case_id = 15
REPORTSFRM.Command1.Caption = " ÊÞÑíÑ ÇáÎÒíäÉ (( ÇáÍÇÝÙÉ )) "

End Sub

Private Sub m27_Click()
operation_from.Show
operation_from.Text8.Visible = False
operation_from.Text24.Visible = False

operation_from.Command6.Visible = False
operation_from.Command7.Visible = False
operation_from.Command10.Visible = False
operation_from.Label20.Visible = False
operation_from.Label29.Visible = False
operation_from.Frame1.Visible = False
operation_from.Caption = "ÔÇÔÉ ÇáÚÖæíÉ ÇáÌÏíÏÉ"
operation_from.Adodc1.CommandType = adCmdText
operation_from.Adodc1.RecordSource = "select *  FROM OPERATIONS where PAYED=0 and (operation_type= 'ÚÖæíÉ ÌÏíÏÉ' OR operation_type= 'äÔÇØ ÌÏíÏ')"
operation_from.Adodc1.Refresh


If operation_from.Adodc1.Recordset.RecordCount > 0 Then
operation_from.Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub m12_Click()
SERVICE_TYPE.Show
End Sub

Private Sub m13_Click()
MASROUF_TYPE.Show
End Sub

Private Sub m34_Click()
reprint_card.Show
End Sub

Private Sub x1_Click()
dalil.Show
End Sub

Private Sub x2_Click()
messages_frm.Show
End Sub

Private Sub x3_Click()
ATTENDANCE.Show
End Sub

Private Sub x4_Click()
ENZAR.Show
End Sub

Private Sub x5_Click()
MEETING.Show
End Sub

Private Sub x6_Click()
TAHWEL.Show
End Sub

Private Sub XC1_Click()
table.Show
End Sub

Private Sub xc111_Click()
hay.Show
End Sub

Private Sub XC20_Click()
MISTER.Show
End Sub

Private Sub xc22_Click()
street.Show
End Sub
