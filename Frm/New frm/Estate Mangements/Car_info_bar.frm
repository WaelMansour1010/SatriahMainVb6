VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Car_info_bar 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Ñ"
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3540
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox d2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   1560
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "æÕÝ ÇÎÑ ÕíÇäÉ"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ÊÇÑíÎ ÇÎÑ ÕíÇäÉ"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÝÑÚ"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ÑÞã ÇááæÍÉ"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -240
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ÊÇÑíÎ ÇäÊåÇÁ ÇáÊÃãíä"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -240
         TabIndex        =   8
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ÊÇÑíÎ ÇäÊåÇÁ ÇáÇÓÊãÇÑÉ"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -240
         TabIndex        =   7
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ÇÎÑ ÞÑÇÉ ááÚÏÇÏ"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -240
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ÚåÏå ÇáÓÇÆÞ"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -240
         TabIndex        =   5
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ÑÞã ÇáãÚÏå/ÇáÓíÇÑÉ"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -240
         TabIndex        =   4
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   4560
      Top             =   0
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   -480
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Last maintenance Descrirtion"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   0
         TabIndex        =   29
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Last maintenance Date"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   0
         TabIndex        =   28
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Car Index"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -240
         TabIndex        =   16
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Driver"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -240
         TabIndex        =   15
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Last Counter"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -240
         TabIndex        =   14
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "End license warrenty "
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -240
         TabIndex        =   13
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "End Insurance Date"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -240
         TabIndex        =   12
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Car #"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -240
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin VB.Label D1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxxxxxxxxx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      TabIndex        =   30
      Top             =   2160
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      Height          =   3495
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxxxxxxxxx"
      DataField       =   "branch_name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      TabIndex        =   23
      Top             =   1920
      Width           =   1245
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxxxxxxxxx"
      DataField       =   "end_insurance_date"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      TabIndex        =   22
      Top             =   1680
      Width           =   1245
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxxxxxxxxx"
      DataField       =   "estmara_date_end"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      TabIndex        =   21
      Top             =   1440
      Width           =   1245
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxxxxxxxxx"
      DataField       =   "KM_counter"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      TabIndex        =   20
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxxxxxxxxx"
      DataField       =   "driver_name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxxxxxxxxx"
      DataField       =   "id"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Car_no 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxxxxx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.Label inventory_id 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ãÚáæãÇÊ Úä ÇáãÚÏå"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Car_info_bar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim first_run As Boolean


Private Sub Form_Activate()
On Error Resume Next

If my_language = "E" Then
Frame2.Visible = True
Label1.Caption = "õEquipments  information"
d2.RightToLeft = False
Else
Frame1.Visible = True
End If

'On Error Resume Next
If Car_no.Caption = "" Then Unload Me
If first_run = False Then
first_run = True

           
Adodc1.RecordSource = "select * from  CARS where  Car_no='" & Car_no.Caption & "'"
Adodc1.Refresh

Adodc2.RecordSource = "select * from  maintenance where  Car_no='" & Car_no.Caption & "' order by opr_id"
Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
Adodc2.Recordset.MoveLast

d1.Caption = Adodc2.Recordset.Fields!opr_date
d2.Text = Adodc2.Recordset.Fields!error_description
Else
d1.Caption = " "
d2.Text = """"
End If

End If
End Sub

Private Sub Form_Load()
On Error Resume Next

   Me.Left = (MDIForm1.Width - Me.Width) / 2
    Me.Top = (MDIForm1.Height - Me.Height) / 2 - 500
    
    Adodc1.ConnectionString = connection_string
Adodc1.CommandType = adCmdText
          
         Adodc2.ConnectionString = connection_string
Adodc2.CommandType = adCmdText


          
          
End Sub

 

Private Sub item_code_Change()
 On Error Resume Next

If Car_no.Caption = "" Then Unload Me


 
Adodc1.RecordSource = "select * from  CARS where  Car_no='" & Car_no.Caption & "'"
Adodc1.Refresh
              
Adodc2RecordSource = "select * from  maintenance where  Car_no='" & Car_no.Caption & "' order by opr_id"
Adodc2.Refresh

If Adodc2.Recordset.RecordCount > 0 Then
Adodc2.Recordset.MoveLast

d1.Caption = Adodc2.Recordset.Fields!opr_date
d2.Text = Adodc2.Recordset.Fields!error_description
Else
d1.Caption = " "
d2.Text = """"
End If
 
End Sub

 

Private Sub Form_Unload(Cancel As Integer)
first_run = False
End Sub

Private Sub Label9_Click()
On Error Resume Next

Unload Me
End Sub

Private Sub Timer1_Timer()
On Error Resume Next

Unload Me

End Sub
