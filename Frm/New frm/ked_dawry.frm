VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ked_dawry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   «‰‘«¡ ﬁÌœ œÊ—Ì"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8415
   Icon            =   "ked_dawry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   8415
   Begin VB.TextBox TXTSERIAL 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1080
      Width           =   2415
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      Caption         =   "ÌœÊÌ"
      Height          =   255
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3960
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ì"
      Height          =   255
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "ked_dawry.frx":000C
      Left            =   4560
      List            =   "ked_dawry.frx":0019
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox desc 
      Alignment       =   1  'Right Justify
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1560
      Width           =   6975
   End
   Begin VB.TextBox id 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   7200
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«‰‘«¡ «·› —« "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "ked_dawry.frx":002C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "ked_dawry.frx":0048
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777215
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "index"
         Caption         =   "„"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "ItemCode"
         Caption         =   "—ﬁ„ «·ﬁÌœ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "ked_date"
         Caption         =   " «—ÌŒ «·ﬁÌœ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "ItemID"
         Caption         =   "—ﬁ„ «·ﬁÿ⁄…"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "SallingPrice"
         Caption         =   "”⁄— «·»Ì⁄"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "akher_s3r_shera"
         Caption         =   "akher_s3r_shera"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "motwaset_taklefa"
         Caption         =   "motwaset_taklefa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "blocked"
         Caption         =   "blocked"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "akher_shera_date"
         Caption         =   "akher_shera_date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "akher_be3_date"
         Caption         =   "akher_be3_date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "akher_sarf_date"
         Caption         =   "akher_sarf_date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "hesab_taklefa_method"
         Caption         =   "hesab_taklefa_method"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   -1  'True
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4995.213
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1484.787
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   330
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   582
      _Version        =   393216
      Format          =   97714177
      CurrentDate     =   38784
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1032
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
      Caption         =   " Õ—Ìﬂ"
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
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Left            =   6120
      TabIndex        =   19
      Top             =   7200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ–›"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   255
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "ked_dawry.frx":005D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton3 
      Height          =   375
      Left            =   4680
      TabIndex        =   20
      Top             =   7200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ–› «·ﬂ·"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   255
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "ked_dawry.frx":0079
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "œ› — «·”‰œ"
      Height          =   255
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "«·› —… »Ì‰ «·ﬁÌÊœ"
      Height          =   255
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«» œ«¡ „‰"
      Height          =   495
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "⁄œœ „—«  «·ﬁÌœ"
      Height          =   495
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Ê’› «·ﬁÌœ  "
      Height          =   495
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·ﬁÌœ "
      Height          =   495
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "   «‰‘«¡ ﬁÌœ œÊ—Ì"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   735
      Index           =   2
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   8400
   End
End
Attribute VB_Name = "ked_dawry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ALLButton1_Click()
    Dim date1 As Date
    Dim interval As String

    If SystemOptions.UserInterface = ArabicInterface Then
        If Not IsNumeric(Text3.Text) Then MsgBox "⁄œœ „—«  «· ﬂ—«— ·«»œ «‰  ﬂÊ‰ «—ﬁ«„", vbCritical: Exit Sub
        If Not IsNumeric(Text4.Text) Then MsgBox "«·› —… »Ì‰ «·ﬁÌÊœ ·«»œ «‰  ﬂÊ‰ «—ﬁ«„", vbCritical: Exit Sub
        If Combo1.ListIndex = -1 Then MsgBox "·«»œ „‰  ÕœÌœ «·› —… »Ì‰ «·ﬁÌÊœ", vbCritical: Exit Sub
    Else

        If Not IsNumeric(Text3.Text) Then MsgBox "Insert Numeric in repeated count", vbCritical: Exit Sub
        If Not IsNumeric(Text4.Text) Then MsgBox "interval must be numeric", vbCritical: Exit Sub
        If Combo1.ListIndex = -1 Then MsgBox "  Specify interval day-month-year", vbCritical: Exit Sub

    End If
 
    If Combo1.ListIndex = 0 Then
        interval = "d"
    Else

        If Combo1.ListIndex = 1 Then
            interval = "m"

            If Combo1.ListIndex = 2 Then
                interval = "yyyy"
            End If
        End If
    End If

    Dim mydate As Date
  
Cn.Execute "delete KED_DAWRY WHERE ked_no=" & val(Me.ID.Text)
  Adodc1.Refresh
  DataGrid2.Refresh
    For I = 1 To val(Text3.Text)
    If I = 1 Then
      date1 = XPDtbBill.Value
      mydate = XPDtbBill.Value
   Else
     date1 = DateAdd(interval, Text4.Text, mydate)
    End If
    
      

        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields!ked_no = Me.ID.Text
        Adodc1.Recordset.Fields!ked_serial = Me.TxtSerial.Text
        Adodc1.Recordset.Fields!des = Me.Desc.Text
        Adodc1.Recordset.Fields!ked_date = date1
        Adodc1.Recordset.Fields![Index] = I
        Adodc1.Recordset.Fields!OK = 0
        mydate = date1
 
        Adodc1.Recordset.update

    Next I

End Sub

Private Sub ALLButton2_Click()
Adodc1.Recordset.delete
End Sub

Private Sub ALLButton3_Click()
Cn.Execute "delete KED_DAWRY WHERE ok=0 and  ked_no=" & val(Me.ID.Text)
    Adodc1.Refresh
 Me.DataGrid2.Refresh

End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    XPDtbBill.Value = Now
    
    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "   SELECT  * FROM KED_DAWRY WHERE id=" & val(Me.ID.Text)
    Adodc1.Refresh
    XPDtbBill.Value = Date


    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    
    End If

End Sub
Public Function loadfunctions(ID As Integer)
 
    Adodc1.RecordSource = "   SELECT  * FROM KED_DAWRY WHERE  ok=0 and  ked_no=" & ID
    Adodc1.Refresh
 Me.DataGrid2.Refresh

End Function
Private Sub ChangeLang()
    Me.Caption = "Create Repeated Voucher"
    Label1(2).Caption = Me.Caption
    Label2.Caption = "Voucher No."
    Label3.Caption = "Voucher Desc."
    Label4.Caption = "Repeated Count"
    Label6.Caption = "Interval"
    Label5.Caption = "Start From"
    Combo1.Clear
    Combo1.AddItem "Day"
    Combo1.AddItem "Month"
    Combo1.AddItem "Year"

    Option1.Caption = "Auto"
    Option2.Caption = "Manual"

    ALLButton1.Caption = "Create Intervals"

    DataGrid2.RightToLeft = False

    DataGrid2.Columns(0).Caption = "index"
    DataGrid2.Columns(2).Caption = "Voucher Date"

End Sub
