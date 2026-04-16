VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMTRansferData 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "äÞá ÇáÈíÇäÇÊ"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1935
   ScaleWidth      =   4155
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
   Begin VB.TextBox TxtPOSDB 
      Height          =   375
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Text            =   "Client DB"
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox TxtServerDataBaseName 
      Height          =   375
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Text            =   "serverdb"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "äÞá ÇáÈíÇäÇÊ"
      Height          =   375
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   7335
      Left            =   11760
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   -360
      Width           =   6135
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   5280
         Top             =   3120
      End
      Begin VB.CommandButton CmdOpen 
         Caption         =   "..."
         Height          =   375
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   5040
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox TxtCHECKTIME 
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Text            =   "CHECKTIME Field Name"
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox TxtUSERID 
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Text            =   "USERID Field Name"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox TxtTableName 
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Text            =   "TableName"
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtDbPath 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Text            =   "Database Path"
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         Caption         =   "äÞá ÇáÈíÇäÇÊ"
         Height          =   375
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   3240
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker DcTime 
         Height          =   330
         Left            =   1800
         TabIndex        =   14
         Top             =   2520
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   90767362
         CurrentDate     =   38784
      End
      Begin VB.Label LblInfo 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   3600
         Width           =   3735
      End
      Begin VB.Label Label5 
         Caption         =   "Update Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Date/Time Field"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Machhine Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Table Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "DB Path"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox TxtLicense 
      Alignment       =   1  'Right Justify
      Height          =   1095
      Left            =   -840
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   7800
      Width           =   7935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   10560
      Top             =   6960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "POS DB"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Server DB"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "FRMTRansferData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
  
 
Private Sub CmdOpen_Click()
cd1.ShowOpen
 
txtDbPath.Text = cd1.FileName


End Sub

Private Sub Command1_Click()
Dim strsql As String
Dim ServerDb As String
Dim POSDb As String

 

'     SysSQLServerUserId = GetSetting("Byte_DBS", "Setting", "SysSQLServerUserId", "salim")
'    SysSQLServerUserpassword = GetSetting("Byte_DBS", "Setting", "SysSQLServerUserpassword", "salim")
 
 


ServerDb = TxtServerDataBaseName.Text
POSDb = TxtPOSDB.Text


 
   Set toConnection = New ADODB.Connection
    With toConnection
        .CommandTimeout = 60
        .CursorLocation = adUseClient
        .ConnectionTimeout = 15
       If SysSQLServerType = 1 Then
        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
        "Persist Security Info=False;Initial Catalog=" & POSDb & _
        ";Data Source=" & SysSQLServerName & ";Port=1433"
        
        ElseIf SysSQLServerType = 2 Then
 
     
                 If SysSQLServerTypeTechnical = "0" Then
                 .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
                    "Persist Security Info=False;Initial Catalog=" & POSDb & _
                    ";Data Source=" & SysSQLServerName & ";Port=1433"
                  Else
                     .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & SysSQLServerName
                End If
          End If

.Open
End With
GoTo Transactions
'ÇáÇÚÏÇÏÇÊ
strsql = "    delete  " & POSDb & "..branches" 'done
Cn.Execute strsql

strsql = "   INSERT INTO " & POSDb & "..branches" 'done
strsql = strsql & "  SELECT * FROM " & ServerDb & "..branches"
Cn.Execute strsql


'ÇáÝÊÑÇÊ
strsql = "    delete  " & POSDb & "..Tblyearsdata" 'done
Cn.Execute strsql

strsql = "   INSERT INTO " & POSDb & "..Tblyearsdata" 'done
strsql = strsql & "  SELECT * FROM " & ServerDb & "..Tblyearsdata"
Cn.Execute strsql

'ÊÝÇÕíá ÇáÝÊÑÇÊ

strsql = "    delete  " & POSDb & "..TblAccountIntervals" 'done
Cn.Execute strsql

strsql = "   INSERT INTO " & POSDb & "..TblAccountIntervals" 'done
strsql = strsql & "  SELECT * FROM " & ServerDb & "..TblAccountIntervals"
Cn.Execute strsql





'ØÇáÈäæß
'strsql = "    delete  " & POSDb & "..BanksData" 'done
'Cn.Execute strsql

strsql = "   INSERT INTO " & POSDb & "..BanksData" 'done
strsql = strsql & "  SELECT * FROM " & ServerDb & "..BanksData"
strsql = strsql & " where BankID not in("
strsql = strsql & " select BankID"
strsql = strsql & " from   " & POSDb & "..BanksData"
strsql = strsql & " )"


Cn.Execute strsql

'ÇáÎÒä
'strsql = "    delete  " & POSDb & "..TblBoxesData" 'done
'Cn.Execute strsql

strsql = "   INSERT INTO " & POSDb & "..TblBoxesData" 'done
strsql = strsql & "  SELECT * FROM " & ServerDb & "..TblBoxesData"
strsql = strsql & " where BoxID not in("
strsql = strsql & " select BoxID"
strsql = strsql & " from   " & POSDb & "..TblBoxesData"
strsql = strsql & " )"
Cn.Execute strsql

'UNITS

strsql = "    delete  " & POSDb & "..TblUnites" 'done
Cn.Execute strsql


strsql = "   INSERT INTO " & POSDb & "..TblUnites" 'done
strsql = strsql & "  SELECT * FROM " & ServerDb & "..TblUnites"
Cn.Execute strsql



'Coding

strsql = "    delete  " & POSDb & "..Coding" 'done
Cn.Execute strsql

'
 


strsql = "   INSERT INTO " & POSDb & "..Coding" 'done
strsql = strsql & " ( [FIELD_no] , [FIELD_NAME], [Name], [departement], [branch_no], [no_of_digit], [Auto], [Zeros], [prifix] )"

strsql = strsql & "  SELECT    "
strsql = strsql & "  [FIELD_no] , [FIELD_NAME], [Name], [departement], [branch_no], [no_of_digit], [Auto], [Zeros], [prifix] "

strsql = strsql & "  FROM " & ServerDb & "..Coding"
 Cn.Execute strsql

'sanad_numbering

strsql = "    delete  " & POSDb & "..sanad_numbering" 'done
Cn.Execute strsql

strsql = "   INSERT INTO " & POSDb & "..sanad_numbering" 'done
strsql = strsql & "( [sanad_no], [sanad_type], [numbering_id], [numbering_type], [branch_no], [no_of_digit], [zeros], [start_at], [departement], [end_at], [BranchName], [Prefix], [StoreCoding], [YearDigit] )"

strsql = strsql & "  SELECT    "
strsql = strsql & " [sanad_no], [sanad_type], [numbering_id], [numbering_type], [branch_no], [no_of_digit], [zeros], [start_at], [departement], [end_at], [BranchName], [Prefix], [StoreCoding], [YearDigit] "

strsql = strsql & "  FROM " & ServerDb & "..sanad_numbering"
Cn.Execute strsql


strsql = "    delete  " & POSDb & "..tblActivitesType" 'done
Cn.Execute strsql

strsql = "INSERT INTO " & POSDb & "..tblActivitesType "  'done
strsql = strsql & " SELECT * FROM " & ServerDb & "..tblActivitesType"
 
Cn.Execute strsql


 
strsql = "INSERT INTO " & POSDb & "..TblBranchesData "  'done
strsql = strsql & " SELECT * FROM " & ServerDb & "..TblBranchesData"
strsql = strsql & " where branch_id not in("
strsql = strsql & " select branch_id"
strsql = strsql & " from   " & POSDb & "..TblBranchesData"
strsql = strsql & " )"

Cn.Execute strsql


strsql = " INSERT INTO " & POSDb & "..TblStore" 'done
strsql = strsql & " SELECT * FROM " & ServerDb & "..TblStore"
strsql = strsql & "  where StoreID not in("
strsql = strsql & "  select StoreID"
strsql = strsql & "  from   " & POSDb & "..TblStore"
strsql = strsql & "  )"
 Cn.Execute strsql
 
 
 strsql = "  INSERT INTO " & POSDb & "..Groups" 'done
strsql = strsql & " SELECT * FROM " & ServerDb & "..Groups"
strsql = strsql & "  where GroupID not in("
strsql = strsql & " select GroupID"
strsql = strsql & " from   " & POSDb & "..Groups"
strsql = strsql & " )"
 Cn.Execute strsql
 
strsql = "   INSERT INTO " & POSDb & "..TblUnites" 'done
strsql = strsql & " SELECT * FROM " & ServerDb & "..TblUnites"
strsql = strsql & " where UnitID not in("
strsql = strsql & " select UnitID"
strsql = strsql & " from   " & POSDb & "..TblUnites"
strsql = strsql & " )"
Cn.Execute strsql
 
strsql = "   INSERT INTO " & POSDb & "..TblItems" 'done
strsql = strsql & " SELECT * FROM " & ServerDb & "..TblItems"
strsql = strsql & " where ItemID not in("
strsql = strsql & " select ItemID"
strsql = strsql & " from   " & POSDb & "..TblItems"
strsql = strsql & " )"
Cn.Execute strsql


strsql = "    delete  " & POSDb & "..TblItemsUnits" 'done
Cn.Execute strsql

strsql = "   INSERT INTO " & POSDb & "..TblItemsUnits" 'done
strsql = strsql & "  SELECT * FROM " & ServerDb & "..TblItemsUnits"
Cn.Execute strsql


strsql = "    INSERT INTO " & POSDb & ".dbo.TblItemsParts" 'done
strsql = strsql & "  ( ItemID,PartItemID,PartItemQty,PartItemPrice,Unitid )"
strsql = strsql & "  SELECT  ItemID,PartItemID,PartItemQty,PartItemPrice,Unitid"
 strsql = strsql & "  From " & ServerDb & ".dbo.TblItemsParts"
Cn.Execute strsql


strsql = "    INSERT INTO " & POSDb & "..TblItemsColors" 'done
strsql = strsql & " SELECT * FROM " & ServerDb & "..TblItemsColors"
strsql = strsql & " where ColorID not in("
strsql = strsql & " select ColorID"
strsql = strsql & " from   " & POSDb & "..TblItemsColors"
strsql = strsql & " )"
Cn.Execute strsql


strsql = "    INSERT INTO " & POSDb & "..TblItemsSizes" 'done
strsql = strsql & " SELECT * FROM " & ServerDb & "..TblItemsSizes"
strsql = strsql & " where SizeId not in("
strsql = strsql & " select SizeId"
strsql = strsql & " from   " & POSDb & "..TblItemsSizes"
strsql = strsql & " )"
Cn.Execute strsql
 
 
 strsql = "    INSERT INTO " & POSDb & "..TblItemsclasses" 'done
strsql = strsql & "  SELECT * FROM " & ServerDb & "..TblItemsclasses"
strsql = strsql & "  where SizeId not in("
strsql = strsql & " select SizeId"
strsql = strsql & " from   " & POSDb & "..TblItemsclasses"
strsql = strsql & " )"
Cn.Execute strsql




strsql = "    delete  " & POSDb & "..TblItemsAttach" 'done
Cn.Execute strsql
strsql = "    INSERT INTO " & POSDb & ".dbo.TblItemsAttach" 'done
strsql = strsql & " ( ItemID,AttachItemID,AttachItemQty,AttachItemPrice )"
strsql = strsql & " SELECT   ItemID,AttachItemID,AttachItemQty,AttachItemPrice"
 strsql = strsql & " From " & ServerDb & ".dbo.TblItemsAttach"
Cn.Execute strsql


strsql = "   INSERT INTO " & POSDb & "..TblCustemers"
strsql = strsql & " SELECT * FROM " & ServerDb & "..TblCustemers"
strsql = strsql & " where CusID not in("
strsql = strsql & " select CusID"
strsql = strsql & " from   " & ServerDb & "..TblCustemers"
strsql = strsql & " )"
Cn.Execute strsql


strsql = strsql & " INSERT INTO " & POSDb & "..ACCOUNTS"
strsql = strsql & " ( [Account_Code], [Account_Name], [Parent_Account_Code], [last_account], [cannot_del], [Account_Serial], [BasicAccount], [DateCreated], [Account_NameEng], [zmam], [mowazna], [currenct_code], [rate], [cost_center], [Sum_account], [opening_balance], [opening_balance_type], [cost_center_id], [cost_center_type], [ShowInBlanceSheet], [ActivityTypeId], [AccountTypes], [AccountTab], [DepitOrCredit], [Differenttype], [Authority], [Block], [UserGroupId], [UserId], [Branch], [Balance], [DepitBalance], [CreditBalance], [interval1], [interval2], [interval3], [interval4], [interval5], [interval6], [interval7], [interval8], [interval9], [interval10], [interval11], [interval12] )"
strsql = strsql & " SELECT   [Account_Code], [Account_Name], [Parent_Account_Code], [last_account], [cannot_del], [Account_Serial], [BasicAccount], [DateCreated], [Account_NameEng], [zmam], [mowazna], [currenct_code], [rate], [cost_center], [Sum_account], [opening_balance], [opening_balance_type], [cost_center_id], [cost_center_type], [ShowInBlanceSheet], [ActivityTypeId], [AccountTypes], [AccountTab], [DepitOrCredit], [Differenttype], [Authority], [Block], [UserGroupId], [UserId], [Branch], [Balance], [DepitBalance], [CreditBalance], [interval1], [interval2], [interval3], [interval4], [interval5], [interval6], [interval7], [interval8], [interval9], [interval10], [interval11], [interval12]"
strsql = strsql & " FROM  " & ServerDb & "..ACCOUNTS"
strsql = strsql & " where Account_Code not in("
strsql = strsql & " select Account_Code"
strsql = strsql & " from   " & POSDb & "..ACCOUNTS"
strsql = strsql & " )"

Cn.Execute strsql



  

 

strsql = strsql & " INSERT INTO " & POSDb & "..TblEmpJobsTypes"
strsql = strsql & " SELECT * FROM " & ServerDb & "..TblEmpJobsTypes"
strsql = strsql & " where JobTypeID not in("
strsql = strsql & " select JobTypeID"
strsql = strsql & " from   " & POSDb & "..TblEmpJobsTypes"
strsql = strsql & " )"
Cn.Execute strsql


strsql = strsql & " INSERT INTO " & POSDb & "..TBLSalesRepGroups"
strsql = strsql & " SELECT * FROM " & ServerDb & "..TBLSalesRepGroups"
strsql = strsql & " where id not in("
strsql = strsql & " select id"
strsql = strsql & " from   " & POSDb & "..TBLSalesRepGroups"
strsql = strsql & " )"
Cn.Execute strsql

strsql = strsql & " INSERT INTO " & POSDb & "..TBLSalesRepData"
strsql = strsql & " SELECT * FROM " & ServerDb & "..TBLSalesRepData"
strsql = strsql & " where id not in("
strsql = strsql & " select id"
strsql = strsql & " from   " & POSDb & "..TBLSalesRepData)"

Cn.Execute strsql

strsql = strsql & " INSERT INTO " & POSDb & "..TBLSalesRepData1"
strsql = strsql & " SELECT * FROM " & ServerDb & "..TBLSalesRepData1"
strsql = strsql & " where id not in("
strsql = strsql & " select id"
strsql = strsql & " from   " & POSDb & "..TBLSalesRepData1)"
Cn.Execute strsql

strsql = strsql & " INSERT INTO " & POSDb & "..TBLSalesRepData2"
strsql = strsql & " SELECT * FROM " & ServerDb & "..TBLSalesRepData2"
strsql = strsql & " where id not in("
strsql = strsql & " select id"
strsql = strsql & " from   " & POSDb & "..TBLSalesRepData2)"
Cn.Execute strsql


Transactions:

'////////////////////////////////////////copy Sales Transactions

  Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Sql As String
    Dim mytext As String
    
 Sql = " select * from Transactions    WHERE  Copied is null And Transaction_Type = 21 "
 
    Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 Dim FromTransaction_ID As Double
 Dim FromBranchID As Integer
 Dim FromTransaction_Date As Date
 Dim FromNots As String
  Dim FromNots2 As String
   Dim fromTransaction_Serial As String
 Dim FromNoteseial1 As String
 Dim FromTransaction_Type As Integer
 
Dim BranchID As Integer
  Dim Transaction_ID As Double
  Dim Transaction_Date As Date
   Dim Transaction_Serial  As String
 Dim Nots As String
Dim Nots2 As String
Dim Transaction_Type As Integer

 'sales
    If Rs3.RecordCount > 0 Then
      
            For I = 1 To Rs3.RecordCount
             FromTransaction_Type = IIf(IsNull(Rs3("Transaction_Type").value), 0, Rs3("Transaction_Type").value)
               FromTransaction_ID = IIf(IsNull(Rs3("Transaction_ID").value), 0, Rs3("Transaction_ID").value)
               FromBranchID = IIf(IsNull(Rs3("BranchID").value), 0, Rs3("BranchID").value)
               fromTransaction_Serial = IIf(IsNull(Rs3("Transaction_Serial").value), 0, Rs3("Transaction_Serial").value)
              
               Noteseial1 = IIf(IsNull(Rs3("Noteserial1").value), 0, Rs3("Noteserial1").value)
               Noteseial = IIf(IsNull(Rs3("Noteserial").value), 0, Rs3("Noteserial").value)
               FromNots = IIf(IsNull(Rs3("Nots").value), 0, Rs3("Nots").value)
               FromNots2 = IIf(IsNull(Rs3("Nots2").value), 0, Rs3("Nots2").value)
               FromTransaction_Date = IIf(IsNull(Rs3("Transaction_Date").value), 0, Rs3("Transaction_Date").value)
              
              Transaction_Date = FromTransaction_Date
              Transaction_Type = FromTransaction_Type
              BranchID = FromBranchID
             Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
             Transaction_Serial = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=21"))
             Noteseial1 = Voucher_coding(BranchID, FromTransaction_Date, 7, 170, , 21)
             Noteseial = Notes_coding(BranchID, FromTransaction_Date)
             NoteID = CStr(new_id("Notes", "NoteID", "", True))

             '" & ServerDb & "
             
Sql = "INSERT INTO  " & POSDb & "..Transactions  (     " & Transaction_ID & ", " & Transaction_Serial & ", " & Transaction_Date & ", Transaction_Type, PaymentType, Transaction_HijriDate, Trans_Discount, Trans_DiscountType, CusID, StoreID, "
Sql = Sql & "  TaxFound, TaxValue, UserID, ReturnID, MaintenanceID, Emp_ID, ExtraOptions, MainOperationID, SaleType, CashCustomerName, CashCustomerPhone,"
Sql = Sql & "                        CashCustomerMobile, CashCustomerAddress, CashCustomerComment, TransactionComment, TaxAddValue, TaxStampValue, TaxServiceValue, WorkOrderID,"
Sql = Sql & "  ItemMaking, ItemMakingQty, ItemMakingCost, " & Nots & ", Shahne, " & Nots2 & ", warsha, project, empid, Shipment_no, order_no, Currency_id, Total, Closed, shipped, project_id,"
Sql = Sql & "                        opr_fullcode, ToTAlELSHahn, opening_balance_voucher_id, Currency_rate, total_expenses, total_payments, NoteSerial, NoteSerial1, NoteId, countryid, remark,"
                      Sql = Sql & "  shipmentMethod, ShipmentPrice, ShipmentArae, StoreID1, Transaction_Type_Sub, Product_Issue_voucher_Serial, Product_Receive_voucher_Serial, WorkHour,"
Sql = Sql & "                        startDate, EndDate, startTime, EndTime, TotalMaterials, LineExpenses, workerTotal, Expenses, FinacilaTotal, FactoryExpenses, ArrivalDate, BranchId, PayedValue,"
                      Sql = Sql & "  NetValue, RemainValue, PaymentNetid, NetTransactionNo, BillBasedOn, ReturnSerial, BoxID, SalesInvoiceDate, WorkOrderNO, EstimatedCost, ReturnType,"
Sql = Sql & "                        GardFromDate, GardTodate, GardEntryType, StartGard, StartSetelment, Account1, Account2, DifferentAccounts, ManualNO, CBoBasedON, OldNoteSerial1, Doctype,"
                      Sql = Sql & "  ExtraAccount, ExtraValue, ManualNo1, ManualNo2, LcNo, ChkInstall, ChkReturnPurcahse, Posted, PostedDate, POSBillType, STableID, SessionD, Printed,"
Sql = Sql & "                        DamageOrSample, Prefix, Fullcode, TicketNO, general_cost_center, IndirectCostForProduction, TotalEstimatedCost, PurchaseBill, CarId, DriverId, ShipmentStatus,"
                      Sql = Sql & "  ContryId, Cityid, Neighborhoodid, Streetid, DetailedAddress, PONo, Approved, PODays, Days1, Days2, ReciveDate, InternalFlag, ProductionPlanno, OPrType,"
Sql = Sql & "                        OrderType, Phone, DepartementID, FixesAssetsID, RegionID, Address, Enterdate, EnterTime, ContactPhone, ContactTime, EqamaNo, KMOut, Transporter, GoogleMap,"
                      Sql = Sql & "  DepartureDate, DepartureTime, oorderdate, ArrivalTime, KMIn, PolicyNo, ReciveOrderO, Transporterdriver, InspectionReport, empID2, ProductionTypeid, empID1,"
Sql = Sql & "                        Station, CompsBill, purchaseType, Tms_Oper_ID, ModeSupply, ModeReceptEq, DeptID, EndShow, DMYEShow, CountRecept, DMYRecept, ReceptMode, NoteSerial2,"
                      Sql = Sql & "  OrderID, PaymentT, order_no1, ShipingID, PaymentID, MixID, MIxCode, ProductionOrder, ResProductionNo, ProkerId, Priod, PriodDMY, SippingDate, DeliverDate,"
Sql = Sql & "                        OrderSupply, BillSupplier, ReasonReturns, NotSeialPO6, requestOrOrder, employeeDiscount, Without, Wait, shipmentType, Inspection, CusID1, CarTypeID,"
                      Sql = Sql & "  ShippingTypeID, ShippingStatus, Shipping_Pos, FromDate, ToDate, ReqStatus, Product_Receive_voucher_Serial22, nots22, opr_Employee, LocationID,"
Sql = Sql & "                        LockedInterval, Sandts, TotalQest, QstValue, QstNo, QestStartDateH, QestEndtDateH, QestStartDate, QestEndtDate, YMD, LawFirmValue, OpOrderID, OldOpOrderID,"
                      Sql = Sql & "  TotalPayed , AdvPay, SpecialOffer, CustomerlocationID, 1)"
'Sql = Sql & "   From dbo.Transactions"
     
     Sql = " SELECT     " & Transaction_ID & ", " & Transaction_Serial & ", Transaction_Date, Transaction_Type, PaymentType, Transaction_HijriDate, Trans_Discount, Trans_DiscountType, CusID, StoreID, "
Sql = Sql & "   TaxFound, TaxValue, UserID, ReturnID, MaintenanceID, Emp_ID, ExtraOptions, MainOperationID, SaleType, CashCustomerName, CashCustomerPhone,"
Sql = Sql & "                        CashCustomerMobile, CashCustomerAddress, CashCustomerComment, TransactionComment, TaxAddValue, TaxStampValue, TaxServiceValue, WorkOrderID,"
Sql = Sql & "  ItemMaking, ItemMakingQty, ItemMakingCost, Nots, Shahne, Nots2, warsha, project, empid, Shipment_no, order_no, Currency_id, Total, Closed, shipped, project_id,"
Sql = Sql & "                        opr_fullcode, ToTAlELSHahn, opening_balance_voucher_id, Currency_rate, total_expenses, total_payments, NoteSerial, NoteSerial1, NoteId, countryid, remark,"
                      Sql = Sql & "  shipmentMethod, ShipmentPrice, ShipmentArae, StoreID1, Transaction_Type_Sub, Product_Issue_voucher_Serial, Product_Receive_voucher_Serial, WorkHour,"
Sql = Sql & "                        startDate, EndDate, startTime, EndTime, TotalMaterials, LineExpenses, workerTotal, Expenses, FinacilaTotal, FactoryExpenses, ArrivalDate, BranchId, PayedValue,"
                      Sql = Sql & "  NetValue, RemainValue, PaymentNetid, NetTransactionNo, BillBasedOn, ReturnSerial, BoxID, SalesInvoiceDate, WorkOrderNO, EstimatedCost, ReturnType,"
Sql = Sql & "                        GardFromDate, GardTodate, GardEntryType, StartGard, StartSetelment, Account1, Account2, DifferentAccounts, ManualNO, CBoBasedON, OldNoteSerial1, Doctype,"
                      Sql = Sql & "  ExtraAccount, ExtraValue, ManualNo1, ManualNo2, LcNo, ChkInstall, ChkReturnPurcahse, Posted, PostedDate, POSBillType, STableID, SessionD, Printed,"
Sql = Sql & "                        DamageOrSample, Prefix, Fullcode, TicketNO, general_cost_center, IndirectCostForProduction, TotalEstimatedCost, PurchaseBill, CarId, DriverId, ShipmentStatus,"
                      Sql = Sql & "  ContryId, Cityid, Neighborhoodid, Streetid, DetailedAddress, PONo, Approved, PODays, Days1, Days2, ReciveDate, InternalFlag, ProductionPlanno, OPrType,"
Sql = Sql & "                        OrderType, Phone, DepartementID, FixesAssetsID, RegionID, Address, Enterdate, EnterTime, ContactPhone, ContactTime, EqamaNo, KMOut, Transporter, GoogleMap,"
                      Sql = Sql & "  DepartureDate, DepartureTime, oorderdate, ArrivalTime, KMIn, PolicyNo, ReciveOrderO, Transporterdriver, InspectionReport, empID2, ProductionTypeid, empID1,"
Sql = Sql & "                        Station, CompsBill, purchaseType, Tms_Oper_ID, ModeSupply, ModeReceptEq, DeptID, EndShow, DMYEShow, CountRecept, DMYRecept, ReceptMode, NoteSerial2,"
                      Sql = Sql & "  OrderID, PaymentT, order_no1, ShipingID, PaymentID, MixID, MIxCode, ProductionOrder, ResProductionNo, ProkerId, Priod, PriodDMY, SippingDate, DeliverDate,"
Sql = Sql & "                        OrderSupply, BillSupplier, ReasonReturns, NotSeialPO6, requestOrOrder, employeeDiscount, Without, Wait, shipmentType, Inspection, CusID1, CarTypeID,"
                      Sql = Sql & "  ShippingTypeID, ShippingStatus, Shipping_Pos, FromDate, ToDate, ReqStatus, Product_Receive_voucher_Serial22, nots22, opr_Employee, LocationID,"
Sql = Sql & "                        LockedInterval, Sandts, TotalQest, QstValue, QstNo, QestStartDateH, QestEndtDateH, QestStartDate, QestEndtDate, YMD, LawFirmValue, OpOrderID, OldOpOrderID,"
                      Sql = Sql & "  TotalPayed , AdvPay, SpecialOffer, CustomerlocationID, Copied"

Sql = Sql & "   From " & ServerDb & "..dbo.Transactions"
Sql = Sql & " where   Transaction_ID=" & FromTransaction_ID
 Cn.Execute Sql
     
     

         
        
        
            Next I
     
      End If
 
    Rs3.Close
 
MsgBox "Done"
'////////////////////////////////////////copy Sales Transactions

End Sub

Private Sub Form_Load()
 On Error Resume Next
txtDbPath = GetSetting("ConvertToAccess", "Setting", "DbPath", "DatabasePath")
TxtTableName = GetSetting("ConvertToAccess", "Setting", "TableName", "TableName")
TxtUSERID = GetSetting("ConvertToAccess", "Setting", "USERID", "USERID")
TxtCHECKTIME = GetSetting("ConvertToAccess", "Setting", "CHECKTIME", "CHECKTIME")
DcTime.value = GetSetting("ConvertToAccess", "Setting", "UpdateHours", "00")

TxtServerDataBaseName = SysSQLServerDataBaseName
Exit Sub
End Sub

 
