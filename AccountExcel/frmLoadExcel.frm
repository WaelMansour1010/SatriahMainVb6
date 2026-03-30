VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmLoadExcel 
   Caption         =   "Form1"
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRowNo 
      Height          =   495
      Left            =   8385
      TabIndex        =   15
      Top             =   2505
      Width           =   2955
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   2460
      Left            =   735
      TabIndex        =   14
      Top             =   5310
      Width           =   9885
      _cx             =   17436
      _cy             =   4339
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton btnTest 
      Caption         =   "Test Data"
      Height          =   525
      Left            =   1575
      TabIndex        =   13
      Top             =   2535
      Width           =   2265
   End
   Begin VB.ComboBox cmbParent 
      Height          =   315
      Left            =   2055
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   4680
      Width           =   1470
   End
   Begin VB.ComboBox cmbEnglishName 
      Height          =   315
      Left            =   2070
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   4290
      Width           =   1470
   End
   Begin VB.ComboBox cmbArabicName 
      Height          =   315
      Left            =   2085
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   4035
      Width           =   1470
   End
   Begin VB.ComboBox cmbCode 
      Height          =   315
      Left            =   2115
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   3495
      Width           =   1470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   525
      Left            =   4275
      TabIndex        =   4
      Top             =   2520
      Width           =   2265
   End
   Begin VB.TextBox TxtFileName 
      Height          =   405
      Left            =   1755
      TabIndex        =   1
      Top             =   1335
      Width           =   11205
   End
   Begin VB.CommandButton btnLoadFile 
      Caption         =   "Load File"
      Height          =   525
      Left            =   1635
      TabIndex        =   0
      Top             =   630
      Width           =   2265
   End
   Begin MSComDlg.CommonDialog CdgOpen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Caption         =   "Start From Row"
      Height          =   225
      Left            =   6750
      TabIndex        =   16
      Top             =   2640
      Width           =   1440
   End
   Begin VB.Label Label5 
      Caption         =   "Parent"
      Height          =   435
      Left            =   270
      TabIndex        =   12
      Top             =   4680
      Width           =   1515
   End
   Begin VB.Label Label4 
      Caption         =   "English Name"
      Height          =   435
      Left            =   270
      TabIndex        =   11
      Top             =   4290
      Width           =   1515
   End
   Begin VB.Label Label3 
      Caption         =   "Arabic Name"
      Height          =   435
      Left            =   270
      TabIndex        =   10
      Top             =   4050
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "Code Colmn"
      Height          =   435
      Left            =   270
      TabIndex        =   9
      Top             =   3495
      Width           =   1515
   End
   Begin VB.Label lblLog 
      Height          =   435
      Left            =   1665
      TabIndex        =   3
      Top             =   1815
      Width           =   11250
   End
   Begin VB.Label Label1 
      Caption         =   "File"
      Height          =   435
      Left            =   105
      TabIndex        =   2
      Top             =   1395
      Width           =   1515
   End
End
Attribute VB_Name = "frmLoadExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tablename As String
Private Sub btnLoadFile_Click()
   If Right(Trim(TxtFileName), 4) = ".xls" And Dir(Trim(TxtFileName)) <> "" Then
        CdgOpen.filename = Trim(TxtFileName)
    Else
        CdgOpen.InitDir = Trim(TxtFileName)
    End If
    CdgOpen.ShowOpen

    If CdgOpen.filename <> "" Then
        TxtFileName = CdgOpen.filename
       PopulateRS TxtFileName
     
    End If
End Sub
'
'Public Function OpenExcelRS(FileName) As ADODB.Recordset
'    On Error GoTo EH
'
'    Dim firstSheetName As String
'    Dim cnn            As ADODB.Connection
'    Dim Rs             As ADODB.Recordset
'    Set cnn = New ADODB.Connection
'    Dim sConn As String
'    Dim sfo   As New Scripting.FileSystemObject
'    Dim ext   As String
'    ext = sfo.GetExtensionName(FileName)
'
'    'sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Mode=Share Exclusive;Data Source=" & FileName & ";Extended Properties='Excel 8.0;IMEX=1'"
'    If ext = "xlsx" Then
'        With cnn
'            .Provider = "Microsoft.ACE.OLEDB.12.0;"
'            .ConnectionString = "Data Source=" & strPOINTDataPath & strSourceFile & _
'               ";Extended Properties=""Excel 12.0 Xml;HDR=NO;IMEX=1"";"
'            .ConnectionTimeout = 40
'            .Open
'        End With
'    Else
'        With cnn
'            .Provider = "Microsoft.ACE.OLEDB.12.0;"
'            .ConnectionString = "Data Source=" & strPOINTDataPath & strSourceFile & _
'               ";Extended Properties=""Excel 12.0 Xml;HDR=NO;IMEX=1"";"
'            .ConnectionTimeout = 40
'            .Open
'        End With
'    End If
'    'cnn.Open sConn
'
'    Set Rs = cnn.OpenSchema(adSchemaTables)
'    If Not Rs.EOF Then
'        firstSheetName = Rs.Fields("table_name").Value
'    Else
'        DisplayError "File Not Valid "
'        Exit Function
'    End If
'    Rs.Close
'    cnn.Close
'    Set Rs = Nothing
'    Set cnn = Nothing
'
'    Dim ExcelRs As ADODB.Recordset
'    Set ExcelRs = New ADODB.Recordset
'    With ExcelRs
'        .CursorLocation = adUseClient
'        .Open "[" & firstSheetName & "]", _
'           "Provider=Microsoft.Jet.OLEDB.4.0;" _
'           & "Data Source='" & FileName & "';" _
'           & "Mode=Share Exclusive;" _
'           & "Extended Properties='Excel 8.0;IMEX=1'", _
'           adOpenStatic, _
'           adLockReadOnly, _
'           adCmdTable
'        Set .ActiveConnection = Nothing
'    End With
'    Set OpenExcelRS = ExcelRs
'    Exit Function
'EH:
'    DisplayError "File My Not Valid " & Err.Description
'    Set OpenExcelRS = Nothing
'End Function

'
'Private Sub cmdLoadFile_Click()
'    On Error GoTo EH
'
'
'
'
'    Dim RsData        As New ADODB.Recordset
'    Dim FileName      As String
'    Dim moConn        As New ADODB.Connection
'    Dim mrs           As ADODB.Recordset
'    Dim tblname       As String
'    Dim shortFileName As String
'    lblTime.Visible = True
'    Me.Enabled = False
'    moConn.CursorLocation = adUseClient
'    Dim rsCheck As New ADODB.Recordset
'
'
'        FileName = TxtFileName
'      '  shortFileName = grdFiles.TextMatrix(i, grdFiles.ColIndex("Name"))
'
'        moConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source='" & FileName & "';" & "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
'        Set mrs = moConn.OpenSchema(adSchemaTables)
'
'        If Not mrs.EOF Then
'            tblname = mrs.Fields("table_name").Value
'            RsData.CursorLocation = adUseClient
'            RsData.Open "Select *   from [" & tblname & "]", moConn, adOpenKeyset, adLockReadOnly
'            Dim RowID          As Integer
'
'            Dim isREcoredSaved As Boolean
'            Dim strQuery       As String
'            Dim OLDSec         As Long
'            Dim Secondes       As Long
'
'            Dim AllSec         As Long
'            lblTime.Visible = True
'            RowID = 0
'            AllFinshedRows = 0
'            Dim currentRows As Long
'            Dim AllFileRows As Long
'            AllFileRows = Val(grdFiles.TextMatrix(i, grdFiles.ColIndex("Rows")))
'
'            Do While Not RsData.EOF
'                RowID = RowID + 1
'                AllFinshedRows = AllFinshedRows + 1
'                currentRows = currentRows + 1
'                ' lbl(32).Caption = "F[" & currentRows & "]>[" & AllFileRows & "] A[" & AllFinshedRows & "]>[" & AllFileRows & "]"
'                DoEvents
'                strQuery = "SELECT Count(*) cnt "
'                strQuery = strQuery & "From notes_all "
'                strQuery = strQuery & "WHERE ExcelFile = '" & shortFileName & "' "
'                strQuery = strQuery & " AND NoteType = 85 "
'                strQuery = strQuery & "  AND ExcelRow =  " & RowID & " ;"
'                ' rsCheck.CursorLocation = adUseClient
'                rsCheck.Open strQuery, Cn, adOpenForwardOnly, adLockReadOnly
'                isREcoredSaved = rsCheck!cnt > 0
'                rsCheck.Close
'
'                '*********************
'                If Not isREcoredSaved Then
'                    StartTime = Now
'
'                    If Trim(RsData.Fields("«· «—ÌŒ").Value & "") <> "" Then
'                        SaveItemsExcelMeth_New RsData, RowID, shortFileName
'                    End If
'
'                    OLDSec = AllSec
'                    Secondes = DateDiff("s", StartTime, Now)
'
'                    AllSec = ((allExcelRows - AllFinshedRows) * Secondes)
'
'                    If AllSec = 0 Then
'                        AllSec = OLDSec
'                    End If
'
'                    lblTime.Caption = StringDotFormat("{0} of {1} Recored(s)  Estimated Time : {2} ", AllFinshedRows, allExcelRows, GetTimeHour(AllSec))
'                End If
'
'                '*********************
'                RsData.MoveNext
'            Loop
'
'            RsData.Close
'        End If
'
'        mrs.Close
'        moConn.Close
'
'   ' Next
'
'
'    MsgBox " „ Õ›Ÿ «·Õ—ﬂ« "
'    Exit Sub
'EH:
'    Me.Enabled = True
'    lblTime.Visible = False
'    MsgBox Err.Description
'End Sub

Function PopulateRS(pFileString As String) As ADODB.Recordset
    On Error GoTo populaters_error
    Dim rst    As New ADODB.Recordset
   
    Dim moConn As New ADODB.Connection
    Dim firstSheetName As String
    Dim sConn  As String
    sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Mode=Share Exclusive;Data Source=" & pFileString & ";Extended Properties='Excel 8.0;IMEX=1'"
    moConn.Open sConn

    Set rst = moConn.OpenSchema(adSchemaTables)
    If Not rst.EOF Then
        firstSheetName = rst.Fields("table_name").Value
    Else
        MsgBox "File Not Valid "
        Exit Function
    End If
    rst.Close
    moConn.Close
    Set rst = Nothing
    Set moConn = Nothing
     
    '    mrs.MoveFirst
    '    tablename = mrs.Fields("table_name").Value
    '    mrs.Close
    '    moConn.Close
    
    StrSQL = "SELECT * FROM [" & firstSheetName & "]"
 
    Dim ExcelRs As ADODB.Recordset
    Set ExcelRs = New ADODB.Recordset
    With ExcelRs
        .CursorLocation = adUseClient
        .Open "[" & firstSheetName & "]", _
           "Provider=Microsoft.Jet.OLEDB.4.0;" _
           & "Data Source='" & pFileString & "';" _
           & "Mode=Share Exclusive;" _
           & "Extended Properties='Excel 8.0;IMEX=1'", _
           adOpenStatic, _
           adLockReadOnly, _
           adCmdTable
        Set .ActiveConnection = Nothing
    End With
    
   
    Dim fld As ADODB.Field
    cmbCode.Clear
    cmbArabicName.Clear
    cmbEnglishName.Clear
    cmbParent.Clear
    For Each fld In ExcelRs.Fields
        cmbCode.AddItem fld.Name
        cmbArabicName.AddItem fld.Name
        cmbEnglishName.AddItem fld.Name
        cmbParent.AddItem fld.Name
    Next
    ExcelRs.Close
    
   
    Exit Function
populaters_error:
    MsgBox Err.Description
End Function

Public Function OpenExcelRS(filename) As ADODB.Recordset
    On Error GoTo EH

    Dim firstSheetName As String
    Dim cnn            As ADODB.Connection
    Dim rs             As ADODB.Recordset
    Set cnn = New ADODB.Connection
    Dim sConn As String
    sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Mode=Share Exclusive;Data Source=" & filename & ";Extended Properties='Excel 8.0;IMEX=1'"
    cnn.Open sConn

    Set rs = cnn.OpenSchema(adSchemaTables)
    If Not rs.EOF Then
        firstSheetName = rs.Fields("table_name").Value
    Else
        MsgBox "File Not Valid "
        Exit Function
    End If
    rs.Close
    cnn.Close
    Set rs = Nothing
    Set cnn = Nothing
    
    Dim ExcelRs As ADODB.Recordset
    Set ExcelRs = New ADODB.Recordset
    With ExcelRs
        .CursorLocation = adUseClient
        .Open "[" & firstSheetName & "]", _
           "Provider=Microsoft.Jet.OLEDB.4.0;" _
           & "Data Source='" & filename & "';" _
           & "Mode=Share Exclusive;" _
           & "Extended Properties='Excel 8.0;IMEX=1'", _
           adOpenStatic, _
           adLockReadOnly, _
           adCmdTable
        Set .ActiveConnection = Nothing
    End With
    Set OpenExcelRS = ExcelRs
    Exit Function
EH:
    MsgBox "File My Not Valid " & Err.Description
    Set OpenExcelRS = Nothing
End Function

Private Sub btnTest_Click()
 
    On Error GoTo populaters_error
  
    
   
    Set VSFlexGrid1.DataSource = OpenExcelRS(TxtFileName)
   
    '    rst.ActiveConnection = Nothing 'disconnect from the connection
    '    cnct32.Close
    '    Set cnct32 = Nothing
    '    Set PopulateRS = rst
    '    Set rst = Nothing
    Exit Sub
populaters_error:
    MsgBox Err.Description
 
End Sub

Private Sub Command1_Click()
    If Trim(txtRowNo) = "" Then
        MsgBox "Insert First Valid Row "
        Exit Sub
    End If
    
    On Error GoTo populaters_error
    
    Dim StrSQL As String
    strDriver = "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)"
    
    Dim codeCol        As String
   
    Dim arabicNameCol  As String
    Dim EnglishNameCol As String
    Dim ParentCol      As String
    
    
    Dim mCode As String
Dim mName As String
Dim mNameP As String
Dim mSql As String
Dim StrNewAccountCode  As String
Dim mBranchName As String
Dim mBranchId As Long
Dim last_account As Boolean
Dim opening_balance As Double
Dim RSDDD As New ADODB.Recordset


    arabicNameCol = cmbArabicName.List(cmbArabicName.ListIndex)
    EnglishNameCol = cmbEnglishName.List(cmbEnglishName.ListIndex)
    codeCol = cmbCode.List(cmbCode.ListIndex)
    ParentCol = cmbParent.List(cmbParent.ListIndex)
    If Trim(arabicNameCol) = "" Then MsgBox "Select All column Names ": Exit Sub
    If Trim(EnglishNameCol) = "" Then MsgBox "Select All column Names ": Exit Sub
    If Trim(codeCol) = "" Then MsgBox "Select All column Names ": Exit Sub
    If Trim(ParentCol) = "" Then MsgBox "Select All column Names ": Exit Sub
    
    StrSQL = "SELECT  * "
   
    StrSQL = StrSQL & "    FROM [" & tablename & "]"
 
    Dim SysSQLServerTypeTechnical, StrAppRegPath, SysSQLServerName, SysSQLServerDataBaseName, SysSQLServerUserId, SysSQLServerUserpassword, SysSQLServerType
    StrAppRegPath = "bisegypt\SimpleAccounting"
    SysSQLServerType = Val(GetSetting(StrAppRegPath, "ServerCon", "ServerType", 0)) '0 loca 1 not 2 rem
    SysSQLServerName = GetSetting(StrAppRegPath, "ServerCon", "ServerName", "")
    SysSQLServerTypeTechnical = GetSetting(StrAppRegPath, "ServerCon", "SysSQLServerTypeTechnical", "0")

    SysSQLServerDataBaseName = GetSetting("Byte_DBS", "Setting", "DBPath", "Byte")

    SysSQLServerUserId = GetSetting("Byte_DBS", "Setting", "SysSQLServerUserId", "salim")
    SysSQLServerUserpassword = GetSetting("Byte_DBS", "Setting", "SysSQLServerUserpassword", "salim")
    
    ConnectionFirst
    Dim rst As New ADODB.Recordset
    Set rst = OpenExcelRS(TxtFileName)
    
    If rst.EOF Then MsgBox "No Data ": Exit Sub
    Dim RecoredCount As Integer
    rst.MoveLast
    rst.MoveFirst
    RecoredCount = rst.RecordCount
  Dim mArabicName As String
    Dim rsInsert As New ADODB.Recordset
    Dim s        As String
    s = "SELECT *  "
    s = s & "FROM ACCOUNTS "
    
 '   cn.BeginTrans
    On Error GoTo lblerro
    Dim mRow As Integer
    Row = 1
    lblLog = ""
    Dim rsParant As ADODB.Recordset
    Do While Not rst.EOF
    
'
'       mCode = Trim(GrdAccount.TextMatrix(i, GrdAccount.ColIndex("Account_Serial")))
'    mName = Trim(GrdAccount.TextMatrix(i, GrdAccount.ColIndex("Account_Name")))
'    mNameP = Trim(GrdAccount.TextMatrix(i, GrdAccount.ColIndex("AccountNamePar")))
'    last_account = Trim(GrdAccount.ValueMatrix(i, GrdAccount.ColIndex("last_account")))
'    If last_account Then
'        last_account = False
'    Else
'        last_account = True
'    End If
'    If last_account = False Then
'        last_account = False
'
'    End If
'
'    mCode = Trim(GrdAccount.TextMatrix(i, GrdAccount.ColIndex("Account_Serial")))
    
    
    
        
        
        lblLog = "Row " & Row & " Of " & RecoredCount
        DoEvents
        If Row < Val(txtRowNo) Then GoTo NextRow
'        rsInsert.Open s + " Where Account_Serial = '" & rst.Fields(codeCol) & "'", cn, adOpenKeyset, adLockOptimistic
        
        
        sql = " select * from ACCOUNTS Where Account_Serial = '" & rst.Fields(ParentCol) & "' "
    
    Set rsParant = New ADODB.Recordset
    rsParant.Open sql, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsParant.EOF Then
        mParantAccount = ""
        mDepitOrCredit = 0
    Else
        mParantAccount = Trim(rsParant!Account_Code & "")
        mDepitOrCredit = Val(rsParant!DepitOrCredit & "")
    End If
    
'        If rsInsert.EOF Then
'            rsInsert.AddNew
'        End If
mBranchId = 0
mArabicName = IIf(rst(arabicNameCol) & "" = "", Trim$(rst.Fields(EnglishNameCol)), Trim$(Trim$(rst.Fields(arabicNameCol) & "")))
        
        StrNewAccountCode = AddNewAccount(Trim(mParantAccount), mArabicName, last_account, False, Trim$(rst.Fields(EnglishNameCol)), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(mDepitOrCredit = 0, 0, 1), 0, 0, 0, 1, False, , True, , , mBranchId, opening_balance, rst.Fields(codeCol))
        
        SaveBransh_UserAccount StrNewAccountCode
        
'        rsInsert!Account_Code = rst.Fields(codeCol)
'        rsInsert!Account_Name = rst.Fields(arabicNameCol)
'        rsInsert!Account_NameEng = rst.Fields(EnglishNameCol)
'        rsInsert!Parent_Account_Code = rst.Fields(ParentCol)
'        rsInsert!BasicAccount = False  ' ÿ Õÿ «‰  «·œÌ›Ê·  ›«·ÌÊ Ì« «»«‘«
'        rsInsert!cannot_del = False
'        rsInsert!last_account = False
'        rsInsert.Update
'        rsInsert.Close
      
NextRow:
        rst.MoveNext
        Row = Row + 1
          
    Loop
  '  cn.CommitTrans
    MsgBox " All Done "
    Exit Sub
lblerro:
    Cn.RollbackTrans
    MsgBox Err.Description
    Exit Sub
populaters_error:
    MsgBox Err.Description
    
End Sub




'
'Public Sub GridFromToExecl(ByRef mGrid As Object, _
'                           frm As Form, _
'                           Optional ByVal mTable As String = "", _
'                           Optional ByVal ExtraFieldsName As String = "", Optional ByVal ExtraFieldsType As String = "", Optional ByVal ExtraFieldsTitle As String = "", Optional ByVal ExtraFieldsColComboList As String = "")
'    On Error Resume Next
'
'
'    Dim myform As New frm
'    Set myform = New frm
'
'  '  myform.mTableName = mTable
'    myform.mGrid.Rows = 1
'    Set myform.SenderObject = frm
'    Set myform.SenderGrid = mGrid
'    'myform.mGrid.Cols = 1
'    myform.mGrid.Rows = mGrid.Rows
'    '    myform.mGrid.Cols = mGrid.Cols
'    '    myform.mGrid2.Cols = mGrid.Cols
'    '    myform.mGrid3.Cols = mGrid.Cols
'
'    myform.MainFormName = frm.Name
'    Dim colKey
'    Dim ColDataType
'    Dim TextMatrix
'    Dim ColComboList
'    '
'    If ExtraFieldsName <> "" Then
'        colKey = Split(ExtraFieldsName, ",")
'        ColDataType = Split(ExtraFieldsType, ",")
'        TextMatrix = Split(ExtraFieldsTitle, ",")
'        ColComboList = Split(ExtraFieldsColComboList, ",")
'        For i = 1 To UBound(colKey)
'            AddToGrid myform.mGrid, colKey(i), TextMatrix(i), ColDataType(i), ColComboList(i)
'            AddToGrid myform.mGrid2, colKey(i), TextMatrix(i), ColDataType(i), ColComboList(i)
'            AddToGrid myform.mGrid3, colKey(i), TextMatrix(i), ColDataType(i), ColComboList(i)
'            AddToGrid myform.mtmpGrd, colKey(i), TextMatrix(i), ColDataType(i), ColComboList(i)
'        Next
'    End If
'    For i = myform.mGrid.Cols - 1 To 1 Step -1
'        If myform.mGrid.ColHidden(i) Then
'            '            myform.mGrid.ColPosition(i) = myform.mGrid.Cols - 1
'            '            myform.mGrid2.ColPosition(i) = myform.mGrid2.Cols - 1
'            '            myform.mGrid3.ColPosition(i) = myform.mGrid3.Cols - 1
'            '            myform.mtmpGrd.ColPosition(i) = myform.mtmpGrd.Cols - 1
'        End If
'    Next
'
'
'    '    myform.mGrid2.Cols = myform.mGrid.Cols
'    '    myform.mGrid3.Cols = myform.mGrid.Cols
'
'    For j = 0 To mGrid.Cols - 1
'        Screen.MousePointer = vbHourglass
'        'If Not myForm.mGrid.ColHidden(j) Then
'        myform.mGrid.colKey(j) = mGrid.colKey(j)
'        myform.mGrid.FixedAlignment(j) = mGrid.FixedAlignment(j)
'        myform.mGrid.ColAlignment(j) = mGrid.ColAlignment(j)
'        myform.mGrid.ColComboList(j) = mGrid.ColComboList(j)
'        myform.mGrid.ColDataType(j) = mGrid.ColDataType(j)
'        myform.mGrid.ColFormat(j) = mGrid.ColFormat(j)
'        myform.mGrid.ColHidden(j) = mGrid.ColHidden(j)
'
'        myform.mGrid2.TextMatrix(0, j) = mGrid.TextMatrix(0, j)
'        myform.mGrid2.colKey(j) = mGrid.colKey(j)
'        myform.mGrid2.FixedAlignment(j) = mGrid.FixedAlignment(j)
'        myform.mGrid2.ColAlignment(j) = mGrid.ColAlignment(j)
'        myform.mGrid2.ColComboList(j) = mGrid.ColComboList(j)
'        myform.mGrid2.ColDataType(j) = mGrid.ColDataType(j)
'        myform.mGrid2.ColFormat(j) = mGrid.ColFormat(j)
'        myform.mGrid2.ColHidden(j) = mGrid.ColHidden(j)
'
'        myform.mGrid3.TextMatrix(0, j) = mGrid.TextMatrix(0, j)
'        myform.mGrid3.colKey(j) = mGrid.colKey(j)
'        myform.mGrid3.FixedAlignment(j) = mGrid.FixedAlignment(j)
'        myform.mGrid3.ColAlignment(j) = mGrid.ColAlignment(j)
'        myform.mGrid3.ColComboList(j) = mGrid.ColComboList(j)
'        myform.mGrid3.ColDataType(j) = mGrid.ColDataType(j)
'        myform.mGrid3.ColFormat(j) = mGrid.ColFormat(j)
'        myform.mGrid3.ColHidden(j) = mGrid.ColHidden(j)
'
'        For i = 0 To mGrid.Rows - 1
'
'            If InStr(1, mGrid.ColComboList(j), "#") And i <> 0 Then
'                myform.mGrid.TextMatrix(i, j) = mGrid.TextMatrix(i, j)
'            Else
'                myform.mGrid.TextMatrix(i, j) = mGrid.TextMatrix(i, j)
'            End If
'            DoEvents
'        Next
'
'        'End If
'    Next
'    Screen.MousePointer = vbNormal
'    'GridSerial myForm.mGrid
'    myform.Show 1
'    If myform.Tag = "OK" Then
'        Screen.MousePointer = vbHourglass
'        mGrid.Rows = 1
'        mGrid.Rows = myform.mGrid.Rows
'        For j = 1 To myform.mGrid.Cols - 1
'            For i = 1 To myform.mGrid.Rows - 1
'                mGrid.TextMatrix(i, j) = myform.mGrid.TextMatrix(i, j)
'                DoEvents
'            Next
'        Next
'    End If
'  '  GridSerial mGrid
'    Screen.MousePointer = vbNormal
'    Unload myform
'
'    '  myform.mGrid.Rows = 1
'    '    myform.mGrid.Cols = 1
'    '    myform.mGrid.Rows = mGrid.Rows
'    '    myform.mGrid.Cols = mGrid.Cols
'    '    myform.mGrid2.Cols = mGrid.Cols
'    '    myform.mGrid3.Cols = mGrid.Cols
'    '    myform.MainFormName = frm.Name
'    '    myform.mTableName = mTable
'    '
'    '    If mTable <> "" Then
'    '        myform.cmdSave.Visible = True
'    '        myform.SSTTab0.TabsPerRow = 3
'    '        myform.SSTTab0.Tab = 0
'    '    Else
'    '        myform.SSTTab0.TabsPerRow = 1
'    '        myform.SSTTab0.TabVisible(0) = True
'    '        myform.SSTTab0.TabVisible(1) = False
'    '        myform.SSTTab0.TabVisible(2) = False
'    '        myform.cmdSave.Visible = False
'    '    End If
'    '    myform.mGrid2.Cols = myform.mGrid.Cols
'    '    myform.mGrid3.Cols = myform.mGrid.Cols
'    '
'    '    For j = 0 To mGrid.Cols - 1
'    '        Screen.MousePointer = vbHourglass
'    '        'If Not myForm.mGrid.ColHidden(j) Then
'    '        myform.mGrid.colKey(j) = mGrid.colKey(j)
'    '        myform.mGrid.FixedAlignment(j) = mGrid.FixedAlignment(j)
'    '        myform.mGrid.ColAlignment(j) = mGrid.ColAlignment(j)
'    '        myform.mGrid.ColComboList(j) = mGrid.ColComboList(j)
'    '        myform.mGrid.ColDataType(j) = mGrid.ColDataType(j)
'    '        myform.mGrid.ColFormat(j) = mGrid.ColFormat(j)
'    '        myform.mGrid.ColHidden(j) = mGrid.ColHidden(j)
'    '
'    '        myform.mGrid2.TextMatrix(0, j) = mGrid.TextMatrix(0, j)
'    '        myform.mGrid2.colKey(j) = mGrid.colKey(j)
'    '        myform.mGrid2.FixedAlignment(j) = mGrid.FixedAlignment(j)
'    '        myform.mGrid2.ColAlignment(j) = mGrid.ColAlignment(j)
'    '        myform.mGrid2.ColComboList(j) = mGrid.ColComboList(j)
'    '        myform.mGrid2.ColDataType(j) = mGrid.ColDataType(j)
'    '        myform.mGrid2.ColFormat(j) = mGrid.ColFormat(j)
'    '        myform.mGrid2.ColHidden(j) = mGrid.ColHidden(j)
'    '
'    '        myform.mGrid3.TextMatrix(0, j) = mGrid.TextMatrix(0, j)
'    '        myform.mGrid3.colKey(j) = mGrid.colKey(j)
'    '        myform.mGrid3.FixedAlignment(j) = mGrid.FixedAlignment(j)
'    '        myform.mGrid3.ColAlignment(j) = mGrid.ColAlignment(j)
'    '        myform.mGrid3.ColComboList(j) = mGrid.ColComboList(j)
'    '        myform.mGrid3.ColDataType(j) = mGrid.ColDataType(j)
'    '        myform.mGrid3.ColFormat(j) = mGrid.ColFormat(j)
'    '        myform.mGrid3.ColHidden(j) = mGrid.ColHidden(j)
'    '        For i = 0 To mGrid.Rows - 1
'    '            myform.mGrid.TextMatrix(i, j) = mGrid.TextMatrix(i, j)
'    '            DoEvents
'    '        Next
'    '        'End If
'    '    Next
'    '    Screen.MousePointer = vbNormal
'    '    'GridSerial myForm.mGrid
'    '    myform.Show 1
'    '    If myform.Tag = "OK" Then
'    '        Screen.MousePointer = vbHourglass
'    '        mGrid.Rows = 1
'    '        mGrid.Rows = myform.mGrid.Rows
'    '        For j = 1 To myform.mGrid.Cols - 1
'    '            For i = 1 To myform.mGrid.Rows - 1
'    '                mGrid.TextMatrix(i, j) = myform.mGrid.TextMatrix(i, j)
'    '                DoEvents
'    '            Next
'    '        Next
'    '    End If
'    '    GridSerial mGrid
'    '    Screen.MousePointer = vbNormal
'    '    Unload myform
'End Sub
'
'



'
''emp_contract_type  ContractID
'MaritalStatus
'    DcbMatrial.AddItem "√⁄“»"
'      DcbMatrial.AddItem "„ “ÊÃ"
'     Dcbsex.AddItem "–ﬂ—"
'      Dcbsex.AddItem "√‰ÀÏ"
'       Dcbsex.AddItem "Male"
'     Dcbsex.AddItem "Female"
'       DcbMatrial.AddItem "Single"
'     DcbMatrial.AddItem "Married"
'
'JobTypeID
'JobTypeID,JobTypeName From TblEmpJobsTypes
'
'pasplace
'select  id,name  from jopstatus
'select  id,name  from Nationality
' select  id,name  from dean
'
'
'SELECT DISTINCT pasplace, pasplace AS pasplaceName"
'sql = sql & " From dbo.TblEmployee"
'sql = sql & " WHERE     (NOT (pasplace IS NULL))"
'
'
'
'sql = "SELECT DISTINCT BankCode, BankCode AS BankCodeName"
'sql = sql & " From dbo.TblEmployee"
'sql = sql & " WHERE     (NOT (BankCode IS NULL)) "
'
'sql = "SELECT DISTINCT BanckName, BanckName AS BanckNameName"
'sql = sql & " From dbo.TblEmployee"
'sql = sql & " WHERE     (NOT (BanckName IS NULL)) "

Public Function GetGridFileName(ByVal G As Object, Optional MainFormName As String = "") As String
    Dim GlobalGridName As String
    Dim IndexS As String
    Dim MainContainerName As String

    On Error Resume Next
    IndexS = G.Index

    MainContainerName = GetMainForm(G.Container)
    GlobalGridName = MainContainerName & "\" & G.Name & IndexS & MainFormName
    GlobalGridName = "Import"
    GetGridFileName = App.Path & GlobalGridName & ".xls"

End Function

Private Function ToHex(ByRef pstrMessage As String) As String

    Dim llngMaxIndex As Long
    Dim llngIndex As Long
    Dim lstrHex As String

    llngMaxIndex = LenB(pstrMessage)

    For llngIndex = 1 To llngMaxIndex
        lstrHex = lstrHex & Right("0" & Hex(AscB(MidB(pstrMessage, llngIndex, 1))), 2)
    Next

    ToHex = lstrHex

End Function

Private Function FromHex(ByRef pstrHex As String) As String

    Dim llngMaxIndex As Long
    Dim llngIndex As Long
    Dim lstrMessage As String

    llngMaxIndex = Len(pstrHex)

    For llngIndex = 1 To llngMaxIndex Step 2
        lstrMessage = lstrMessage & ChrB("&h" & Mid(pstrHex, llngIndex, 2))
    Next

    FromHex = lstrMessage

End Function

Private Function Translate(ByRef pstrMessage As String, ByVal Key As String) As String

    Dim llngIndex As Long
    Dim llngMessageLength As Long
    Dim llngKeyLength As Long
    Dim lstrKey As String
    Dim llngKeyIndex As Long
    Dim lbytMessageByte As Byte
    Dim lbytKeyByte As Byte
    Dim llngMessageIndex As Long
    Dim lstrTranslation As String

    lstrKey = ToHex(Key)
    llngKeyLength = Len(lstrKey) \ 2

    If llngKeyLength = 0 Then Exit Function

    llngMessageLength = Len(pstrMessage) \ 2

    For llngIndex = 1 To llngMessageLength

        llngKeyIndex = (((llngIndex - 1) Mod llngKeyLength) * 2) + 1
        llngMessageIndex = ((llngIndex - 1) * 2) + 1

        lbytMessageByte = Int("&h" & Mid(pstrMessage, llngMessageIndex, 2))
        lbytKeyByte = Int("&h" & Mid(lstrKey, llngKeyIndex, 2))

        lstrTranslation = lstrTranslation & ToHex(ChrB(lbytMessageByte Xor lbytKeyByte))

    Next

    Translate = lstrTranslation

End Function

Public Function HexDecrypt(ByVal pstrMessage As String, ByVal Key As String) As String
    pstrMessage = Translate(pstrMessage, Key)
    pstrMessage = FromHex(pstrMessage)
    HexDecrypt = pstrMessage
End Function
Public Function HexEncrypt(ByVal pstrMessage As String, ByVal Key As String) As String
    pstrMessage = ToHex(pstrMessage)
    pstrMessage = Translate(pstrMessage, Key)
    HexEncrypt = pstrMessage
End Function

Public Function GetMainForm(ByVal Obj) As String
    Dim n As String
    On Error Resume Next
    n = Obj.Container.Name

    If n = "" Then
        GetMainForm = Obj.Name
    Else
        GetMainForm = GetMainForm(Obj.Container)
    End If
End Function


Public Sub ToExcel(ByRef mGrid As Object, _
                   Frm As Form, _
                   Optional MainFormName As String = "")
    On Error GoTo EH

    Screen.MousePointer = vbHourglass
    For i = 1 To mGrid.Cols - 1

        If Not mGrid.ColHidden(i) Then
            If mGrid.ColDataType(i) = 0 Then
                If mGrid.ColComboList(i) <> "" Then
                    mGrid.ColDataType(i) = flexDTSingle
                Else
                    mGrid.ColDataType(i) = flexDTString
                End If

            End If




        End If


    Next
  '  ExportToExcel Frm, mGrid, , , MainFormName

    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    MsgBox MyErrorHandler(Err.Number)

End Sub

Public Function MyErrorHandler(ErrNo As Long) As String
    mMsg = ""
    Select Case ErrNo
    
    Case 0
        MyErrorHandler = ""
        Exit Function
 
    Case -2147217864

        If ArabicInterface Then
            mMsg = " „ ≈Ã—«¡  ⁄œÌ·«  ⁄·Ï Â–Â «·‘«‘Â „‰ ÃÂ«“ ¬Œ—- „‰ ›÷·ﬂ «⁄œ  Õ„Ì· «·Õ—ﬂÂ À„ Õ«Ê· „—Â «Œ—Ï" & " - Optimistic concurrency erorr "
        Else
            mMsg = "This Form is editing from another computer- realod and Try again " & " - Optimistic concurrency erorr "
        End If

    Case -2147467259
        If ArabicInterface Then
            mMsg = "«·ÃÂ«“ «·Œ«œ„ «·—∆Ì”Ì „€·ﬁ √Ê €Ì— „ÊÃÊœ ⁄·Ï Â–Â «·‘»ﬂ…" & " - " & ErrNo
        Else
            mMsg = "Couldn't close file , Because it's already in use by another person or process" & " - " & ErrNo
        End If
    Case -2147352567
        'If ArabicInterface Then
        '    mMsg = "ÌÃ»  Œ’Ì’ «·ÿ«»⁄«  „‰ ≈œ«—… «·‰Ÿ«„" & " - " & ErrNo
        'Else
        '    mMsg = "Select Correct Report Printer Device" & " - " & ErrNo
        'End If
    Case 3155, 3022, -2147217873, -2147217900    ' insert fail
        If ArabicInterface Then
            mMsg = " ·«Ì„ﬂ‰ «÷«›… Â–« «·”Ã· ° Â–Â «·»Ì«‰«   „  ”ÃÌ·Â« „‰ ﬁ»·" & " - " & ErrNo
        Else
            mMsg = "You Can not Add this Record , May be there is Dublicated values" & " - " & ErrNo
        End If
    Case 3200    ' Change Or Delete Failed
        If ArabicInterface Then
            mMsg = " ·«Ì„ﬂ‰ «·€«¡ √Ê  ⁄œÌ· Â–« «·”Ã·  »”»» ÊÃÊœ »Ì«‰«  √Œ—Ï „— »ÿ… »Â ÊÌÃ» «·€«¡Â« √Ê·«" & " - " & ErrNo
        Else
            mMsg = "You Can not Delete Or Modify this Record , Because There Is Some Data Depends On It " & " - " & ErrNo
        End If
    Case 3157, 3046, 3202, 3218    ' Update Fail
        If ArabicInterface Then
            mMsg = " Â‰«ﬂ ›‘· ›Ï  Œ“Ì‰ «· ⁄œÌ·«  ° ﬁœ ÌﬂÊ‰ «·”Ã· „ﬁ›· »Ê«”ÿ… „” Œœ„ ¬Œ—° Õ«Ê· „—… √Œ—Ï" & " - " & ErrNo
        Else
            mMsg = "Update Failed , May be The record is locked by another User , Try Again " & " - " & ErrNo
        End If
    Case 3186, 3187, 3188
        If ArabicInterface Then
            mMsg = "”Ã· „€·ﬁ »Ê«”ÿ… „” Œœ„ ¬Œ—" & " - " & ErrNo
        Else
            mMsg = "Current Record locked by Another user" & " - " & ErrNo
        End If
    Case 3167
        If ArabicInterface Then
            mMsg = " „ «·€«¡ Â–« «·”Ã· »«·›⁄· " & " - " & ErrNo
        Else
            mMsg = "Record Already Deleted" & " - " & ErrNo
        End If
    Case 3314
        If ArabicInterface Then
            mMsg = "„‰ ›÷·ﬂ √ﬂ„· «·»Ì«‰«  ﬁ»· «· Œ“Ì‰" & " - " & ErrNo
        Else
            mMsg = "Please Complete the data before saving" & " - " & ErrNo
        End If
    Case 3262, 3211, 3212    ' Locked by another user and wait
        If ArabicInterface Then
            mMsg = "·« Ì„ﬂ‰ ≈€·«ﬁ «·„·› »”»» ÊÃÊœ „” Œœ„ ¬Œ— ÌﬁÊ„ »≈” Œœ«„Â √Ê ﬁ«„ »≈€·«ﬁÂ" & " - " & ErrNo
        Else
            mMsg = "Couldn't close file , Because it's already in use by another person or process" & " - " & ErrNo
        End If
    Case 3197    ' Couldn't repaire this files
        If ArabicInterface Then
            mMsg = "√ﬂÀ— „‰ „” Œœ„ Õ«Ê·Ê«  €ÌÌ— ‰›” «·»Ì«‰«  ›Ï ‰›” «·Êﬁ " & " - " & ErrNo
        Else
            mMsg = "Another Users are attempting to change the same data at the same time" & " - " & ErrNo
        End If
    Case 3056    ' Couldn't repaire this files
        If ArabicInterface Then
            mMsg = "·« Ì„ﬂ‰  ’·ÌÕ «·„·›«  «·„” Œœ„…" & " - " & ErrNo
        Else
            mMsg = "Couldn't repaire this files" & " - " & ErrNo
        End If
    Case 3014, 3037    ' Can't open any more files
        If ArabicInterface Then
            mMsg = "·« Ì„ﬂ‰ › Õ „·›«  √Œ—Ï" & " - " & ErrNo
        Else
            mMsg = "Can't open any more files" & " - " & ErrNo
        End If
    Case 3356, 3260, 3261, 3189, 3008, 3164, 3006    ' Table or Database Locked
        If ArabicInterface Then
            mMsg = "«·„·› „€·ﬁ »Ê«”ÿ… „” Œœ„ ¬Œ—" & " - " & ErrNo
        Else
            mMsg = "The File is Locked by Another User" & " - " & ErrNo
        End If
    Case 3201    ' Add Or Edit Fail
        If ArabicInterface Then
            mMsg = " ·«Ì„ﬂ‰ «÷«›… Â–« «·”Ã· √Ê «· ⁄œÌ· ›ÌÂ ° ·√‰Â „— »ÿ »„·› ·„ Ì „ «·≈÷«›… √Ê «· ⁄œÌ· ›ÌÂ Õ Ï «·¬‰" & " - " & ErrNo
        Else
            mMsg = "You Can not Add this Record or Change it , Because it's Linked to a File that has not been Added or Changed till Now" & " - " & ErrNo
        End If
    Case -2147217887
        If ArabicInterface Then
            mMsg = "Œÿ√ €Ì— „⁄—Ê› ° Õ«Ê·  ‰›Ì– ‰›” «·⁄„·Ì… „—… √Œ—Ï" & " - " & ErrNo
        Else
            mMsg = "Undefined Error , Try again : " & ErrNo
        End If
    Case 3704
        
        On Error Resume Next
        db.Close
        Exit Function
    Case -1000000001
       
        MyErrorHandler = ""
        Exit Function
    End Select
    '*************************
    If Err.Number = vbObjectError + 1000 Then
        

        mMsg = mMsg & vbNewLine & Err.Description
    Else
        mMsg = mMsg & vbNewLine & Err.Description & " : " & Err.Number
    End If
    '*************************
    If ErrNo <> -2147217864 Then
        If db.Errors.count > 0 Then
            ss = ""
            Dim adoErr As ADODB.Error
            j = 1
            On Error GoTo EEE
            For Each adoErr In db.Errors
                If adoErr.Number <> 0 Then
                    If j = 1 Then ss = vbNewLine & "-------SQL Errors-------"
                    ss = ss & vbNewLine & "Error (" & j & ")=" & adoErr.Description & " : " & adoErr.Number & ":" & adoErr.SQLState & ":" & adoErr.NativeError
                    j = j + 1
                End If
            Next adoErr
EEE:
            ' for this rand error Not enough storage is available to process this command.
            If Err.Number = 48 Then
                Set adoErr = db.Errors(0)
                ss = ss & vbNewLine & "Error (48)=" & adoErr.Description & " : " & adoErr.Number & ":" & adoErr.SQLState & ":" & adoErr.NativeError
            End If
            On Error GoTo 0
            mMsg = mMsg & vbNewLine & ss
        End If
    End If
    '*************************
    'If Trim(mMsg) <> "()(0)" Then MyErrorHandler = mMsg Else MyErrorHandler = ""
    MyErrorHandler = mMsg & ":" & Erl
    IsAboutError = True

End Function




Private Sub GetIDCombo(ByVal mTableColID As String, ByVal mRow As Long, ByVal mCol As Long, ByVal mGrid As Object)
Dim mTxt As String
mTxt = Trim(mGrid.TextMatrix(mRow, mCol - 1))
Select Case mTableColID
Case "sexID"
    If mTxt = "Male" Or mTxt = "–ﬂ—" Then
        mTxt = 1
    Else
        mTxt = 2
    End If
Case "MaritalStatusID"
'    DcbMatrial.AddItem "√⁄“»"
'      DcbMatrial.AddItem "„ “ÊÃ"
    If mTxt = "√⁄“»" Or mTxt = "Single" Then
        mTxt = 0
    ElseIf mTxt = "„ “ÊÃ" Or UCase(mTxt) = "MARRIED" Then
        mTxt = 1
    ElseIf mTxt = "„ÿ·ﬁ/„ÿ·›…" Or UCase(mTxt) = "DIVORCED" Then
        mTxt = 2
    ElseIf mTxt = "«—„·/√—„·…" Or UCase(mTxt) = "WIDOWED" Then
        mTxt = 3
        
    End If
    
Case "Status_id"
'    DcbMatrial.AddItem "√⁄“»"
'      DcbMatrial.AddItem "„ “ÊÃ"
    If mTxt = "Ã«—Ì «·«Â·«ﬂ" Or mTxt = "Ã«—Ï «·«Â·«ﬂ" Then
        mTxt = 0
    ElseIf mTxt = "„ Êﬁ›" Or UCase(mTxt) = "Stoped" Then
        mTxt = 1
    ElseIf mTxt = " „ «· Œ·’ »«·»Ì⁄" Or UCase(mTxt) = " „ «· Œ·’ »«·»Ì⁄" Then
        mTxt = 2
    ElseIf mTxt = " „ «·«Â·«ﬂ »«· Œ—Ìœ" Or UCase(mTxt) = " „ «·«Â·«ﬂ »«· Œ—Ìœ" Then
        mTxt = 3
        
    End If
    
 Case "Depreciation_Type_id"
'    DcbMatrial.AddItem "√⁄“»"
'      DcbMatrial.AddItem "„ “ÊÃ"
    If mTxt = "«·ﬁ”ÿ «·À«» " Or mTxt = "«·ﬁ”ÿ «·À«» " Then
        mTxt = 0
    ElseIf mTxt = "«·ﬁ”ÿ  «·„ ‰«ﬁ’" Or UCase(mTxt) = "«·ﬁ”ÿ  «·„ ‰«ﬁ’" Then
        mTxt = 1

    End If
       
Case "Emp_Name1.Emp_Name2.Emp_Name3.Emp_Name4"
    mTxt = mGrid.TextMatrix(mRow, mCol - 4) + " " + mGrid.TextMatrix(mRow, mCol - 3) + " " + mGrid.TextMatrix(mRow, mCol - 2) + " " + mGrid.TextMatrix(mRow, mCol - 1)
Case ""
End Select
mGrid.TextMatrix(mRow, mCol) = mTxt
End Sub

Public Function ToHijriDate(ByVal GregorianDate As String) As String
    Dim HijriDate As String, DateFormat As String
    ' DateFormat = "long date"
    
    DateFormat = "dd-mm-yyyy"
    HijriDate = ConvertDate(GregorianDate, vbCalGreg, vbCalHijri, DateFormat)
    ToHijriDate = HijriDate
    
End Function
Private Function ConvertDate(ByRef StringIn As String, _
                             ByRef OldCalender As Integer, _
                             ByVal NewCalender As Integer, _
                             ByRef NewFormat As String) As String
                             If StringIn = "" Then Exit Function
On Error Resume Next
    Dim SavedCal As Integer
    Dim d As Date, s As String
    SavedCal = Calendar
    Calendar = OldCalender
    d = CDate(StringIn)
    Calendar = NewCalender
    s = CStr(d)
    ConvertDate = Format(s, NewFormat)
    Calendar = SavedCal
End Function

Public Function ToGregorianDate(ByVal HijriDate As String) As Date
    Dim GregorianDate As String, DateFormat As String
  If HijriDate = "" Then Exit Function
    DateFormat = "dd/mm/yyyy"
    
    GregorianDate = ConvertDate(HijriDate, vbCalHijri, vbCalGreg, DateFormat)
    If DateDiff("D", "01/01/1900", GregorianDate) < 0 Then
    GregorianDate = Date
    End If
    ToGregorianDate = GregorianDate
End Function

Public Function CheckDateIsHij(ByVal mDate As String) As Integer
    If Not IsDate(mDate) Then CheckDateIsHij = 3: Exit Function
    
    If Trim(mDate) = "" Then CheckDateIsHij = 3: Exit Function
    
    If Year(mDate) < 1800 Then
        CheckDateIsHij = 1
    Else
        CheckDateIsHij = 2
    End If
End Function


Private Sub GetFieldID(ByVal mTableColName As String, ByVal mRow As Long, ByVal mCol As Long, ByVal mGrid As Object, Optional ByVal MainTableName As String = "")
    Dim mTableName As String
    Dim mFieldIDName As String
    Dim mFieldName As String
    Dim xx As Variant
    Dim mValue As String
    Dim rsDummy As New ADODB.Recordset
    Dim rsDummy2 As New ADODB.Recordset
    If mCol = 67 Then
        mCol = 67
    End If
    If mGrid.ColKey(mCol) = "NationlID" Then
        mCol = mCol
    End If
    Dim mValue2 As String
    If mGrid.ColKey(mCol) = "DeanID" Then
        mCol = mCol
    End If
    If mGrid.ColKey(mCol) = "DOBH" Then
        mCol = mCol
    End If
    If mTableColName = "Date" Then
        If CheckDateIsHij(Trim(mGrid.TextMatrix(mRow, mCol - 1))) = 1 Then
            'If Trim(mGrid.TextMatrix(mRow, mCol - 1)) <> "" Then
                mGrid.TextMatrix(mRow, mCol) = Trim(mGrid.TextMatrix(mRow, mCol - 1))
                mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(mGrid.TextMatrix(mRow, mCol))
            'Else
            'End If
        ElseIf CheckDateIsHij(Trim(mGrid.TextMatrix(mRow, mCol - 1))) = 2 Then
            If Trim(mGrid.TextMatrix(mRow, mCol - 1)) = "" Then
                mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(Trim(mGrid.TextMatrix(mRow, mCol)))
            Else
                mGrid.TextMatrix(mRow, mCol) = ToHijriDate(Trim(mGrid.TextMatrix(mRow, mCol - 1)))
            End If
        ElseIf CheckDateIsHij(Trim(mGrid.TextMatrix(mRow, mCol - 1))) = 3 Then
            If mGrid.TextMatrix(mRow, mCol) <> "" Then
                mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(Trim(mGrid.TextMatrix(mRow, mCol)))
            End If
            'mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(mGrid.TextMatrix(mRow, mCol))
        Else
        
        End If
        Exit Sub
    End If
    xx = Split(mTableColName, ",")
    mTableName = xx(0)
    mFieldIDName = xx(1)
    mFieldName = xx(2)
    
 If mRow = 50 Then
 mRow = mRow
 End If
    mValue = Trim(mGrid.TextMatrix(mRow, mCol - 1))
Dim strValue As String
strValue = ""
Dim mValue3 As String

mValue3 = mValue
If (Right(mValue, 1)) = "Â" Then
    strValue = "…"
ElseIf (Right(mValue, 1)) = "…" Then
    strValue = "Â"
    
End If
If strValue <> "" Then
    mValue3 = Replace(mValue3, Right(mValue3, 1), strValue)
End If
Dim mEngLett As String
mEngLett = "e"
    Dim s As String
    mValue2 = mValue
    Select Case mTableName
    Case "jopstatus"
        If UCase(mValue) = "ACTIVE" Then
            mValue2 = "⁄·Ï ﬁÊ… «·⁄„·"
            
        End If
    Case "dean"
      If UCase(mValue) = "ISLAM" Then
            mValue2 = "„”·„"
       ElseIf UCase(mValue) = "CHRISTIAN" Then
            mValue2 = "„”ÌÕÏ"
        End If
    Case "Nationality"
        If UCase(mValue) = "JORDAN" Then
            mValue2 = "«—œ‰"
        ElseIf UCase(mValue) = "INDIA" Then
            mValue2 = "Â‰œ"
        ElseIf Trim(UCase(mValue)) = "" Then
            mValue2 = "”⁄ÊœÌ"
        ElseIf UCase(mValue) = "EGYPT" Then
            mValue2 = "„’—"
        ElseIf UCase(mValue) = "PAKISTAN" Then
            mValue2 = "»«ﬂ” «‰"
        ElseIf UCase(mValue) = "BANGLADESH" Then
            mValue2 = "»‰Ã·«œÌ‘"
        ElseIf UCase(mValue) = "SUDAN" Then
            mValue2 = "”Êœ«‰"
        ElseIf UCase(mValue) = "ETHIOPIA" Then
            mValue2 = "«ÀÌÊ»Ì«"
            
        ElseIf UCase(mValue) = "CAMEROON" Then
            mValue2 = "ﬂ«„Ì—Ê‰"
        ElseIf UCase(mValue) = "PALESTINE" Then
            mValue2 = "›·”ÿÌ‰"
        ElseIf UCase(mValue) = "SYRIA" Then
            mValue2 = "”Ê—Ì«"
        ElseIf UCase(mValue) = "JORDANIAN" Then
            mValue2 = "«—œ‰"
        ElseIf UCase(mValue) = "AMERICA" Then
            mValue2 = "«„—Ìﬂ«"
        ElseIf UCase(mValue) = "EGYPTIAN" Then
            mValue2 = "„’—"
        ElseIf UCase(mValue) = "KENYA" Then
            mValue2 = "ﬂÌ‰Ì«"
        ElseIf UCase(mValue) = "LEBANON" Then
            mValue2 = "·»‰«‰"
        ElseIf UCase(mValue) = "SIRLANKIAN" Then
            mValue2 = "”Ì—·«‰ﬂ"
        ElseIf UCase(mValue) = "YEMEN" Then
            mValue2 = "Ì„‰"
        ElseIf UCase(mValue) = "TUNIS" Then
            mValue2 = " Ê‰”"
        ElseIf UCase(mValue) = "MALAYSIA" Then
            mValue2 = "„«·Ì“Ì«"
         Else
            mValue2 = mValue
         
            
        End If
        If mValue = "" Then mValue2 = "”⁄ÊœÌ"
    Case Else
    End Select
    If mValue = "" Then
        Exit Sub
    End If
    mEngLett = "e"
    If UCase(mTableName) = "ACCOUNTS" Then
         mEngLett = "Eng"
    End If
    If UCase(mTableName) = "TBLCOUNTRIESGOVERNMENTS" Then
         mEngLett = ""
    End If

    
    s = "Select " & mFieldName & " ," & mFieldIDName & " ," & Trim(mFieldName) & mEngLett & "   "
    If UCase(mTableName) = "GROUPS" Or UCase(mTableName) = "GROUPSCUSTOMERS" Or UCase(mTableName) = "FIXEDASSETSGROUP" Then
        s = s & " ,ParentID,FullCode,GroupCode,Code,LastGroup "
    End If
    
    s = s & " from  " & mTableName
    s = s & " Where (" & mFieldName & " = '" & Trim(mValue2) & "' Or " & Trim(mFieldName) & mEngLett & "    = '" & Trim(mValue) & "')"
    s = s & " or (" & mFieldName & " = '" & Trim(mValue3) & "' Or " & Trim(mFieldName) & mEngLett & "   = '" & Trim(mValue3) & "')"
    If UCase(mTableName) = "GROUPS" Or UCase(mTableName) = "GROUPSCUSTOMERS" Or UCase(mTableName) = "FIXEDASSETSGROUP" Then
        s = s & " Or FullCode = '" & Trim(mValue3) & "' "
        If mFieldName = "GroupName" And mGrid.ColKey(mCol - 2) = "MainGroupCode" And mValue <> "" Then
        'If mFieldName = "GroupName" And mGrid.ColKey(mCol) = "ParentID2" And mValue <> "" Then
            
            
            s = "Select " & mFieldName & " ," & mFieldIDName & " ," & Trim(mFieldName) & mEngLett & "   "
            If UCase(mTableName) = "GROUPS" Or UCase(mTableName) = "GROUPSCUSTOMERS" Or UCase(mTableName) = "FIXEDASSETSGROUP" Then
                s = s & " ,ParentID,FullCode,GroupCode,Code,LastGroup "
            End If
            
            s = s & " from  " & mTableName
            s = s & " Where           "
            
            s = s & "  FullCode = '" & Trim(Trim(mGrid.TextMatrix(mRow, mCol - 2))) & "' "
        End If
    End If
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
    
    If rsDummy.EOF Then
        s = s & " Or ( " & mFieldName & " Like '%" & Trim(mValue2) & "%' Or " & Trim(mFieldName) & mEngLett & "    Like '%" & Trim(mValue) & "%')"
    
    End If
    
    If rsDummy.EOF And UCase(mTableName) = "ACCOUNTS" Then
        MsgBox "Â–« «·Õ”«» €Ì— „ÊÃÊœ ›Ï «·œ·Ì· " & mValue
        Exit Sub
    End If
    rsDummy.Close
    
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
   ' If Trim(mGrid.TextMatrix(mRow, mCol - 4)) <> "" And UCase(mTableName) = "GROUPS" Then GoTo 11
    If UCase(mTableName) = "GROUPS" And rsDummy.EOF Then
11:
        rsDummy.Close
             s = "Select " & mFieldName & " ," & mFieldIDName & " ," & Trim(mFieldName) & "e   "
        If UCase(mTableName) = "GROUPS" Or UCase(mTableName) = "GROUPSCUSTOMERS" Or UCase(mTableName) = "FIXEDASSETSGROUP" Then
            s = s & " ,ParentID,FullCode,GroupCode,Code,LastGroup "
        End If
        Dim mValue4  As String
        mValue4 = Trim(mGrid.TextMatrix(mRow, mCol - 2))
        
        s = s & " from  " & mTableName
        s = s & " Where " & mFieldName & " Like '%" & Trim(mValue2) & "%' Or " & Trim(mFieldName) & "e Like '%" & Trim(mValue) & "%'"
        s = s & " Or Fullcode   Like '%" & Trim(mValue4) & "%' Or Code Like '%" & Trim(mValue4) & "%'"
        rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
        If rsDummy.EOF Then
            mValue4 = mValue4
        End If
    End If
    
    If Not rsDummy.EOF Then
        If UCase(mTableName) = "ACCOUNTS" Then
            mGrid.TextMatrix(mRow, mCol) = Trim(rsDummy.Fields.Item(Trim(mFieldIDName)) & "")
        Else
            mGrid.TextMatrix(mRow, mCol) = Val(rsDummy.Fields.Item(Trim(mFieldIDName)) & "")
        End If
        If mGrid.ColKey(mCol) = "ParentID" Or mGrid.ColKey(mCol) = "ParentID2" Then
            mGrid.TextMatrix(mRow, mGrid.ColIndex("Code")) = Trim(mGrid.TextMatrix(mRow, mGrid.ColIndex("FullCode")))
            
            If mGrid.ColKey(mCol) = "ParentID2" And Val(mGrid.TextMatrix(mRow, mCol)) <> 0 Then
            
                mGrid.TextMatrix(mRow, mGrid.ColIndex("ParentID")) = Val(mGrid.TextMatrix(mRow, mCol))
            End If
            Dim mmm As String
            mmm = SearchInGrid(mGrid, mValue, "GroupName")
            If mmm <> "" Then
                'mGrid.TextMatrix(mRow, mGrid.ColIndex("GroupCode")) = GetNewGroupCode(Val(mGrid.TextMatrix(CLng(mmm), mGrid.ColIndex("NewId"))))
            End If
            mGrid.TextMatrix(mRow, mGrid.ColIndex("LastGroup")) = 0
        End If

    Else
       
        rsDummy.AddNew
        rsDummy(Trim(mFieldName)) = mValue
        rsDummy(Trim(mFieldName) & mEngLett) = mValue
        If mGrid.ColKey(mCol) = "ParentID" Or (mGrid.ColKey(mCol) = "ParentID2" And Val(mGrid.TextMatrix(mRow, mCol)) <> 0) Then
            'rsDummy("ParentID") = mValue
            Dim mm As String
            mm = SearchInGrid(mGrid, mValue, "GroupName")
            If mm <> "" Then
                rsDummy("ParentID") = Val(mGrid.TextMatrix(CLng(mm), mCol))
                rsDummy("FullCode") = Trim(mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("FullCode")))
                rsDummy("Code") = Trim(mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("FullCode")))
            Else
                xx = Split(Trim(mGrid.TextMatrix(mRow, mGrid.ColIndex("FullCode"))), "-")
                rsDummy("ParentID") = 1
                rsDummy("FullCode") = xx(0)
                rsDummy("Code") = xx(0)
            End If
            rsDummy("GroupCode") = GetNewGroupCode(Val(rsDummy("ParentID") & ""), mTableName)
            
            rsDummy("LastGroup") = 0
            If mm <> "" Then
                mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("Code")) = Trim(mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("FullCode")))
                mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("GroupCode3")) = rsDummy("GroupCode") & ""
                mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("LastGroup")) = 0
            End If
        End If
        s = "Select Max(" & mFieldIDName & ")  as MaxID  from  " & mTableName
        
        rsDummy2.Open s, Cn, adOpenKeyset, adLockOptimistic
        Dim mMaxId As Long
        If Not rsDummy2.EOF Then
            mMaxId = Val(rsDummy2!MaxId & "") + 1
        Else
            mMaxId = 1
        End If
        If UCase(mTableName) <> "GROUPSCUSTOMERS" Then
            rsDummy(Trim(mFieldIDName)) = mMaxId
        End If
        rsDummy(Trim(mFieldName)) = mValue
        rsDummy.Update
       ' mGrid.TextMatrix(mRow, mGrid.ColIndex("NewId")) = mMaxId
        mGrid.TextMatrix(mRow, mCol) = rsDummy(Trim(mFieldIDName) & "")
    End If

End Sub

Private Function SearchInGrid(ByVal mGrd As Object, ByVal mTxt As String, ByVal mFldName As String) As String
Dim i As Long
For i = 1 To mGrd.Rows - 1
    If Trim(mGrd.TextMatrix(i, mGrd.ColIndex(mFldName))) = mTxt Then
        SearchInGrid = i
        Exit Function
    End If
Next
SearchInGrid = ""
End Function
Function FileExists(filename) As Boolean
    On Error GoTo CheckError        ' Turn on error trapping so error handler                            ' responds if any error is detected.
    FileExists = (Dir(filename) <> "")
    Exit Function            ' Avoid executing error handler                             ' if no error occurs.

CheckError:        ' Branch here if error occurs.    ' Define constants to represent Visual Basic error code.
    FileExists = False
    Resume Next
End Function



Public Function ConnectionFirst() As Boolean

On Error GoTo ErrTrap
'«” ›”«—
'ServerDb = TxtServerDataBaseName.Text
'wael
'ServerDb = DestinationServer
' POSDb = TxtServerDataBaseName.Text


If Cn.State = 1 Then Exit Function
ServerDb = SysSQLServerDataBaseName
     Set Cn = New ADODB.Connection
    With Cn
        .CommandTimeout = 60
        .CursorLocation = adUseClient
        .ConnectionTimeout = 15
       If SysSQLServerType = 1 Then
        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
        "Persist Security Info=False;Initial Catalog=" & ServerDb & _
        ";Data Source=" & SysSQLServerName & ";Port=1433"
        
        ElseIf SysSQLServerType = 2 Then
 
     
                 If SysSQLServerTypeTechnical = "0" Then
                 .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
                    "Persist Security Info=False;Initial Catalog=" & ServerDb & _
                    ";Data Source=" & SysSQLServerName & ";Port=1433"
                    '";Data Source=" & ServerDb & ";Port=1433"
                    
                  Else
                     .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & ServerDb & ";Data Source=" & SysSQLServerName 'SysSQLServerName
                End If
          End If

.Open
End With
ConnectionFirst = True
Exit Function
ErrTrap:
ConnectionFirst = False
End Function





 
 
Public Function GetNewGroupCode(LngParentGroupID As Long, _
                                Optional ByVal mTableName As String = "") As String
    Dim rs               As ADODB.Recordset
    Dim StrSQL           As String
    Dim StrParentCode    As String
    Dim StrNewGroupCode  As String
    Dim StrLastGroupCode As String
    Dim IntTemp          As String
    If mTableName = "" Then
        mTableName = "Groups"
    End If
    On Error GoTo ErrTrap
    StrSQL = "Select GroupCode From " & mTableName & "  Where GroupID=" & LngParentGroupID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF Or rs.BOF) Then
        StrParentCode = IIf(IsNull(rs("GroupCode").Value), "", rs("GroupCode").Value)
    End If

    rs.Close
    Set rs = New ADODB.Recordset
    StrSQL = "Select * From " & mTableName & "  Where ParentID=" & LngParentGroupID & " Order By GroupID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        StrNewGroupCode = StrParentCode & "1"
    Else
        rs.MoveLast
        StrLastGroupCode = IIf(IsNull(rs("GroupCode").Value), "", rs("GroupCode").Value)
        IntTemp = Val(Mid(StrLastGroupCode, Len(StrParentCode) + 1))
        StrNewGroupCode = StrParentCode & CStr(IntTemp + 1)
    End If

    rs.Close
    Set rs = Nothing
    GetNewGroupCode = StrNewGroupCode
    Exit Function
ErrTrap:
End Function



Private Function GetNewCode(LngParentGroupID As Long, Optional ByVal mTableName As String = "", Optional ByVal mTableGroupName As String = "Groups", Optional ByVal mFieldGroup As String = "GroupID") As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim StrParentCode  As String
    Dim StrNewGroupCode As String
    Dim StrLastGroupCode As String
    Dim IntTemp As String
    If mTableName = "" Then
        mTableName = "Groups"
    End If
    On Error GoTo ErrTrap
    StrSQL = "Select Max(Code) Code From " & mTableName & "  Where " & mFieldGroup & " =" & LngParentGroupID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF Or rs.BOF) Then
        StrParentCode = IIf(IsNull(rs("Code").Value), "", rs("Code").Value)
    Else
        StrParentCode = "000"
    End If

     Set rs = New ADODB.Recordset
    StrSQL = "Select * From " & mTableGroupName & "   Where GroupID=" & LngParentGroupID & " Order By GroupID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Dim mTmpGroup2  As String
    If Not rs.BOF Then
        StrNewGroupCode = rs!code & ""
        mTmpGroup2 = Replace(StrParentCode, StrNewGroupCode, "")
    End If
    If Trim(mTmpGroup2) = "" Then mTmpGroup2 = "000"
    rs.Close
    Dim mTmp As Long
    mTmp = Val(mTmpGroup2) + 1
    If Len(CStr(mTmp)) = 1 Then
        StrParentCode = "00" & mTmp
    ElseIf Len(CStr(mTmp)) = 2 Then
        StrParentCode = "0" & mTmp
    ElseIf Len(CStr(mTmp)) = 3 Then
        StrParentCode = "" & mTmp
    End If
    Set rs = New ADODB.Recordset
    StrSQL = "Select * From " & mTableGroupName & "   Where GroupID=" & LngParentGroupID & " Order By GroupID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        StrNewGroupCode = StrParentCode & "1"
    Else
        rs.MoveLast
        StrLastGroupCode = IIf(IsNull(rs("Code").Value), "", rs("Code").Value)
        IntTemp = Val(Mid(StrLastGroupCode, Len(StrParentCode)))
        If IntTemp = 0 Then
            IntTemp = Val(Mid(StrLastGroupCode, Len(StrParentCode)))
        End If
        IntTemp = Val(Mid(StrLastGroupCode, Len(StrParentCode) - 1))
        StrNewGroupCode = StrLastGroupCode & StrParentCode & IntTemp
    End If

    rs.Close
    Set rs = Nothing
    GetNewCode = StrNewGroupCode
    Exit Function
ErrTrap:
End Function





Public Sub SaveBransh_UserAccount(Optional StrNewAccountCode As String)
Dim i As Integer
Dim sql As String
Dim Rs3 As ADODB.Recordset
'If ListGroupSelected.ListCount >= 0 Then
'sql = "Select * from  TblAccountBranch where 1=-1"
'Set Rs3 = New ADODB.Recordset
'Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'For i = 0 To ListGroupSelected.ListCount - 1
'Rs3.AddNew
'Rs3("BranchID").Value = ListGroupSelected.ItemData(i)
'Rs3("Account_Code").Value = Trim(StrNewAccountCode)
'Rs3.Update
'Next i
'End If
'
'If ListUserSelect.ListCount >= 0 Then
'sql = "Select * from  TblAccountUser where 1=-1"
'Set Rs3 = New ADODB.Recordset
'Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'For i = 0 To ListUserSelect.ListCount - 1
'Rs3.AddNew
'Rs3("UserID").Value = ListUserSelect.ItemData(i)
'Rs3("Account_Code").Value = Trim(StrNewAccountCode)
'Rs3.Update
'Next i
'End If
End Sub




Public Function AddNewAccount(StrParentAccCode As String, _
                              StrNewAccountName As String, _
                              Optional BolLastAcc As Boolean = True, _
                              Optional BolCannotDel As Boolean = False, _
                              Optional StrNewAccountNamee As String = "", _
                              Optional currenct_code As String = 1, _
                              Optional budget As Boolean = False, _
                              Optional cost_center As Boolean = False, _
                              Optional Sum_account As Boolean = False, _
                              Optional Branch As String = "0", _
                              Optional serial As String, _
                              Optional cost_center_type As Integer = 0, _
                              Optional cost_center_id As String, _
                              Optional ActivityTypeId As Integer = 0, _
                              Optional AccountTypes As Integer = 0, _
                              Optional AccountTab As Integer = 0, _
                              Optional DepitOrCreditv As Integer = 0, _
                              Optional Differenttypev As Integer = 0, _
                              Optional Authorityv As Integer = 0, _
                              Optional UserGroupIdv As Integer = 0, _
                              Optional UserIdv As Integer = 0, _
                              Optional ChKBlock As Boolean = False, Optional BasicAccount As Boolean = False, Optional last_account As Boolean = False, Optional ByVal mSerialAcc As Long = 0, Optional ByVal mLevel As Long = 0, Optional ByVal mBranchId As Long = 1, Optional ByVal opening_balance As Double = 0, Optional mSeialCode As String = "")
      
    If StrParentAccCode = "" Then StrParentAccCode = "r"
    If mSerialAcc = 0 Then
     
            ParentAccountPrperties StrParentAccCode, AccountTypes, AccountTab, DepitOrCreditv, Differenttypev, Authorityv
              If CHECK_LAST_ACCOUNT(StrParentAccCode) = True Then MsgBox "·«Ì„ﬂ‰ «‰‘«¡ Õ”«»  Õ  «·Õ”«» «·‰Â«∆Ì :  " & Get_Account_Serial(StrParentAccCode): AddNewAccount = "": Exit Function
     
    Else
        Select Case mSerialAcc
        Case 1
            AccountTypes = 1
            AccountTab = 0
            DepitOrCreditv = 1
            Differenttypev = 1
            Authorityv = 0
        Case 2
            AccountTypes = 1
            AccountTab = 1
            DepitOrCreditv = 1
            Differenttypev = 1
            Authorityv = 0
        Case 3
            AccountTypes = 2
            AccountTab = 2
            DepitOrCreditv = 1
            Differenttypev = 1
            Authorityv = 0
        Case 4
            AccountTypes = 2
            AccountTab = 3
            DepitOrCreditv = 1
            Differenttypev = 1
            Authorityv = 0
        Case 5
            AccountTypes = 0
            AccountTab = 4
            DepitOrCreditv = 1
            Differenttypev = 1
            Authorityv = 0
        End Select
    End If
    Dim StrSQL  As String
    Dim rs As ADODB.Recordset
    Dim StrNewAccCode As String
    Dim VarTemp As Variant

    Dim i As Integer, j As Integer
    If StrParentAccCode <> "r" Then
        StrSQL = "SELECT  ACCOUNTS.ActivityTypeId , ACCOUNTS.Branch , ACCOUNTS.Sum_account ,ACCOUNTS.cost_center ,ACCOUNTS.mowazna,ACCOUNTS.currenct_code,ACCOUNTS.Account_ID,Account_Code,Account_Name,Parent_Account_Code,ACCOUNTS.cost_center_type,  ACCOUNTS.cost_center_id" & ",last_account,cannot_del,Account_Serial,BasicAccount,DateCreated,Account_NameEng "
        StrSQL = StrSQL + " From ACCOUNTS "
        StrSQL = StrSQL + " Where ((ACCOUNTS.Parent_Account_Code='" & StrParentAccCode & "'))"
        StrSQL = StrSQL + " ORDER by ACCOUNTS.Account_ID"
    
        StrSQL = " select * "
        StrSQL = StrSQL + " From ACCOUNTS "
        StrSQL = StrSQL + " Where ((ACCOUNTS.Parent_Account_Code='" & StrParentAccCode & "'))"
        StrSQL = StrSQL + " ORDER by ACCOUNTS.Account_ID"
    
        StrSQL = " select * "
        StrSQL = StrSQL + " From ACCOUNTS where Parent_Account_Code='-1' "
    Else
        StrParentAccCode = ""
        StrSQL = " Select * From ACCOUNTS where Parent_Account_Code='-1' "
    End If
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    'If Rs.BOF Or Rs.EOF Then
    '    StrNewAccCode = StrParentAccCode & "a" & 1
    'Else
    '    Rs.MoveLast
    '    VarTemp = Split(Rs("Account_Code").Value, "a", , vbTextCompare)
    '    I = VarTemp(UBound(VarTemp))
    '    StrNewAccCode = StrParentAccCode & "a" & I + 1
    '
    'End If
    Dim Count_ACCOUNT_digit As Integer
    Dim NoOfAs As Integer
   ' If StrParentAccCode = "" Then StrParentAccCode = "r"
    StrNewAccCode = GetNewAcountCode(StrParentAccCode)
    NoOfAs = CountAs(StrParentAccCode) + 1

    'Count_ACCOUNT_digit = GetAccountsLevel(NoOfAs)

    'If NoOfAs = 1 Or NoOfAs = 2 Then
    'Count_ACCOUNT_digit = 0
    'ElseIf NoOfAs = 3 And NoOfAs = 4 Then
    'Count_ACCOUNT_digit = 2
    'Else
    'Count_ACCOUNT_digit = SystemOptions.Count_ACCOUNT_digit ' GetSetting(StrAppRegPath, "Setting", "COUNT_ACCOUNT_digit", 0)
    'End If
    Count_ACCOUNT_digit = GetAccountsLevel(NoOfAs)

    rs.AddNew
    rs("AccountTypes").Value = AccountTypes
    rs("AccountTab").Value = AccountTab
    rs("DepitOrCredit").Value = DepitOrCreditv
    rs("Differenttype").Value = Differenttypev
    rs("Authority").Value = Authorityv
    rs("UserGroupId").Value = UserGroupIdv
    rs("Userid").Value = UserIdv
    rs("Block").Value = ChKBlock
    rs("Level").Value = mLevel
    rs("Account_Code").Value = StrNewAccCode
    rs("Account_Name").Value = StrNewAccountName
    rs("Parent_Account_Code").Value = IIf(StrParentAccCode = "", "r", StrParentAccCode)
    rs("last_account").Value = BolLastAcc 'IIf(BolLastAcc, BolLastAcc, last_account)
    rs("cannot_del").Value = BolCannotDel
    rs("Branch").Value = Branch
    rs("BranchId").Value = mBranchId
serial = mSeialCode
    If Branch <> "" Then
    
        If Len(Branch) = 1 Then Branch = "00" & Branch
        If Len(Branch) = 2 Then Branch = "0" & Branch
             
        If serial = "" Then
            'If BolLastAcc = False Then
            'rs("Account_Serial").value = branch & Get_Account_Serial(StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(2, "0"))
            'Else
            'rs("Account_Serial").value = branch & Get_Account_Serial(StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(COUNT_ACCOUNT_digit, "0"))
            'End If
            
            If BolLastAcc = False Then
                rs("Account_Serial").Value = Get_Account_Serial(StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(Count_ACCOUNT_digit, "0"))
            Else
                rs("Account_Serial").Value = Get_Account_Serial(StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(Count_ACCOUNT_digit, "0"))
            End If

        Else
            rs("Account_Serial").Value = serial
        
        End If

    Else

        If serial = "" Then

            '   If get_account_max(Get_Account_Serial(StrParentAccCode)) >= 9 Then
            If BolLastAcc = False Then
                rs("Account_Serial").Value = Get_Account_Serial(StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(Count_ACCOUNT_digit, "0"))
            Else
                rs("Account_Serial").Value = Get_Account_Serial(StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(Count_ACCOUNT_digit, "0"))
            End If

            '   Else
            '        rs("Account_Serial").value = Get_Account_Serial(StrParentAccCode) & "00" & get_account_max(Get_Account_Serial(StrParentAccCode)) + 1 ' Replace(StrNewAccCode, "a", "", , , vbTextCompare)
            '   End If
          
        Else
            rs("Account_Serial").Value = serial
        End If
        
    End If
    
    rs("BasicAccount").Value = IIf(StrParentAccCode = "", 1, BasicAccount)
    rs("DateCreated").Value = Date

    If StrNewAccountNamee = "" Then
        rs("Account_NameEng").Value = StrNewAccountName
    Else
        rs("Account_NameEng").Value = StrNewAccountNamee
    End If
    If opening_balance <> 0 Then
        rs!opening_balance = opening_balance
    End If
    rs("currenct_code").Value = currenct_code
    rs("mowazna").Value = budget
    rs("cost_center").Value = cost_center
    rs("Sum_account").Value = Sum_account
   
    rs("cost_center_type").Value = cost_center_type
    rs("cost_center_id").Value = cost_center_id
    rs("ActivityTypeId").Value = ActivityTypeId
    
    rs.Update
    rs.Close
    Set rs = Nothing
    AddNewAccount = StrNewAccCode
    Exit Function
End Function

Public Function GetAccountsLevel(AccountsLevelsid) As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select * from AccountsLevelsDetails where " & "Level" & "=" & AccountsLevelsid
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        GetAccountsLevel = 1
    Else
        GetAccountsLevel = IIf(IsNull(rs("NoOfDigits").Value), 0, rs("NoOfDigits").Value)
    End If

End Function



Public Function CHECK_LAST_ACCOUNT(account As String) As Boolean
    Dim rs As ADODB.Recordset
  ' If account = "r" Then CHECK_LAST_ACCOUNT = False: Exit Function
    StrSQL = "Select * From Accounts Where Account_Code='" & account & "'"
    Set rs = New ADODB.Recordset
    
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic

    If Not (rs.BOF Or rs.EOF) Then
        
        If rs("last_account").Value = True Then
            CHECK_LAST_ACCOUNT = True: Exit Function
        Else
            CHECK_LAST_ACCOUNT = False: Exit Function
        End If

    Else
        CHECK_LAST_ACCOUNT = False: Exit Function
    End If
    
  
End Function



Public Function ParentAccountPrperties(StrParentAccCode As String, _
                                       Optional ByRef AccountTypes As Integer = 0, _
                                       Optional ByRef AccountTab As Integer = 0, _
                                       Optional ByRef DepitOrCreditv As Integer = 0, _
                                       Optional ByRef Differenttypev As Integer = 0, _
                                       Optional ByRef Authorityv As Integer = 0, _
                                       Optional ByRef UserGroupIdv As Integer = 0, _
                                       Optional ByRef UserIdv As Integer = 0)

    Dim StrSQL As String
    Dim rs As New ADODB.Recordset
    Dim rs22 As New ADODB.Recordset
    If StrParentAccCode = "r" Then
        AccountTypes = 1
    End If
    StrSQL = "Select * from ACCOUNTS where Account_Code='" & StrParentAccCode & "'"
    rs22.Open StrSQL, Cn, adOpenStatic, adLockOptimistic

    If rs22.RecordCount > 0 Then
        AccountTypes = rs22("AccountTypes").Value
        AccountTab = rs22("AccountTab").Value
        DepitOrCreditv = IIf(IsNull(rs22("DepitOrCredit").Value), 0, rs22("DepitOrCredit").Value)
        Differenttypev = IIf(IsNull(rs22("Differenttype").Value), 1, rs22("Differenttype").Value)
        Authorityv = rs22("Authority").Value

    Else

    End If

End Function






Public Function Get_Account_Serial(AccCode As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from ACCOUNTS where Account_Code='" & AccCode & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Get_Account_Serial = "": Exit Function
    If IsNull(Rs3("Account_Serial").Value) Then Get_Account_Serial = "": Exit Function
    If Not IsNull(Rs3("Account_Serial").Value) Then Get_Account_Serial = Rs3("Account_Serial").Value: Exit Function
    Rs3.Close

End Function



Private Function GetNewAcountCode(StrParentAccountCode As String) As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Long, j As Long
    Dim LngMax As Long
    Dim pos As Integer
    If StrParentAccountCode = "" Then
        StrSQL = "SELECT Max(Account_Serial) Account_Serial "
        StrSQL = StrSQL + " From ACCOUNTS Where BasicAccount = 1"
    Else
        StrSQL = "SELECT Account_Code "
        StrSQL = StrSQL + " From ACCOUNTS  Where ((ACCOUNTS.Parent_Account_Code='" & StrParentAccountCode & "'))"
        StrSQL = StrSQL + " ORDER by ACCOUNTS.Account_ID "
    End If
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        GetNewAcountCode = StrParentAccountCode & "a" & 1
        Exit Function
    Else
        If StrParentAccountCode <> "" Then
        
            pos = find_a_pos(rs("Account_Code").Value)
    
            LngMax = Mid(rs("Account_Code").Value, pos + 1, Len(rs("Account_Code").Value) - pos)
    
            For i = 0 To rs.RecordCount - 1
                pos = find_a_pos(rs("Account_Code").Value)
    
                If Mid(rs("Account_Code").Value, pos + 1, Len(rs("Account_Code").Value) - pos) > LngMax Then
                    LngMax = Mid(rs("Account_Code").Value, pos + 1, Len(rs("Account_Code").Value) - pos)
                End If
         
                rs.MoveNext
            Next i
            GetNewAcountCode = StrParentAccountCode & "a" & (LngMax + 1)
        Else
            GetNewAcountCode = "a" & IIf(rs!account_serial & "" = "", 1, Val(rs!account_serial & "") + 1)
        
        End If
        
    End If

End Function



Public Function CountAs(str As String) As Integer
    Dim count As Integer

    For i = 1 To Len(str)

        If Mid$(str, i, 1) = "a" Then count = count + 1
    Next

    CountAs = count
End Function



Public Function get_account_max(account_serial As String, _
                                Optional StrParentAccCode As String) As Double
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset

    Dim sql As String

    Dim max_no As String
    Dim i As Integer
    Dim ACCOUNT_CODE_AS As Integer
    ACCOUNT_CODE_AS = CountA(StrParentAccCode) + 1
    'Sql = "Select * from ACCOUNTS where Account_Serial like'" & account_serial & "%'"
    ' Sql = "Select max(cast(Account_Serial as float))  as max_no from ACCOUNTS where Account_Serial like'" & account_serial & "%'"
    '  Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    Dim account_root_lenght As Integer
    Dim max_no_lenght As Double
  
    'max_no_lenght = IIf(IsNull(Rs3("max_no").value) = False, Len(Rs3("max_no").value), 0)
  
    account_root_lenght = Len(account_serial)
 
    'Sql = "Select max(cast(right(Account_Serial , " & max_no_lenght - account_root_lenght & ") as float ))  as max_no from ACCOUNTS where Account_Serial like'" & account_serial & "%'" & "AND LEN(REPLACE(ACCOUNT_CODE, 'a', ''))=" & 1 'ACCOUNT_CODE_AS
    sql = "Select max(cast(account_serial as float) )  as max_no , max(account_serial   )  as max_no1  from ACCOUNTS where Account_Serial like'" & account_serial & "%' AND LEN(account_code) -LEN(REPLACE(ACCOUNT_CODE, 'a', ''))=" & ACCOUNT_CODE_AS
    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
    Dim max_lenght As Double
 
    If Rs4.RecordCount = 0 Or IsNull(Rs4("max_no").Value) Then get_account_max = 0: Exit Function
   
    Dim start_zero  As Integer
    start_zero = 0
    start_zero = 0

    If IsNull(Rs4("max_no1").Value) Then
   
    Else

        For i = 1 To Len(Rs4("max_no1").Value)

            If Mid(Rs4("max_no1").Value, i, 1) = "0" Then
                start_zero = start_zero + 1
                Else: GoTo mm
            End If
                    
        Next i

    End If

mm:
    max_no = IIf(IsNull(Rs4("max_no").Value), 0, Rs4("max_no").Value)
   
    max_lenght = Len(max_no) - account_root_lenght + start_zero

    If max_lenght <= 0 Then GoTo ll
    max_no = Right(max_no, max_lenght)
   
ll:
    get_account_max = max_no

End Function




Private Function find_a_pos(X As String) As Integer
    Dim pos As Integer
    Dim i As Integer

    For i = 1 To Len(X)

        If Mid(X, i, 1) = "a" Then
            pos = i
        End If

    Next i

    find_a_pos = pos

End Function




Public Function CountA(ByVal sText As String) As Long
    Dim bArr() As Byte
    Dim i As Long
    Dim count As Long
 
    For i = 1 To Len(sText)

        ' if this char is a space, increase the counter
        If Mid(sText, i, 1) = "a" Then count = count + 1
    Next

    CountA = count
End Function





Private Sub Form_Load()
'Con
End Sub
