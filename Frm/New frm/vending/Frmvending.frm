VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frmvending 
   ClientHeight    =   9060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15390
   Icon            =   "Frmvending.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   15390
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   9960
      TabIndex        =   16
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox Txtmachine_id 
      Height          =   375
      Left            =   18000
      TabIndex        =   15
      Top             =   480
      Width           =   1335
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   8625
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   19995
      _cx             =   35269
      _cy             =   15214
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
      BackColorFixed  =   14871017
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"Frmvending.frx":000C
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
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
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   15000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.TextBox txtIn 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   21360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   12480
      TabIndex        =   7
      Top             =   8640
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "„“«„‰…"
      Height          =   375
      Left            =   13200
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton CMDSelectFile 
      Caption         =   "Õœœ «·„·›"
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   375
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "C:\Users\Dynamic\Desktop\vendon_product-sales-2016-10-16_142308.xlsx"
      Top             =   0
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.CommandButton CmdImport 
      Caption         =   "«” Ì—«œ «·„·›"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   375
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dbToDate 
      Height          =   315
      Left            =   15600
      TabIndex        =   5
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   94240769
      CurrentDate     =   38784
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8355
      Left            =   13680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   6015
   End
   Begin MSComCtl2.DTPicker DbFromDate 
      Height          =   315
      Left            =   18000
      TabIndex        =   12
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   94240769
      CurrentDate     =   38784
   End
   Begin VB.Label Label1 
      Caption         =   "„«ﬂÌ‰… „Õœœ…"
      Height          =   375
      Left            =   19320
      TabIndex        =   14
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·› —… «·Ì"
      Height          =   210
      Index           =   1
      Left            =   16800
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Tag             =   "53"
      Top             =   120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·› —… „‰"
      Height          =   210
      Index           =   0
      Left            =   19080
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Tag             =   "53"
      Top             =   120
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   7125
      Left            =   13080
      Picture         =   "Frmvending.frx":0277
      Top             =   2520
      Visible         =   0   'False
      Width           =   11205
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„·›"
      Height          =   210
      Index           =   15
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Tag             =   "53"
      Top             =   15
      Visible         =   0   'False
      Width           =   1020
   End
End
Attribute VB_Name = "Frmvending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim BillTOTAL As Double
Dim CostTOTAL As Double
 Dim Account_Code_dynamic As String
            Dim StrTempAccountCode As String
            Dim CostAccount As String
  Dim StoreAccount As String
  Dim TxtNoteSerial1V As String
  
 
Dim strXML As String
Private Sub CMDSelectFile_Click()
CD1.ShowOpen
txtFile.Text = CD1.filename
 End Sub

Function CREATE_VOUCHER_GE(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchID As Integer, StoreID As Double, Transaction_Date As Date, BoxID As Double)
Dim LngDevID As Long
Dim LngDevNO As Integer
 Dim StrTempDes As String
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·œ«∆‰
     
    my_branch = BranchID
LngDevNO = 1
    StrTempDes = "„»Ì⁄«  «·Ì…"

 
Account_Code_dynamic = get_account_code_branch(2, my_branch)
   StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", CLng(BoxID))  '????????

If BillTOTAL > 0 Then
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, BillTOTAL, 0, StrTempDes, general_noteid, , , , Transaction_Date, val(Use_Id), Transaction_ID, , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
        LngDevNO = LngDevNO + 1
        
                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code_dynamic, BillTOTAL, 1, StrTempDes, general_noteid, , , , Transaction_Date, val(Use_Id), Transaction_ID, , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            
End If

     
   ' Dim StrSQL  As String
   ' StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & Transaction_ID
   ' Cn.Execute StrSQL
ErrTrap:
End Function

Function CREATE_VOUCHER_GE1(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchID As Integer, StoreID As Double, Transaction_Date As Date, BoxID As Double)
Dim LngDevID As Long
Dim LngDevNO As Integer
 Dim StrTempDes As String
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·œ«∆‰
     
    my_branch = BranchID
LngDevNO = 1
    StrTempDes = "”‰œ ’—› »‰«¡ ⁄·Ì „»Ì⁄«  «·Ì…"

 
Account_Code_dynamic = get_account_code_branch(2, my_branch)
   StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", CLng(BoxID))  '????????

If CostTOTAL > 0 Then
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, CostAccount, CostTOTAL, 0, StrTempDes, general_noteid, , , , Transaction_Date, val(Use_Id), Transaction_ID, , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
        LngDevNO = LngDevNO + 1
        
                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StoreAccount, CostTOTAL, 1, StrTempDes, general_noteid, , , , Transaction_Date, val(Use_Id), Transaction_ID, , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            
End If

     
   ' Dim StrSQL  As String
   ' StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & Transaction_ID
   ' Cn.Execute StrSQL
ErrTrap:
End Function



Private Sub createVoucher(BranchID As Double, _
BoxID As Double, _
Transaction_Date As Date, _
Transaction_Type As Double, _
CBoBasedON As Double, _
UserID As Double, _
Trans_DiscountType As Double, _
CusID As Double, _
StoreID As Double, _
PaymentType As Double, _
Emp_id As Double, _
TransactionComment As String)
Dim Sql As String
Dim Msg As String
Dim NoteID As Long
Dim Transaction_ID As Long
Dim Transaction_ID1 As Long
Dim Transaction_serial As String
Dim NoteSerial As String
Dim NoteSerial1 As String
Dim StrSQL As String

BillTOTAL = 0
CostTOTAL = 0
'Check
  NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 7, 170, , 21)
  
        If NoteSerial1 = "" Then
                 If NoteSerial1 = "error" Then
                     MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—› „Œ“‰Ì ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                 ElseIf NoteSerial1 = "" Then
                         MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
        
                 End If
        End If

NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
 If NoteSerial = "" Then
            If NoteSerial = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            ElseIf NoteSerial = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                 
            End If
End If
           
            Account_Code_dynamic = get_account_code_branch(2, CInt(BranchID))
        
            If Account_Code_dynamic = "NO branch" Or Account_Code_dynamic = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ —»ÿ «·Õ”«»«  ··„»Ì⁄« ", vbCritical
                Else
                    MsgBox "Sales Not Created", vbCritical
                End If

                Exit Sub
              End If
              
  
  
 
           CostAccount = get_account_code_branch(1, CInt(BranchID))
        
            If CostAccount = "NO branch" Or CostAccount = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ —»ÿ  ﬂ·›… «·„»Ì⁄«   ", vbCritical
                Else
                    MsgBox "Sales Not Created", vbCritical
                End If

             Exit Sub
              End If
              
              

    StoreAccount = get_store_Account(CInt(StoreID), "Account_Code")
      If StoreAccount = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
           Exit Sub
            End If


 'end Check
        
 NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 7, 170, , 21)
Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        
 Sql = "INSERT INTO  Transactions (  "
Sql = Sql & " Transaction_ID ,"
Sql = Sql & " BranchID ,"
Sql = Sql & " NoteSerial ,"
Sql = Sql & " NoteSerial1 ,"
Sql = Sql & " boxId ,"
Sql = Sql & " Transaction_serial ,"
Sql = Sql & " Transaction_Date ,"
Sql = Sql & " Transaction_Type ,"
Sql = Sql & " CBoBasedON ,"
Sql = Sql & " UserID ,"
Sql = Sql & " Trans_DiscountType ,"
Sql = Sql & " CusID ,"
Sql = Sql & " StoreId ,"
Sql = Sql & " PaymentType ,"
Sql = Sql & " Emp_id ,"
 Sql = Sql & " TransactionComment )"
 
 Sql = Sql & " VALUES("
Sql = Sql & " " & Transaction_ID & " ,"
Sql = Sql & " " & BranchID & " ,"
Sql = Sql & "'" & NoteSerial & "' ,"
Sql = Sql & "'" & NoteSerial1 & "' ,"
Sql = Sql & " " & BoxID & " ,"
Sql = Sql & "'" & Transaction_serial & "',"
Sql = Sql & " " & SQLDate(Transaction_Date, True) & " ,"
Sql = Sql & " " & 21 & " ,"
Sql = Sql & " 0 ,"
Sql = Sql & " " & user_id & " ,"
Sql = Sql & " 0 ,"
Sql = Sql & " " & CusID & " ,"
Sql = Sql & " " & StoreID & " ,"
Sql = Sql & " 0 ,"
Sql = Sql & " " & Emp_id & " ,"
 Sql = Sql & "'" & TransactionComment & "')"
 

         Cn.Execute Sql
 



 
        Dim RSTransDetails As New ADODB.Recordset
     
StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   Dim RowNum As Integer
        For RowNum = 1 To FG.Rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("ItemId")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = Transaction_ID
             
                RSTransDetails("ColorID").value = 1
                RSTransDetails("ItemSize").value = 1
                RSTransDetails("ClassId").value = 1
                
         
        
                RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemId")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemId"))))
                RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Quantity")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Quantity"))))
                RSTransDetails("SHOWQTY").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Quantity")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Quantity"))))
               RSTransDetails("Price").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
               RSTransDetails("CostPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("CostPrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("CostPrice"))))
               
               BillTOTAL = BillTOTAL + (RSTransDetails("Price").value * RSTransDetails("SHOWQTY").value)
           CostTOTAL = CostTOTAL + (RSTransDetails("CostPrice").value * RSTransDetails("SHOWQTY").value)
                
                RSTransDetails("CostPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("CostPrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("CostPrice"))))
               RSTransDetails("showPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
             RSTransDetails("UnitID").value = 1
     RSTransDetails("SavedItemType").value = 0
            
                RSTransDetails.update
            End If

        Next RowNum
NoteSerial = Notes_coding(val(BranchID), Transaction_Date)

CreateNotes NoteID, Transaction_Date, CInt(BranchID), 170, BillTOTAL, NoteSerial, NoteSerial1, "Transactions", "Transaction_ID", Transaction_ID, NoteSerial1, ToHijriDate(Transaction_Date)

CREATE_VOUCHER_GE Transaction_ID, NoteSerial, NoteSerial1, NoteID, CInt(BranchID), StoreID, Transaction_Date, BoxID
        ' MsgBox " „ «‰‘«¡ «·”‰œ"
        
     '   StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
     '   Cn.Execute StrSQL
        
'******************************************************issueVoucher

 Transaction_ID1 = CStr(new_id("Transactions", "Transaction_ID", "", True))
TxtNoteSerial1V = Voucher_coding(val(BranchID), Transaction_Date, 10, 180, , 19)


Dim TxtNoteSerialV As String


        Sql = "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,nots2,NoteSerial,NoteSerial1,NoteId,BranchId,Closed,ManualNO)SELECT " & Transaction_ID1 & ", '',Transaction_Date,Transaction_Type = 19,CusID,StoreID,UserID,Emp_ID,nots=" & Transaction_ID & ",nots2='" & NoteSerial1 & "' ,NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & 0 & ",BranchId,1,ManualNO From Transactions Where  Transaction_ID =" & Transaction_ID
        Cn.Execute Sql
        '
        Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ProductionDate,ExpiryDate,LotNO)SELECT  costprice,guaranteeTime," & Transaction_ID1 & ",Item_ID,ItemCase,ItemSerial , Quantity, costprice/ QtyBySmalltUnit ,ColorID,ItemSize, UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID ,ProductionDate,ExpiryDate,LotNO From dbo.Transaction_Details Where SavedItemType=0 and   Transaction_ID = " & Transaction_ID
NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
       CreateNotes NoteID, Transaction_Date, CInt(BranchID), 180, CostTOTAL, NoteSerial, TxtNoteSerial1V, "Transactions", "Transaction_ID", Transaction_ID1, TxtNoteSerial1V, ToHijriDate(Transaction_Date)

CREATE_VOUCHER_GE1 Transaction_ID1, NoteSerial, TxtNoteSerial1V, NoteID, CInt(BranchID), StoreID, Transaction_Date, BoxID
  
'***********************
         StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID1 & " WHERE Transaction_ID=" & Transaction_ID
         Cn.Execute StrSQL
'***********************

   
       
 
        'StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
        'Cn.Execute StrSQL
  MsgBox " „   «·‰ﬁ·"
  
'******************************************************issueVoucher








     
 
    '
 
ErrTrap:

End Sub

 
Private Sub CmdImport_Click()
On Error Resume Next
If txtFile.Text = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
Dim currentvalue As String

Dim BranchID As Double
Dim account_serial As String
Dim des As String
Dim DebitValue As String
Dim CreditValue As String
  

    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")

    ExcelObj.Workbooks.Open txtFile.Text    ' App.Path & "\TrialBalance.xls"
DoEvents
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
 Dim Device As String
 Dim Lines As String
 Dim Product As String
 Dim Quantity As String
 Dim PayType As String
 
 Dim DeviceId As Double
'Dim BtanchID As Double
Dim ItemID As Double
Dim CashType As Integer
Dim SalesPersonId As Double
Dim BoxID As Double
Dim CurrencyId As Double
Dim Price As Double
Dim costprice As Double

Dim CusID As Double
Dim StoreID As Double
 'Dim BranchID As Double

 
    With ExcelSheet
    i = 8
  FIXEDROW = 7
    Do Until .cells(i, 1) & "" = ""
 '       Set l = lvwList.ListItems.Add(, , .Cells(i, 1))
   Device = .cells(i, 1)
   Lines = .cells(i, 2)
   Product = .cells(i, 3)
   Quantity = .cells(i, 4)
   PayType = .cells(i, 6)
 GetStoreData Device, DeviceId, BranchID
    
     
 With FG
       
  .TextMatrix(i - FIXEDROW, .ColIndex("Device")) = (Device)
    .TextMatrix(i - FIXEDROW, .ColIndex("DeviceId")) = (DeviceId)
        .TextMatrix(i - FIXEDROW, .ColIndex("BtanchID")) = (BranchID)
        
  .TextMatrix(i - FIXEDROW, .ColIndex("Lines")) = (Lines)
  .TextMatrix(i - FIXEDROW, .ColIndex("Product")) = (Product)
  GetItemsData Product, ItemID
  .TextMatrix(i - FIXEDROW, .ColIndex("ItemID")) = (ItemID)
  
  Price = GetItemPrice(CLng(ItemID), 1, 1)
costprice = ModItemCostPrice.GetCostItemPrice(CLng(ItemID), 0, "", , SystemOptions.SysMainStockCostMethod, , , dbTodate, , 1)

       .TextMatrix(i - FIXEDROW, .ColIndex("Price")) = (Price)
       .TextMatrix(i - FIXEDROW, .ColIndex("CostPrice")) = (costprice)
  
  
 .TextMatrix(i - FIXEDROW, .ColIndex("Quantity")) = (Quantity)
 .TextMatrix(i - FIXEDROW, .ColIndex("Type")) = (PayType)
If .TextMatrix(i - FIXEDROW, .ColIndex("Type")) = "CASH" Then
.TextMatrix(i - FIXEDROW, .ColIndex("CashType")) = "0"
Else
.TextMatrix(i - FIXEDROW, .ColIndex("CashType")) = "1"
End If
SalesPersonId = 1
BoxID = 1
CurrencyId = 1
CusID = 2
AddItems val(Lines), CusID, DeviceId, BranchID, SalesPersonId, Me.dbTodate, ItemID, val(Quantity), Price, costprice
' Lines, 1
         .Row = i - FIXEDROW
                             .Col = .ColIndex("Serial")
                             .ShowCell i - FIXEDROW, .ColIndex("Serial")
                            
                             .SetFocus


 End With
 If .cells(i, 1) & "" = "" Then Exit Sub
        i = i + 1
    Loop

    End With

       ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing
createVoucher BranchID, BoxID, dbTodate, 21, 0, val(user_id), 0, CusID, DeviceId, 0, SalesPersonId, "»Ì«‰«  «·Ì… „‰ ‰ﬁ«ÿ «·»Ì⁄"

End Sub
Function AddItems(Lines As Double, CusID As Double, StoreID As Double, BranchID As Double, Emp_id As Double, Transaction_Date As Date, Item_ID As Double, ShowQty As Double, Price As Double, costprice As Double)

Dim StrSQL As String
StrSQL = "insert into TblVending (Lines,CusID,StoreID,BranchID,Emp_id,Transaction_Date,Item_ID,ShowQty,Price,CostPrice)   "
StrSQL = StrSQL & "  Values (" & Lines & "," & CusID & "," & StoreID & "," & BranchID & "," & Emp_id & "," & SQLDate(Transaction_Date, True) & "," & Item_ID & "," & ShowQty & "," & Price & "," & costprice & ")   "
Cn.Execute StrSQL
End Function

Private Sub Command1_Click()
Dim Result As String
Dim URL As String
Dim BoxID As Double
'users
'url = url & "https://v3.vendon.net/rest/v1.3.6/user"

'sales
Dim unixfrom As String
Dim unixto As String
unixfrom = TimeStamp(dbFromDate.value)

If dbFromDate.value <> dbTodate.value Then
unixto = TimeStamp(dbTodate.value)
 Else
 unixto = TimeStamp(DateAdd("d", 1, dbTodate.value))
 End If
'URL = "http://v3.vendon.net/rest/v1.3.7/stats/machineSales?from_timestamp=1476316800&to_timestamp=1476731103"
URL = "http://v3.vendon.net/rest/v1.3.7/stats/machineSales?from_timestamp=" & unixfrom & "&to_timestamp=" & unixto & ""

'Me.WbHelp.Navigate websiteurl

  FG.Clear flexClearScrollable, flexClearEverything
            FG.Rows = 2
            FG.Enabled = True
            
'alldata what we need
'URL = "http://v3.vendon.net/rest/v1.3.7/stats/vends?from_timestamp=1476316800&to_timestamp=1476731103"
If Txtmachine_id = "" Then
URL = "http://v3.vendon.net/rest/v1.3.7/stats/vends?from_timestamp=" & unixfrom & "&to_timestamp=" & unixto
Else
URL = "http://v3.vendon.net/rest/v1.3.7/stats/vends?from_timestamp=" & unixfrom & "&to_timestamp=" & unixto & "& machine_id=" & Txtmachine_id & ""
End If
Result = WebRequest(URL)

Text1 = Result
If Text1 = "" Then Exit Sub
DoEvents
'******************************jSON
 If Text1.Text = "" Then Exit Sub
  Dim F As Integer
    Dim JSON As String
    Dim Candles As JsonBag
    Dim i As Long
    Dim DateValue As Date
         Dim StoreName As String
        Dim StoreID As Double
    Dim BranchID As Double
    Dim ItemID As Double
    Dim StoreFullCode As String
    Dim ProductName As String
    Dim ItemFullcode As String
    Dim CusID As Double
    Dim DeviceId As Double
Dim SalesPersonId As Double
Dim Price As Double
Dim costprice As Double

  JSON = Text1.Text
    With New JsonBag
        .DecimalMode = True
        .JSON = JSON
   ' .Whitespace = True
        txtIn.Text = .JSON
        Dim Strdata As String
        Set Candles = .Item("result")
        Dim Transaction_ID As Double
       Dim Transaction_Date As Date
       Dim FIXEDROW As Integer
       Dim Quantity As Double
       Dim Lines As Double
       Dim payment_method As String
       FIXEDROW = 0
       If Candles.count = 0 Then Exit Sub
        For i = 1 To Candles.count

    Strdata = Candles(i).Item("transaction_id")
      Transaction_ID = Candles(i).Item("transaction_id")
     '  Transaction_Date = UnixToDate()
     Transaction_Date = DateAdd("s", Candles(i).Item("transaction_dt"), DateSerial(1970, 1, 1))
        'ToUnixTime = DateDiff("s", DateSerial(1970, 1, 1), time)
        'FromUnixTime = DateAdd("s", UnixTime, DateSerial(1970, 1, 1))
       StoreFullCode = Candles(i).Item("machine_id")
       Quantity = Candles(i).Item("quantity")
        Price = Candles(i).Item("price")
        ItemFullcode = Candles(i).Item("stock_id")
        Lines = Candles(i).Item("selection")
        payment_method = Candles(i).Item("payment_method")
        GetStoreData StoreName, StoreID, BranchID, StoreFullCode
        With Me.FG
  .TextMatrix(i - FIXEDROW, .ColIndex("Transaction_ID")) = Transaction_ID
  .TextMatrix(i - FIXEDROW, .ColIndex("Transaction_Date")) = Transaction_Date
  .TextMatrix(i - FIXEDROW, .ColIndex("DeviceId")) = StoreID
.TextMatrix(i - FIXEDROW, .ColIndex("Device")) = (StoreFullCode)
.TextMatrix(i - FIXEDROW, .ColIndex("storename")) = (StoreName)
  .TextMatrix(i - FIXEDROW, .ColIndex("BtanchID")) = (BranchID)
        
  .TextMatrix(i - FIXEDROW, .ColIndex("Lines")) = (Lines)
  GetItemsData ProductName, ItemID, ItemFullcode
  .TextMatrix(i - FIXEDROW, .ColIndex("ItemID")) = (ItemID)
  .TextMatrix(i - FIXEDROW, .ColIndex("ItemFullcode")) = (ItemFullcode)
  .TextMatrix(i - FIXEDROW, .ColIndex("Product")) = (ProductName)
   'Price = GetItemPrice(CLng(ItemID), 1, 1)
costprice = ModItemCostPrice.GetCostItemPrice(CLng(ItemID), 0, "", , SystemOptions.SysMainStockCostMethod, , , dbTodate, , 1)

       .TextMatrix(i - FIXEDROW, .ColIndex("Price")) = (Price)
       .TextMatrix(i - FIXEDROW, .ColIndex("CostPrice")) = (costprice)
  
  
 .TextMatrix(i - FIXEDROW, .ColIndex("Quantity")) = (Quantity)
 .TextMatrix(i - FIXEDROW, .ColIndex("Type")) = payment_method
If .TextMatrix(i - FIXEDROW, .ColIndex("Type")) = "CASH" Then
.TextMatrix(i - FIXEDROW, .ColIndex("CashType")) = "0"
Else
.TextMatrix(i - FIXEDROW, .ColIndex("CashType")) = "1"
End If
SalesPersonId = 1
BoxID = 1
'CurrencyId = 1
CusID = 2
ProductName = ""
StoreName = ""
If StoreID <> 0 Then
AddItems val(Lines), CusID, StoreID, BranchID, SalesPersonId, Me.dbTodate, ItemID, val(Quantity), Price, costprice
End If
' Lines, 1
      '   .Row = I - fixedRow
                          '   .Col = .ColIndex("Serial")
                          '   .ShowCell I - fixedRow, .ColIndex("Serial")
                            
                          '   .SetFocus

 .Rows = .Rows + 1
    End With
 
'  I = I + 1
           
        Next
        
'        txtOut.Text = .JSON
    End With
 


'******************************jSON
Dim StrSQL As String
StrSQL = "DELETE TblVending "
 
Cn.Execute StrSQL

createVoucher BranchID, BoxID, dbTodate, 21, 0, val(user_id), 0, CusID, StoreID, 0, SalesPersonId, "»Ì«‰«  «·Ì… „‰ ‰ﬁ«ÿ «·»Ì⁄"

End Sub
Function AddItemToGrid(Device As String, DeviceId As Double, BranchID As Double, Lines)

    
End Function
Private Sub Command2_Click()
  Dim F As Integer
    Dim JSON As String
    Dim Candles As JsonBag
    Dim i As Long
    Dim DateValue As Date
    
  '  F = FreeFile(0)
  '  Open App.Path & "\sample.txt" For Input As #F
  '  JSON = Input$(LOF(F), #F)
  '  Close #F
  JSON = Text1.Text
    With New JsonBag
        .DecimalMode = True
        .JSON = JSON
   ' .Whitespace = True
        txtIn.Text = .JSON
        
        Set Candles = .Item("result")
        For i = 1 To Candles.count
        MsgBox Candles(i).Item("transaction_id")
        MsgBox Candles(i).Item("machine_id")
        Next
        
        txtOut.Text = .JSON
    End With
 
End Sub

Private Sub Command3_Click()
Dim Result As String
Dim URL As String
Dim BoxID As Double
'users
'url = url & "https://v3.vendon.net/rest/v1.3.6/user"

'sales
Dim unixfrom As String
Dim unixto As String
unixfrom = TimeStamp(dbFromDate.value)

If dbFromDate.value <> dbTodate.value Then
unixto = TimeStamp(dbTodate.value)
 Else
 unixto = TimeStamp(DateAdd("d", 1, dbTodate.value))
 End If
'URL = "http://v3.vendon.net/rest/v1.3.7/stats/machineSales?from_timestamp=1476316800&to_timestamp=1476731103"
URL = "http://v3.vendon.net/rest/v1.3.7/stats/machineSales?from_timestamp=" & unixfrom & "&to_timestamp=" & unixto & ""
URL = "http://run-code.com/ServiceSales.aspx?type=SelectBranch"
'Me.WbHelp.Navigate websiteurl

  FG.Clear flexClearScrollable, flexClearEverything
            FG.Rows = 2
            FG.Enabled = True
            
'alldata what we need
'URL = "http://v3.vendon.net/rest/v1.3.7/stats/vends?from_timestamp=1476316800&to_timestamp=1476731103"
If Txtmachine_id = "" Then
URL = "http://v3.vendon.net/rest/v1.3.7/stats/vends?from_timestamp=" & unixfrom & "&to_timestamp=" & unixto
Else
URL = "http://v3.vendon.net/rest/v1.3.7/stats/vends?from_timestamp=" & unixfrom & "&to_timestamp=" & unixto & "& machine_id=" & Txtmachine_id & ""
End If
URL = "http://run-code.com/ServiceSales.aspx?type=SelectBranch"
Result = WebRequest(URL)

Text1 = Result
If Text1 = "" Then Exit Sub
DoEvents
'******************************jSON
 If Text1.Text = "" Then Exit Sub
  Dim F As Integer
    Dim JSON As String
    Dim Candles As JsonBag
    Dim i As Long
    Dim DateValue As Date
         Dim StoreName As String
        Dim StoreID As Double
    Dim BranchID As Double
    Dim ItemID As Double
    Dim StoreFullCode As String
    Dim ProductName As String
    Dim ItemFullcode As String
    Dim CusID As Double
    Dim DeviceId As Double
Dim SalesPersonId As Double
Dim Price As Double
Dim costprice As Double

  JSON = Text1.Text
    With New JsonBag
        .DecimalMode = True
        .JSON = JSON
   ' .Whitespace = True
        txtIn.Text = .JSON
        Dim Strdata As String
        Set Candles = .Item("branch_id")
        Dim Transaction_ID As Double
       Dim Transaction_Date As Date
       Dim FIXEDROW As Integer
       Dim Quantity As Double
       Dim Lines As Double
       Dim payment_method As String
       FIXEDROW = 0
       If Candles.count = 0 Then Exit Sub
        For i = 1 To Candles.count

    Strdata = Candles(i).Item("transaction_id")
      Transaction_ID = Candles(i).Item("transaction_id")
     '  Transaction_Date = UnixToDate()
     Transaction_Date = DateAdd("s", Candles(i).Item("transaction_dt"), DateSerial(1970, 1, 1))
        'ToUnixTime = DateDiff("s", DateSerial(1970, 1, 1), time)
        'FromUnixTime = DateAdd("s", UnixTime, DateSerial(1970, 1, 1))
       StoreFullCode = Candles(i).Item("machine_id")
       Quantity = Candles(i).Item("quantity")
        Price = Candles(i).Item("price")
        ItemFullcode = Candles(i).Item("stock_id")
        Lines = Candles(i).Item("selection")
        payment_method = Candles(i).Item("payment_method")
        GetStoreData StoreName, StoreID, BranchID, StoreFullCode
        With Me.FG
  .TextMatrix(i - FIXEDROW, .ColIndex("Transaction_ID")) = Transaction_ID
  .TextMatrix(i - FIXEDROW, .ColIndex("Transaction_Date")) = Transaction_Date
  .TextMatrix(i - FIXEDROW, .ColIndex("DeviceId")) = StoreID
.TextMatrix(i - FIXEDROW, .ColIndex("Device")) = (StoreFullCode)
.TextMatrix(i - FIXEDROW, .ColIndex("storename")) = (StoreName)
  .TextMatrix(i - FIXEDROW, .ColIndex("BtanchID")) = (BranchID)
        
  .TextMatrix(i - FIXEDROW, .ColIndex("Lines")) = (Lines)
  GetItemsData ProductName, ItemID, ItemFullcode
  .TextMatrix(i - FIXEDROW, .ColIndex("ItemID")) = (ItemID)
  .TextMatrix(i - FIXEDROW, .ColIndex("ItemFullcode")) = (ItemFullcode)
  .TextMatrix(i - FIXEDROW, .ColIndex("Product")) = (ProductName)
   'Price = GetItemPrice(CLng(ItemID), 1, 1)
costprice = ModItemCostPrice.GetCostItemPrice(CLng(ItemID), 0, "", , SystemOptions.SysMainStockCostMethod, , , dbTodate, , 1)

       .TextMatrix(i - FIXEDROW, .ColIndex("Price")) = (Price)
       .TextMatrix(i - FIXEDROW, .ColIndex("CostPrice")) = (costprice)
  
  
 .TextMatrix(i - FIXEDROW, .ColIndex("Quantity")) = (Quantity)
 .TextMatrix(i - FIXEDROW, .ColIndex("Type")) = payment_method
If .TextMatrix(i - FIXEDROW, .ColIndex("Type")) = "CASH" Then
.TextMatrix(i - FIXEDROW, .ColIndex("CashType")) = "0"
Else
.TextMatrix(i - FIXEDROW, .ColIndex("CashType")) = "1"
End If
SalesPersonId = 1
BoxID = 1
'CurrencyId = 1
CusID = 2
ProductName = ""
StoreName = ""
If StoreID <> 0 Then
AddItems val(Lines), CusID, StoreID, BranchID, SalesPersonId, Me.dbTodate, ItemID, val(Quantity), Price, costprice
End If
' Lines, 1
      '   .Row = I - fixedRow
                          '   .Col = .ColIndex("Serial")
                          '   .ShowCell I - fixedRow, .ColIndex("Serial")
                            
                          '   .SetFocus

 .Rows = .Rows + 1
    End With
 
'  I = I + 1
           
        Next
        
'        txtOut.Text = .JSON
    End With
 


'******************************jSON
Dim StrSQL As String
StrSQL = "DELETE TblVending "
 
Cn.Execute StrSQL

createVoucher BranchID, BoxID, dbTodate, 21, 0, val(user_id), 0, CusID, StoreID, 0, SalesPersonId, "»Ì«‰«  «·Ì… „‰ ‰ﬁ«ÿ «·»Ì⁄"


End Sub

Private Sub Form_Load()
dbTodate.value = Date
Me.dbFromDate.value = Date

Me.Caption = TimeStamp(dbTodate.value)
      FG.Clear flexClearScrollable, flexClearEverything
            FG.Rows = 2
            FG.Enabled = True
End Sub

