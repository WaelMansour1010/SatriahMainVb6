VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMTRansferData2 
   Caption         =   "‰ﬁ· «·»Ì«‰« "
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9315
   Icon            =   "FRMTRansferData2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   9315
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
   Begin VB.CommandButton Command7 
      Caption         =   " ÕœÌÀ Ãœ«Ê· «·Õ—ﬂ«  «·«”«”Ì…"
      Height          =   375
      Left            =   4260
      TabIndex        =   36
      Top             =   2190
      Width           =   1935
   End
   Begin VB.TextBox TxtLicense 
      Alignment       =   1  'Right Justify
      Height          =   1095
      Left            =   0
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   26
      Top             =   7800
      Width           =   6435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "‰ﬁ· «·»Ì«‰«  „‰ «·‰ﬁÿÂ ··”Ì—›—"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   6960
      Width           =   2895
   End
   Begin VB.ComboBox ServersName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   810
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   7350
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.ComboBox DbName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1050
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   7230
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.CommandButton Command2 
      Caption         =   " ÕœÌÀ „·› «·⁄„·«¡  "
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   4830
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "«Œ »«— "
      Height          =   405
      Left            =   240
      TabIndex        =   21
      Top             =   5490
      Width           =   2625
   End
   Begin VB.CommandButton cmdUpdateEmp 
      Caption         =   " ÕœÌÀ „·› «·„ÊŸ›Ì‰ „‰ «·‰ﬁÿÂ"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   6480
      Width           =   2835
   End
   Begin VB.CommandButton cmdUdateFiles 
      Caption         =   " ÕœÌÀ «·„·›«  «·«”«”Ì… „‰ «·”Ì—›— "
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   6060
      Width           =   2655
   End
   Begin VB.Frame ServerData 
      Caption         =   "POS Data"
      Height          =   1935
      Left            =   30
      TabIndex        =   10
      Top             =   1440
      Width           =   3495
      Begin VB.TextBox POSlServer 
         Height          =   375
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox TxtPOSDB 
         Height          =   375
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Text            =   "LOCALPOS"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox POSname 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   3345
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Server name"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DBname"
         Height          =   375
         Left            =   -240
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server Data"
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3375
      Begin VB.TextBox DestinationServer 
         Height          =   375
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TxtServerDataBaseName 
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Text            =   "byte"
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DBname"
         Height          =   375
         Left            =   -360
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Server name"
         Height          =   375
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3360
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.TextBox Text2 
      Height          =   1365
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "FRMTRansferData2.frx":058A
      Top             =   3480
      Width           =   5775
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   90
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   4140
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.TextBox txtStartTime 
      Height          =   375
      Left            =   3420
      TabIndex        =   1
      Top             =   180
      Width           =   2205
   End
   Begin VB.TextBox txtEndTime 
      Height          =   375
      Left            =   3420
      TabIndex        =   0
      Top             =   630
      Width           =   2205
   End
   Begin MSComCtl2.DTPicker dbRecordDate 
      Height          =   285
      Left            =   90
      TabIndex        =   16
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   229507073
      CurrentDate     =   41640
   End
   Begin VSFlex8UCtl.VSFlexGrid grd 
      Height          =   4110
      Left            =   6840
      TabIndex        =   29
      Top             =   540
      Width           =   6795
      _cx             =   11986
      _cy             =   7250
      Appearance      =   2
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
      Rows            =   12
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FRMTRansferData2.frx":0590
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
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
      WallPaperAlignment=   0
      AccessibleName  =   "ReCostDet"
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8UCtl.VSFlexGrid Grd2 
      Height          =   4110
      Left            =   6840
      TabIndex        =   30
      Top             =   5220
      Width           =   7005
      _cx             =   12356
      _cy             =   7250
      Appearance      =   2
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
      Rows            =   12
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FRMTRansferData2.frx":0660
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
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
      WallPaperAlignment=   0
      AccessibleName  =   "ReCostDet"
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComCtl2.DTPicker txtToDate 
      Height          =   285
      Left            =   4530
      TabIndex        =   37
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   229572609
      CurrentDate     =   41640
   End
   Begin MSComCtl2.DTPicker txtFromDate 
      Height          =   285
      Left            =   4530
      TabIndex        =   38
      Top             =   1380
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   229572609
      CurrentDate     =   41640
   End
   Begin VB.Label Label4 
      Caption         =   "‰ﬁ· »Ì«‰«  «·„⁄œ« /«·”Ì«—«  Ê«·»’„Â „‰ «·‰ﬁÿÂ ··”Ì—›—"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3240
      TabIndex        =   35
      Top             =   6960
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   " ÕœÌÀ „·› «·„ÊŸ›Ì‰ „‰ «·‰ﬁÿÂ"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3120
      TabIndex        =   34
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "„Ê«ﬁ⁄ Ê „” Œœ„Ì‰ Ê’·«ÕÌ ⁄„ Ê⁄„·«¡"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3120
      TabIndex        =   33
      Top             =   6000
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "”Ã·«  „‰ﬁÊ·… „‰ «·”Ì—›—"
      Height          =   255
      Index           =   1
      Left            =   8850
      TabIndex        =   32
      Top             =   180
      Width           =   2865
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "”Ã·«  „‰ﬁÊ·… „‰ «·«ÃÂ“… «·›—⁄Ì…"
      Height          =   255
      Index           =   0
      Left            =   8910
      TabIndex        =   31
      Top             =   4830
      Width           =   2865
   End
   Begin VB.Label lblCount 
      Height          =   315
      Left            =   5610
      TabIndex        =   28
      Top             =   6060
      Width           =   945
   End
   Begin VB.Label lblWait 
      BackStyle       =   0  'Transparent
      Caption         =   "Ì—ÃÏ «·«‰ Ÿ«— Ã«—Ì ‰ﬁ· «·»Ì«‰« "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   3090
      TabIndex        =   27
      Top             =   5520
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Êﬁ  «·»œ«Ì…"
      Height          =   255
      Index           =   0
      Left            =   5760
      TabIndex        =   18
      Top             =   210
      Width           =   795
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Êﬁ  «·‰Â«Ì…"
      Height          =   255
      Left            =   5790
      TabIndex        =   17
      Top             =   690
      Width           =   795
   End
End
Attribute VB_Name = "FRMTRansferData2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 Dim CountItems As Long, CountSales As Long, CountSalesReturn As Long, CountPurchase As Long, CountPurchaseReturn As Long, mCounCountRec
 Dim cProgress As ClsProgress
 Dim BolFrmLoaded As Boolean

 
 


Private Sub CmdOpen_Click()
cd1.ShowOpen
 
txtDbPath.Text = cd1.FileName


End Sub
Function CopyIssueTtransaction(invoiceTransaction_ID As Double, invoiceNoteserial1 As String, Transaction_ID As Double, Transaction_Type As Double, issuenoteserial As String, issuenoteserial1 As String, SessionCode As String)
'////////////////////////////////////////copy   Transactions

  Dim Rs3 As ADODB.Recordset
  Dim rsDouble_Entry As ADODB.Recordset
  
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim mytext As String
    
 sql = " select * from Transactions    WHERE Transaction_ID =" & Transaction_ID
 
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
 Dim FromTransaction_ID As Double
 Dim FromBranchID As Integer
 Dim FromTransaction_Date As Date
 Dim FromNots As String
  Dim FromNots2 As String
   Dim fromTransaction_Serial As String
 Dim FromNoteseial1 As String
 Dim FromTransaction_Type As Integer
 
Dim BranchID As Integer
 ' Dim Transaction_ID As Double
  Dim Transaction_Date As Date
   Dim Transaction_Serial  As String
 Dim Nots As String
Dim Nots2 As String
'Dim Transaction_Type As Integer
Dim FromNoteId As Double
 'sales
    If Rs3.RecordCount > 0 Then
      
            For i = 1 To Rs3.RecordCount
             FromTransaction_Type = IIf(IsNull(Rs3("Transaction_Type").Value), 0, Rs3("Transaction_Type").Value)
               FromTransaction_ID = IIf(IsNull(Rs3("Transaction_ID").Value), 0, Rs3("Transaction_ID").Value)
                
               
               
               FromBranchID = IIf(IsNull(Rs3("BranchID").Value), 0, Rs3("BranchID").Value)
               fromTransaction_Serial = IIf(IsNull(Rs3("Transaction_Serial").Value), 0, Rs3("Transaction_Serial").Value)
        
              
               FromNoteSerial1 = IIf(IsNull(Rs3("Noteserial1").Value), 0, Rs3("Noteserial1").Value)
                FromNoteSerial = IIf(IsNull(Rs3("Noteserial").Value), 0, Rs3("Noteserial").Value)
                FromNoteId = IIf(IsNull(Rs3("NoteId").Value), 0, Rs3("NoteId").Value) ' —ﬁ„ ﬁÌœ «·”‰œ
                
               FromNots2 = IIf(IsNull(Rs3("Nots2").Value), 0, Rs3("Nots2").Value) '—ﬁ„ «·›« Ê—… «·«’·ÌÏ…
               FromTransaction_Date = IIf(IsNull(Rs3("Transaction_Date").Value), 0, Rs3("Transaction_Date").Value)
              
                      Dim FromEmp_ID As Double

 Dim FromStoreID As Double
Dim FromCusID As Double
               Dim FromBoxid As Double
            Dim PayMentType As Integer
               Dim BillBasedOn
           'Dim BillBasedOn As Integer
              Dim VATYou As Double
               Dim VAT As Double
               Dim FromUserID As Double
               Dim POSBillType As Double
               FromUserID = IIf(IsNull(Rs3("UserID").Value), 0, Rs3("UserID").Value)
               FromEmp_ID = IIf(IsNull(Rs3("Emp_ID").Value), 0, Rs3("Emp_ID").Value)
               FromStoreID = IIf(IsNull(Rs3("storeID").Value), 0, Rs3("storeID").Value)
               FromCusID = IIf(IsNull(Rs3("CusID").Value), 0, Rs3("CusID").Value)
               
               FromBoxid = IIf(IsNull(Rs3("Boxid").Value), 0, Rs3("Boxid").Value)
               POSBillType = IIf(IsNull(Rs3("POSBillType").Value), 0, Rs3("POSBillType").Value)
               
               
                FromPaymentType = IIf(IsNull(Rs3("PaymentType").Value), 0, Rs3("PaymentType").Value)
                FromBillBasedOn = IIf(IsNull(Rs3("BillBasedOn").Value), 0, Rs3("BillBasedOn").Value)
                FromVATYou = IIf(IsNull(Rs3("VATYou").Value), 0, Rs3("VATYou").Value)
                FromVAT = IIf(IsNull(Rs3("VAT").Value), 0, Rs3("VAT").Value)
               '


              Transaction_Date = FromTransaction_Date
              Transaction_Type3 = FromTransaction_Type
              BranchID = FromBranchID
             Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
             Transaction_Serial = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=" & Transaction_Type & ""))
             If Transaction_Type = 19 Then
 
              NoteSerial1 = Voucher_coding(FromBranchID, FromTransaction_Date, 10, 180, , 19)
             ElseIf Transaction_Type = 20 Then
 
              NoteSerial1 = Voucher_coding(FromBranchID, FromTransaction_Date, 9, 160, , 20)
              
             End If
             
             
             NoteSerial = Notes_coding(FromBranchID, FromTransaction_Date)
             NoteId = CStr(new_id("Notes", "NoteID", "", True))
            TransactionComment = " ”‰œ „‰ﬁÊ· „‰ ﬁ«⁄œ…  " & POSname.Text & "   "
            TransactionComment = TransactionComment & "  —ﬁ„ «·”‰œ «·«’·Ì   " & FromNoteSerial1
             '" & ServerDb & "
             
 
              
'ÂÌœ— «·”‰œ
'*****************************************************************************************

'*****************************************************************************************
 sql = " INSERT INTO  [" & ServerDb & "].[dbo].[Transactions]  (    "
sql = sql & "  Transaction_ID,Transaction_Date, Transaction_Serial , Transaction_Type, PaymentType, CusID, StoreID, UserID, Emp_ID, BranchId, BoxID  "
sql = sql & " , BillBasedOn, VAT, VATYou, NoteSerial,NoteSerial1,NoteId,Copied,TransactionComment,closed,SessionCode,OldNoteSerial1,OldNoteSerial,OldNoteId,OldTransaction_ID)"
 
sql = sql & "   values (" & Transaction_ID & "," & SQLDate(Transaction_Date, True) & ", " & Transaction_Serial & "," & Transaction_Type & "," & FromPaymentType & "," & FromCusID & "," & FromStoreID & ",1," & FromEmp_ID & "," & FromBranchID & "," & FromBoxid
sql = sql & "," & FromBillBasedOn & "," & FromVAT & "," & FromVATYou & ",'" & NoteSerial & "','" & NoteSerial1 & "'," & NoteId & ",1,'" & TransactionComment & "',1,'" & SessionCode & "',"

sql = sql & "'" & FromNoteSerial1 & "' , "
sql = sql & "'" & FromNoteSerial & "' , " & FromNoteId & " , " & FromTransaction_ID & " )"

            '   fromTransaction_Serial
 Cn.Execute sql

Text2.Text = sql
      
      
     ' ›«’Ì· «·”‰œ
  
 
 sql = " select * from Transaction_Details   where  Transaction_ID=" & FromTransaction_ID
    Set rsDouble_Entry = New ADODB.Recordset
    '
   rsDouble_Entry.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    Dim j As Double
    For j = 1 To rsDouble_Entry.RecordCount
    Item_ID = IIf(IsNull(rsDouble_Entry("Item_ID").Value), 0, rsDouble_Entry("Item_ID").Value)
     ItemCase = IIf(IsNull(rsDouble_Entry("ItemCase").Value), 0, rsDouble_Entry("ItemCase").Value)
      Quantity = IIf(IsNull(rsDouble_Entry("Quantity").Value), 0, rsDouble_Entry("Quantity").Value)
       Price = IIf(IsNull(rsDouble_Entry("Price").Value), 0, rsDouble_Entry("Price").Value)
        ItemDiscountType = IIf(IsNull(rsDouble_Entry("ItemDiscountType").Value), 0, rsDouble_Entry("ItemDiscountType").Value)
         ItemDiscount = IIf(IsNull(rsDouble_Entry("ItemDiscount").Value), 0, rsDouble_Entry("ItemDiscount").Value)
         ShowQty = IIf(IsNull(rsDouble_Entry("ShowQty").Value), 0, rsDouble_Entry("ShowQty").Value)
         showPrice = IIf(IsNull(rsDouble_Entry("showPrice").Value), 0, rsDouble_Entry("showPrice").Value)
         UnitId = IIf(IsNull(rsDouble_Entry("UnitId").Value), 0, rsDouble_Entry("UnitId").Value)
         ColorID = IIf(IsNull(rsDouble_Entry("ColorID").Value), 0, rsDouble_Entry("ColorID").Value)
         ItemSize = IIf(IsNull(rsDouble_Entry("ItemSize").Value), 0, rsDouble_Entry("ItemSize").Value)
         ClassId = IIf(IsNull(rsDouble_Entry("ClassId").Value), 0, rsDouble_Entry("ClassId").Value)
         
         
 
    sql = " INSERT INTO  [" & ServerDb & "].[dbo].[Transaction_Details]  (    "
sql = sql & "  Transaction_ID,  Item_ID, ItemCase, Quantity, Price, ItemDiscountType, ItemDiscount, ShowQty, showPrice,UnitId , ColorID, ItemSize, ClassId,SessionCode)"
 sql = sql & "   values (" & Transaction_ID & "," & Item_ID & ", " & ItemCase & "," & Quantity & "," & Price & "," & ItemDiscountType & "," & ItemDiscount & "," & ShowQty & "," & showPrice
 sql = sql & "," & UnitId & "," & ColorID & "," & ItemSize & "," & ClassId & "" & ",'" & SessionCode & "')"
 
           Cn.Execute sql
           rsDouble_Entry.MoveNext
    Next j
    
 
 
         
         
'ﬁÌœ «·”‰œ
  

sql = " INSERT INTO [" & ServerDb & "].[dbo].[Notes]([NoteID], [NoteDate], [NoteType], [NoteSerial], [NoteSerial1] ,branch_no,Transaction_ID,SessionCode)"
 sql = sql & " values( " & NoteId & ", " & SQLDate(Transaction_Date, True) & " , " & mNoteType & ", " & NoteSerial & ", " & NoteSerial1 & "," & BranchID & "," & Transaction_ID & ",'" & SessionCode & "')"
 Cn.Execute sql
' MsgBox "ﬁÌœ «·”‰œ"
 DEVID = CStr(new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", ""))
 
 
 'Dim rsDouble_Entry As ADODB.Recordset
  Set rsDouble_Entry = New ADODB.Recordset
     sql = " select * from DOUBLE_ENTREY_VOUCHERS   where   Notes_ID=" & FromNoteId
   rsDouble_Entry.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    Dim w As Double
    For w = 1 To rsDouble_Entry.RecordCount
    Account_Code = IIf(IsNull(rsDouble_Entry("Account_Code").Value), 0, rsDouble_Entry("Account_Code").Value)
    Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
    Credit_Or_Debit = IIf(IsNull(rsDouble_Entry("Credit_Or_Debit").Value), 0, rsDouble_Entry("Credit_Or_Debit").Value)
    Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
    Double_Entry_Vouchers_Description = IIf(IsNull(rsDouble_Entry("Double_Entry_Vouchers_Description").Value), 0, rsDouble_Entry("Double_Entry_Vouchers_Description").Value) & Chr(13) & "  ”‰œ ’—› " & TransactionComment
    'RecordDate = IIf(IsNull(rsDouble_Entry("RecordDate").Value), 0, rsDouble_Entry("RecordDate").Value)
    DEV_ID_Line_No = IIf(IsNull(rsDouble_Entry("DEV_ID_Line_No").Value), 0, rsDouble_Entry("DEV_ID_Line_No").Value)
    branch_id = IIf(IsNull(rsDouble_Entry("branch_id").Value), 0, rsDouble_Entry("branch_id").Value)
    sql = "  INSERT INTO [" & ServerDb & "].[dbo].[DOUBLE_ENTREY_VOUCHERS]([Double_Entry_Vouchers_ID], [DEV_ID_Line_No], [Account_Code], [Value], [Credit_Or_Debit], [Double_Entry_Vouchers_Description], [RecordDate], [Notes_ID] ,branch_id,UserID,Transaction_ID,SessionCode) "
    sql = sql & " values (  " & DEVID & ", " & DEV_ID_Line_No & ", '" & Account_Code & "', " & Value & ", " & Credit_Or_Debit & ", '" & Double_Entry_Vouchers_Description & "',  " & SQLDate(Transaction_Date, True) & ", " & NoteId & " ," & branch_id & ",1 ," & Transaction_ID & ",'" & SessionCode & "')"
  Cn.Execute sql


    rsDouble_Entry.MoveNext
    Next w
   
  
  
 
'*****************************************************************
'**********************************************************
 


  
         
         
        
     Next i
     
     
      End If
 
    Rs3.Close
  'Sql = Sql & "[" & POSDb & "].dbo.Transactions"
  '„‰⁄ «·‰ﬁ· „—… «Œ—Ì
  sql = "update   [" & POSDb & "].dbo.Transactions" & "  set  Copied =1,SessionCode = '" & SessionCode & "' WHERE   Transaction_ID =" & FromTransaction_ID
 POSConnection.Execute sql
 

  sql = "update   [" & POSDb & "].dbo.Transaction_Details" & "  set  Copied =1,SessionCode = '" & SessionCode & "' where Transaction_ID =" & FromTransaction_ID
 POSConnection.Execute sql

 
     StrSQL = "UPDATE  [" & ServerDb & "].dbo. Transactions SET NOTS=" & invoiceTransaction_ID & ",NOTS2= '" & invoiceNoteserial1 & "' ,SessionCode = '" & SessionCode & "' WHERE Transaction_ID=" & Transaction_ID
        Cn.Execute StrSQL
             StrSQL = "UPDATE  [" & ServerDb & "].dbo. Transactions SET NOTS=" & Transaction_ID & ",NOTS2= '" & NoteSerial1 & "',SessionCode = '" & SessionCode & "' WHERE Transaction_ID=" & invoiceTransaction_ID
        Cn.Execute StrSQL
        
        
End Function
Function ConnectionFirst() As Boolean

On Error GoTo ErrTrap
'«” ›”«—
'ServerDb = TxtServerDataBaseName.Text
'wael
'ServerDb = DestinationServer
' POSDb = TxtServerDataBaseName.Text


ServerDb = TxtServerDataBaseName.Text

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


'ServerDb = TxtServerDataBaseName.Text
'wael

POSDb = TxtPOSDB.Text
POSServer = POSlServer.Text


     Set POSConnection = New ADODB.Connection
    With POSConnection
        .CommandTimeout = 60
        .CursorLocation = adUseClient
        .ConnectionTimeout = 15
       If SysSQLServerType = 1 Then
        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
        "Persist Security Info=False;Initial Catalog=" & POSDb & _
        ";Data Source=" & POSServer & ";Port=1433"
        
        ElseIf SysSQLServerType = 2 Then
 
     
                 If SysSQLServerTypeTechnical = "0" Then
                 .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
                    "Persist Security Info=False;Initial Catalog=" & POSDb & _
                    ";Data Source=" & POSServer & ";Port=1433"
                    '";Data Source=" & ServerDb & ";Port=1433"
                    
                  Else
                     .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & POSServer 'SysSQLServerName
                End If
          End If

.Open

End With
ConnectionFirst = True
Dim mPosD  As String
Dim mServerD  As String
mPosD = "[" & POSlServer & "]" & ".Master.dbo."
mServerD = "[" & SysSQLServerName & "]" & ".Master.dbo."

Dim s As String
Dim ss As String
    
    s = " USE MASTER " & vbNewLine
    s = s & " DECLARE @sql NVARCHAR(4000) " & vbNewLine

    s = s & " DECLARE db_cursor CURSOR FOR " & vbNewLine
    s = s & "         select 'sp_dropserver ''' + [srvName] + '''' from sysservers " & vbNewLine

    s = s & "     OPEN db_cursor " & vbNewLine
    s = s & "     FETCH NEXT FROM db_cursor INTO @sql " & vbNewLine

    s = s & "     WHILE @@FETCH_STATUS = 0 " & vbNewLine
    s = s & "     BEGIN " & vbNewLine

    s = s & "            EXEC (@sql) " & vbNewLine

    s = s & "            FETCH NEXT FROM db_cursor INTO @sql " & vbNewLine
    s = s & "     End " & vbNewLine

    s = s & "     Close db_cursor " & vbNewLine
    s = s & "     DEALLOCATE db_cursor " & vbNewLine
    
    ss = "     USE " & ServerDb & vbNewLine
    
    Cn.Execute s & ss
    ss = "USE " & POSDb & vbNewLine
    POSConnection.Execute s & ss
   
Dim rsDummy As New ADODB.Recordset
's = "select * from " & mServerD & "sysservers Where srvName Like '" & SysSQLServerName & "'"
'rsDummy.Open s, Cn, adOpenStatic
'If rsDummy.EOF Then
'    Cn.Execute "EXEC sp_addlinkedserver [" & SysSQLServerName & "]"
'   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
'End If
'rsDummy.Close

's = "select * from sys.servers Where name Like '" & SysSQLServerName & "'"


's = "select * from sys.servers Where name Like '" & POSServer & "'"
s = "select * from sysservers Where srvName Like '" & POSServer & "'"
rsDummy.Open s, Cn, adOpenStatic
If rsDummy.EOF Then
    Cn.Execute "EXEC sp_addlinkedserver [" & POSServer & "]"
   ' Cn.Execute "EXEC sp_addlinkedserver [" & SysSQLServerName & "]"
   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
End If
  
's = "select * from " & mServerD & "sysservers Where srvName Like '" & SysSQLServerName & "'"
s = "select * from sysservers Where srvName Like '" & SysSQLServerName & "'"
rsDummy.Close
rsDummy.Open s, Cn, adOpenStatic
If rsDummy.EOF Then
   
    Cn.Execute "EXEC sp_addlinkedserver [" & SysSQLServerName & "]"
   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
End If


'rsDummy.Close
s = " Use Master "
POSConnection.Execute s

's = "select * from " & mPosD & "sysservers Where srvName Like '" & SysSQLServerName & "'"
s = "select * from sysservers Where srvName Like '" & SysSQLServerName & "'"
rsDummy.Close
rsDummy.Open s, POSConnection, adOpenStatic
If rsDummy.EOF Then
    POSConnection.Execute " EXEC sp_addlinkedserver [" & SysSQLServerName & "]"

   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
End If

rsDummy.Close

s = "select * from sysservers Where srvName Like '" & POSServer & "'"

rsDummy.Open s, POSConnection, adOpenStatic
If rsDummy.EOF Then
    
    POSConnection.Execute " EXEC sp_addlinkedserver [" & POSServer & "]"
   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
End If
rsDummy.Close



s = "select * from " & mPosD & "sysservers Where srvName Like '" & POSServer & "'"
rsDummy.Open s, POSConnection, adOpenStatic
If rsDummy.EOF Then

    POSConnection.Execute " EXEC sp_addlinkedserver [" & POSServer & "]"
   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
End If
rsDummy.Close


s = "Select * from TblOptions "
rsDummy.Open s, Cn, adOpenStatic
If Not rsDummy.EOF Then
    NoOFDigitUserTrans = Val(rsDummy!NoOFDigitUserTrans & "")
    StoreDigit = Val(rsDummy!StoreDigit & "")
    BranchDigit = Val(rsDummy!BranchDigit & "")
    IsSerialByUserTrans = Val(rsDummy!IsSerialByUserTrans & "")
    ExpensesCoding = Val(rsDummy!ExpensesCoding & "")
    InstallmntsvchrCoding = Val(rsDummy!InstallmntsvchrCoding & "")
    ExpensesCoding2 = Val(rsDummy!ExpensesCoding2 & "")
    AllowProjectBill2Serial = Val(rsDummy!AllowProjectBill2Serial & "")
    ExpensesCoding = Val(rsDummy!ExpensesCoding & "")
    AllowProjectBill2Serial = Val(rsDummy!AllowProjectBill2Serial & "")
    NoOFDigitUserVouc = Val(rsDummy!NoOFDigitUserVouc & "")
    JLCodeBasedOnBranch = Val(rsDummy!JLCodeBasedOnBranch & "")
    IsSerialByUserVouch = Val(rsDummy!IsSerialByUserVouch & "")
    JLCodeBasedOnBranch = Val(rsDummy!JLCodeBasedOnBranch & "")
    
End If

rsDummy.Close
'
's = "select * from sys.servers Where name Like '" & POSServer & "'"
'rsDummy.Open s, POSConnection, adOpenStatic
'If rsDummy.EOF Then
'    POSConnection.Execute " EXEC sp_addlinkedserver [" & POSServer & "]"
'   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
'End If



'Do While Not rsDummy.EOF
'
'
'    rsDummy.MoveNext
'Loop




Exit Function
ErrTrap:
Text1 = Cn.ConnectionString
Text2 = POSConnection.ConnectionString
MsgBox "Õÿ√ ›Ì «·« ’«·"
 ConnectionFirst = False


End Function


Private Function DeleteLinkedServer()
 

    
    
End Function

Private Sub cmdUdateFiles_Click()




'   ************************************'check items here first wael*******************
 Dim StrSQL As String
If POSlServer.Text = "" Then
MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
Exit Sub
End If

Command4_Click
lblWait.Visible = True
   
   
  UpdateFiles POSlServer, TxtPOSDB, "TblBranchesData", "branch_id"
  UpdateFiles POSlServer, TxtPOSDB, "TblCustemers", "CusID"
  UpdateFiles POSlServer, TxtPOSDB, "TblUsers", "UserID"
  
 ' UpdateFiles POSlServer, TxtPOSDB, "TblUsersBranches", "userid", True, , True
  
   Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " select * from TblUsersBranches             "
    
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 Dim mSql As String
 POSConnection.Execute "Delete TblUsersBranches"
  mSql = GetSqlQueryInsert(Rs3, mPosD, "TblUsersBranches", "Account_ID", "", "", 0, 0, 0)
  POSConnection.Execute mSql
  
    
    
    Set Rs3 = New ADODB.Recordset
sql = " select * from ACCOUNTS     Where Account_Code Not In (Select Account_Code from " & mPosD & "Accounts  )        "
    
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 mSql = GetSqlQueryInsert(Rs3, mPosD, "ACCOUNTS", "Account_ID", "", "", 0, 0, 0)
  POSConnection.Execute mSql
 
  
  UpdateFiles POSlServer, TxtPOSDB, "TblUserScreen", "ID", True
  UpdateFiles POSlServer, TxtPOSDB, "ScreenJuncUser", "JuncID", True
  
  DoEvents
  
   lblWait.Visible = False
   MsgBox " „ ‰ﬁ· «·»Ì«‰« "

End Sub

Private Sub UpdateFiles(ByVal POSlServer As String, ByVal POSDb As String, ByVal mTableName As String, Optional ByVal mFieldName As String = "Id", Optional ByVal IsResetData As Boolean = False, Optional ByVal IsFromServer As Boolean = True, Optional ByVal isIdent As Boolean = False)
    
   Dim NoOFItem_POS As Double
   Dim NoOFItem_Server As Double
   
   Dim Rs3 As New ADODB.Recordset
   Dim MaxItem_POS As Double
   Dim MaxItem_Server As Double
   Dim ss As String
    If IsResetData Then
    
        
    End If
    
    ss = "     USE " & ServerDb & vbNewLine
    
    Cn.Execute ss
    ss = "USE " & POSDb & vbNewLine
    POSConnection.Execute ss

    sql = " select count (" & mFieldName & " ) As NoOfitems ,max(" & mFieldName & " ) as MaxItemid from " & mTableName
     
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    If Rs3.RecordCount > 0 Then
        NoOFItem_POS = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
        MaxItem_POS = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
   
    End If
    Rs3.Close
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   ' MsgBox "Step 1"
    If Rs3.RecordCount > 0 Then
        NoOFItem_Server = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
        MaxItem_Server = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
    End If
    Rs3.Close
    
   ' MsgBox "Item Server" & NoOFItem_Server
   ' MsgBox "Item Pos" & NoOFItem_POS
    'step 2
   ' Exit Sub
   ' If NoOFItem_Server > NoOFItem_POS Then
       
         'MsgBox "Step 3"
        Dim s As String
        
        'If NoOFItem_Server > NoOFItem_POS Then
            
             
            s = ""
             
           
  
            
         
           ' Text4 = s
           ' Exit Sub
           
            If IsResetData Then
                If IsFromServer Then
                    ss = "Delete " & mTableName & " Where " & mFieldName & " In "
                    ss = "  (SELECT " & mFieldName & " "
                    ss = ss & "                                      FROM   " & mServerD & mTableName & " );"
                    POSConnection.Execute ss
                Else
                    ss = "Delete " & mTableName & " Where " & mFieldName & " In "
                    ss = "  (SELECT " & mFieldName & " "
                    ss = ss & "                                      FROM   " & mPosD & mTableName & " );"
                    Cn.Execute ss
                
                End If
            End If
'            If isIdent Then
'                ss = "SET IDENTITY_INSERT  " & mTableName & "   Off"
'                POSConnection.Execute ss
'            End If
           If IsFromServer Then
            s = GetSql(mServerD, mPosD, mTableName, mFieldName)
            Else
            s = GetSql(mPosD, mServerD, mTableName, mFieldName)
            End If
            Cn.Execute s
           ' MsgBox "Step 4"
            
            
           '  MsgBox " „ ‰ﬁ· »Ì«‰«  «·»Ì«‰« "
           '  cmdUdateFiles.Enabled = False
    
       ' End If
    ' Else
     '   MsgBox "    Ã„Ì⁄ «·»Ì«‰«  „ÕœÀ…"
       ' lblWait.Visible = False
 
   ' End If
    'ss = "SET IDENTITY_INSERT " & mTableName & "   On"
    'POSConnection.Execute ss
End Sub



Private Sub cmdUpdateEmp_Click()





'   ************************************'check items here first wael*******************
 Dim StrSQL As String
If POSlServer.Text = "" Then
MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
Exit Sub
End If

Command4_Click
lblWait.Visible = True
   
   
UpdateFiles POSlServer, TxtPOSDB, "TblEmpData", "HafizaNo", , False
'UpdateFiles POSlServer, TxtPOSDB, "TblEmpDataFingerPrint", "HafizaNo", True, False
  
  lblWait.Visible = False
  MsgBox " „ ‰ﬁ· «·»Ì«‰« "
   
   


End Sub

Private Sub Command1_Click()
On Error GoTo ErrTrap
Dim StrSQL As String
'On Error GoTo ErrTrap
If POSlServer.Text = "" Then
    MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
    Exit Sub
End If


If ConnectionFirst = False Then
Exit Sub
End If
Dim X As Date
Dim mTimeStart As String
'ServerDb = DestinationServer
 'POSDb = TxtServerDataBaseName.Text
    lblWait.Visible = True
  ' Command2_Click
    Dim rs As New ADODB.Recordset
    rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

    JLCodeBasedOnBranch = IIf(rs("JLCodeBasedOnBranch").Value = 0 Or IsNull(rs("JLCodeBasedOnBranch").Value), False, True)
    StoreDigit = IIf(IsNull(rs("StoreDigit").Value), 1, (rs("StoreDigit").Value))
    BranchDigit = IIf(IsNull(rs("BranchDigit").Value), 1, (rs("BranchDigit").Value))
    

    Set POSConnection = New ADODB.Connection
    With POSConnection
        .CommandTimeout = 60
        .CursorLocation = adUseClient
        .ConnectionTimeout = 15
       If SysSQLServerType = 1 Then
            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
            "Persist Security Info=False;Initial Catalog=" & POSDb & _
            ";Data Source=" & POSlServer & ";Port=1433"
        
        ElseIf SysSQLServerType = 2 Then
             If SysSQLServerTypeTechnical = "0" Then
             .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
                "Persist Security Info=False;Initial Catalog=" & POSDb & _
                ";Data Source=" & POSlServer & ";Port=1433"
              Else
                 .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & POSlServer 'SysSQLServerName
            End If
        End If
       '   Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Adnan;Data Source=WAELPC\SQLEXPRESS;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=WAELPC;Use Encryption for Data=False;Tag with column collation when possible=False;

        .Open
    End With


GoTo Transactions

Transactions:
    CopyData "TblTripReg", "ID", "BranchId", "Recorddate", 1101, 82, False
    CopyData "TblEmpDataInOut", "ID", "BranchId", "Recorddate", 0, 0, False
ErrTrap:
lblWait.Visible = False
MsgBox " „ ‰ﬁ· «·»Ì«‰« "
End Sub


Private Sub CopyData(mMainTableName As String, mTransActionIDName As String, mBranchIdName As String, mFieldDate As String, mNoteType As Integer, mSanadNo As Integer, isFiterDate As Boolean)
Dim SessionCode As String
Dim mMaxNo As Long
Dim ss As String
Dim rsDummyMax As New ADODB.Recordset
 Dim BeginTrans As Boolean
Dim isFoundData As Boolean
On Error GoTo ErrTrap
'ss = "Select Max(SessionCode ) MaxNo from TblOffline"
'rsDummyMax.Open ss, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'If rsDummyMax.EOF Then
'    mMaxNo = Val(rsDummyMax!MaxNo & "") + 1
    
'End If

SessionCode = CStr(Now) '& mMaxNo


'////////////////////////////////////////copy Sales Transactions

    Dim Rs3 As ADODB.Recordset
    Dim rsDouble_Entry As ADODB.Recordset
    
    Set Rs3 = New ADODB.Recordset
    
   
    
    Dim sql As String
    Dim mytext As String
    
 
    

    
   ' sql = " select * from Transactions    WHERE  Copied is null And " & GetQuery
    'sql = " select * from " & mMainTableName & "     WHERE  Copied is null and " & GetQuery(mMainTableName)
    'salim here
    sql = " select * from " & mMainTableName & "     WHERE  Copied is null "

    
    
    
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    mTimeStart = Now
    txtStartTime = mTimeStart
    Text3 = sql
    Dim FromTransaction_ID As Double
    Dim FromBranchID As Integer
    Dim FromTransaction_Date As Date
    Dim DateRec As Date
    Dim last_changed As Date
    
    Dim FromNots As String
    Dim FromNots2 As String
    Dim fromTransaction_Serial As String
    Dim FromNoteseial1 As String
    Dim FromTransaction_Type As Integer
    
    Dim BranchID As Integer
    Dim Transaction_ID As Double
    Dim Transaction_Type As Integer
    Dim Transaction_Date As Date
    Dim Transaction_Serial  As String
    Dim Nots As String
    Dim Nots2 As String
    Dim mOldNoteSerial1 As String
    
     
'eee
    'Dim Transaction_Type As Integer
    Dim FromNoteId As Double
   
    
   
 'sales
    
        If Rs3.RecordCount > 0 Then
'            Set cProgress = New ClsProgress
'            BolFrmLoaded = True
'            cProgress.ProgressType = Waiting
'            cProgress.StartProgress

'            Do While Rs3.State = adStateExecuting
'                DoEvents
'            Loop
            
'            If BolFrmLoaded = True Then
'                cProgress.StopProgess
'                Set cProgress = Nothing
'            End If
                Cn.BeginTrans
                BeginTrans = True
                Dim mFieldString As String
                Dim mFieldValue As String
               ' MsgBox Rs3.RecordCount
               Dim mValuex As String
                Dim j As Long
                
                For i = 1 To Rs3.RecordCount
                    sql = " INSERT INTO  [" & ServerDb & "].[dbo]." & mMainTableName & "   ("
                     mFieldString = ""
                     mFieldValue = ""
                     isFoundData = True
                     
                    For j = 0 To Rs3.Fields.Count - 1
                        If j = Rs3.Fields.Count - 1 Then
                            mFieldString = mFieldString & UCase(Rs3.Fields.Item(j).Name)
                        Else
                            mFieldString = mFieldString & UCase(Rs3.Fields.Item(j).Name) & ","
                        End If
                    
                    Next j
                    j = 0
                    For j = 0 To Rs3.Fields.Count - 1
                         FromBranchID = IIf(IsNull(Rs3(mBranchIdName).Value), 0, Rs3(mBranchIdName).Value)
                         FromTransaction_Date = IIf(IsNull(Rs3(mFieldDate).Value), Date, Rs3(mFieldDate).Value)
                        If UCase(Rs3.Fields.Item(j).Name) = "ID" Then
                            If j = Rs3.Fields.Count - 1 Then
                                mFieldValue = mFieldValue & CStr(new_id(mMainTableName, mTransActionIDName, "", True))
                            Else
                                mFieldValue = mFieldValue & CStr(new_id(mMainTableName, mTransActionIDName, "", True)) & ","
                            End If
                            
                        ElseIf UCase(Rs3.Fields.Item(j).Name) = "NOTESERIAL1" Then
                            If j = Rs3.Fields.Count - 1 Then
                                mFieldValue = mFieldValue & Voucher_coding(FromBranchID, FromTransaction_Date, mSanadNo, mNoteType, , , , , , , mMainTableName)
                            Else
                                mFieldValue = mFieldValue & Voucher_coding(FromBranchID, FromTransaction_Date, mSanadNo, mNoteType, , , , , , , mMainTableName) & ","
                            End If
                            mOldNoteSerial1 = Rs3.Fields.Item(j).Value & ""
                        ElseIf UCase(Rs3.Fields.Item(j).Name) = "OLDNOTESERIAL1" Then
                         If j = Rs3.Fields.Count - 1 Then
                                mFieldValue = mFieldValue & mOldNoteSerial1
                            Else
                                mFieldValue = mFieldValue & mOldNoteSerial1 & ","
                            End If
                            
                        ElseIf UCase(Rs3.Fields.Item(j).Name) = "OLDID" Then
                            If j = Rs3.Fields.Count - 1 Then
                                mFieldValue = mFieldValue & Rs3.Fields.Item(mTransActionIDName).Value & ""
                            Else
                                mFieldValue = mFieldValue & Rs3.Fields.Item(mTransActionIDName).Value & "" & ","
                            End If
                        Else
                            If Rs3.Fields.Item(j).Type = adInteger Or Rs3.Fields.Item(j).Type = adCurrency Or Rs3.Fields.Item(j).Type = adBoolean Or Rs3.Fields.Item(j).Type = adSmallInt Or Rs3.Fields.Item(j).Type = adBigInt Or Rs3.Fields.Item(j).Type = adTinyInt Or Rs3.Fields.Item(j).Type = adUnsignedTinyInt Or Rs3.Fields.Item(j).Type = adNumeric Or Rs3.Fields.Item(j).Type = adDouble Or Rs3.Fields.Item(j).Type = adDecimal Then
                                mValuex = Val(Rs3.Fields.Item(j).Value & "")
                            ElseIf Rs3.Fields.Item(j).Type = adDBTimeStamp Or Rs3.Fields.Item(j).Type = adDBTime Or Rs3.Fields.Item(j).Type = adDBDate Then
                                If Not IsDate(Rs3.Fields.Item(j).Value & "") Then
                                    mValuex = "Null"
                                Else
                                    mValuex = SQLDate(Rs3.Fields.Item(j).Value & "", True)
                                End If
                            Else
                                mValuex = "N'" & Trim(Rs3.Fields.Item(j).Value & "") & "'"
                            End If
                            
                            If j = Rs3.Fields.Count - 1 Then

                                mFieldValue = mFieldValue & mValuex
                            Else
                                mFieldValue = mFieldValue & mValuex & ","
                            End If
                        End If
                    Next j
                    
                   sql = sql & mFieldString & " ) values " & "(" & mFieldValue & ")"
                   Cn.Execute sql
                    
                    

                    DoEvents
                    Rs3.MoveNext
        Next i
        Rs3.Close
      'Sql = Sql & "[" & POSDb & "].dbo.Transactions"
      '„‰⁄ «·‰ﬁ· „—… «Œ—Ì
      
    
            sql = "update   [" & POSDb & "].dbo." & mMainTableName & "  set  Copied =1,SessionCode = '" & SessionCode & "' "
      sql = sql & "  Where Copied Is Null  "
      
      If isFiterDate Then
        sql = sql & " And " & GetQuery(mMainTableName) & "  "
        sql = sql & " and dbo." & mMainTableName & "." & mFieldDate & " ='" & SQLDate(dbRecordDate.Value, False) & "'"
    End If
      
     POSConnection.Execute sql
    ' MsgBox "8"
    
'      sql = "update   [" & POSDb & "].dbo.Transaction_Details" & "  set  Copied =1,SessionCode = '" & SessionCode & "' WHERE  Copied is null    "
'     POSConnection.Execute sql
     

  


 End If
If isFoundData Then
     Dim rsOffline As New ADODB.Recordset
    Dim mEndTime22 As String
    mEndTime22 = Now
    s = "Select * from TblOffline2 where 1 = -1"
    rsOffline.Open s, Cn, adOpenKeyset, adLockOptimistic
    'MsgBox s
    rsOffline.AddNew
    'MsgBox s & "Save"
    'rsOffline!Id = mMaxId
    rsOffline!Recorddate = Date
 '   rsOffline!EndTime = mEndTime22
 '   rsOffline!StartTime = mTimeStart
    rsOffline!SessionCode = SessionCode
    rsOffline!POSname = POSlServer
    
'    rsOffline!CountSales = CountSales
'    rsOffline!CountSalesReturn = CountSalesReturn
'    rsOffline!CountPurchase = CountPurchase
'    rsOffline!CountPurchaseReturn = CountPurchaseReturn
'    rsOffline!CountRec = CountRec
    rsOffline.Update
    
    Cn.CommitTrans
    BeginTrans = False
End If



 
    





'Dim mMaxId As Long
's = "Select Max(Id) as MaxID  from TblOffline"
'rsOffline.Open s, Cn, adOpenKeyset, adLockOptimistic
'mMaxId = 1
'If Not rsOffline.EOF Then
'    mMaxId = Val(rsOffline!MaxID & "") + 1
'
'End If
'rsOffline.Close

lblWait.Visible = False

txtEndTime = mEndTime22
txtCountSalesReturn = CountSalesReturn
txtCountSales = CountSales
Exit Sub



ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'Resume Next
'MsgBox "Done"
'////////////////////////////////////////copy Sales Transactions

End Sub


Private Function GetSqlQueryInsert(ByVal Rs3 As ADODB.Recordset, ByVal mServer As String, mMainTableName As String, mTransActionIDName As String, mBranchIdName As String, mFieldDate As String, mNoteType As Integer, mSanadNo As Integer, ByVal isIdent As Long) As String
   Dim FromTransaction_ID As Double
    Dim FromBranchID As Integer
    Dim FromTransaction_Date As Date
    Dim DateRec As Date
    Dim last_changed As Date
    
    Dim FromNots As String
    Dim FromNots2 As String
    Dim fromTransaction_Serial As String
    Dim FromNoteseial1 As String
    Dim FromTransaction_Type As Integer
    
    Dim BranchID As Integer
    Dim Transaction_ID As Double
    Dim Transaction_Type As Integer
    Dim Transaction_Date As Date
    Dim Transaction_Serial  As String
    Dim Nots As String
    Dim Nots2 As String
    Dim mOldNoteSerial1 As String
    Dim SessionCode As String
Dim mMaxNo As Long
Dim mNoteID As Long
Dim ss As String
Dim rsDummyMax As New ADODB.Recordset
 Dim BeginTrans As Boolean
Dim isFoundData As Boolean
       Dim mFieldString As String
                Dim mFieldValue As String
               ' MsgBox Rs3.RecordCount
               Dim mValuex As String
     Dim s As String
     Dim rsDummyCheckID As ADODB.Recordset
     
'eee
    'Dim Transaction_Type As Integer
    Dim FromNoteId As Double
                For i = 1 To Rs3.RecordCount
                    sql = " INSERT INTO  " & mServer & "" & mMainTableName & "   (" & vbNewLine
                     mFieldString = ""
                     mFieldValue = ""
                     isFoundData = True
                    For j = 0 To Rs3.Fields.Count - 1
                        If isIdent = 0 Then
                            If UCase(Rs3.Fields.Item(j).Name) <> "ID" And UCase(Rs3.Fields.Item(j).Name) <> "ACCOUNT_ID" Then
                                If j = Rs3.Fields.Count - 1 Then
                                
                                    mFieldString = mFieldString & UCase(Rs3.Fields.Item(j).Name) & vbNewLine
                                Else
                                    mFieldString = mFieldString & UCase(Rs3.Fields.Item(j).Name) & "," & vbNewLine
                                End If
                            End If
                        Else
                            If j = Rs3.Fields.Count - 1 Then
                            
                                mFieldString = mFieldString & UCase(Rs3.Fields.Item(j).Name) & vbNewLine
                            Else
                                mFieldString = mFieldString & UCase(Rs3.Fields.Item(j).Name) & "," & vbNewLine
                            End If
                        End If
                    Next j
                         

                    j = 0
                    For j = 0 To Rs3.Fields.Count - 1
                        If j = 47 Then
                            j = j
                        End If
                        If mBranchIdName <> "" Then
                            FromBranchID = IIf(IsNull(Rs3(mBranchIdName).Value), 0, Rs3(mBranchIdName).Value)
                        End If
                        If mFieldDate <> "" Then
                            FromTransaction_Date = IIf(IsNull(Rs3(mFieldDate).Value), Date, Rs3(mFieldDate).Value)
                        End If
                        If UCase(Rs3.Fields.Item(j).Name) = "ID" Or UCase(Rs3.Fields.Item(j).Name) = UCase(mTransActionIDName) Then
                            If isIdent = 1 Then
                                If j = Rs3.Fields.Count - 1 Then
                                    mFieldValue = mFieldValue & CStr(new_id(mMainTableName, mTransActionIDName, "", True, , POSConnection)) & vbNewLine
                                Else
                                    mFieldValue = mFieldValue & CStr(new_id(mMainTableName, mTransActionIDName, "", True, , POSConnection)) & "," & vbNewLine
                                End If
                            ElseIf isIdent = 2 Then
                                Set rsDummyCheckID = New ADODB.Recordset
                                s = "Select " & mTransActionIDName & " NoteID from " & mMainTableName & " Where " & mTransActionIDName & "= " & Trim(Rs3.Fields.Item(j).Value & "")
                                rsDummyCheckID.Open s, POSConnection, adOpenKeyset, adLockReadOnly
                                If rsDummyCheckID.EOF Then
                                    mNoteID = Val(Trim(Rs3.Fields.Item(j).Value & ""))
                                Else
                                    mNoteID = (new_id(mMainTableName, mTransActionIDName, "", True, , POSConnection))
                                End If
                                If j = Rs3.Fields.Count - 1 Then
                                    mFieldValue = mFieldValue & Trim(mNoteID) & vbNewLine
                                 Else
                                    mFieldValue = mFieldValue & Trim(mNoteID) & "," & vbNewLine
                                 End If
                            
                            End If
                            
                        ElseIf UCase(Rs3.Fields.Item(j).Name) = "NOTESERIAL1" Then
                            If mSanadNo = 0 And mNoteType = 0 Then
                                If j = Rs3.Fields.Count - 1 Then
                                    If mMainTableName = "Transactions" Then
                                        mFieldValue = mFieldValue & "'" & Trim(Rs3.Fields.Item(j).Value & "") & "'" & vbNewLine
                                    Else
                                        mFieldValue = mFieldValue & Trim(Rs3.Fields.Item(j).Value & "") & vbNewLine
                                    End If
                                 Else
                                    If mMainTableName = "Transactions" Then
                                        mFieldValue = mFieldValue & "'" & Trim(Rs3.Fields.Item(j).Value & "") & "'" & "," & vbNewLine
                                    Else
                                        mFieldValue = mFieldValue & Trim(Rs3.Fields.Item(j).Value & "") & "," & vbNewLine
                                    End If
                                 End If
                            Else
                                If j = Rs3.Fields.Count - 1 Then
                                    mFieldValue = mFieldValue & Voucher_coding(FromBranchID, FromTransaction_Date, mSanadNo, mNoteType, , , , , , , mMainTableName) & vbNewLine
                                Else
                                    mFieldValue = mFieldValue & Voucher_coding(FromBranchID, FromTransaction_Date, mSanadNo, mNoteType, , , , , , , mMainTableName) & "," & vbNewLine
                                End If
                            End If
                           
                            mOldNoteSerial1 = Rs3.Fields.Item(j).Value & ""
                        ElseIf UCase(Rs3.Fields.Item(j).Name) = "OLDNOTESERIAL1" Then
                         If j = Rs3.Fields.Count - 1 Then
                                mFieldValue = mFieldValue & mOldNoteSerial1 & vbNewLine
                            Else
                                mFieldValue = mFieldValue & mOldNoteSerial1 & "," & vbNewLine
                            End If
                            
                        ElseIf UCase(Rs3.Fields.Item(j).Name) = "OLDID" Then
                            If j = Rs3.Fields.Count - 1 Then
                                mFieldValue = mFieldValue & Rs3.Fields.Item(mTransActionIDName).Value & "" & vbNewLine
                            Else
                                mFieldValue = mFieldValue & Rs3.Fields.Item(mTransActionIDName).Value & "" & "," & vbNewLine
                            End If
                        Else
                            If Rs3.Fields.Item(j).Type = adInteger Or Rs3.Fields.Item(j).Type = adCurrency Or Rs3.Fields.Item(j).Type = adBoolean Or Rs3.Fields.Item(j).Type = adSmallInt Or Rs3.Fields.Item(j).Type = adBigInt Or Rs3.Fields.Item(j).Type = adTinyInt Or Rs3.Fields.Item(j).Type = adUnsignedTinyInt Or Rs3.Fields.Item(j).Type = adNumeric Or Rs3.Fields.Item(j).Type = adDouble Or Rs3.Fields.Item(j).Type = adDecimal Then
                                mValuex = Val(Rs3.Fields.Item(j).Value & "")
                            ElseIf Rs3.Fields.Item(j).Type = adDBTimeStamp Or Rs3.Fields.Item(j).Type = adDBTime Or Rs3.Fields.Item(j).Type = adDBDate Then
                                If Not IsDate(Rs3.Fields.Item(j).Value & "") Then
                                    mValuex = "Null"
                                Else
                                    mValuex = "'" & Rs3.Fields.Item(j).Value & "'"
                                End If
                            Else
                                mValuex = "N'" & Trim(Rs3.Fields.Item(j).Value & "") & "'"
                            End If
                            
                            If j = Rs3.Fields.Count - 1 Then

                                mFieldValue = mFieldValue & mValuex & vbNewLine
                            Else
                                mFieldValue = mFieldValue & mValuex & "," & vbNewLine
                            End If
                        End If
                        If mValuex = "" Then
                            mValuex = mValuex
                        End If
                    Next j
                    
                   sql = sql & mFieldString & " ) values " & "(" & mFieldValue & ")" & vbNewLine
                 '  Cn.Execute sql
                    GetSqlQueryInsert = GetSqlQueryInsert & " ; " & sql
                    Rs3.MoveNext

                    DoEvents
        Next i
        Rs3.Close
End Function


Private Sub Command2_Click()

'   ************************************'check items here first wael*******************
 Dim StrSQL As String
If POSlServer.Text = "" Then
MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
Exit Sub
End If

Command4_Click
lblWait.Visible = True
   Dim NoOFItem_POS As Double
   Dim NoOFItem_Server As Double
   
   Dim Rs3 As New ADODB.Recordset
   Dim MaxItem_POS As Double
   Dim MaxItem_Server As Double
   'step one check item
       
    ss = "     USE " & ServerDb & vbNewLine
    
    Cn.Execute ss
    ss = "USE " & POSDb & vbNewLine
    POSConnection.Execute ss
    
    sql = " select count (CusID ) As NoOfitems ,max(CusID) as MaxItemid from TblCustemers  "
     
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    If Rs3.RecordCount > 0 Then
        NoOFItem_POS = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
        MaxItem_POS = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
   
    End If
    Rs3.Close
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   ' MsgBox "Step 1"
    If Rs3.RecordCount > 0 Then
        NoOFItem_Server = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
        MaxItem_Server = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
    End If
    Rs3.Close
    
   ' MsgBox "Item Server" & NoOFItem_Server
   ' MsgBox "Item Pos" & NoOFItem_POS
    'step 2
   ' Exit Sub
    If NoOFItem_Server > NoOFItem_POS Then
             'checkGroup
        Dim NoOfGroups_pos As Double
        Dim NoOfGroups_server As Double
             
        Dim MaxGroupid_pos As Double
        Dim MaxGroupidserver As Double
        
                       
       
         'MsgBox "Step 3"
        Dim s As String
        
        If NoOFItem_Server > NoOFItem_POS Then
            
   
            BolFrmLoaded = True
    
              
         Do While Rs3.State = adStateExecuting
                DoEvents
            Loop
     
              
            s = ""
             
            Dim mPosD As String
            Dim mServerD As String
             mPosD = "[" & POSlServer & "]" & "." & POSDb & ".dbo."
             mServerD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
             mServerD = ServerDb & ".dbo."
            
       
            
            s = " INSERT INTO " & mPosD & "TblCustemers"
            s = s & " SELECT *"
            s = s & " FROM   " & mServerD & "TblCustemers T2"
            s = s & " WHERE  T2.CusID NOT IN (SELECT CusID"
            s = s & "                                      FROM   " & mPosD & "TblCustemers);"
            
            
            Cn.Execute s

            'MsgBox "Step 7"

             
            'Copy  remains Groups
            'Copy  remains Items
            'Copy itemsunits
            
            
             MsgBox " „ ‰ﬁ· »Ì«‰«  «·⁄„·«¡"
             Command2.Enabled = False
    
        End If
     Else
   '   MsgBox "    „·›   «·⁄„·«¡ „ÕœÀ"
      lblWait.Visible = False
 
End If
    
    
   '************************************'check items here first*******************

End Sub

Private Sub Command4_Click()
    If ConnectionFirst = False Then
        Exit Sub
    End If
    Dim StrSQL As String
If POSlServer.Text = "" Then
MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
Exit Sub
End If



   Dim NoOFItem_POS As Double
   Dim NoOFItem_Server As Double
   
   Dim Rs3 As New ADODB.Recordset
   Dim MaxItem_POS As Double
   Dim MaxItem_Server As Double
   'step one check item
       
    ss = "     USE " & ServerDb & vbNewLine
    
    Cn.Execute ss
    ss = "USE " & POSDb & vbNewLine
    POSConnection.Execute ss
    Text2 = ss & " " & POSConnection.ConnectionString
    sql = " select count (CusID ) As NoOfitems ,max(CusID) as MaxItemid from TblCustemers  "
     
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    If Rs3.RecordCount > 0 Then
        NoOFItem_POS = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
        MaxItem_POS = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
   
    End If
'   MsgBox "⁄œœ ⁄„·«¡  «·‰ﬁÿ…" & NoOFItem_POS
    Rs3.Close
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Rs3.RecordCount > 0 Then
'        NoOFItem_Server = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
'        MaxItem_Server = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
'    End If
'    Rs3.Close
'
'
    'step 2
    
End Sub

Public Sub CreatLog_File_for_error(str As String)
    Dim StrLogFileName As String
    Dim IntFreeFile As Integer
    Dim ss As String

    StrLogFileName = App.Path & "\employee_account_error.txt"

    If Dir(StrLogFileName) <> "" Then
        Kill StrLogFileName
    End If
    ss = str
   
    IntFreeFile = FreeFile
    
    Open StrLogFileName For Output As #IntFreeFile
    Print #IntFreeFile, ss
    Close #IntFreeFile
End Sub



Private Sub Command7_Click()
'   ************************************'check items here first wael*******************
 Dim StrSQL As String
If POSlServer.Text = "" Then
MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
Exit Sub
End If

'Command4_Click
lblWait.Visible = True
   Dim mWhere As String
   Dim mWhere2 As String
   Dim mWhere3 As String
   Dim mWhere4 As String
   If txtFromDate.Value = txtToDate.Value Then
        mWhere = " NoteDate > '" & SQLDate(txtFromDate.Value, False) & "'"
        mWhere = " NoteDate > '" & SQLDate(txtFromDate.Value, False) & "'"
        mWhere2 = " Transaction_Date > '" & SQLDate(txtFromDate.Value, False) & "'"
        mPosD = "[" & POSlServer & "]" & "." & POSDb & ".dbo."
        mWhere3 = " Notes_ID IN (SELECT NoteId FROM " & mPosD & ".Notes Where " & mWhere & ")"
        mWhere4 = " Transaction_ID IN (SELECT Transaction_ID FROM " & mPosD & ".Transactions Where " & mWhere2 & ")"
   End If
   
   
        
'    Set Rs3 = New ADODB.Recordset
'sql = " select * from notes_all   Where " & mWhere
'
'Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
' mSql = GetSqlQueryInsert(Rs3, mPosD, "notes_all", "NoteID", "", "NoteDate", 0, 0, 2)
'
'  Text2 = mSql
'  POSConnection.Execute mSql
'
'
'     Set Rs3 = New ADODB.Recordset
'sql = " select * from notes Where " & mWhere
'
'Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
' mSql = GetSqlQueryInsert(Rs3, mPosD, "notes", "NoteID", "", "NoteDate", 0, 0, 2)
'
'  Text2 = mSql
'  POSConnection.Execute mSql
'
   
   
       Set Rs3 = New ADODB.Recordset
StrSQL = " select * from Transactions   Where " & mWhere2
StrSQL = StrSQL & " Order By Transaction_ID"
    
Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 StrSQL = GetSqlQueryInsert(Rs3, mPosD, "Transactions", "Transaction_ID", "", "Transaction_Date", 0, 0, 0)
  
  Text2 = StrSQL
  'POSConnection.Execute StrSQL
  CreatLog_File_for_error StrSQL
  
    
'Set Rs3 = New ADODB.Recordset
'StrSQL = " select * from Transaction_Details   Where " & mWhere4
'
'Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
' StrSQL = GetSqlQueryInsert(Rs3, mPosD, "Transaction_Details", "ID", "", "Transaction_Date", 0, 0, 2)
'
'  Text2 = StrSQL
'  CreatLog_File_for_error StrSQL
' ' POSConnection.Execute mSql
'
'
'
'
'Set Rs3 = New ADODB.Recordset
'StrSQL = " select * from DOUBLE_ENTREY_VOUCHERS Where " & mWhere3
'
'Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
' StrSQL = GetSqlQueryInsert(Rs3, mPosD, "DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", "RecordDate", 0, 0, 2)
'
'  Text2 = StrSQL
'  POSConnection.Execute StrSQL
'
'
  
'  UpdateFiles POSlServer, POSDb, "notes_all", "NoteID", mWhere
'  UpdateFiles POSlServer, POSDb, "notes", "NoteID", mWhere2
'  UpdateFiles POSlServer, POSDb, "DOUBLE_ENTREY_VOUCHERS", "", mWhere3
'  UpdateFiles POSlServer, POSDb, "Transactions", "Transaction_ID", mWhere2
'  UpdateFiles POSlServer, POSDb, "Transaction_Details", "Transaction_ID", mWhere4
  
  
   
End Sub

Private Sub Form_Load()
'21 11 2017
' „  ‰›Ì– «·„»Ì⁄«  ﬂ«„·… „⁄ ﬁÌœÂ«  „⁄ ”‰œ «·’—› „⁄ ﬁÌœ…
'
'
'
 On Error Resume Next
txtDbPath = GetSetting("ConvertToAccess", "Setting", "DbPath", "DatabasePath")
TxtTableName = GetSetting("ConvertToAccess", "Setting", "TableName", "TableName")
TxtUSERID = GetSetting("ConvertToAccess", "Setting", "USERID", "USERID")
TxtCHECKTIME = GetSetting("ConvertToAccess", "Setting", "CHECKTIME", "CHECKTIME")
'DcTime.Value = GetSetting("ConvertToAccess", "Setting", "UpdateHours", "00")
dbRecordDate = Date
TxtServerDataBaseName = SysSQLServerDataBaseName
DestinationServer = SysSQLServerName
'BranchDigit = 1
Dim Msg As String
If Dir(App.Path & "\pos.txt", vbNormal) = "" Then
            Msg = "„·›  ”ÃÌ· «·ﬁÊ«⁄œ €Ì— „ÊÃÊœ ...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            
           End
           
        End If
        
    Open App.Path & "\pos.txt" For Input As #1
    POSname.Clear

    Do Until EOF(1)
        Line Input #1, a
        'subsequent lines
 
        If a <> "" Then
            VarSet = Split(a, "*", , vbTextCompare)

            If VarSet(0) <> Empty Or VarSet(0) <> "" Then
            
             POSname.AddItem (VarSet(0))
                ServersName.AddItem (VarSet(1))
            DbName.AddItem (VarSet(2))
                            
            End If
        End If
    
    Loop

    Close #1


 
End Sub

 
Private Sub grd_Click()
If grd.Row <> 0 Then
    dbRecordDate = grd.TextMatrix(grd.Row, grd.ColIndex("Transaction_Date"))
End If
End Sub

Private Sub POSname_Change()
  If ConnectionFirst = False Then
        Exit Sub
    End If
    Dim StrSQL As String
    If POSlServer.Text = "" Then
        MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
    Exit Sub
End If



   Dim NoOFItem_POS As Double
   Dim NoOFItem_Server As Double
   
   Dim Rs3 As New ADODB.Recordset
   Dim MaxItem_POS As Double
   Dim MaxItem_Server As Double
   'step one check item
       
    ss = "     USE " & ServerDb & vbNewLine
    
    Cn.Execute ss
    ss = "USE " & POSDb & vbNewLine
    POSConnection.Execute ss
    
    sql = " "
    
    sql = sql & "     SELECT SUM(CountSales) CountSales ,SUM(CountReturn) CountReturn,Transaction_Date FROM ("
    sql = sql & "         SELECT COUNT(t.Transaction_ID)     CountTotal,"
    sql = sql & "                CountSales       = ("
    sql = sql & "                    Case t.Transaction_Type"
    sql = sql & "                         WHEN 21 THEN COUNT(t.Transaction_ID)"
    sql = sql & "                         ELSE 0"
    sql = sql & "                    End"
    sql = sql & "                ),"
    sql = sql & "                CountReturn     = ("
    sql = sql & "                    Case t.Transaction_Type"
    sql = sql & "                         WHEN 9 THEN COUNT(t.Transaction_ID)"
    sql = sql & "                         ELSE 0"
    sql = sql & "                    End"
    sql = sql & "                ),"
    sql = sql & "                t.Transaction_Date,"
    sql = sql & "                Transaction_Type"
    sql = sql & "         FROM   Transactions             AS t"
    sql = sql & "         Where IsNull(t.Copied, 0) = 0"
    sql = sql & "                AND (t.Transaction_Type = 9 OR t.Transaction_Type = 21)"
    sql = sql & "         Group By"
    sql = sql & "                Transaction_Date,"
    sql = sql & "                Transaction_Type"
        
    sql = sql & "         ) T"
    sql = sql & "         Group By"
    sql = sql & "                Transaction_Date"
    sql = sql & "         Order By"
    sql = sql & "                Transaction_Date"

     
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    grd.Rows = 1
    grd.Rows = 2
    Do While Not Rs3.EOF
        grd.TextMatrix(grd.Rows - 1, grd.ColIndex("CountSales")) = Rs3!CountSales & ""
        grd.TextMatrix(grd.Rows - 1, grd.ColIndex("CountReturn")) = Rs3!CountReturn & ""
        grd.TextMatrix(grd.Rows - 1, grd.ColIndex("Transaction_Date")) = Rs3!Transaction_Date & ""
        Rs3.MoveNext
        grd.Rows = grd.Rows + 1
    Loop
    Rs3.Close
               mPosD = "[" & POSlServer & "]" & "." & POSDb & ".dbo."
             mServerD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
             mServerD = ServerDb & ".dbo."
End Sub

Private Sub POSname_Click()
On Error Resume Next
    DbName.ListIndex = POSname.ListIndex
    ServersName.ListIndex = POSname.ListIndex
     
   POSlServer.Text = ServersName.Text
    TxtPOSDB.Text = DbName.Text
    
    POSname_Change
    
    
    
End Sub
Private Function GetQuery(ByVal mTableName As String) As String
    Dim s As String
'    s = "(1 = 1)  "
'    If chkSales.Value = vbChecked Then
'        s = s & " Or Transaction_Type = 21 "
'    End If
'
'    If chkSalesReturn.Value = vbChecked Then
'        s = s & " Or Transaction_Type = 9 "
'    End If
'
'    If chkPurchaseReturn.Value = vbChecked Then
'        s = s & " Or Transaction_Type = 5 "
'    End If
'
'    If chkPurchase.Value = vbChecked Then
'        s = s & " Or Transaction_Type = 22"
'    End If
'
'
'    If chkIn.Value = vbChecked Then
'        s = s & " Or Transaction_Type = 20"
'    End If
'
'    If chkOut.Value = vbChecked Then
'        s = s & " Or Transaction_Type = 19"
'    End If


    Dim tempString As String
    Dim i As Integer

    
    
    'GetTransIds = tempString
    
    s = s & "  ( " & mTableName & ".RecordDate ='" & SQLDate(dbRecordDate.Value, False) & "')"
    
GetQuery = s
End Function

