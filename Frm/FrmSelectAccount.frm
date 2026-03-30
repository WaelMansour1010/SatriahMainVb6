VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmSelectAccount 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
   Icon            =   "FrmSelectAccount.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5955
   Begin VB.TextBox TxtItem 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   3165
   End
   Begin VB.TextBox TxtCodeItem 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Width           =   2205
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "„Ê«›ﬁ"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid1 
      Height          =   3795
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5865
      _cx             =   10345
      _cy             =   6694
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSelectAccount.frx":000C
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   765
      Index           =   5
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   5955
      _cx             =   10504
      _cy             =   1349
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "FrmSelectAccount.frx":0137
      Caption         =   " ÕœÌœ «·Õ”«»« "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   6
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   0
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.CheckBox Check17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   195
         Left            =   4560
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·Õ”«»"
      Height          =   345
      Index           =   0
      Left            =   1440
      TabIndex        =   7
      Top             =   840
      Width           =   930
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ"
      Height          =   345
      Index           =   32
      Left            =   4200
      TabIndex        =   6
      Top             =   840
      Width           =   930
   End
End
Attribute VB_Name = "FrmSelectAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public IndexType As Integer
Private Sub ChangeLang()
  '  Dim XPic As IPictureDisp
  '  Set XPic = Me.btnFirst.ButtonImage
  '  Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
  '  Set Me.btnLast.ButtonImage = XPic
  '  Set XPic = Me.btnPrevious.ButtonImage
  '  Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
  '  Set Me.btnNext.ButtonImage = XPic

    Me.Caption = "Select Account "
    Ele(5).Caption = Me.Caption
    Check17.Caption = "Select All "
    Check17.RightToLeft = False
    lbl(32).Caption = "Account Code"
     lbl(0).Caption = "Account Name"
CMDOK.Caption = "OK"

    With Me.Grid1
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("Account_Serial")) = "Account Code"
        .TextMatrix(0, .ColIndex("Account_Name")) = "Account Name"
        
    End With

End Sub
Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.Grid1
 
            For i = 1 To .Rows - 1
        
           .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
            Next i

        End With

    Else

        With Me.Grid1

            For i = 1 To .Rows - 1
        
              .Cell(flexcpChecked, i, .ColIndex("Select")) = flexUnchecked
            Next i

        End With

    End If


End Sub

Sub save()

  Dim Roww1 As Long
  Dim rs2 As ADODB.Recordset
  Dim typ As Integer
  Dim strsql As String
  Dim fulcode As String
  Dim name As String
  Dim empid As Integer
Set rs2 = New ADODB.Recordset

Roww1 = FrmAccEditJournal3.Roww

  With FrmAccEditJournal3.Fg_Journal
 typ = val(.TextMatrix(Roww1, .ColIndex("ToTrans")))
  empid = val(.TextMatrix(Roww1, .ColIndex("NEmpid")))
 fulcode = .TextMatrix(Roww1, .ColIndex("Fullcode"))
 name = .TextMatrix(Roww1, .ColIndex("NEmpName"))

        For i = 1 To Grid1.Rows - 1

            If (Grid1.TextMatrix(i, Grid1.ColIndex("Account_Serial")) <> "") And (Grid1.Cell(flexcpChecked, i, Grid1.ColIndex("Select")) = flexChecked) Then
         
                   
                .TextMatrix(Roww1, .ColIndex("BranchId")) = val(FrmAccEditJournal3.dcBranch.BoundText)
                .TextMatrix(Roww1, .ColIndex("BranchName")) = FrmAccEditJournal3.dcBranch.text
               
                .TextMatrix(Roww1, .ColIndex("ToTrans")) = typ
                .TextMatrix(Roww1, .ColIndex("NEmpName")) = name
                 .TextMatrix(Roww1, .ColIndex("Fullcode")) = fulcode
                .TextMatrix(Roww1, .ColIndex("NEmpid")) = empid
                .TextMatrix(Roww1, .ColIndex("userid")) = user_id
                .TextMatrix(Roww1, Col) = Trim(Grid1.TextMatrix(i, Grid1.ColIndex("Account_Serial")))

              '  If .TextMatrix(Row, Col) = "" Then
              '      Exit Sub
              '  End If

                strsql = "SELECT ACCOUNTS.cost_center_id,ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(Grid1.TextMatrix(i, Grid1.ColIndex("Account_Serial"))) & "'"
                Set rs = Nothing
                Set rs = New ADODB.Recordset
                rs.Open strsql, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    If BolEditOnMainAccounts = False Then
                        If FrmAccEditJournal3.LastAccount(rs("Account_Code").value) = False Then
                            .TextMatrix(Roww1, Col) = ""
                            .TextMatrix(Roww1, .ColIndex("AccountCode")) = ""
                            .TextMatrix(Roww1, .ColIndex("AccountName")) = ""
                            Exit Sub
                        End If
                    End If
.TextMatrix(Roww1, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
                    .TextMatrix(Roww1, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                    .TextMatrix(Roww1, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    
                    .TextMatrix(Roww1, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), 0, rs("cost_center").value)
                    
                    .TextMatrix(Roww1, .ColIndex("cost_center_id")) = IIf((rs("cost_center").value) = False, "", rs("cost_center_id").value)
                    
                   ' Dim rs2 As ADODB.Recordset
                    Dim My_SQL As String

                    If IsNull(rs("currenct_code").value) Then

                        .TextMatrix(Roww1, .ColIndex("currenct_code")) = ""
                    
                        .TextMatrix(Roww1, .ColIndex("rate")) = "1"
                    
                        GoTo xx
                    End If

                    My_SQL = "  select * from currency WHERE id=" & val(rs("currenct_code").value)

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
                    .TextMatrix(Roww1, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Roww1, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), 1, rs2.Fields("rate").value)
xx:
                Else
                 '   GetMsgs 130, vbExclamation
                
                  If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "—ﬁ„ Õ”«» €Ì— ’ÕÌÕ", vbCritical
                  Else
                        MsgBox "Account Code  not Exist ", vbCritical
                  End If
                  
                    .TextMatrix(Roww1, Col) = ""
                    .TextMatrix(Roww1, .ColIndex("AccountCode")) = ""
                     .TextMatrix(Roww1, .ColIndex("AccountName")) = ""
                     
                    Exit Sub
                End If

                rs.Close
                Set rs = Nothing
            
           FrmAccEditJournal3.Fg_Journal_AfterEdit (Roww1), 0
           Roww1 = Roww1 + 1
         End If
        Next i
       
  End With
End Sub
Private Sub CmdOk_Click()
save
  Unload Me
End Sub





Sub LoadGride()
  Dim RsDev As ADODB.Recordset
    Dim strsql As String
    Dim i As Integer
    Dim Roww1 As Integer
    Dim StrItemID As String
    Dim ac1, ac2, ac3, ac4, ac5, ac6 As String
   Dim RsDev1 As ADODB.Recordset
   Dim strsql2 As String
   Roww1 = FrmAccEditJournal3.Roww
   If IndexType = 1 Or IndexType = 2 Or IndexType = 3 Then
strsql = "SELECT     Account_Code, Account_Serial, Account_NameEng, Account_Name"
strsql = strsql & "  from dbo.ACCOUNTS where (last_account=1)"

If TxtCodeItem.text <> "" Then
strsql = strsql & " and   Account_Serial like'%" & TxtCodeItem.text & "%'"
End If
If TxtItem.text <> "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
                strsql = strsql & " and  Account_Name like'%" & TxtItem.text & "%'"
        Else
            strsql = strsql & " and  Account_NameEng like'%" & TxtItem.text & "%'"
        End If
  End If

  If IndexType = 1 Then
 strsq2 = " SELECT     Emp_ID, Account_code, Account_code1, Account_Code2, Account_Code3, Account_Code4, Account_Code5"
strsq2 = strsq2 & " From dbo.TblEmployee"
strsq2 = strsq2 & " Where (Emp_id = " & val(FrmAccEditJournal3.Fg_Journal.TextMatrix(Roww1, FrmAccEditJournal3.Fg_Journal.ColIndex("NEmpid"))) & ")"
Set RsDev1 = New ADODB.Recordset
    RsDev1.Open strsq2, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If RsDev1.RecordCount > 0 Then
    RsDev1.MoveFirst
   ac1 = IIf(IsNull(RsDev1("Account_code").value), "", (RsDev1("Account_code").value))
   ac2 = IIf(IsNull(RsDev1("Account_code1").value), "", (RsDev1("Account_code1").value))
   ac3 = IIf(IsNull(RsDev1("Account_Code2").value), "", (RsDev1("Account_Code2").value))
   ac4 = IIf(IsNull(RsDev1("Account_Code3").value), "", (RsDev1("Account_Code3").value))
  ac5 = IIf(IsNull(RsDev1("Account_Code5").value), "", (RsDev1("Account_Code5").value))
  ac6 = IIf(IsNull(RsDev1("Account_Code4").value), "", (RsDev1("Account_Code4").value))
End If
 ElseIf IndexType = 2 Then
 strsq2 = " SELECT     CusID, Account_Code_As_Client, Account_Code_As_Supplier, Account_Code, Account_Code1, Account_Code2"
strsq2 = strsq2 & " From dbo.TblCustemers"
strsq2 = strsq2 & " Where (CusID = " & val(FrmAccEditJournal3.Fg_Journal.TextMatrix(Roww1, FrmAccEditJournal3.Fg_Journal.ColIndex("NEmpid"))) & ")"
Set RsDev1 = New ADODB.Recordset
    RsDev1.Open strsq2, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If RsDev1.RecordCount > 0 Then
    RsDev1.MoveFirst
   ac1 = IIf(IsNull(RsDev1("Account_Code_As_Client").value), "", (RsDev1("Account_Code_As_Client").value))
   ac2 = IIf(IsNull(RsDev1("Account_Code_As_Supplier").value), "", (RsDev1("Account_Code_As_Supplier").value))
   ac3 = IIf(IsNull(RsDev1("Account_Code").value), "", (RsDev1("Account_Code").value))
   ac4 = IIf(IsNull(RsDev1("Account_Code1").value), "", (RsDev1("Account_Code1").value))
   ac5 = IIf(IsNull(RsDev1("Account_Code2").value), "", (RsDev1("Account_Code2").value))

End If
End If
  If IndexType = 1 Or IndexType = 2 Then
  strsql = strsql & " and (Account_Code='" & ac1 & "' or Account_Code='" & ac2 & "'or Account_Code='" & ac3 & "'or Account_Code='" & ac4 & "' or Account_Code='" & ac5 & "' or Account_Code='" & ac6 & "' ) "
  End If
         ' ItemID, ItemCode, ItemName, ItemNamee
         
    Set RsDev = New ADODB.Recordset
    RsDev.Open strsql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
  '
                ' .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsDev("ItemID").value), 0, val(RsDev("ItemID").value))
                 
                .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(RsDev("Account_Serial").value), "", (RsDev("Account_Serial").value))
             
             
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                Else
                    .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
                End If
              
             
                
              '  .Cell(flexcpChecked, i, .ColIndex("Select")) = flexUnchecked
            
                RsDev.MoveNext
            Next i
    .AutoSize 0, .Cols - 1, False
        End With
 End If
    End If
    
End Sub

Private Sub Form_Load()
    Resize_Form Me
      If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
LoadGride

End Sub

'Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'Dim rs As ADODB.Recordset
'Dim strsql As String
'With Grid1
'Select Case .ColKey(Col)
'Case "Select"
'
'                    strsql = "SELECT  ACCOUNTS.Account_Code From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(Grid1.TextMatrix(Row, Grid1.ColIndex("Account_Serial"))) & "'"
'                Set rs = Nothing
'                Set rs = New ADODB.Recordset
'                rs.Open strsql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'                If Not (rs.BOF Or rs.EOF) Then
'                    If BolEditOnMainAccounts = False Then
'                        If FrmAccEditJournal3.LastAccount(rs("Account_Code").value) = False Then
'                        .Cell(flexcpChecked, Row, .ColIndex("Select")) = flexUnchecked
'
'                            Exit Sub
'                        End If
'                    End If
'                  End If
'
'
'End Select
'End With
'End Sub

Private Sub Grid1_Click()

'Form_Load
'TxtCodeItem.text = ""
'TxtItem.text = ""
End Sub



Private Sub TxtCodeItem_Change()
LoadGride
End Sub

Private Sub TxtItem_Change()
LoadGride
End Sub
