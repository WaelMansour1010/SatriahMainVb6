VERSION 5.00
Begin VB.Form FrmEditUsers22 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»Ì«‰«  «·„” Œœ„Ì‰"
   ClientHeight    =   11550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   21570
   Icon            =   "frmEditUsers22.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   11550
   ScaleWidth      =   21570
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "FrmEditUsers22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim RsTemp As New ADODB.Recordset
Private Sub BtnCancel_Click()
    Unload Me
End Sub
Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ  «·„” Œœ„ " & TXTCode.Text & CHR(13) & "   «”„ «·„” Œœ„  " & XPTxtUserName.Text & CHR(13) & "   «·›—⁄ " & DcBranches.Text
        LogTexte = "  Screen  " & ScreenNameEnglish & CHR(13) & " User Code " & TXTCode.Text & CHR(13) & "   User Name  " & XPTxtUserName.Text & CHR(13) & "   Branch " & DcBranches
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), , Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , TXTCode, TXTCode
    Else
        AddToLogFile CInt(user_id), , Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TXTCode, TXTCode
    End If
End Function
Private Sub btnDelete_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    If TxtVac_ID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
     Else
        MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
   End If
        If MSGType = vbYes Then
           Cn.Execute "Delete from TblUsersStores where userid = " & val(TxtVac_ID.Text) & ""
           Cn.Execute "Delete from TblUsersBranches where userid = " & val(TxtVac_ID.Text) & ""
           Cn.Execute "Delete from TblUsersBoxes where userid = " & val(TxtVac_ID.Text) & ""
           Cn.Execute "Delete from TblUserAccount where UserID = " & val(TxtVac_ID.Text) & ""
           Cn.Execute "Delete from TblUsersProductLine where UserID = " & val(TxtVac_ID.Text) & ""
            RsSavRec.Find "userid=" & val(TxtVac_ID.Text), , adSearchForward, 1
            CuurentLogdata ("D")
            Dim StrSQL As String
            StrSQL = "Delete From TblUsersStores Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            
            StrSQL = "Delete From TblUsersProductLine Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
                       
            
            StrSQL = "Delete From TblUsersBranches Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            Set Me.ImgPic.Picture = Nothing
            RsSavRec.delete
            If SystemOptions.UserInterface = ArabicInterface Then
               MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
               MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            End If
            '------------------------------ Move Next ---------------------------.
            FillGridWithData
            BtnNext_Click
        End If
    End If

    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            'StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
           If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry ... This record can not be deleted because it is linked to other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select
End Sub
Private Sub BtnFirst_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec.MoveFirst
    FiLLTXT

    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
           ' Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
           ' Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
           ' Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
      If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
     Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MoveLast
    FiLLTXT
    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
          '  Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
          '  Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
          '  Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
    If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
     Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnModify_Click()
    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID.Text <> "" Then
        TxtModFlg = "E"
        Frm2.Enabled = True
        'Me.TXTDiscounts.SetFocus
        CuurentLogdata
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
          '  Msg = "⁄›Ê«" & Chr(13)
          '  Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
          '  Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
           Else
            Msg = "Sorry..." & CHR(13)
            Msg = Msg & " This record can not be edited at this time" & CHR(13)
            Msg = Msg & "Because it was modified by another user on the network"
         
           End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    '-----------------------------------
    Me.TxtVac_ID.Text = ""
 
    Me.DcBranches.BoundText = ""
    Me.DCEmP.BoundText = ""
    Me.DCJob.BoundText = ""
    Me.DCSalesRepGroups.BoundText = ""
    
    clear_all Me
    FillGridWithData
    CboPriv.ListIndex = 0
    '-----------------------------------
    TxtModFlg.Text = "N"

    My_SQL = "TBLSalesRepData"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.Text = rs.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If

    rs.Close
    CmbType.ListIndex = 0

    ListGroupSelected.Clear
    ListBoxesSelected.Clear
    ListStoreSelected.Clear
    
 
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext

        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If

    FiLLTXT
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
            'Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            'Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            'Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
    If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
           Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MovePrevious

    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If

    FiLLTXT
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
          '  Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
          '  Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
          '  Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
         If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
           Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnQuery_Click()
FrmUserSearch.show
FrmUserSearch.lblSearchtype = 0

End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
    If CboPriv.ListIndex = -1 Then
    
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify Privligies"
        Else
              Msg = "Õœœ «·’·«ÕÌ« "
        End If
        
        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboPriv.SetFocus
        Exit Sub
    End If
 
    If Trim(DcBranches.BoundText) = "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify Branch"
        Else
            Msg = "Õœœ «·›—⁄ «·«› —«÷Ì  "
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcBranches.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
 
    If Trim(Me.DCEmP.BoundText) = "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify Employee"
        Else
            Msg = "Õœœ «·„ÊŸ›    "
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCEmP.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
 
    If XPTxtUserName.Text = "" Then
    
         If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify User"
        Else
           Msg = "√œŒ· «”„ «·„” Œœ„"
        End If
        
        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtUserName.SetFocus
        Exit Sub
    End If

 '   If TxtPassWord.text = "" Then
 '       Msg = "√œŒ· ﬂ·„… «·„—Ê—"
 '       MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 '       TxtPassWord.SetFocus
 '       Exit Sub
 '   End If
 '
 '   If XPTxtPassConfirm.text = "" Then
 '       Msg = "√œŒ·  √ﬂÌœ ﬂ·„… «·„—Ê—"
 '       MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 '       XPTxtPassConfirm.SetFocus
 '       Exit Sub
 '   End If
Dim StrSQL As String
    If StrComp(TxtPassWord.Text, XPTxtPassConfirm.Text, vbTextCompare) <> 0 Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Passwords not matched"
            Else
                Msg = "ﬂ·„… «·„—Ê— Ê √ﬂÌœ ﬂ·„… «·„—Ê— " & CHR(13)
                Msg = Msg + "€Ì— „ ÿ«»ﬁ Ì‰"
             End If
        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtPassConfirm.SetFocus
        Exit Sub
    End If

    StrSQL = "select * From TblUsers where UserName='" & Trim(XPTxtUserName.Text) & "'" & " and UserID<>" & val(TxtVac_ID.Text)
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
    If SystemOptions.UserInterface = EnglishInterface Then
    Msg = "Another user already Exist with the same name"
    Else
        Msg = "ÌÊÃœ „” Œœ„ „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ" & CHR(13)
        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„” Œœ„"
    End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtUserName.SetFocus
        RsTemp.Close
        Exit Sub
    End If

 
    '------------------------------ check if Empcode exist ----------------------
 
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text

            '------------------------------ new record ----------------------------
        Case "N"
            '------------------------- save record -----------------------------
            AddNewRec
            BtnLast_Click

        Case "E"
            '----------------------------- save edit -------------------------------
            'RsEmployee("userid").value = RsSavRec("UserID").value
            StrSQL = "Delete From TblUsersStores Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From TblUsersProductLine Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From TblUsersBranches Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            
            FiLLRec
    End Select

    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = EnglishInterface Then
MsgBox "error during saving", vbOKOnly + vbMsgBoxRight, App.title
Else
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
End If
End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtVac_ID.Text)
    Me.TxtModFlg.Text = "R"
End Sub

Private Sub BtnUpdate_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    FristCount = RsSavRec.RecordCount
    RsSavRec.Requery
    LastCount = RsSavRec.RecordCount
    BtnUndo_Click
If SystemOptions.UserInterface = ArabicInterface Then
    If FristCount = LastCount Then
        Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
    Else
        Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
        End If
    End If
Else
    If FristCount = LastCount Then
        Msg = "No new data"
    Else
        Msg = "Number of records before update" & vbCrLf & FristCount & vbCrLf & "Number of records after  update" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "Number of new records" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "Number of records deleted" & vbCrLf & FristCount - LastCount
        End If
    End If
End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub

Private Sub CmdPic_Click(Index As Integer)
On Error GoTo ErrTrap
    Select Case Index

        Case 0

            With cdg
               
                .CancelError = False
                .DialogTitle = " ≈Œ Ì«— ’Ê—…"
                'Set The Filter to show pictures only
                .filter = "Bitmap (*.bmp)|*.bmp|JPEG(*.JPG,*.JPEG,*.JPE,*.JFIF)|*.jpg;*.jpeg;*.jpe;*.jfif|"  ' choose formats to include
          
                .ShowOpen

                If .filename <> "" Then
                    Set Me.ImgPic.Picture = LoadPicture(.filename)
                End If

            End With

        Case 1
            Set Me.ImgPic.Picture = Nothing
    End Select
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " ÕÃ„ «·’Ê—… €Ì— „œ⁄Ê„", vbCritical
Else
MsgBox " image Size Not Siutable, vbCritical"
End If


End Sub

Private Sub DBCboClientName_Change()
Dim Fullcode As String
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode ', 1, DepitIntervalID, DepitInterval, , creditlocked
    TxtSearchCode.Text = Fullcode

End Sub

Private Sub ImgPic_DblClick()
  Load FrmViewPic
    Set FrmViewPic.MainView.Picture = ImgPic.Picture
    FrmViewPic.show vbModal
End Sub

Private Sub Label15_Click()
'    If ListBoxesSelected.ListIndex > -1 Then
'        ListBoxesSelected.RemoveItem ListBoxesSelected.ListIndex
'    End If
'
'
Dim i As Long

For i = 0 To ListBoxesSelected.ListCount - 1
    If i > ListBoxesSelected.ListCount - 1 Then Exit For
    If ListBoxesSelected.Selected(i) Then
        ListBoxesSelected.RemoveItem i
        'ListStoreSelected.ListIndex
        i = i - 1
    End If
Next

End Sub

Private Sub Label16_Click()
    ListBoxesSelected.Clear
End Sub
Private Sub Label17_Click()
    Dim i As Integer
    
    ListBoxesSelected.Clear

    For i = 0 To ListBoxesAll.ListCount - 1
        ListBoxesSelected.AddItem ListBoxesAll.List(i)
        ListBoxesSelected.ItemData(i) = ListBoxesAll.ItemData(i)
    Next i

End Sub

Private Sub Label18_Click()
'    If ListBoxesAll.ListIndex = -1 Then Exit Sub
'    ListBoxesSelected.AddItem ListBoxesAll.List(ListBoxesAll.ListIndex)
'    ListBoxesSelected.ItemData(ListBoxesSelected.NewIndex) = ListBoxesAll.ItemData(ListBoxesAll.ListIndex)

    Dim i As Long
    
    For i = 0 To ListBoxesAll.ListCount - 1
        If ListBoxesAll.Selected(i) Then
            ListBoxesSelected.AddItem ListBoxesAll.List(i)
            ListBoxesSelected.ItemData(ListBoxesSelected.NewIndex) = ListBoxesAll.ItemData(i)
            
        End If
    Next
            

End Sub

Private Sub Label21_Click()
'    If ListAllAccount.ListIndex = -1 Then Exit Sub
'    ListAccountSelect.AddItem ListAllAccount.List(ListAllAccount.ListIndex)
'    ListAccountSelect.ItemData(ListAccountSelect.NewIndex) = ListAllAccount.ItemData(ListAllAccount.ListIndex)
'
    
            If ListStoreall.ListIndex = -1 Then Exit Sub
'    ListStoreSelected.AddItem ListStoreall.List(ListStoreall.ListIndex)
'    ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(ListStoreall.ListIndex)
'
    
Dim i As Long

For i = 0 To ListAllAccount.ListCount - 1
    If ListAllAccount.Selected(i) Then
        ListAccountSelect.AddItem ListAllAccount.List(i)
        ListAccountSelect.ItemData(ListAccountSelect.NewIndex) = ListAllAccount.ItemData(i)
        
    End If
Next

'ItemData (i)

End Sub

Private Sub Label23_Click()
    Dim i As Integer
    ListAccountSelect.Clear
    For i = 0 To ListAllAccount.ListCount - 1
        ListAccountSelect.AddItem ListAllAccount.List(i)
        ListAccountSelect.ItemData(i) = ListAllAccount.ItemData(i)
    Next i
End Sub

Private Sub Label24_Click()
 ListAccountSelect.Clear
End Sub

Private Sub Label25_Click()
 
        

Dim i As Long

For i = 0 To ListAccountSelect.ListCount - 1
    If i > ListAccountSelect.ListCount - 1 Then Exit For
    If ListAccountSelect.Selected(i) Then
        ListAccountSelect.RemoveItem i
        'ListStoreSelected.ListIndex
        i = i - 1
    End If
Next

End Sub

Private Sub ListAllAccount_Click()

End Sub

Private Sub ListAllAccount_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
                      Account_search.show
                     Account_search.case_id = 78912

                   End If
End Sub

Private Sub ListBoxesAll_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        FrmExpensesSearch.Indx = 2
        FrmExpensesSearch.RetrunType = 986
        FrmExpensesSearch.show
    End If
End Sub

Private Sub ListStoreall_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        FrmStoreSearch.mIndex = 1
        Set FrmStoreSearch.RetrunFrm = Me
        FrmStoreSearch.show
    End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer
    
    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 1
        DBCboClientName.BoundText = CUSTID
    End If
End Sub
Private Sub dcEmp_Change()
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
        If (Me.DCEmP.BoundText) = "" Then Exit Sub
        Me.TXTCode.Text = get_EMPLOYEE_Data(val(Me.DCEmP.BoundText), "Fullcode")
        'DCEmp.text = DCEmp.text
    End If
End Sub
Private Sub Dcemp_Click(Area As Integer)
    dcEmp_Change
End Sub
Private Sub DCEmP_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 2911
        Set FrmEmployeeSearch.RetrunFrm = Me
        FrmEmployeeSearch.show
    End If
End Sub
Private Sub Form_Load()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos

    My_SQL = "TblUsers"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient

    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    
    
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches Me.DcBranches
    Dcombos.GetEmployees Me.DCEmP
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
    Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName1
    Dcombos.GetStores Me.DCStore
    Dcombos.GetStores Me.DCStore1
    Dcombos.GetStores Me.DCStore3
    Dcombos.GetStores Me.DCStore2
    Dcombos.GetBoxes Me.DCBoxes
    Dcombos.GetBoxes Me.DCBoxes1
    Dcombos.GetBanks Me.Dbanks

    Set cSearch = New clsDCboSearch
    Set cSearch.Client = Me.DCEmP
    Set cSearch.Client = Me.DcBranches
    Set cSearch.Client = Me.DCStore
    Set cSearch.Client = Me.DCBoxes
    Set cSearch.Client = Me.DCBoxes1
    Set cSearch.Client = Me.Dbanks


    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("EmpName"), Me.DCEmP
    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("BranchId"), Me.DcBranches
    
    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("StoreID"), Me.DCStore
    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("BoxID"), Me.DCBoxes
    'ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("BoxID"), Me.DCBoxes
    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("BankID"), Me.Dbanks

    FillGridWithData

    With Me.Grid
        .Cell(flexcpPicture, 0, .ColIndex("DiscountValue")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
        For i = 0 To .Cols - 1
            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
        Next
   
        .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
        .RowHeight(-1) = 300
    End With

    BtnFirst_Click
    ShowTip

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    FillMylist

    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If

ErrTrap:
End Sub
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    Label20.Caption = "All Accounts"
    Label19.Caption = "Selected Accounts"
    btnQuery.Caption = "Search"
    Ele(6).Caption = "Electronic Signature"
    Me.Caption = "Users Data"
    Me.Label1(2).Caption = Me.Caption
    Label1(3).Caption = "Code"
    Label1(1).Caption = "Name"
    lbl(4).Caption = "Privligies"
    Label1(9).Caption = "User Name"
    Label1(5).Caption = "Branch"
    Label1(4).Caption = "Box for Sale"
    Label1(14).Caption = "Box for Purchase"
    Label1(5).Caption = "Branch"
    Label1(0).Caption = "Password"
    Label1(8).Caption = "Re. password"
    Label1(7).Caption = "Sale Store"
    Label1(12).Caption = "Store Purchase"
    Label1(10).Caption = "Default Client"
    Label1(11).Caption = "Default Supplier"
    CmdPic(0).Caption = "Add Picture"
    CmdPic(1).Caption = "Delete Picture"
    Label1(6).Caption = "Bank"
    chkNextLogin.Caption = "Change password at login"
    
    chkHidLowering.Caption = "Hide the subtraction of output alerts"
Frame1.Caption = "Selected Boxes"
Frame2.Caption = "Selected Accounts"
    Frame11.Caption = "Selected Branch"
    Frame10.Caption = "Selected Stores"
    Label11.Caption = "All Branch"
    Label12.Caption = "Selected Branch"

    Label9.Caption = "All Stores"
    Label10.Caption = "Selected Stores"

    Label2(0).Caption = "Current Record"
    Label2(1).Caption = "NO. Recordes"

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
 
    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "Ser"
        .TextMatrix(0, .ColIndex("EmpCode")) = "Code"
        .TextMatrix(0, .ColIndex("EmpName")) = "Emp Name"
        .TextMatrix(0, .ColIndex("JobID")) = "Job"
        .TextMatrix(0, .ColIndex("groupid")) = "Group"
        .TextMatrix(0, .ColIndex("BranchId")) = "Branch"
        .TextMatrix(0, .ColIndex("discountvalue")) = "Discount%"
        .TextMatrix(0, .ColIndex("UserName")) = "UserName"
        .TextMatrix(0, .ColIndex("StoreId")) = "Store Name"
        .TextMatrix(0, .ColIndex("boxId")) = "Box Name"
        .TextMatrix(0, .ColIndex("BankID")) = "Bank"
    End With
    
    '######### khaled was here ############
    isDeactivatedchk.Caption = "Deactivate User"
    Label14.Caption = "All Boxes"
    Label13.Caption = "Selected Boxes"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
        Select Case Me.TxtModFlg.Text
            Case "N"
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & CHR(13)
                    StrMSG = StrMSG & " the new data  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
                End If
            Case "E"
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & CHR(13)
                    StrMSG = StrMSG & " the Modifications  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
                End If
        End Select

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)

        Select Case IntResult
            Case vbYes
                Cancel = True
                btnSave_Click
            Case vbCancel
                Cancel = True
        End Select
    End If

    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
    Set FrmVacancy = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo ErrTrap

    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If

        RsSavRec.Close
        Set RsSavRec = Nothing
    End If

    Set cSearch = Nothing
ErrTrap:
End Sub
Private Sub Label5_Click()
    If ListGroupSelected.ListIndex > -1 Then
        ListGroupSelected.RemoveItem ListGroupSelected.ListIndex
    End If
End Sub
Private Sub Label6_Click()
    ListGroupSelected.Clear
End Sub
Private Sub Label7_Click()
    Dim i As Integer
    ListGroupSelected.Clear

    For i = 0 To ListGroupAll.ListCount - 1
        ListGroupSelected.AddItem ListGroupAll.List(i)
        ListGroupSelected.ItemData(i) = ListGroupAll.ItemData(i)
    Next i
End Sub
Private Sub Label8_Click()
    If ListGroupAll.ListIndex = -1 Then Exit Sub
    ListGroupSelected.AddItem ListGroupAll.List(ListGroupAll.ListIndex)
    ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(ListGroupAll.ListIndex)
End Sub
Private Sub LblSelect_Click()
'    If ListStoreall.ListIndex = -1 Then Exit Sub
'    ListStoreSelected.AddItem ListStoreall.List(ListStoreall.ListIndex)
'    ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(ListStoreall.ListIndex)
'
    
        If ListStoreall.ListIndex = -1 Then Exit Sub
'    ListStoreSelected.AddItem ListStoreall.List(ListStoreall.ListIndex)
'    ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(ListStoreall.ListIndex)
'
    
Dim i As Long

For i = 0 To ListStoreall.ListCount - 1
    If ListStoreall.Selected(i) Then
        ListStoreSelected.AddItem ListStoreall.List(i)
        ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(i)
        
    End If
Next

'ItemData (i)
End Sub
Private Sub Label22_Click()
    Dim i As Integer
    ListStoreSelected.Clear
    For i = 0 To ListStoreall.ListCount - 1
        ListStoreSelected.AddItem ListStoreall.List(i)
        ListStoreSelected.ItemData(i) = ListStoreall.ItemData(i)
    Next i
End Sub
Private Sub Label3_Click()
    ListStoreSelected.Clear
End Sub
Private Sub Label4_Click()
'    If ListStoreSelected.ListIndex > -1 Then
'        ListStoreSelected.RemoveItem ListStoreSelected.ListIndex
'    End If
    

Dim i As Long

For i = 0 To ListStoreSelected.ListCount - 1
    If i > ListStoreSelected.ListCount - 1 Then Exit For
    If ListStoreSelected.Selected(i) Then
        ListStoreSelected.RemoveItem i
        'ListStoreSelected.ListIndex
        i = i - 1
    End If
Next
    
    
End Sub
Function createlistString(mylist As ListBox, Optional ByRef Listitems As String)
    Dim i As Integer
    Dim str As String
    str = "0"
    Listitems = ""
    For i = 0 To mylist.ListCount - 1
        str = str & "," & mylist.ItemData(i)
        Listitems = Listitems & "," & mylist.List(i)
    Next i
    createlistString = str
End Function


Function FillMylist(Optional ByVal mIndexd As Long = 0)

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
    
    If mIndexd = 1 Or mIndexd = 0 Then
        sql = " SELECT * from  TblStore "
    
        If SystemOptions.UserInterface = ArabicInterface Then
            sql = sql & " order by  StoreName"
        Else
            sql = sql & " order by  StoreNamee"
        End If
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListStoreall.Clear
        'ListStoreSelected.Clear
    
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    ListStoreall.AddItem IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
                Else
                    ListStoreall.AddItem IIf(IsNull(rs("StoreNamee").value), "", rs("StoreNamee").value)
                End If
    
                ListStoreall.ItemData(ListStoreall.NewIndex) = rs("StoreID").value
                rs.MoveNext
            Next i
        End If
    
        rs.Close
    End If
    If mIndexd = 0 Then
        sql = " SELECT * from  TblBranchesData "
     
        If SystemOptions.UserInterface = ArabicInterface Then
            sql = sql & " order by  branch_name"
        Else
            sql = sql & " order by  branch_namee"
        End If
     
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListGroupAll.Clear
        'ListGroupSelected.Clear
    
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    ListGroupAll.AddItem IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                    ListGroupAll.AddItem IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                End If
    
                ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("branch_id").value
                rs.MoveNext
            Next i
        End If
        rs.Close
    End If
'    sql = "select* from TblBoxesData where Type = 0 "
If mIndexd = 2 Or mIndexd = 0 Then
        sql = "select* from TblBoxesData    "
        ' sql = "select* from TblBoxesData where  "
        If SystemOptions.UserInterface = ArabicInterface Then
            sql = sql & " order by  BoxName"
        Else
            sql = sql & " order by  BoxNameE"
        End If
     
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListBoxesAll.Clear
        
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    ListBoxesAll.AddItem IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
                Else
                    ListBoxesAll.AddItem IIf(IsNull(rs("BoxNameE").value), "", rs("BoxNameE").value)
                End If
    
                ListBoxesAll.ItemData(ListBoxesAll.NewIndex) = rs("BoxID").value
                rs.MoveNext
            Next i
        End If
        rs.Close
      ''/////Account
   End If
   If mIndexd = 3 Or mIndexd = 0 Then
        sql = " SELECT * from  ACCOUNTS "
        sql = sql & " where   last_account=0"
        If SystemOptions.UserInterface = ArabicInterface Then
            sql = sql & " order by  Account_Name"
        Else
            sql = sql & " order by  Account_NameEng"
        End If
     
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListAllAccount.Clear
        'ListGroupSelected.Clear
    
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    ListAllAccount.AddItem IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                Else
                    ListAllAccount.AddItem IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                End If
    
                ListAllAccount.ItemData(ListAllAccount.NewIndex) = rs("Account_ID").value
                rs.MoveNext
            Next i
        End If
        rs.Close
        
    End If
    If mIndexd = 0 Then
      sql = "select * from TblProductLine "
        ' sql = "select* from TblBoxesData where  "
       
        sql = sql & " order by  Name"
        
     
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListProductLineAll.Clear
        
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                ListProductLineAll.AddItem IIf(IsNull(rs("Name").value), "", rs("Name").value)
    
                ListProductLineAll.ItemData(ListProductLineAll.NewIndex) = rs("ID").value
                rs.MoveNext
            Next i
        End If
    End If
End Function
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub AddNewRec()

    On Error GoTo ErrTrap
    
    Dim StrRecID As String
    
    'StrRecID = new_id("TBLSalesRepData", "id", "")
    
    RsSavRec.AddNew
    RsSavRec("UserID").value = CStr(new_id("TblUsers", "UserID", "", True))
    'RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Public Sub FiLLRec()

    On Error GoTo ErrTrap

    RsSavRec.Fields("PassWord").value = IIf((TxtPassWord.Text) <> "", (TxtPassWord.Text), "")
    RsSavRec.Fields("EmpID").value = IIf(val(Me.DCEmP.BoundText) <> 0, val(Me.DCEmP.BoundText), Null)
    RsSavRec.Fields("BranchId").value = IIf(val(Me.DcBranches.BoundText) <> 0, val(Me.DcBranches.BoundText), Null)
    RsSavRec.Fields("BoxID").value = IIf(val(Me.DCBoxes.BoundText) <> 0, val(Me.DCBoxes.BoundText), Null)
    RsSavRec.Fields("BoxID1").value = IIf(val(Me.DCBoxes1.BoundText) <> 0, val(Me.DCBoxes1.BoundText), Null)
    
    RsSavRec.Fields("BankID").value = IIf(val(Me.Dbanks.BoundText) <> 0, val(Me.Dbanks.BoundText), Null)
    RsSavRec.Fields("StoreID").value = IIf(val(Me.DCStore.BoundText) <> 0, val(Me.DCStore.BoundText), Null)
    RsSavRec.Fields("Custid").value = IIf(val(Me.DBCboClientName.BoundText) <> 0, val(Me.DBCboClientName.BoundText), Null)
    
    RsSavRec.Fields("StoreID1").value = IIf(val(Me.DCStore1.BoundText) <> 0, val(Me.DCStore1.BoundText), Null)
    RsSavRec.Fields("StoreID3").value = IIf(val(Me.DCStore3.BoundText) <> 0, val(Me.DCStore3.BoundText), Null)
    RsSavRec.Fields("StoreID2").value = IIf(val(Me.DCStore2.BoundText) <> 0, val(Me.DCStore2.BoundText), Null)
    RsSavRec.Fields("Custid1").value = IIf(val(Me.DBCboClientName1.BoundText) <> 0, val(Me.DBCboClientName1.BoundText), Null)
    
    RsSavRec("UserName").value = Trim(XPTxtUserName.Text)
    If ImgPic.Picture = 0 Then
        RsSavRec("UserSign").value = Null
    Else
        If SavePictureToDB(ImgPic, RsSavRec, "UserSign") = False Then
            GoTo ErrTrap
        End If
    End If

    If Me.CboPriv.ListIndex = 0 Then
        RsSavRec("UserType").value = 2
    Else
        RsSavRec("InvPrices").value = 1
        RsSavRec("InvPrices1").value = 1
        RsSavRec("InvPrices2").value = 1
        
        RsSavRec("ShowInvProfit").value = 1
        RsSavRec("AllowOverMax").value = 1

        RsSavRec("FullPremis").value = 1
        RsSavRec("UserType").value = 0
    End If
 
    RsSavRec("PassConfirm").value = Trim(XPTxtPassConfirm.Text)
   
    RsSavRec("IsActive").value = 1
    
 
    If chkNextLogin.value = vbChecked Then
        RsSavRec("ChangePW").value = 1
    Else
        RsSavRec("ChangePW").value = 0
    End If
    
    If chkHidLowering.value = vbChecked Then
        RsSavRec("HidLowering").value = 1
    Else
        RsSavRec("HidLowering").value = 0
    End If
        
    
    
    '########## Khaled's was here #################
    If isDeactivatedchk.value = vbChecked Then
        RsSavRec("isDeactivated").value = 1
    Else
        RsSavRec("isDeactivated").value = 0
    End If
    '###############################################
 
    
    'RsSavRec.Fields("JobID").value = IIf(Me.DCJob.BoundText <> 0, Val(Me.DCJob.BoundText), Null)

    RsSavRec.update
    Dim UsrID As Double
   UsrID = IIf(IsNull(RsSavRec("UserID").value), 0, RsSavRec("UserID").value)
    If Me.TxtModFlg.Text = "E" Then
    Cn.Execute "Delete from TblUsersStores where userid = " & UsrID & ""
    Cn.Execute "Delete from TblUsersBranches where userid = " & UsrID & ""
    Cn.Execute "Delete from TblUsersBoxes where userid = " & UsrID & ""
    Cn.Execute "Delete from TblUserAccount where UserID = " & UsrID & ""
    Cn.Execute "Delete from TblUsersProductLine where UserID = " & UsrID & ""
    
    End If
    Dim i As Integer
    Dim RsEmployee As New ADODB.Recordset
    
        If ListStoreSelected.ListCount <> 0 Then
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open "TblUsersStores", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
            For i = 0 To ListStoreSelected.ListCount - 1
                RsEmployee.AddNew
                RsEmployee("storeId").value = ListStoreSelected.ItemData(i)
                RsEmployee("userid").value = UsrID
                RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If

        If ListGroupSelected.ListCount <> 0 Then
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open "TblUsersBranches", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
                For i = 0 To ListGroupSelected.ListCount - 1
                    RsEmployee.AddNew
                    RsEmployee("BranchID").value = ListGroupSelected.ItemData(i)
                    RsEmployee("userid").value = UsrID
                    RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If
        
        If ListBoxesSelected.ListCount <> 0 Then
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open "TblUsersBoxes", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
                For i = 0 To ListBoxesSelected.ListCount - 1
                    RsEmployee.AddNew
                    RsEmployee("BoxId").value = ListBoxesSelected.ItemData(i)
                    RsEmployee("userid").value = UsrID
                    RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If
        
    If ListProductLineSelected.ListCount <> 0 Then
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open "TblUsersProductLine", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
                For i = 0 To ListProductLineSelected.ListCount - 1
                    RsEmployee.AddNew
                    RsEmployee("ProductLineId").value = ListProductLineSelected.ItemData(i)
                    RsEmployee("userid").value = UsrID
                    'RsEmployee("ShowAlarm").value = FG.ValueMatrix(i, FG.ColIndex("ShowAlarm"))
                    RsEmployee("TypeLine").value = 0
                    
                    RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If
        Dim sql As String
        
        sql = "Select * from TblUsersProductLine Where  TypeLine = 1 "
        
        saveGrid sql, FG, "ShowAlarm", "", "userId", UsrID, "TypeLine", 1
        
        
         If ListAccountSelect.ListCount <> 0 Then
         sql = "select * from TblUserAccount   where 1=-1"
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
                For i = 0 To ListAccountSelect.ListCount - 1
                    RsEmployee.AddNew
                    RsEmployee("Account_ID").value = ListAccountSelect.ItemData(i)
                    RsEmployee("UserID").value = UsrID
                    RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If
        
        CuurentLogdata
        
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Else
            MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        End If
    
        FillGridWithData
        TxtModFlg = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub
Public Sub FiLLTXT()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    Frm2.Enabled = False
    TxtVac_ID.Text = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    TxtPassWord.Text = IIf(IsNull(RsSavRec.Fields("PassWord").value), "", RsSavRec.Fields("PassWord").value)
    XPTxtPassConfirm.Text = IIf(IsNull(RsSavRec.Fields("PassConfirm").value), "", RsSavRec.Fields("PassConfirm").value)
    XPTxtUserName.Text = IIf(IsNull(RsSavRec.Fields("UserName").value), 0, RsSavRec.Fields("UserName").value)
    If Not IsNull(RsSavRec("UserType").value) Then
        If RsSavRec("UserType").value = 2 Then
            CboPriv.ListIndex = 0
        Else
            CboPriv.ListIndex = 1
        End If
    End If
      
    If Not IsNull(RsSavRec("ChangePW").value) Then
        If RsSavRec("ChangePW").value = 0 Then
            chkNextLogin.value = vbUnchecked
        Else
            chkNextLogin.value = vbChecked
        End If
    Else
        chkNextLogin.value = vbUnchecked
    End If
    
   
    If Not IsNull(RsSavRec("HidLowering").value) Then
        If RsSavRec("HidLowering").value = 0 Then
            chkHidLowering.value = vbUnchecked
        Else
            chkHidLowering.value = vbChecked
        End If
    Else
        chkHidLowering.value = vbUnchecked
    End If
    
     
    
    
    '################# khaled was here #####################
    If Not IsNull(RsSavRec("isDeactivated").value) Then
        If RsSavRec("isDeactivated").value = 0 Then
            isDeactivatedchk.value = vbUnchecked
        Else
            isDeactivatedchk.value = vbChecked
        End If
    Else
        isDeactivatedchk.value = vbUnchecked
    End If
    '#######################################################
       
    Me.DCEmP.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
    TXTCode.Text = get_EMPLOYEE_Data(val(Me.DCEmP.BoundText), "fullcode")
    Me.DcBranches.BoundText = IIf(IsNull(RsSavRec.Fields("BranchId").value), "", RsSavRec.Fields("BranchId").value)
    Me.DCStore.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID").value), "", RsSavRec.Fields("StoreID").value)
    Me.DCStore1.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID1").value), "", RsSavRec.Fields("StoreID1").value)
    Me.DCStore3.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID3").value), "", RsSavRec.Fields("StoreID3").value)
        Me.DCStore2.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID2").value), "", RsSavRec.Fields("StoreID2").value)
    Me.DCBoxes1.BoundText = IIf(IsNull(RsSavRec.Fields("BoxID1").value), "", RsSavRec.Fields("BoxID1").value)
    Me.DCBoxes.BoundText = IIf(IsNull(RsSavRec.Fields("BoxID").value), "", RsSavRec.Fields("BoxID").value)
    Me.Dbanks.BoundText = IIf(IsNull(RsSavRec.Fields("BankID").value), "", RsSavRec.Fields("BankID").value)
    Me.DBCboClientName.BoundText = IIf(IsNull(RsSavRec.Fields("Custid").value), "", RsSavRec.Fields("Custid").value)
    Me.DBCboClientName1.BoundText = IIf(IsNull(RsSavRec.Fields("Custid1").value), "", RsSavRec.Fields("Custid1").value)
    If Not IsNull(RsSavRec("UserSign").value) Then
        If LenB(RsSavRec("UserSign")) Then
            LoadPictureFromDB ImgPic, RsSavRec, "UserSign"
        Else
            Set ImgPic.Picture = Nothing
        End If
    Else
        Set ImgPic.Picture = Nothing
    End If
    

'********************************************************************
     
    ListStoreSelected.Clear

    Dim RsEmployee As ADODB.Recordset
    Dim StrSQL As String
 
    StrSQL = " SELECT     TOP 100 PERCENT dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.StoreID"
    StrSQL = StrSQL & "  FROM         dbo.TblUsersStores INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblStore ON dbo.TblUsersStores.StoreID = dbo.TblStore.StoreID"
    StrSQL = StrSQL & "  Where (dbo.TblUsersStores.UserID = " & val(TxtVac_ID.Text) & ")"
    StrSQL = StrSQL & "  ORDER BY dbo.TblUsersStores.id"

    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            If SystemOptions.UserInterface = ArabicInterface Then
                ListStoreSelected.AddItem IIf(IsNull(RsEmployee("StoreName").value), "", RsEmployee("StoreName").value)
            Else
                ListStoreSelected.AddItem IIf(IsNull(RsEmployee("StoreNameE").value), "", RsEmployee("StoreNameE").value)
            End If
            ListStoreSelected.ItemData(i) = val(IIf(IsNull(RsEmployee("StoreID").value), 0, (RsEmployee("StoreID").value)))
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If


'*********************************************************************************
     
    ListGroupSelected.Clear

    StrSQL = " SELECT     TOP 100 PERCENT dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
    StrSQL = StrSQL & " FROM         dbo.TblUsersBranches INNER JOIN"
    StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.TblUsersBranches.BranchID = dbo.TblBranchesData.branch_id"
    StrSQL = StrSQL & " Where (dbo.TblUsersBranches.UserID = " & val(TxtVac_ID.Text) & ")"
    StrSQL = StrSQL & " ORDER BY dbo.TblUsersBranches.id"

    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroupSelected.AddItem IIf(IsNull(RsEmployee("branch_name").value), "", RsEmployee("branch_name").value)
            Else
                ListGroupSelected.AddItem IIf(IsNull(RsEmployee("branch_nameE").value), "", RsEmployee("branch_nameE").value)
            End If
            ListGroupSelected.ItemData(i) = val(IIf(IsNull(RsEmployee("branch_id").value), 0, (RsEmployee("branch_id").value)))
                
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If
'*********************************************************************************
    ListBoxesSelected.Clear
    
    StrSQL = "SELECT TblUsersBoxes.id, TblBoxesData.BoxName, TblUsersBoxes.BoxId, TblUsersBoxes.userid, TblBoxesData.BoxNameE"
    StrSQL = StrSQL & " FROM TblUsersBoxes INNER JOIN"
    StrSQL = StrSQL & " TblBoxesData ON TblUsersBoxes.BoxId = TblBoxesData.BoxID"
    StrSQL = StrSQL & " Where (TblUsersBoxes.UserID = " & val(TxtVac_ID.Text) & ")"

    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            If SystemOptions.UserInterface = ArabicInterface Then
                ListBoxesSelected.AddItem IIf(IsNull(RsEmployee("BoxName").value), "", RsEmployee("BoxName").value)
            Else
                ListBoxesSelected.AddItem IIf(IsNull(RsEmployee("BoxNameE").value), "", RsEmployee("BoxNameE").value)
            End If
            ListBoxesSelected.ItemData(i) = val(IIf(IsNull(RsEmployee("BoxId").value), 0, (RsEmployee("BoxId").value)))
                
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If
    
  ListProductLineSelected.Clear
    
    StrSQL = "SELECT TblUsersProductLine.id,TblUsersProductLine.ShowAlarm, TblProductLine.Name, TblUsersProductLine.ProductLineId, TblUsersProductLine.userid "
    StrSQL = StrSQL & " FROM TblUsersProductLine INNER JOIN"
    StrSQL = StrSQL & " TblProductLine ON TblUsersProductLine.ProductLineId = TblProductLine.ID"
    StrSQL = StrSQL & " Where (TblUsersProductLine.UserID = " & val(TxtVac_ID.Text) & ") and IsNull( TypeLine,0) = 0"

    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    StrSQL = "SELECT TblProductLine.id,TblProductLine.id as ProductLineID ,TblUsersProductLine.ShowAlarm, TblProductLine.Name,  TblUsersProductLine.userid "
    StrSQL = StrSQL & " FROM TblUsersProductLine RIGHT outer JOIN "
    StrSQL = StrSQL & " TblProductLine ON TblUsersProductLine.ProductLineId = TblProductLine.ID and (TblUsersProductLine.UserID = " & val(TxtVac_ID.Text) & ") and TypeLine = 1"
    
    loadgrid StrSQL, FG, True, False
    
    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            
                ListProductLineSelected.AddItem IIf(IsNull(RsEmployee("Name").value), "", RsEmployee("Name").value)
            
            ListProductLineSelected.ItemData(i) = val(IIf(IsNull(RsEmployee("ProductLineId").value), 0, (RsEmployee("ProductLineId").value)))
                
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If
        
'''//////////////
    ListAccountSelect.Clear
    
    StrSQL = " SELECT     dbo.TblUserAccount.UserID, dbo.TblUserAccount.Account_ID, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng"
    StrSQL = StrSQL & " FROM         dbo.TblUserAccount LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.ACCOUNTS ON dbo.TblUserAccount.Account_ID = dbo.ACCOUNTS.Account_ID"
    StrSQL = StrSQL & "     Where (dbo.TblUserAccount.UserID = " & val(TxtVac_ID.Text) & ")"
    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            If SystemOptions.UserInterface = ArabicInterface Then
                ListAccountSelect.AddItem IIf(IsNull(RsEmployee("Account_Name").value), "", RsEmployee("Account_Name").value)
            Else
                ListAccountSelect.AddItem IIf(IsNull(RsEmployee("Account_NameEng").value), "", RsEmployee("Account_NameEng").value)
            End If
            ListAccountSelect.ItemData(i) = val(IIf(IsNull(RsEmployee("Account_ID").value), 0, (RsEmployee("Account_ID").value)))
                
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If
'*********************************************************************************
    'Me.DCJob.BoundText = IIf(IsNull(RsSavRec.Fields("JobID").value), "", RsSavRec.Fields("JobID").value)

    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
    With Grid
        For i = 1 To .Rows - 1
            If Trim(TxtVac_ID.Text) = .TextMatrix(i, .ColIndex("UserID")) Then
                TxtSerial.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:

End Sub
Public Sub EditRec(StrTable As String, RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec
End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("UserID")))
ErrTrap:
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
    DCEmP.BoundText = GeTEmpIDByEmpCode(TXTCode.Text)
End If
End Sub
Private Sub TxtPassWord_DblClick()
'     If user_id = 1 Then
'     MsgBox txtPassword
'     End If
End Sub
Private Sub TxtSearchCode1_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer
    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 1
        DBCboClientName.BoundText = CUSTID
    End If
End Sub
Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "UserID=" & RecId, , adSearchForward, 1

    If Not (RsSavRec.EOF) Then
        FiLLTXT
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If

    'RsSavRec.Filter = adFilterNone
End Function

'Private Sub TxtVacCode_KeyPress(KeyAscii As Integer)
'KeyAscii = DataFormat(ChrOnly, KeyAscii)
'End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.Text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        'btnNext.Enabled = False
        'btnPrevious.Enabled = False
        'btnFirst.Enabled = False
        'btnLast.Enabled = False
        ListGroupAll.Enabled = True
        ListStoreall.Enabled = True
        ListBoxesAll.Enabled = True
        ListAllAccount.Enabled = True
        ListProductLineAll.Enabled = True
    ElseIf TxtModFlg.Text = "R" Then
        ListAllAccount.Enabled = False
        ListProductLineAll.Enabled = False
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtVac_ID.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
        ListGroupAll.Enabled = False
        ListStoreall.Enabled = False
        ListBoxesAll.Enabled = False
    ElseIf TxtModFlg.Text = "E" Then
        ListAllAccount.Enabled = True
                ListProductLineAll.Enabled = True
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
        ListGroupAll.Enabled = True
        ListStoreall.Enabled = True
        ListBoxesAll.Enabled = True
    End If

End Sub
Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblUsers order by userid"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("UserID")) = IIf(IsNull(rs.Fields("UserID").value), "", rs.Fields("UserID").value)
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(rs.Fields("EmpID").value), "", rs.Fields("EmpID").value)
                .TextMatrix(i, .ColIndex("EmpCode")) = get_EMPLOYEE_Data(val(.TextMatrix(i, .ColIndex("EmpID"))), "fullcode")
                .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs.Fields("UserName").value), "", rs.Fields("UserName").value)
                .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(rs.Fields("EmpID").value), "", rs.Fields("EmpID").value)
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("BranchId").value), "", rs.Fields("BranchId").value)
                .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(rs.Fields("StoreID").value), "", rs.Fields("StoreID").value)
                .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(rs.Fields("BoxID").value), "", rs.Fields("BoxID").value)
                .TextMatrix(i, .ColIndex("BankID")) = IIf(IsNull(rs.Fields("BankID").value), "", rs.Fields("BankID").value)
                rs.MoveNext
            Next i
            rs.Close
        End If
        .AutoSize 0, .Cols - 1, False
        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
            
        .AddControl btnNew, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With

ErrTrap:
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            btnNew_Click
        Else
            SendKeys "{TAB}"
        End If
    End If
    'New -------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If
    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If
    'save ------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If
    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If
    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If
    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If
    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
        BtnFirst_Click
    End If
    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If
    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If
    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
    End If

    Exit Sub
ErrTrap:
End Sub
Private Function CheckDelCountry(Lngid As Long) As Boolean
    'Dim Rs As ADODB.Recordset
    'Dim StrSQL As String
    'StrSQL = "Select * From TblEmployee Where GovernmentID=" & Lngid & ""
    'Set Rs = New ADODB.Recordset
    'Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If Not (Rs.BOF Or Rs.EOF) Then
    '    CheckDelCountry = False
    'Else
    '    CheckDelCountry = True
    'End If
    'Rs.Close
    'Set Rs = Nothing
End Function

Private Sub Label28_Click()
    If ListProductLineAll.ListIndex = -1 Then Exit Sub
    ListProductLineSelected.AddItem ListProductLineAll.List(ListProductLineAll.ListIndex)
    ListProductLineSelected.ItemData(ListProductLineSelected.NewIndex) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
'    FG.Rows = ListProductLineSelected.ListCount + 1
'    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Name")) = ListProductLineAll.List(ListProductLineAll.ListIndex)
'    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("ProductLineID")) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
End Sub

Private Sub Label29_Click()
    Dim i As Integer
    ListProductLineSelected.Clear
'    FG.Rows = 1
'    FG.Rows = ListProductLineSelected.ListCount + 1
    For i = 0 To ListProductLineAll.ListCount - 1
        ListProductLineSelected.AddItem ListProductLineAll.List(i)
        ListProductLineSelected.ItemData(i) = ListProductLineAll.ItemData(i)
'        FG.TextMatrix(i + 1, FG.ColIndex("Name")) = ListProductLineAll.List(ListProductLineAll.ListIndex)
'        FG.TextMatrix(i + 1, FG.ColIndex("ProductLineID")) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
        
    Next i

End Sub

Private Sub Label30_Click()
 ListProductLineSelected.Clear
' FG.Rows = 1
End Sub

Private Sub Label31_Click()
    If ListProductLineSelected.ListIndex > -1 Then
      ListProductLineSelected.RemoveItem ListProductLineSelected.ListIndex
        'FG.RemoveItem
    End If

End Sub



