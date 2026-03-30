Attribute VB_Name = "ModPremis"

Public Enum DoOperation

    Do_New
    Do_Edit
    Do_save
    Do_Delete
    Do_Search
    Do_Preview
    Do_Print
    DO_UNDO
    Do_Tools
    Do_Open
    Do_Close
    Do_Copy
    Do_Paste
    Do_Refresh
Do_Attach
End Enum

Public Enum UsersPremis
    NoPremis '·« ÊÃœ ’·«ÕÌ«  ‰Â«∆Ì«
    FullPremis '’·«ÕÌ…  «„…
    CustomePremis '’·«ÕÌ… „Õœœ…
End Enum

Public Function AllowToSeeItems(Trans_ID) As Boolean
    Dim StrSQL As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset

    If SystemOptions.usertype = UserNourCo Or SystemOptions.usertype = UserAdminAll Or SystemOptions.UserItemsPremis = FullPremis Then
        AllowToSeeItems = True
        Exit Function
    End If

    StrSQL = "Select TRANSACTION_DETAILS.Item_ID, ITEMS.Group_Code " & "  FROM ITEMS INNER JOIN TRANSACTION_DETAILS ON ITEMS.Item_ID =" & " TRANSACTION_DETAILS.Item_ID "
    StrSQL = StrSQL + " Where TRANSACTION_DETAILS.Transaction_Header_ID='" & Trans_ID & "'"
    StrSQL = StrSQL + " AND Group_Code NOT IN(Select Users_Junk_Groups.Group_Code From Users_Junk_Groups " & "Where Users_Junk_Groups.User_ID=" & LngUserID & ")"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        If rs.RecordCount > 0 Then
            Msg = "⁄‹‹‹‹›‹‹‹‹Ê«"
            Msg = "Â–Â «·Õ—ﬂ…  Õ ÊÏ ⁄·Ï √’‰«› ·Ì” ·ﬂ Õﬁ «·√ÿ·«⁄ ⁄·ÌÂ«"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            AllowToSeeItems = False
        End If

    Else
        AllowToSeeItems = True
    End If

End Function

Public Function DoPremis(DoMode As DoOperation, _
                         Optional frmname As String = "", _
                         Optional AlarmMsgBox As Boolean = True) As Boolean
    '  DoPremis = True
    '    Exit Function
    Dim StrSQL As String
    Dim rs As New ADODB.Recordset
    Dim Msg  As String
    Dim FrmCap As String
'If frmname = "frmsalebill5" Then

'        DoPremis = True
'        Exit Function
'End If

    If frmname = "FrmShowPrice" And GeneralPriceType <> 0 Then

        frmname = frmname & GeneralPriceType
    End If
If frmname = "frmsalebill2" Then frmname = "frmsalebill"
If frmname = "frmsalebill1" Then frmname = "frmsalebill"

If user_id = 1 Or DoMode = 9 Then ' power admin
    If SystemOptions.usertype = UserNourCo Or SystemOptions.usertype = UserAdminAll Or SystemOptions.UserScreenPremis = FullPremis Then

        DoPremis = True
        Exit Function
    End If
End If
    'If SystemOptions.UserScreenPremis = NoPremis Then
    '    If AlarmMsgBox <> False Then
    '        Msg = "⁄›Ê«" & Chr(13)
    '        Msg = Msg & "·Ì”  ·ﬂ «Ì… ’·«ÕÌ… ··⁄„· ›Ï «Ï ‘«‘… „‰ ‘«‘«  «·»—‰«„Ã..."
    '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    End If
    '    DoPremis = False
    '    Exit Function
    'End If
    On Error GoTo ErrTrap
   ' StrSQL = "SELECT  canshow,ScreenJuncUser.User_ID, ScreenJuncUser.ScreenName," & "ScreenJuncUser.CanAdd,ScreenJuncUser.CanEdit,ScreenJuncUser.CanDelete," & "ScreenJuncUser.CanPrint,ScreenJuncUser.CanSearch"
   ' StrSQL = StrSQL + ",Screens.ScreenCaption,Screens.ScreenTitleEng "
   ' StrSQL = StrSQL + " FROM ScreenJuncUser RIGHT OUTER JOIN Screens ON " & "ScreenJuncUser.ScreenName = Screens.ScreenName "
   ' StrSQL = StrSQL + "Where .ScreenJuncUser.User_ID=" & user_id & "  And ScreenJuncUser.ScreenName='" & frmname & "'"
    
    
    StrSQL = "SELECT  canshow,ScreenJuncUser.User_ID, ScreenJuncUser.ScreenName," & "ScreenJuncUser.CanAdd,ScreenJuncUser.CanEdit,ScreenJuncUser.CanDelete," & "ScreenJuncUser.CanPrint,ScreenJuncUser.CanSearch,ScreenJuncUser.Attachments"
 '   StrSQL = StrSQL + ",Screens.ScreenCaption,Screens.ScreenTitleEng "
    StrSQL = StrSQL + " FROM ScreenJuncUser  "
    StrSQL = StrSQL + "Where .ScreenJuncUser.User_ID=" & user_id & "  And ScreenJuncUser.ScreenName='" & frmname & "'"
        
        
   Select Case DoMode

        Case Do_New
        StrSQL = StrSQL + " order by ScreenJuncUser.CanAdd Desc "
         Case Do_Edit
         StrSQL = StrSQL + " order by ScreenJuncUser.CanEdit Desc "
        Case Do_Delete
        StrSQL = StrSQL + " order by ScreenJuncUser.CanDelete Desc "
        Case Do_Search
        StrSQL = StrSQL + " order by ScreenJuncUser.CanSearch Desc "
 Case Do_Print
 StrSQL = StrSQL + " order by ScreenJuncUser.Canprint Desc "
Case Do_Open
StrSQL = StrSQL + " order by ScreenJuncUser.CanShow Desc "
Case Do_Attcah
StrSQL = StrSQL + " order by ScreenJuncUser.Attachments Desc "

End Select
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If AlarmMsgBox <> False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & "·Ì” ·ﬂ «Ï ’·«ÕÌ… ··⁄„· ›Ï Â–Â «·‘«‘…. "
        Else
        Msg = "Sorry" & CHR(13)
            Msg = Msg & "Not Authorized to do that. "
        
        End If
            
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        End If

        DoPremis = False
        Exit Function
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
     '   FrmCap = IIf(IsNull(rs("ScreenCaption").value), "", rs("ScreenCaption").value)
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
       ' FrmCap = IIf(IsNull(rs("ScreenTitleEng").value), "", rs("ScreenTitleEng").value)
    End If

    Select Case DoMode

        Case Do_New

            If rs("CanAdd").value = True Then

                DoPremis = True
            Else

                If AlarmMsgBox <> False Then
                              If SystemOptions.UserInterface = ArabicInterface Then
                                  Msg = "⁄›Ê«..." & CHR(13)
                                  Msg = Msg & " ·Ì”  ·ﬂ ’·«ÕÌ… ≈÷«›… ”Ã· ÃœÌœ " & CHR(13)
                                  Msg = Msg & "›Ï ‘«‘… " & FrmCap
                              Else
                                Msg = "Sorry" & CHR(13)
                          Msg = Msg & "Not Authorized to do that. "
                    
                              End If
                          MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

                DoPremis = False
                Exit Function
            End If

        Case Do_Edit

            If rs("CanEdit").value = True Then

                DoPremis = True
            Else

                If AlarmMsgBox <> False Then
                     
                    
                                                  If SystemOptions.UserInterface = ArabicInterface Then
                   Msg = "⁄›Ê«..." & CHR(13)
                    Msg = Msg & " ·Ì”  ·ﬂ ’·«ÕÌ… «· ⁄œÌ· " & CHR(13)
                    Msg = Msg & "›Ï ‘«‘… " & FrmCap
                              Else
                                Msg = "Sorry" & CHR(13)
                          Msg = Msg & "Not Authorized to do that. "
                    
                              End If



                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

                DoPremis = False
                Exit Function
            End If

        Case Do_Delete

            If rs("CanDelete").value = True Then

                DoPremis = True
            Else

                If AlarmMsgBox <> False Then
                    
                    
                    
                       
                                 If SystemOptions.UserInterface = ArabicInterface Then
                          Msg = "⁄›Ê«..." & CHR(13)
                    Msg = Msg & " ·Ì”  ·ﬂ ’·«ÕÌ… Õ–› ”Ã· " & CHR(13)
                    Msg = Msg & "›Ï ‘«‘… " & FrmCap
                                                                                     
                    Else
                                Msg = "Sorry" & CHR(13)
                          Msg = Msg & "Not Authorized to do that. "
                    
                              End If


                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

                DoPremis = False
                Exit Function
            End If

        Case Do_Search

            If rs("CanSearch").value = True Then

                DoPremis = True
            Else

                If AlarmMsgBox <> False Then
                     
                    
                                If SystemOptions.UserInterface = ArabicInterface Then
                                                                                      
                    Msg = "⁄›Ê«..." & CHR(13)
                    Msg = Msg & " ·Ì”  ·ﬂ ’·«ÕÌ… «·»ÕÀ " & CHR(13)
                    Msg = Msg & "›Ï ‘«‘… " & FrmCap
                                                                                     
                    Else
                                Msg = "Sorry" & CHR(13)
                          Msg = Msg & "Not Authorized to do that. "
                    
                              End If



                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

                DoPremis = False
                Exit Function
            End If

        Case Do_Print

            If rs("CanPrint").value = True Then

                DoPremis = True
            Else

                If AlarmMsgBox <> False Then
                    
                    
                                     If SystemOptions.UserInterface = ArabicInterface Then
                     Msg = "⁄›Ê«..." & CHR(13)
                    Msg = Msg & " ·Ì”  ·ﬂ ’·«ÕÌ… «·ÿ»«⁄… " & CHR(13)
                    Msg = Msg & "›Ï ‘«‘… " & FrmCap
                                                                                     
                                                                                     
                    Else
                                Msg = "Sorry" & CHR(13)
                          Msg = Msg & "Not Authorized to do that. "
                    
                              End If



                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

                DoPremis = False
                Exit Function
            End If

        Case Do_Open

       '     If rs("CanAdd").value = False And rs("CanEdit").value = False And rs("CanDelete").value = False And rs("CanPrint").value = False And rs("CanSearch").value = False And rs("CanShow").value = False Then
    If rs("CanShow").value = False Then
                If AlarmMsgBox <> False Then
                    
                    
                                        
                                     If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "⁄›Ê«" & CHR(13)
                    Msg = Msg & "«·„” Œœ„ ·Ì” ·Â √Ì… ’·«ÕÌ«  ··⁄„· ›Ï " & CHR(13)
                    Msg = Msg & FrmCap
                                                                                     
                                                                                     
                    Else
                                Msg = "Sorry" & CHR(13)
                          Msg = Msg & "Not Authorized to do that. "
                    
                              End If



                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

                DoPremis = False
                Exit Function
            Else

                DoPremis = True
            End If

    
       Case Do_Attach

            If rs("Attachments").value = True Then

                DoPremis = True
            Else

                If AlarmMsgBox <> False Then
                    
                    
                                     If SystemOptions.UserInterface = ArabicInterface Then
                     Msg = "⁄›Ê«..." & CHR(13)
                    Msg = Msg & " ·Ì”  ·ﬂ ’·«ÕÌ… «·„—›ﬁ«  " & CHR(13)
                    Msg = Msg & "›Ï ‘«‘… " & FrmCap
                                                                                     
                                                                                     
                    Else
                                Msg = "Sorry" & CHR(13)
                          Msg = Msg & "Not Authorized to do that. "
                    
                              End If



                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

                DoPremis = False
                Exit Function
            End If

     
    End Select

    Exit Function
ErrTrap:

    If AlarmMsgBox <> False Then
        
                                        
                                     If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÕœÀ Œÿ√ ›Ï «·Ã“¡ «·Œ«’ »’·«ÕÌ«  «·„” Œœ„Ì‰" & CHR(13)
        Msg = Msg & "»—Ã«¡ «·√ ’«· »«·„”∆Ê·."
                                                                                     
                                                                                     
                    Else
        Msg = "ERROR IN Privliges" & CHR(13)
        Msg = Msg & " Contact Administrator"
                    
                              End If
        
        
        
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

    DoPremis = False
End Function

Public Sub ShowScreenPermission(StrScreenName As String)
    Dim Msg As String

    If SystemOptions.usertype = UserNormal Then
        Msg = "Â–Â «·‘«‘… „ «Õ… ›ﬁÿ .. ·„œÌ— «·‰Ÿ«„"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Load FrmPermissionScreen
    FrmPermissionScreen.DcboScreens.BoundText = StrScreenName
    FrmPermissionScreen.show vbModal
End Sub

