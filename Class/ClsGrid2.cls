VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGrid2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Grid As vsFlexGrid
Attribute Grid.VB_VarHelpID = -1
Public formname As String
Public WithEvents LblDone As Label
Attribute LblDone.VB_VarHelpID = -1
Public WithEvents LblunDone As Label
Attribute LblunDone.VB_VarHelpID = -1
Public WithEvents LblAccCreitria As Label
Attribute LblAccCreitria.VB_VarHelpID = -1



 


Public Sub Grid_AfterEdit(ByVal Row As Long, _
                          ByVal Col As Long)

    Dim StrAccountCode As String
 
    Dim LngRow As Long
 Dim RsDetails As New ADODB.Recordset
    
 
Dim StrSQL As String
    With Me.Grid

        Select Case .ColKey(Col)
 
            Case "Enteredvalue", "remarks", "remarksE"
        
                StrAccountCode = .ComboData
    '            LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Subdes"), False, True)
       StrSQL = "Select * From  tblScreenCriteriaValues Where id=" & val(.ComboData)
        RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If RsDetails.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("remarks")) = IIf(IsNull(RsDetails("SubDes").value), 0, RsDetails("SubDes").value)
                .TextMatrix(Row, .ColIndex("remarksE")) = IIf(IsNull(RsDetails("SubDese").value), 0, RsDetails("SubDese").value)
                
            .TextMatrix(Row, .ColIndex("Enteredvalue")) = IIf(IsNull(RsDetails("SubValue").value), 0, RsDetails("SubValue").value)
 RsDetails.Close
 Set RsDetails = Nothing
 End If
           
        End Select

 
    End With
 

 

checkDataCreteria
End Sub



Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid

        Select Case .ColKey(Col)

            Case "PlainMessageID"
              
                Cancel = True
 
            Case "Name"
 
                Cancel = True
            
            Case "typeName"
     
                Cancel = True
            
            Case "value"
                Cancel = True
           Case "Done"
                Cancel = True
                   
                    Case "doneid"
                Cancel = True
                
                 
                
                
        End Select
End With
End Sub
Public Function ShowCreteria() As Boolean
Dim i As Integer
     Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Set RsDev = New ADODB.Recordset
    StrSQL = " SELECT     dbo.tblCriteriaSettingsDetails.PlainMessageID, dbo.tblScreenCriteria.Name, dbo.tblScreenCriteria.Namee, dbo.tblScreenCriteria.[value], "
StrSQL = StrSQL & "   dbo.tblScreenCriteria.typeid, dbo.tblScreenCriteria.Remarks"
StrSQL = StrSQL & "  FROM         dbo.tblCriteriaSettingsDetails INNER JOIN"
StrSQL = StrSQL & "   dbo.tblScreenCriteria ON dbo.tblCriteriaSettingsDetails.PlainMessageID = dbo.tblScreenCriteria.CriteriaID INNER JOIN"
StrSQL = StrSQL & "   dbo.tblCriteriaSettings ON dbo.tblCriteriaSettingsDetails.lMessageDefID = dbo.tblCriteriaSettings.id"
StrSQL = StrSQL & "  WHERE     (dbo.tblCriteriaSettings.ScreenName = N'" & formname & "')"
    
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    ShowCreteria = True
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
  
                .TextMatrix(i, .ColIndex("PlainMessageID")) = IIf(IsNull(RsDev("PlainMessageID").value), 0, val(RsDev("PlainMessageID").value))

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsDev("Name").value), "", (RsDev("Name").value))
                Else
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsDev("Namee").value), "", (RsDev("Namee").value))
            
                End If
  
   
             .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("value").value), 0, val(RsDev("value").value))
               .TextMatrix(i, .ColIndex("typeid")) = IIf(IsNull(RsDev("typeid").value), 0, val(RsDev("typeid").value))
      
                       .TextMatrix(i, .ColIndex("typeName")) = ShowOperator(val(.TextMatrix(i, .ColIndex("typeid"))))
               .TextMatrix(i, .ColIndex("remarks")) = IIf(IsNull(RsDev("remarks").value), "", (RsDev("remarks").value))
          
                        
               RsDev.MoveNext
            Next i
 
            .AutoSize 0, .Cols - 1, False
        End With

Else
ShowCreteria = False

    End If
 

checkDataCreteria
End Function




Public Function checkDataCreteria()
Dim typeid As Integer
Dim Enteredvalue As Double
Dim value As Double
Dim i As Integer
LblDone.Caption = 0
LblunDone.Caption = 0
LblAccCreitria.Caption = 0
  With Me.Grid
            For i = .FixedRows To .Rows - 1
                 typeid = val(.TextMatrix(i, .ColIndex("typeid")))
                 Enteredvalue = val(.TextMatrix(i, .ColIndex("Enteredvalue")))
                 value = val(.TextMatrix(i, .ColIndex("value")))
                 LblAccCreitria = LblAccCreitria + 1
                      If typeid = 0 Then
                          If Enteredvalue >= value Then
                          .TextMatrix(i, .ColIndex("Done")) = "„Õﬁﬁ"
                          .TextMatrix(i, .ColIndex("doneid")) = 1
                             .Cell(flexcpBackColor, i, 0, i, 12) = &HFF00&
                             LblDone = LblDone + 1
                          Else
                           .TextMatrix(i, .ColIndex("Done")) = "€Ì— „Õﬁﬁ"
                           .TextMatrix(i, .ColIndex("doneid")) = 0
                             LblunDone = LblunDone + 1
                            .Cell(flexcpBackColor, i, 0, i, 12) = &HFF&
                          End If
                          
                          
                      ElseIf typeid = 1 Then
                      
                           If Enteredvalue >= value Then
                          .TextMatrix(i, .ColIndex("Done")) = "„Õﬁﬁ"
                          .TextMatrix(i, .ColIndex("doneid")) = 1
                            LblDone = LblDone + 1
                             .Cell(flexcpBackColor, i, 0, i, 12) = &HFF00&
                          Else
                           .TextMatrix(i, .ColIndex("Done")) = "€Ì— „Õﬁﬁ"
                           .TextMatrix(i, .ColIndex("doneid")) = 0
                            .Cell(flexcpBackColor, i, 0, i, 12) = &HFF&
                            LblunDone = LblunDone + 1
                          End If
                          
                      ElseIf typeid = 2 Then
                               If Enteredvalue < value Then
                          .TextMatrix(i, .ColIndex("Done")) = "„Õﬁﬁ"
                          .TextMatrix(i, .ColIndex("doneid")) = 1
                            LblDone = LblDone + 1
                             .Cell(flexcpBackColor, i, 0, i, 12) = &HFF00&
                          Else
                           .TextMatrix(i, .ColIndex("Done")) = "€Ì— „Õﬁﬁ"
                           .TextMatrix(i, .ColIndex("doneid")) = 0
                            .Cell(flexcpBackColor, i, 0, i, 12) = &HFF&
                            LblunDone = LblunDone + 1
                          End If
                      ElseIf typeid = 3 Then
                      
                               If Enteredvalue <= value Then
                          .TextMatrix(i, .ColIndex("Done")) = "„Õﬁﬁ"
                          .TextMatrix(i, .ColIndex("doneid")) = 1
                            LblDone = LblDone + 1
                             .Cell(flexcpBackColor, i, 0, i, 12) = &HFF00&
                          Else
                           .TextMatrix(i, .ColIndex("Done")) = "€Ì— „Õﬁﬁ"
                           .TextMatrix(i, .ColIndex("doneid")) = 0
                             LblunDone = LblunDone + 1
                            .Cell(flexcpBackColor, i, 0, i, 12) = &HFF&
                            LblunDone = LblunDone + 1
                          End If
                          
                      ElseIf typeid = 4 Then
                               If Enteredvalue = value Then
                          .TextMatrix(i, .ColIndex("Done")) = "„Õﬁﬁ"
                          .TextMatrix(i, .ColIndex("doneid")) = 1
                            LblDone = LblDone + 1
                             .Cell(flexcpBackColor, i, 0, i, 12) = &HFF00&
                          Else
                           .TextMatrix(i, .ColIndex("Done")) = "€Ì— „Õﬁﬁ"
                           .TextMatrix(i, .ColIndex("doneid")) = 0
                            .Cell(flexcpBackColor, i, 0, i, 12) = &HFF&
                            LblunDone = LblunDone + 1
                          End If
                      ElseIf typeid = 5 Then
                               If Enteredvalue <> value Then
                          .TextMatrix(i, .ColIndex("Done")) = "„Õﬁﬁ"
                          .TextMatrix(i, .ColIndex("doneid")) = 1
                            LblDone = LblDone + 1
                             .Cell(flexcpBackColor, i, 0, i, 12) = &HFF00&
                          Else
                           .TextMatrix(i, .ColIndex("Done")) = "€Ì— „Õﬁﬁ"
                           .TextMatrix(i, .ColIndex("doneid")) = 0
                            .Cell(flexcpBackColor, i, 0, i, 12) = &HFF&
                            LblunDone = LblunDone + 1
                          End If
                          
                                        ElseIf typeid = 6 Then
        
                          
                          .TextMatrix(i, .ColIndex("Done")) = "„Ê«’›« "
                          .TextMatrix(i, .ColIndex("doneid")) = 1
                            LblDone = LblDone + 1
                             .Cell(flexcpBackColor, i, 0, i, 12) = &HFF00&
                
                          
                          
                          
                      End If
             
            
            
            
            Next i
       .AutoSize 0, .Cols - 1, False
   End With

End Function



Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String

    Dim Msg As String

     With Me.Grid

        Select Case .ColKey(Col)

            Case "Enteredvalue"
                StrSQL = "select Distinct * from tblScreenCriteriaValues where CriteriaID=" & val(.TextMatrix(Row, .ColIndex("PlainMessageID")))
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = .BuildComboList(rs, "Subvalue", "Id")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
         
            Case "remarks"
                StrSQL = "select Distinct Id,Subdes  from tblScreenCriteriaValues where CriteriaID=" & val(.TextMatrix(Row, .ColIndex("PlainMessageID")))
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = .BuildComboList(rs, "Subdes", "Id")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
                
                            Case "remarksE"
                StrSQL = "select  Distinct * from tblScreenCriteriaValues where CriteriaID=" & val(.TextMatrix(Row, .ColIndex("PlainMessageID")))
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = .BuildComboList(rs, "Subdese", "Id")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList


        End Select

    End With
End Sub
