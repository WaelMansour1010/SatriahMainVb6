Attribute VB_Name = "newModule"
Public Function ShowOperator(typeid As Integer) As String

 
                        If SystemOptions.UserInterface = ArabicInterface Then
                                    If typeid = 0 Then
                                    ShowOperator = "«ﬂ»— „‰"
                                    ElseIf typeid = 1 Then
                                    ShowOperator = "«ﬂ»— „‰ «Ê Ì”«ÊÌ"
                                    ElseIf typeid = 2 Then
                                    ShowOperator = "«ﬁ· „‰"
                                    ElseIf typeid = 3 Then
                                    ShowOperator = "«ﬁ· „‰ «ÊÌ”«ÊÌ"
                                    ElseIf typeid = 4 Then
                                    ShowOperator = "Ì”«ÊÌ"
                                    ElseIf typeid = 5 Then
                                    ShowOperator = "·« Ì”«ÊÌ"

                                    End If
                        Else
                        
                                    If typeid = 0 Then
                                    ShowOperator = "Greater Than"
                                    ElseIf typeid = 1 Then
                                    ShowOperator = "«Greater Than or Equal"
                                    ElseIf typeid = 2 Then
                                    ShowOperator = "Less than"
                                    ElseIf typeid = 3 Then
                                    ShowOperator = "Less than or equal"
                                    ElseIf typeid = 4 Then
                                    ShowOperator = "Equal"
                                    ElseIf typeid = 5 Then
                                    ShowOperator = "Not Equal"

                                    End If
                        
                        End If
                        
 
 


End Function
Public Function checkDataCreteria(Grid As vsFlexGrid)
Dim typeid As Integer
Dim Enteredvalue As Double
Dim value As Double
  With Grid
            For i = .FixedRows To .Rows - 1
                 typeid = val(.TextMatrix(i, .ColIndex("typeid")))
                 Enteredvalue = val(.TextMatrix(i, .ColIndex("Enteredvalue")))
                 value = val(.TextMatrix(i, .ColIndex("value")))
                 
                      If typeid = 0 Then
                          If Enteredvalue > value Then
                          .TextMatrix(i, .ColIndex("Done")) = "„Õﬁﬁ"
                          .TextMatrix(i, .ColIndex("doneid")) = 1
                             .Cell(flexcpBackColor, i, 0, i, 28) = &HFF00&
                          Else
                           .TextMatrix(i, .ColIndex("Done")) = "€Ì— „Õﬁﬁ"
                           .TextMatrix(i, .ColIndex("doneid")) = 0
                            .Cell(flexcpBackColor, i, 0, i, 28) = &HFF&
                          End If
                          
                          
                      ElseIf typeid = 1 Then
                      
                           If Enteredvalue >= value Then
                          .TextMatrix(i, .ColIndex("Done")) = "„Õﬁﬁ"
                          .TextMatrix(i, .ColIndex("doneid")) = 1
                             .Cell(flexcpBackColor, i, 0, i, 28) = &HFF00&
                          Else
                           .TextMatrix(i, .ColIndex("Done")) = "€Ì— „Õﬁﬁ"
                           .TextMatrix(i, .ColIndex("doneid")) = 0
                            .Cell(flexcpBackColor, i, 0, i, 28) = &HFF&
                          End If
                          
                      ElseIf typeid = 2 Then
                               If Enteredvalue < value Then
                          .TextMatrix(i, .ColIndex("Done")) = "„Õﬁﬁ"
                          .TextMatrix(i, .ColIndex("doneid")) = 1
                             .Cell(flexcpBackColor, i, 0, i, 28) = &HFF00&
                          Else
                           .TextMatrix(i, .ColIndex("Done")) = "€Ì— „Õﬁﬁ"
                           .TextMatrix(i, .ColIndex("doneid")) = 0
                            .Cell(flexcpBackColor, i, 0, i, 28) = &HFF&
                          End If
                      ElseIf typeid = 3 Then
                      
                               If Enteredvalue <= value Then
                          .TextMatrix(i, .ColIndex("Done")) = "„Õﬁﬁ"
                          .TextMatrix(i, .ColIndex("doneid")) = 1
                             .Cell(flexcpBackColor, i, 0, i, 28) = &HFF00&
                          Else
                           .TextMatrix(i, .ColIndex("Done")) = "€Ì— „Õﬁﬁ"
                           .TextMatrix(i, .ColIndex("doneid")) = 0
                            .Cell(flexcpBackColor, i, 0, i, 28) = &HFF&
                          End If
                          
                      ElseIf typeid = 4 Then
                               If Enteredvalue = value Then
                          .TextMatrix(i, .ColIndex("Done")) = "„Õﬁﬁ"
                          .TextMatrix(i, .ColIndex("doneid")) = 1
                             .Cell(flexcpBackColor, i, 0, i, 28) = &HFF00&
                          Else
                           .TextMatrix(i, .ColIndex("Done")) = "€Ì— „Õﬁﬁ"
                           .TextMatrix(i, .ColIndex("doneid")) = 0
                            .Cell(flexcpBackColor, i, 0, i, 28) = &HFF&
                          End If
                      ElseIf typeid = 5 Then
                               If Enteredvalue <> value Then
                          .TextMatrix(i, .ColIndex("Done")) = "„Õﬁﬁ"
                          .TextMatrix(i, .ColIndex("doneid")) = 1
                             .Cell(flexcpBackColor, i, 0, i, 28) = &HFF00&
                          Else
                           .TextMatrix(i, .ColIndex("Done")) = "€Ì— „Õﬁﬁ"
                           .TextMatrix(i, .ColIndex("doneid")) = 0
                            .Cell(flexcpBackColor, i, 0, i, 28) = &HFF&
                          End If
                      End If
             
            
            
            
            Next i
       .AutoSize 0, .Cols - 1, False
   End With

End Function

