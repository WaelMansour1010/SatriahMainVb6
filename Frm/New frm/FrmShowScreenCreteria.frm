VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmShowScreenCreteria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄—÷ „Õœœ«  «·‘«‘« "
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13425
   Icon            =   "FrmShowScreenCreteria.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   13425
   Begin VB.Frame Frame1 
      Caption         =   "œ·«·«  «·«·Ê«‰"
      Height          =   735
      Left            =   9600
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4680
      Width           =   3615
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "€Ì— „Õﬁﬁ"
         Height          =   255
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "„Õﬁﬁ"
         Height          =   255
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1080
         Top             =   240
         Width           =   735
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2640
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Œ—ÊÃ"
      Height          =   375
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton CMDOk 
      Caption         =   "„Ê«›ﬁ"
      Height          =   375
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4920
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4635
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      _cx             =   23310
      _cy             =   8176
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
      Cols            =   30
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmShowScreenCreteria.frx":000C
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
End
Attribute VB_Name = "FrmShowScreenCreteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function ShowCreteria(scrrenname As String) As Boolean
Dim i As Integer
     Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Set RsDev = New ADODB.Recordset
    StrSQL = " SELECT     dbo.tblCriteriaSettingsDetails.PlainMessageID, dbo.tblScreenCriteria.Name, dbo.tblScreenCriteria.Namee, dbo.tblScreenCriteria.[value], "
StrSQL = StrSQL & "   dbo.tblScreenCriteria.typeid"
StrSQL = StrSQL & "  FROM         dbo.tblCriteriaSettingsDetails INNER JOIN"
StrSQL = StrSQL & "   dbo.tblScreenCriteria ON dbo.tblCriteriaSettingsDetails.PlainMessageID = dbo.tblScreenCriteria.CriteriaID INNER JOIN"
StrSQL = StrSQL & "   dbo.tblCriteriaSettings ON dbo.tblCriteriaSettingsDetails.lMessageDefID = dbo.tblCriteriaSettings.id"
StrSQL = StrSQL & "  WHERE     (dbo.tblCriteriaSettings.ScreenName = N'" & scrrenname & "')"
    
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
      
                        If SystemOptions.UserInterface = ArabicInterface Then
                                    If val(.TextMatrix(i, .ColIndex("typeid"))) = 0 Then
                                    .TextMatrix(i, .ColIndex("typeName")) = "«ﬂ»— „‰"
                                    ElseIf val(.TextMatrix(i, .ColIndex("typeid"))) = 1 Then
                                    .TextMatrix(i, .ColIndex("typeName")) = "«ﬂ»— „‰ «Ê Ì”«ÊÌ"
                                    ElseIf val(.TextMatrix(i, .ColIndex("typeid"))) = 2 Then
                                    .TextMatrix(i, .ColIndex("typeName")) = "«ﬁ· „‰"
                                    ElseIf val(.TextMatrix(i, .ColIndex("typeid"))) = 3 Then
                                    .TextMatrix(i, .ColIndex("typeName")) = "«ﬁ· „‰ «ÊÌ”«ÊÌ"
                                    ElseIf val(.TextMatrix(i, .ColIndex("typeid"))) = 4 Then
                                    .TextMatrix(i, .ColIndex("typeName")) = "Ì”«ÊÌ"
                                    ElseIf val(.TextMatrix(i, .ColIndex("typeid"))) = 5 Then
                                    .TextMatrix(i, .ColIndex("typeName")) = "·« Ì”«ÊÌ"

                                    End If
                        Else
                        
                                    If val(.TextMatrix(i, .ColIndex("typeid"))) = 0 Then
                                    .TextMatrix(i, .ColIndex("typeName")) = "Greater Than"
                                    ElseIf val(.TextMatrix(i, .ColIndex("typeid"))) = 1 Then
                                    .TextMatrix(i, .ColIndex("typeName")) = "«Greater Than or Equal"
                                    ElseIf val(.TextMatrix(i, .ColIndex("typeid"))) = 2 Then
                                    .TextMatrix(i, .ColIndex("typeName")) = "Less than"
                                    ElseIf val(.TextMatrix(i, .ColIndex("typeid"))) = 3 Then
                                    .TextMatrix(i, .ColIndex("typeName")) = "Less than or equal"
                                    ElseIf val(.TextMatrix(i, .ColIndex("typeid"))) = 4 Then
                                    .TextMatrix(i, .ColIndex("typeName")) = "Equal"
                                    ElseIf val(.TextMatrix(i, .ColIndex("typeid"))) = 5 Then
                                    .TextMatrix(i, .ColIndex("typeName")) = "Not Equal"

                                    End If
                        
                        End If
                        
               RsDev.MoveNext
            Next i
 
            .AutoSize 0, .Cols - 1, False
        End With

Else
ShowCreteria = False

    End If
 


End Function
Public Function checkData()
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
 

Private Sub CMDOK_Click()
   checkData
End Sub

Private Sub Form_Load()
        Me.top = (mdifrmmain.ScaleHeight - Me.Height) / 2
        Me.left = (mdifrmmain.ScaleWidth - Me.Width) / 2
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
checkData
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
                
                    Case "remarks"
                Cancel = True
                
                
        End Select
End With
End Sub

