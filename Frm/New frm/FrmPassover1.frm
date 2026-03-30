VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form Frmpassover1 
   Caption         =   "«·„” ‰œ«   ﬁÌœ «·«⁄ „«œ"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17085
   HelpContextID   =   440
   Icon            =   "FrmPassover1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   17085
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8205
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   17085
      _cx             =   30136
      _cy             =   14473
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   2
      ChildSpacing    =   1
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   3
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmPassover1.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   990
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   7185
         Width           =   17025
         _cx             =   30030
         _cy             =   1746
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   2
         ChildSpacing    =   1
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
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
         Begin VB.Frame Frame1 
            Caption         =   "œ·«·«  «·«·Ê«‰"
            Height          =   615
            Left            =   12825
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   360
            Width           =   4215
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "„ √Œ—"
               Height          =   255
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   240
               Width           =   1215
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H000000FF&
               FillColor       =   &H000000FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Left            =   3240
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.CheckBox ChkShow 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·«  ŸÂ— Â–Â «·‰«›–… ⁄‰œ  ‘€Ì· «·»—‰«„Ã"
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   6885
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   450
            Visible         =   0   'False
            Width           =   9870
         End
         Begin ImpulseButton.ISButton CmdExit 
            Cancel          =   -1  'True
            Height          =   375
            Left            =   105
            TabIndex        =   5
            Top             =   495
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   0
            ButtonImage     =   "FrmPassover1.frx":03E6
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton CmdPrint 
            Height          =   375
            Left            =   1590
            TabIndex        =   6
            Top             =   495
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄…"
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmPassover1.frx":0780
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ALLButtonS.ALLButton cmdAdd 
            Height          =   420
            Left            =   3345
            TabIndex        =   8
            Tag             =   "Delete Row"
            Top             =   480
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   741
            BTYPE           =   3
            TX              =   " ÕœÌÀ «·»Ì«‰« "
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   65280
            BCOLO           =   65280
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmPassover1.frx":0B1A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ì „  ÕœÌœ Â–Â «·»Ì«‰«  »‰«¡« ⁄·Ï «· «—ÌŒ «·Õ«·Ì ›Ì «·ÃÂ«“"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   7590
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   60
            Width           =   9330
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   6540
         Left            =   30
         TabIndex        =   2
         Top             =   630
         Width           =   17025
         _cx             =   30030
         _cy             =   11536
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         BackColorBkg    =   16777215
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   19
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmPassover1.frx":0B36
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
         ExplorerBar     =   1
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
      Begin VB.Image Image1 
         Height          =   585
         Left            =   30
         Picture         =   "FrmPassover1.frx":0E32
         Top             =   30
         Width           =   1125
      End
      Begin VB.Label LblCaption 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„” ‰œ«   ﬁÌœ «·«⁄ „«œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   585
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   30
         Width           =   17025
      End
   End
End
Attribute VB_Name = "Frmpassover1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Askinterval As String
Dim Askcount As Integer
Public ScreenName As String
 

Private Sub cmdAdd_Click()
loadFlexGrid
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    On Error GoTo ErrTrap
    Dim Reports As ClsRepoerts
    Dim StrSQL As String

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_InstallmentMustPayed", True)
    Askcount = GetSetting(StrAppRegPath, "Setting", "count_InstallmentMustPayed", True)
    
    'StrSQL = "select * From QestNotReceipted where  DueDate<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
    ' StrSQL = StrSQL + " order by CusName,Transaction_ID,QeqtNum"

    StrSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Transactions.Posted, "
    StrSQL = StrSQL + "  dbo.TblUsers.UserName , dbo.Transactions.order_no,  dbo.Transactions.PostedDate"
    StrSQL = StrSQL + " FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL + " dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL + " dbo.TblUsers ON dbo.Transactions.Posted = dbo.TblUsers.UserID"
    StrSQL = StrSQL + " WHERE     (NOT (dbo.Transactions.Posted IS NULL)) AND (dbo.Transactions.order_no NOT IN"
    StrSQL = StrSQL + " (SELECT     order_no"
    StrSQL = StrSQL + " From Transactions"
    StrSQL = StrSQL + " WHERE     Transaction_Type = 21 AND NOT (order_no IS NULL))) AND (dbo.Transactions.Transaction_Type = 17)"
    StrSQL = StrSQL + " ORDER BY dbo.Transactions.PostedDate"
   
    Set Reports = New ClsRepoerts
    Reports.AccreditOrders StrSQL, , LblCaption.Caption
    Exit Sub
ErrTrap:
End Sub

 

Private Sub Fg_CellButtonClick(ByVal Row As Long, _
                               ByVal Col As Long)
Dim currrentScreenName As String
Dim newapprovalno As Double
  
    With Me.Fg
 currrentScreenName = (.TextMatrix(Row, .ColIndex("ScreenName")))
        Select Case .ColKey(Col)

            Case "Show"

           

                    If currrentScreenName = "FrmPO" Then
                        FrmPO.show
                        FrmPO.Retrive val(.TextMatrix(Row, .ColIndex("Transaction_ID")))
                    ElseIf currrentScreenName = "FrmPO1" Then
                        FrmPO1.show
                        FrmPO1.Retrive val(.TextMatrix(Row, .ColIndex("Transaction_ID")))
                        
                      ElseIf currrentScreenName = "FrmPO2" Then
                        FrmPO2.show
                        FrmPO2.Retrive val(.TextMatrix(Row, .ColIndex("Transaction_ID")))
                   ElseIf currrentScreenName = "FrmPO3" Then
                        FrmPO3.show
                        FrmPO3.Retrive val(.TextMatrix(Row, .ColIndex("Transaction_ID")))
                     ElseIf currrentScreenName = "FrmEmpsAdvanceRequest" Then
                       FrmEmpsAdvanceRequest.show
                        FrmEmpsAdvanceRequest.Retrive val(.TextMatrix(Row, .ColIndex("Transaction_ID")))
                        
                   End If
                   

            Case "Approve"
                Dim sql As String
                Dim x As Integer
                If SystemOptions.UserInterface = ArabicInterface Then
                        x = MsgBox(" √ﬂÌœ «·«⁄ „«œ", vbExclamation + vbYesNoCancel)
                Else
                        x = MsgBox(" Confirm Approval", vbExclamation + vbYesNoCancel)
                End If
                If x = vbYes Then
                sql = "update ApprovalData set  Currcursor=null, Remarks='" & (.TextMatrix(Row, .ColIndex("Remarks"))) & "',ApprovDate=getdate()  where id=" & val(.TextMatrix(Row, .ColIndex("id")))
                Cn.Execute sql
                newapprovalno = GetCurrentApprovalForTransactions(val(.TextMatrix(Row, .ColIndex("Transaction_ID"))), currrentScreenName)
                If newapprovalno > 0 Then
              sql = "update ApprovalData set   SendTime=getdate() , Currcursor=1 , FromUser='" & user_name & "'    where id=" & newapprovalno
                Cn.Execute sql
                End If
                
              If CheckLastApprovLevel(currrentScreenName, val(.TextMatrix(Row, .ColIndex("Transaction_ID")))) = 0 Then
                          If currrentScreenName = "FrmEmpsAdvanceRequest" Then
                                    
                                      sql = "update TblEmpAdvanceRequest set   Approved=1      where AdvanceID=" & val(.TextMatrix(Row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                      
                            Else
                                        sql = "update Transactions set   Approved=1      where Transaction_ID=" & val(.TextMatrix(Row, .ColIndex("Transaction_ID")))
                                          Cn.Execute sql
            
                            End If
                
              End If
              
              
              
                loadFlexGrid
                End If
                
        End Select

    End With

End Sub

Public Function loadFlexGrid()
    Dim My_SQL As String
    Dim RowNum As Integer
    Dim ReCount As Integer
    Dim RsTemp As New ADODB.Recordset
Dim screenName1 As String
Dim timecatlog As Double
Dim hours As Integer
Dim minutes As Integer
Dim LateType As String
    If SystemOptions.SysDataBaseType = AccessDataBase Then
    
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
     
        Dim StrSQL As String

        StrSQL = "SELECT     TOP 100 PERCENT    dbo.ApprovalData.ExpectedtimeTime , dbo.ApprovalData.SendTime ,dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, dbo.ApprovalData.currorder, "
        StrSQL = StrSQL + "  dbo.ApprovalData.FromUser,  dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks, dbo.TblEmployee.Emp_Code,"
        StrSQL = StrSQL + "   dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TbLLevels.Name, dbo.TbLLevels.Namee, dbo.Screens.ScreenCaption,"
        StrSQL = StrSQL + "   dbo.Screens.ScreenTitleEng, dbo.ApprovalData.Currcursor, dbo.ApprovalData.id AS searchid"
        StrSQL = StrSQL + "  ,  dbo.ApprovalData.NoteSerial , dbo.ApprovalData.Transaction_Date  FROM         dbo.ApprovalData INNER JOIN"
        StrSQL = StrSQL + "    dbo.TblEmployee ON dbo.ApprovalData.EmpID = dbo.TblEmployee.Emp_ID INNER JOIN"
        StrSQL = StrSQL + "    dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
       StrSQL = StrSQL + "   dbo.Screens ON dbo.ApprovalData.ScreenName = dbo.Screens.ScreenName"
        StrSQL = StrSQL + "   Where (dbo.ApprovalData.Currcursor = 1) And (dbo.ApprovalData.EmpID = " & user_id & ")"
      If ScreenName <> "" Then
      StrSQL = StrSQL & "  AND (ApprovalData.ScreenName = N'" & ScreenName & "') "
      End If
      
        StrSQL = StrSQL + "   ORDER BY dbo.ApprovalData.currorder"
         

    End If

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        With Fg
            .Rows = .FixedRows

            For ReCount = 1 To RsTemp.RecordCount
                .Rows = .Rows + 1
                RowNum = .Rows - 1

                .TextMatrix(RowNum, .ColIndex("Ser")) = ReCount
                'Transaction_ID
                screenName1 = IIf(IsNull(RsTemp("ScreenName").value), "", RsTemp("ScreenName").value)
               .TextMatrix(RowNum, .ColIndex("id")) = IIf(IsNull(RsTemp("searchid").value), "", RsTemp("searchid").value)
                .TextMatrix(RowNum, .ColIndex("Transaction_ID")) = IIf(IsNull(RsTemp("Transaction_ID").value), "", RsTemp("Transaction_ID").value)
                 .TextMatrix(RowNum, .ColIndex("ScreenName")) = screenName1 ' IIf(IsNull(RsTemp("ScreenName").value), "", RsTemp("ScreenName").value)
                .TextMatrix(RowNum, .ColIndex("NoteSerial")) = IIf(IsNull(RsTemp("NoteSerial").value), "", RsTemp("NoteSerial").value)
            
                .TextMatrix(RowNum, .ColIndex("Transaction_Date")) = IIf(IsNull(RsTemp("Transaction_Date").value), "", Format(RsTemp("Transaction_Date").value, "YYYY/MM/DD"))
                  
             .TextMatrix(RowNum, .ColIndex("FromUser")) = IIf(IsNull(RsTemp("FromUser").value), "", RsTemp("FromUser").value)
             .TextMatrix(RowNum, .ColIndex("SendTime")) = IIf(IsNull(RsTemp("SendTime").value), "", Format(RsTemp("SendTime").value, "YYYY/MM/DD  HH:MM AM/PM"))
             .TextMatrix(RowNum, .ColIndex("ExpectedtimeTime")) = IIf(IsNull(RsTemp("ExpectedtimeTime").value), "", RsTemp("ExpectedtimeTime").value)
        '     .TextMatrix(RowNum, .ColIndex("LateType")) = GetTimeforTransaction(screenName1, timecatlog)
             If SystemOptions.UserInterface = ArabicInterface Then
             .TextMatrix(RowNum, .ColIndex("show")) = "«÷€ÿ ··⁄—÷ "
             .TextMatrix(RowNum, .ColIndex("Approve")) = "«÷€ÿ ··≈⁄ „«œ "
             Else
             .TextMatrix(RowNum, .ColIndex("show")) = "Show"
             .TextMatrix(RowNum, .ColIndex("Approve")) = "Approve"
             End If
             
             



             Dim timediff As String
             timediff = DateDiff("N", .TextMatrix(RowNum, .ColIndex("ExpectedtimeTime")), Now)
             
             '"YYYY/MM/DD  HH:MM AM/PM"
             .TextMatrix(RowNum, .ColIndex("ExpectedtimeTime")) = Format(.TextMatrix(RowNum, .ColIndex("ExpectedtimeTime")), "YYYY/MM/DD  HH:MM AM/PM")
If timediff > 0 Then


hours = timediff \ 60
minutes = timediff - (hours * 60)
LateType = hours & ":" & minutes
  .TextMatrix(RowNum, .ColIndex("LateType")) = LateType
      .Cell(flexcpBackColor, RowNum, 0, RowNum, 16) = &HFF&
 Else
 LateType = ""
 End If
 ' If timecatlog = 0 Then
 '           If SystemOptions.UserInterface = ArabicInterface Then
 '           .TextMatrix(RowNum, .ColIndex("LateType")) = "œﬁÌﬁ…"
 '           Else
 '           .TextMatrix(RowNum, .ColIndex("LateType")) = "Minute"
 '           End If
 '
 '
 ' ElseIf timecatlog = 1 Then
 '             If SystemOptions.UserInterface = ArabicInterface Then
 '           .TextMatrix(RowNum, .ColIndex("LateType")) = "”«⁄Â"
 '           Else
 '           .TextMatrix(RowNum, .ColIndex("LateType")) = "Hour"
 '           End If
 ' ElseIf timecatlog = 2 Then
 '             If SystemOptions.UserInterface = ArabicInterface Then
 '           .TextMatrix(RowNum, .ColIndex("LateType")) = "ÌÊ„"
 '           Else
 '           .TextMatrix(RowNum, .ColIndex("LateType")) = "Day/s"
 '           End If
 ' End If
   
             
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(RowNum, .ColIndex("ScreenCaption")) = IIf(IsNull(RsTemp("ScreenCaption").value), "", RsTemp("ScreenCaption").value)
                    .TextMatrix(RowNum, .ColIndex("LevelName")) = IIf(IsNull(RsTemp("name").value), "", RsTemp("name").value)
                Else
                    .TextMatrix(RowNum, .ColIndex("ScreenCaption")) = IIf(IsNull(RsTemp("ScreenTitleEng").value), "", RsTemp("ScreenTitleEng").value)
                   .TextMatrix(RowNum, .ColIndex("LevelName")) = IIf(IsNull(RsTemp("namee").value), "", RsTemp("namee").value)
                End If
             
 
            
                .ColComboList(.ColIndex("Show")) = "..."
                .ColComboList(.ColIndex("Approve")) = "..."
             
                RsTemp.MoveNext
            Next ReCount

            .AutoSize 0, .Cols - 1, False
        End With
Else
   Fg.Rows = 1
    End If

End Function

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim BGround As New ClsBackGroundPic
    Dim BolShowRequest As Boolean

    FormPostion Me, GetPostion
    LoadIcons
    Fg.WallPaper = BGround.Picture
 '   BolShowRequest = GetSetting(StrAppRegPath, "View_Type", "InstallmentMustPayed", True)

    If BolShowRequest = True Then
        chkShow.value = Unchecked
    Else
        chkShow.value = Checked
    End If

    loadFlexGrid
'    Resize_Form Me, ReportSize

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Me.Caption = "Accredit Sales Order"
    LblCaption.Caption = Me.Caption
    chkShow.Caption = "Dont Show at Start"
    Label1.Caption = "Data Based in your System Date"
    Me.CmdExit.Caption = "Exit"
    Me.CmdPrint.Caption = "Print"

    With Me.Fg
        .TextMatrix(0, .ColIndex("order_no")) = "Order No"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Trans Date"
        .TextMatrix(0, .ColIndex("CusName")) = "Cus. Name"
        .TextMatrix(0, .ColIndex("PostedDate")) = "PostedDate"
        .TextMatrix(0, .ColIndex("UserName")) = "By User"
        .TextMatrix(0, .ColIndex("Convert")) = "Convert To Bill"

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If chkShow.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "InstallmentMustPayed", False
    Else
        SaveSetting StrAppRegPath, "View_Type", "InstallmentMustPayed", True
    End If

    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub

Private Sub LoadIcons()
    On Error GoTo ErrTrap

    With Fg
        .Cell(flexcpPicture, 0, .ColIndex("Name")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillIID")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("TransDate")) = mdifrmmain.ImgLstTree.ListImages("qty").Picture
        .Cell(flexcpPicture, 0, .ColIndex("QestNum")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("DueDate")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Value")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub LblCaption_Click()
    On Error GoTo ErrTrap

    If Me.WindowState = vbNormal Then
        Me.WindowState = vbMaximized
    Else
        Me.WindowState = vbNormal
    End If

    Exit Sub
ErrTrap:
End Sub

