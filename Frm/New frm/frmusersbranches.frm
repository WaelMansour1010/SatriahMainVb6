VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmUsersBranches 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ÕœÌœ «·„” Œœ„Ì‰ ›Ì «·›—Ê⁄"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
   Icon            =   "frmusersbranches.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7275
      _cx             =   12832
      _cy             =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   " ÕœÌœ «·„” Œœ„Ì‰ ›Ì «·›—Ê⁄"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   2
      ChildSpacing    =   1
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   7
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
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   2
      Left            =   825
      TabIndex        =   2
      Top             =   5550
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "Õ›Ÿ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   3
      Top             =   5550
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   5340
      Index           =   3
      Left            =   -120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   7425
      _cx             =   13097
      _cy             =   9419
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
      BorderWidth     =   6
      ChildSpacing    =   4
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
      Begin VB.Frame Frame10 
         Caption         =   "«”„«¡ «·„” Œœ„Ì‰"
         Height          =   5055
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   150
         Width           =   7335
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   4020
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   7080
            _cx             =   12488
            _cy             =   7091
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
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmusersbranches.frx":000C
            ScrollTrack     =   0   'False
            ScrollBars      =   2
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
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Tag             =   "Delete Row"
            Top             =   4560
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
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
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmusersbranches.frx":01D2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄·Ìﬁ:"
         Height          =   150
         Index           =   16
         Left            =   975
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Tag             =   "22"
         Top             =   255
         Width           =   390
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Code"
         Height          =   225
         Index           =   8
         Left            =   975
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   3270
         Width           =   180
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Height          =   510
         Index           =   4
         Left            =   675
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   4365
         Width           =   690
      End
   End
End
Attribute VB_Name = "FrmUsersBranches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public branch_id As Integer
Public Row As Integer
Public branches As String

Private Sub Cmd_Click(Index As Integer)
Select Case Index

Case 2
SaveData
Unload Me
Case 6
Unload Me

End Select

End Sub

Function SaveData() As String
Dim str As String
str = ""
    With Me.VSFlexGrid1

        For I = 1 To .Rows - 1

            If .TextMatrix(I, .ColIndex("UserID")) <> "" Then
         
         str = str & .TextMatrix(I, .ColIndex("UserID")) & ","
            End If
            
            '
        Next I

    End With
    If Len(str) > 1 Then str = Mid(str, 1, Len(str) - 1)
    SaveData = str
   With FrmBranchesData.Grid
     .TextMatrix(Me.Row, .ColIndex("Users")) = str
     .TextMatrix(Me.Row, .ColIndex("Updated")) = 1
     
    End With
    
End Function

Function LoadData(BranchId1 As Integer)
    Dim RsDev    As New ADODB.Recordset
'    BranchId = 0
 '   If BranchId = 0 Then
    
 '      StrSQL = " SELECT     dbo.TblUsers.UserName, dbo.TblUsers.UserID, dbo.TblEmployee.Fullcode, dbo.TblUsersBranches.BranchID"
'StrSQL = StrSQL & "  FROM         dbo.TblUsersBranches INNER JOIN"
'StrSQL = StrSQL & "  dbo.TblUsers ON dbo.TblUsersBranches.userid = dbo.TblUsers.UserID LEFT OUTER JOIN"
'StrSQL = StrSQL & "   dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
'StrSQL = StrSQL & "  WHERE      dbo.TblUsersBranches.BranchID in(" & branches & ") "
    
'    Else
'   StrSQL = " SELECT     dbo.TblUsers.UserName, dbo.TblUsers.UserID, dbo.TblEmployee.Fullcode, dbo.TblUsersBranches.BranchID"
'StrSQL = StrSQL & "  FROM         dbo.TblUsersBranches INNER JOIN"
'StrSQL = StrSQL & "  dbo.TblUsers ON dbo.TblUsersBranches.userid = dbo.TblUsers.UserID LEFT OUTER JOIN"
'StrSQL = StrSQL & "   dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
'StrSQL = StrSQL & "  WHERE     (dbo.TblUsersBranches.BranchID = " & BranchId & ") "
'
   
'   End If
   
StrSQL = " SELECT     dbo.TblUsers.UserName, dbo.TblUsers.UserID, dbo.TblEmployee.Fullcode"
StrSQL = StrSQL & "  FROM         dbo.TblUsers LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
   
 StrSQL = StrSQL & "  WHERE      dbo.TblUsers.UserID in(" & branches & ")"
   
   If branch_id <> 0 Then
 StrSQL = " SELECT     TOP 100 PERCENT dbo.TblUsersBranches.userid, dbo.TblEmployee.Fullcode, dbo.TblUsers.UserName"
StrSQL = StrSQL & "  FROM         dbo.TblUsersBranches LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblUsers ON dbo.TblUsersBranches.userid = dbo.TblUsers.UserID LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & "   Where (dbo.TblUsersBranches.BranchId = " & branch_id & ")"
StrSQL = StrSQL & " ORDER BY dbo.TblUsersBranches.id"

End If

    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For I = .FixedRows To .Rows - 1
 
                .TextMatrix(I, .ColIndex("UserID")) = IIf(IsNull(RsDev("UserID").value), "", RsDev("UserID").value)
            
                .TextMatrix(I, .ColIndex("UserName")) = IIf(IsNull(RsDev("UserName").value), "", RsDev("UserName").value)
                         
                .TextMatrix(I, .ColIndex("Fullcode")) = IIf(IsNull(RsDev("Fullcode").value), "", RsDev("Fullcode").value)
            
 
                RsDev.MoveNext
            Next I
 .Rows = .Rows + 1
        End With

    End If
 RsDev.Close
 
 
 
 Set RsDev = Nothing
    ReLineGrid

End Function
Private Sub Form_Load()
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
      '  ChangeLang
    End If

    
   ' Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
   ' Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
   ' Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
  '  Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
 
 LoadData (0)
    Exit Sub
ErrTrap:
End Sub

Private Sub CmdRemove_Click()
    Dim x As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        x = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If x = vbNo Then Exit Sub
     
    If VSFlexGrid1.Rows > 1 Then
        If VSFlexGrid1.Rows = 2 Then
            Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.VSFlexGrid1.Rows > 1 Then
                If Me.VSFlexGrid1.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                    Me.VSFlexGrid1.RemoveItem (Me.VSFlexGrid1.Row)
                End If
            End If
        End If
    End If
            
    ReLineGrid

End Sub

 Private Sub ReLineGrid()
 
 End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid1

        Select Case .ColKey(Col)
 
            Case "UserName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("UserID"), False, True)
                .TextMatrix(Row, .ColIndex("UserID")) = StrAccountCode
             
          
  StrSQL = "SELECT     TOP 100 PERCENT dbo.TblUsers.UserName, dbo.TblEmployee.Fullcode, dbo.TblUsers.UserID"
StrSQL = StrSQL & "  FROM         dbo.TblUsersStores RIGHT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblUsers ON dbo.TblUsersStores.userid = dbo.TblUsers.UserID LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & "  Where (dbo.TblUsers.UserID = " & val(StrAccountCode) & ")"
StrSQL = StrSQL & "  ORDER BY dbo.TblUsersStores.id"
                    
                    Set rs = Nothing
                
                    If StrAccountCode <> "" Then
                            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                            If Not (rs.BOF Or rs.EOF) Then
                                     .TextMatrix(Row, .ColIndex("fullcode")) = _
                                    IIf(IsNull(rs("fullcode").value), "", rs("fullcode").value)
                             
                            End If
                    End If
            
                 
         
        End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid1

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
        Select Case .ColKey(Col)
            
            Case "name"
                Exit Sub
        End Select

    End With

    VSFlexGrid1.ComboList = ""

End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid1

        Select Case .ColKey(Col)

            Case "UserName"
                StrSQL = "select * from TblUsers"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "UserName", "UserID")
                Else
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "UserName", "UserID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub


