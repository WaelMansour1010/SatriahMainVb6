VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmPermissionScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "’·«ÕÌ«  «·„” Œœ„Ì‰ ⁄·Ï «·‘«‘…"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   7650
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7860
      _cx             =   13864
      _cy             =   13494
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
      BorderWidth     =   1
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmPermissionScreen.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   690
         Index           =   2
         Left            =   15
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   15
         Width           =   7830
         _cx             =   13811
         _cy             =   1217
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
         AutoSizeChildren=   0
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
         Begin MSDataListLib.DataCombo DcboScreens 
            Height          =   315
            Left            =   4530
            TabIndex        =   4
            Top             =   330
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label LblIntScreenType 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            Height          =   285
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   330
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·‘«‘…"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   1
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   90
            Width           =   1995
         End
         Begin VB.Label LblScreenType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   330
            Width           =   1995
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·‘«‘…"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   0
            Left            =   4530
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   90
            Width           =   3255
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   600
         Index           =   1
         Left            =   15
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   7035
         Width           =   7830
         _cx             =   13811
         _cy             =   1058
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
         AutoSizeChildren=   0
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
         Begin MSComctlLib.ProgressBar PrgBar 
            Height          =   285
            Left            =   2190
            TabIndex        =   12
            Top             =   150
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.TextBox TxtModFlag 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   660
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   60
            Visible         =   0   'False
            Width           =   585
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   1110
            TabIndex        =   6
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   661
            ButtonStyle     =   1
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
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Cancel          =   -1  'True
            Height          =   375
            Index           =   0
            Left            =   30
            TabIndex        =   7
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   661
            ButtonStyle     =   1
            Caption         =   "≈·€«¡"
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
            DrawFocusRectangle=   0   'False
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Fg 
         Height          =   6300
         Left            =   15
         TabIndex        =   1
         Top             =   720
         Width           =   7830
         _cx             =   13811
         _cy             =   11112
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
         BackColorBkg    =   -2147483643
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
         Rows            =   3
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmPermissionScreen.frx":0081
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
         OutlineBar      =   1
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
         WordWrap        =   -1  'True
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
End
Attribute VB_Name = "FrmPermissionScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FormateGrid()

    Dim i As Integer

    With Me.Fg

        For i = .FixedRows To .Rows - 1

            If i Mod 2 = 0 Then
                .Cell(flexcpBackColor, i, .ColIndex("UserName"), i, .ColIndex("FullAccess")) = &HE2E9E9
            Else
                .Cell(flexcpBackColor, i, .ColIndex("UserName"), i, .ColIndex("FullAccess")) = vbWhite
            End If

        Next i

    End With

End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            Unload Me

        Case 1

            If Me.TxtModFlag.text = "R" Then
                Me.TxtModFlag.text = "E"
            Else
                SavePremissions
            End If

    End Select

End Sub

Private Sub DcboScreens_Change()
    LoadData
End Sub

Private Sub DcboScreens_Click(Area As Integer)
    LoadData
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
    Dim i As Integer
    Dim IntScreenType As Integer

    With Fg
        IntScreenType = val(Me.LblIntScreenType.Caption)

        Select Case .ColKey(Col)

            Case "FullAccess"

                If Row = .FixedRows - 1 Then

                    For i = Row To .Rows - 1

                        If Not .IsSubtotal(i) And IntScreenType < 5 Then
                            .Cell(flexcpChecked, i, .ColIndex("AddNew"), i, .ColIndex("FullAccess")) = .Cell(flexcpChecked, Row, Col)
                        ElseIf Not .IsSubtotal(i) And IntScreenType >= 5 Then
                            .Cell(flexcpChecked, i, .ColIndex("FullAccess")) = .Cell(flexcpChecked, Row, Col)
                        End If

                    Next i

                End If
            
                If IntScreenType < 5 Then
                    .Cell(flexcpChecked, Row, .ColIndex("AddNew"), Row, .ColIndex("Search")) = .Cell(flexcpChecked, Row, Col)
                Else
                    .Cell(flexcpChecked, Row, .ColIndex("FullAccess")) = .Cell(flexcpChecked, Row, Col)
                End If

            Case "NoAccess"

            Case "AddNew"

                If Row = .FixedRows - 1 Then

                    For i = Row To .Rows - 1

                        If Not .IsSubtotal(i) And IntScreenType < 5 Then
                            .Cell(flexcpChecked, i, .ColIndex("AddNew"), i, .ColIndex("AddNew")) = .Cell(flexcpChecked, Row, Col)
                        End If

                    Next i

                End If

            Case "Edit"

                If Row = .FixedRows - 1 Then

                    For i = Row To .Rows - 1

                        If Not .IsSubtotal(i) And IntScreenType < 5 Then
                            .Cell(flexcpChecked, i, .ColIndex("Edit"), i, .ColIndex("Edit")) = .Cell(flexcpChecked, Row, Col)
                        End If

                    Next i

                End If

            Case "Print"

                If Row = .FixedRows - 1 Then

                    For i = Row To .Rows - 1

                        If Not .IsSubtotal(i) And IntScreenType < 5 Then
                            .Cell(flexcpChecked, i, .ColIndex("Print"), i, .ColIndex("Print")) = .Cell(flexcpChecked, Row, Col)
                        End If

                    Next i

                End If

            Case "Delete"

                If Row = .FixedRows - 1 Then

                    For i = Row To .Rows - 1

                        If Not .IsSubtotal(i) And IntScreenType < 5 Then
                            .Cell(flexcpChecked, i, .ColIndex("Delete"), i, .ColIndex("Delete")) = .Cell(flexcpChecked, Row, Col)
                        End If

                    Next i

                End If

            Case "Search"

                If Row = .FixedRows - 1 Then

                    For i = Row To .Rows - 1

                        If Not .IsSubtotal(i) And IntScreenType < 5 Then
                            .Cell(flexcpChecked, i, .ColIndex("Search"), i, .ColIndex("Search")) = .Cell(flexcpChecked, Row, Col)
                        End If

                    Next i

                End If

        End Select

        CellCheck CInt(Row)
    End With

End Sub

Private Sub Fg_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)

    With Me.Fg

        If .Col = -1 Then Exit Sub
        If .Row <= 0 Then Exit Sub
        If .Col = .ColIndex("UserName") Then
            Cancel = True
            Exit Sub
        End If

    End With

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim GrdBack As ClsBackGroundPic
    Set Dcombos = New ClsDataCombos
    Dcombos.GetScreens Me.DcboScreens
    CenterForm Me

    FormPostion Me, GetPostion
    Set Me.Icon = mdifrmmain.ImgLstMenuIcons.ListImages("Users").ExtractIcon
    Set Me.Cmd(0).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").ExtractIcon
    Cmd(0).ButtonPositionImage = impRightOfText
    Cmd(1).ButtonPositionImage = impRightOfText
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        .Cell(flexcpPicture, 0, .ColIndex("AddNew"), 0) = mdifrmmain.ImgLstMenuIcons.ListImages("New").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Edit"), 0) = mdifrmmain.ImgLstMenuIcons.ListImages("Edit").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Delete"), 0) = mdifrmmain.ImgLstMenuIcons.ListImages("Delete").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Print"), 0) = mdifrmmain.ImgLstMenuIcons.ListImages("Print").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Search"), 0) = mdifrmmain.ImgLstMenuIcons.ListImages("Find").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("FullAccess"), 0) = mdifrmmain.ImgLstMenuIcons.ListImages("Tick").ExtractIcon
        .Cell(flexcpPicture, 1, .ColIndex("NoAccess"), 1) = mdifrmmain.ImgLstMenuIcons.ListImages("Stop").ExtractIcon

        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpPictureAlignment, 0, .ColIndex("AddNew"), 0, .ColIndex("FullAccess")) = flexAlignRightCenter
        Else
            .Cell(flexcpPictureAlignment, 0, .ColIndex("AddNew"), 0, .ColIndex("FullAccess")) = flexAlignLeftCenter
        End If

        .Cell(flexcpPictureAlignment, 1, .ColIndex("AddNew"), 1, .ColIndex("FullAccess")) = flexAlignCenterCenter
        .Cell(flexcpChecked, 1, .ColIndex("AddNew"), 1, .ColIndex("FullAccess")) = flexUnchecked
    
        .RowHeightMin = 300
        Set .WallPaper = GrdBack.Picture
        .FixedRows = 2
        .ExtendLastCol = True
        '.AutoSize 0, .Cols - 1, False
    End With

    Me.TxtModFlag.text = "R"
End Sub

Public Sub LoadData()
    Dim rs As ADODB.Recordset
    Dim StrSQL  As String
    Dim i As Integer
    Dim LngFindRow As Long

    If Me.DcboScreens.BoundText = "" Then
        Exit Sub
    End If

    StrSQL = "SELECT  UserID, UserName, UserType, IsActive"
    StrSQL = StrSQL + " From TblUsers Where IsActive=1 "
    StrSQL = StrSQL + " AND UserType <>0 "
    StrSQL = StrSQL + " Order BY  UserName "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Fg
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = .FixedRows + rs.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("UserID")) = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
                .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
                .Cell(flexcpChecked, i, .ColIndex("AddNew"), i, .ColIndex("FullAccess")) = flexUnchecked
                rs.MoveNext
            Next i

        End If

        .Cell(flexcpPictureAlignment, 1, .ColIndex("AddNew"), Fg.Rows - 1, .ColIndex("FullAccess")) = flexAlignCenterCenter
    End With

    StrSQL = "SELECT dbo.TblUsers.UserID, dbo.TblUsers.UserName, dbo.TblUsers.IsActive, dbo.Screens.ScreenName," & "dbo.Screens.ScreenCaption,dbo.Screens.ScreenTitleEng, dbo.Screens.ScreenType, dbo.Screens.ScreenOrder," & "dbo.Screens.ScreenVisible, dbo.ScreenJuncUser.CanAdd,dbo.ScreenJuncUser.CanEdit, dbo.ScreenJuncUser.CanDelete," & "dbo.ScreenJuncUser.CanPrint, dbo.ScreenJuncUser.CanSearch,dbo.ScreenJuncUser.FullAccess"
    StrSQL = StrSQL + " FROM         dbo.TblUsers LEFT OUTER JOIN dbo.ScreenJuncUser ON dbo.TblUsers.UserID = " & "dbo.ScreenJuncUser.User_ID INNER JOIN dbo.Screens ON dbo.ScreenJuncUser.ScreenName = dbo.Screens.ScreenName "
    StrSQL = StrSQL + " Where dbo.Screens.ScreenName='" & Me.DcboScreens.BoundText & "'"
    StrSQL = StrSQL + " AND dbo.TblUsers.IsActive=1"
    StrSQL = StrSQL + " Order BY dbo.TblUsers.UserName"

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Fg

        If Not (rs.BOF Or rs.EOF) Then
            Me.LblIntScreenType.Caption = rs("ScreenType").value
            WriteScreenType IIf(IsNull(rs("ScreenType").value), 0, rs("ScreenType").value)
        
            For i = 1 To rs.RecordCount
                LngFindRow = .FindRow(rs("UserName").value, .FixedRows, .ColIndex("UserName"), False, True)

                If LngFindRow <> -1 Then
                    If rs("ScreenType").value >= 5 Then
                        .Cell(flexcpChecked, LngFindRow, .ColIndex("FullAccess")) = IIf(rs("CanAdd").value = False, flexUnchecked, flexChecked)
                    Else
                        .Cell(flexcpChecked, LngFindRow, .ColIndex("AddNew")) = IIf(rs("CanAdd").value = False, flexUnchecked, flexChecked)
                        .Cell(flexcpChecked, LngFindRow, .ColIndex("Edit")) = IIf(rs("CanEdit").value = False, flexUnchecked, flexChecked)
                        .Cell(flexcpChecked, LngFindRow, .ColIndex("Delete")) = IIf(rs("CanDelete").value = False, flexUnchecked, flexChecked)
                        .Cell(flexcpChecked, LngFindRow, .ColIndex("Print")) = IIf(rs("CanPrint").value = False, flexUnchecked, flexChecked)
                        .Cell(flexcpChecked, LngFindRow, .ColIndex("Search")) = IIf(rs("CanSearch").value = False, flexUnchecked, flexChecked)
                        CellCheck CInt(LngFindRow)
                    End If
                End If

                rs.MoveNext
            Next i

        End If

        .Cell(flexcpPictureAlignment, 1, .ColIndex("AddNew"), Fg.Rows - 1, .ColIndex("FullAccess")) = flexAlignCenterCenter
    End With

    FormateGrid
End Sub

Private Sub CellCheck(LngRow As Integer)
    Dim IntScreenType As Integer

    With Fg
        IntScreenType = val(Me.LblIntScreenType.Caption)

        If IntScreenType <> 5 And IntScreenType <> 6 And IntScreenType <> 7 Then
            If .Cell(flexcpChecked, LngRow, .ColIndex("AddNew")) = flexChecked And .Cell(flexcpChecked, LngRow, .ColIndex("Edit")) = flexChecked And .Cell(flexcpChecked, LngRow, .ColIndex("Delete")) = flexChecked And .Cell(flexcpChecked, LngRow, .ColIndex("Print")) = flexChecked And .Cell(flexcpChecked, LngRow, .ColIndex("Search")) = flexChecked Then
                .Cell(flexcpChecked, LngRow, .ColIndex("FullAccess")) = flexChecked
            Else
                .Cell(flexcpChecked, LngRow, .ColIndex("FullAccess")) = flexUnchecked
            End If
        End If

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub WriteScreenType(IntScreenType As Integer)
    Dim StrTemp As String

    If IntScreenType = 1 Then
        StrTemp = "»Ì«‰«  √”«”Ì…"
    ElseIf IntScreenType = 2 Then
        StrTemp = "„⁄«„·«  «· Ã«—Ì…"
    ElseIf IntScreenType = 3 Then
        StrTemp = "„⁄«„·«  „«·Ì…"
    ElseIf IntScreenType = 4 Then
        StrTemp = "‘∆Ê‰ «·„ÊŸ›Ì‰"
    ElseIf IntScreenType = 5 Then
        StrTemp = " ﬁ«—Ì—"
    ElseIf IntScreenType = 6 Then
        StrTemp = "≈” ⁄·«„« "
    ElseIf IntScreenType = 7 Then
        StrTemp = "√œÊ« "
    Else
        Me.LblIntScreenType.Caption = ""
        Me.LblScreenType.Caption = ""
        Exit Sub
    End If

    Me.LblIntScreenType.Caption = IntScreenType
    Me.LblScreenType.Caption = StrTemp
End Sub

Private Sub TxtModFlag_Change()

    Select Case Me.TxtModFlag.text

        Case "R"
            Me.DcboScreens.Enabled = True
            Me.Fg.Editable = flexEDNone
            Me.Cmd(1).Caption = " ⁄œÌ·"
            Me.PrgBar.Visible = False
            Set Me.Cmd(1).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Edit").ExtractIcon

        Case "E"
            Me.DcboScreens.Enabled = False
            Me.Fg.Editable = flexEDKbdMouse
            Me.Cmd(1).Caption = "Õ›Ÿ"
            Me.PrgBar.Visible = True
            Set Me.Cmd(1).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Save").ExtractIcon
    End Select

End Sub

Private Sub SavePremissions()
    Dim IntRes As Integer
    Dim i  As Integer
    Dim StrSQL As String
    Dim TransBegine As Boolean
    Dim BolAdd As Boolean, BolEdit As Boolean, BolDelete As Boolean
    Dim BolPrint As Boolean, BolSearch  As Boolean, BolFullAccess As Boolean
    Dim Msg As String
    Dim IntFullPre As Integer
    Dim IntScreenType As Integer
    Dim rs As ADODB.Recordset

    On Error GoTo ErrTrap
    Cn.BeginTrans
    TransBegine = True

    StrSQL = "Delete From ScreenJuncUser "
    StrSQL = StrSQL + " Where ScreenName='" & Me.DcboScreens.BoundText & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
    Set rs = New ADODB.Recordset
    rs.Open "ScreenJuncUser", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Fg

        For i = 2 To .Rows - 1

            DoEvents
            Me.PrgBar.Visible = True
            Me.PrgBar.Max = .Rows - 1
            Me.PrgBar.value = i
       
            IntScreenType = val(Me.LblIntScreenType.Caption)

            If IntScreenType <> 5 And IntScreenType <> 6 And IntScreenType <> 7 Then
                BolAdd = IIf(.Cell(flexcpChecked, i, .ColIndex("AddNew")) = flexChecked, True, False)
                BolEdit = IIf(.Cell(flexcpChecked, i, .ColIndex("Edit")) = flexChecked, True, False)
                BolDelete = IIf(.Cell(flexcpChecked, i, .ColIndex("Delete")) = flexChecked, True, False)
                BolPrint = IIf(.Cell(flexcpChecked, i, .ColIndex("Print")) = flexChecked, True, False)
                BolSearch = IIf(.Cell(flexcpChecked, i, .ColIndex("Search")) = flexChecked, True, False)
            ElseIf IntScreenType = 5 Or IntScreenType = 6 Or IntScreenType = 7 Then
                BolAdd = IIf(.Cell(flexcpChecked, i, .ColIndex("FullAccess")) = flexChecked, True, False)
                BolEdit = IIf(.Cell(flexcpChecked, i, .ColIndex("FullAccess")) = flexChecked, True, False)
                BolDelete = IIf(.Cell(flexcpChecked, i, .ColIndex("FullAccess")) = flexChecked, True, False)
                BolPrint = IIf(.Cell(flexcpChecked, i, .ColIndex("FullAccess")) = flexChecked, True, False)
                BolSearch = IIf(.Cell(flexcpChecked, i, .ColIndex("FullAccess")) = flexChecked, True, False)
            End If

            BolFullAccess = IIf(.Cell(flexcpChecked, i, .ColIndex("FullAccess")) = flexChecked, True, False)
            rs.AddNew
            rs("ScreenName").value = Me.DcboScreens.BoundText
            rs("User_ID").value = val(.TextMatrix(i, .ColIndex("UserID")))

            If SystemOptions.SysDataBaseType = AccessDataBase Then
                rs("CanAdd").value = BolAdd
                rs("CanEdit").value = BolEdit
                rs("CanDelete").value = BolDelete
                rs("CanPrint").value = BolPrint
                rs("CanSearch").value = BolSearch
                rs("FullAccess").value = BolFullAccess
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                rs("CanAdd").value = IIf(BolAdd = True, 1, 0)
                rs("CanEdit").value = IIf(BolEdit = True, 1, 0)
                rs("CanDelete").value = IIf(BolDelete = True, 1, 0)
                rs("CanPrint").value = IIf(BolPrint = True, 1, 0)
                rs("CanSearch").value = IIf(BolSearch = True, 1, 0)
                rs("FullAccess").value = IIf(BolFullAccess = True, 1, 0)
            End If

            rs.update

            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = "Update TblUsers Set IsActive=True,FullPremis=2 Where TblUsers.UserID=" & val(.TextMatrix(i, .ColIndex("UserID"))) & ""
            Else
                StrSQL = "Update TblUsers Set IsActive=1,FullPremis=2 Where TblUsers.UserID=" & val(.TextMatrix(i, .ColIndex("UserID"))) & ""
            End If

            Cn.Execute StrSQL, , adExecuteNoRecords

            DoEvents
            Me.PrgBar.value = i

            DoEvents
        Next i

    End With

Exit_Sub:
    Msg = " „  ⁄„·Ì… «·Õ›Ÿ...!!!"
    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Me.PrgBar.value = 0
    Me.PrgBar.Visible = False
    Cn.CommitTrans
    TransBegine = False
    Me.TxtModFlag.text = "R"
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«... ÕœÀ Œÿ« √À‰«¡ Õ›Ÿ «·’·«ÕÌ« ...!!!"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

    If TransBegine Then
        Cn.RollbackTrans
    End If

End Sub
