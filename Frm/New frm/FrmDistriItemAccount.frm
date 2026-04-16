VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmDistriItemAccount 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "‘«‘…  Ê“Ì⁄ «·„’—Ê›«  ⁄·Ï «·«’‰«›"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12120
   Icon            =   "FrmDistriItemAccount.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2745
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   12075
      _cx             =   21299
      _cy             =   4842
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
      Rows            =   1
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmDistriItemAccount.frx":038A
      ScrollTrack     =   -1  'True
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
      Editable        =   0
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   1
      Top             =   4080
      Width           =   765
      _ExtentX        =   1349
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
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   810
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "„”Õ"
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
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   30
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   ""
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
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   825
      Index           =   13
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   12135
      _cx             =   21405
      _cy             =   1455
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
      Begin VB.TextBox TxtAccountCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10560
         TabIndex        =   18
         Top             =   315
         Width           =   1455
      End
      Begin VB.ComboBox DcbTypevalue 
         Height          =   315
         ItemData        =   "FrmDistriItemAccount.frx":04F4
         Left            =   5550
         List            =   "FrmDistriItemAccount.frx":04F6
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   315
         Width           =   1185
      End
      Begin VB.TextBox TxtValue 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4440
         TabIndex        =   16
         Top             =   315
         Width           =   975
      End
      Begin VB.TextBox TxtRemark 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2145
         TabIndex        =   15
         Top             =   315
         Width           =   2160
      End
      Begin MSDataListLib.DataCombo DcbAccount 
         Height          =   315
         Left            =   6885
         TabIndex        =   5
         Top             =   315
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   360
         Index           =   24
         Left            =   1020
         TabIndex        =   6
         Top             =   285
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   635
         Caption         =   "≈÷«›…"
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
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   360
         Index           =   25
         Left            =   135
         TabIndex        =   7
         Top             =   285
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   635
         Caption         =   "Õ–›"
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
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ«·Õ”«»"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   2
         Left            =   9720
         TabIndex        =   19
         Top             =   0
         Width           =   1890
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„·«ÕŸ« "
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
         Height          =   240
         Index           =   0
         Left            =   2310
         TabIndex        =   14
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·Õ”«»"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   1
         Left            =   7530
         TabIndex        =   12
         Top             =   0
         Width           =   1890
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   38
         Left            =   135
         TabIndex        =   11
         ToolTipText     =   "⁄œœ «·√’‰«› «·„ﬂÊ‰… ·Â–« «·’‰› «·„Ã„⁄"
         Top             =   30
         Width           =   255
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ "
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
         Height          =   240
         Index           =   39
         Left            =   990
         TabIndex        =   10
         ToolTipText     =   "⁄œœ «·√’‰«› «·„ﬂÊ‰… ·Â–« «·’‰› «·„Ã„⁄"
         Top             =   30
         Width           =   1380
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁÌ„Â"
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
         Height          =   240
         Index           =   41
         Left            =   3960
         TabIndex        =   9
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·ﬁÌ„Â"
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
         Height          =   240
         Index           =   51
         Left            =   5430
         TabIndex        =   8
         Top             =   0
         Width           =   1290
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  Ê“Ì⁄ «·„’—Ê›«  ⁄·Ï «·«’‰«›"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   -45
      TabIndex        =   13
      Top             =   0
      Width           =   12150
   End
End
Attribute VB_Name = "FrmDistriItemAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
    save
    Unload Me
' GetData
           
      '  Case 1
           ' clear_all Me
'Me.DtpDateFrom.value = ""
'Me.DtpDateTo.value = ""
      '      If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
      '      Else
               ' Me.lbl(0).Caption = "Search Results"
      '      End If

      '  Case 2
      '      Unload Me
       Case 24
       AddNewFgRowother
       Case 25
            DeleteFgRowAther
    End Select

End Sub
Sub save()
Dim str As String
Dim i As Integer
str = ""

With Me.Fg
For i = 1 To .Rows - 1
 If .TextMatrix(i, .ColIndex("Account_Name")) <> "" Then
 str = str & Trim(.TextMatrix(i, .ColIndex("Account_Code"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("TypeValue"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("Vlue"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("Remark"))) & "#"
 str = str & Trim("@")
  str = str & Chr(13)
  str = Trim(str)
 End If
Next
FrmDistriExpensItems.Fg.TextMatrix(FrmDistriExpensItems.LngRow, FrmDistriExpensItems.Fg.ColIndex("Account1")) = str

End With
End Sub

Private Sub DeleteFgRowAther()

    With Me.Fg

        If .Row = -1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        .RemoveItem .Row
        '.AutoSize 0, .Cols - 1, False
     Me.lbl(38).Caption = ModFgLib.GetItemsInFg(Fg, Fg.ColIndex("Account_Code"))
    End With

End Sub
Private Sub AddNewFgRowother()

    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long

    If Me.DcbAccount.BoundText = "" Then
        Msg = "  ÌÃ»  ÕœÌœ «”„ «·Õ”«»"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.DcbAccount.SetFocus
        MsgBox Me.DcbAccount.BoundText
        Exit Sub
    End If



    If val(Me.TxtValue.text) = 0 Then
        Msg = " ÌÃ» «œŒ«· «·ﬁÌ„Â «Ê «·‰”»Â "
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.TxtValue.SetFocus
        Exit Sub
    End If

 

    If val(Me.DcbTypevalue.ListIndex) = -1 Then
        Msg = " ÌÃ»  ÕœÌœ  ‰Ê⁄ «·ﬁÌ„Â"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.DcbTypevalue.SetFocus
        Exit Sub
    End If

   ' With Me.Fg
   '     LngFindRow = .FindRow(val(Me.Dcbiteem.BoundText), .FixedRows, .ColIndex("ItemID"), False, True)
'
'        If LngFindRow <> -1 Then
'            Msg = "Â–« «·’‰› „ÊÃÊœ ›⁄·«"
'            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            .SetFocus
'            Exit Sub
'        End If
''
 '   End With
 LngNewRow = ModFgLib.SetFgForNewRow(Fg, Fg.ColIndex("Account_Code"))

    With Me.Fg
    
    .TextMatrix(LngNewRow, .ColIndex("Serial")) = LngNewRow
    .TextMatrix(LngNewRow, .ColIndex("Account_code1")) = Me.TxtAccountCode.text
        .TextMatrix(LngNewRow, .ColIndex("Account_Code")) = Trim(Me.DcbAccount.BoundText)
        .TextMatrix(LngNewRow, .ColIndex("Account_Name")) = Me.DcbAccount.text
    
        .TextMatrix(LngNewRow, .ColIndex("TypeValue")) = Me.DcbTypevalue.ListIndex
        .TextMatrix(LngNewRow, .ColIndex("TypeValuename")) = Me.DcbTypevalue.text
        .TextMatrix(LngNewRow, .ColIndex("Vlue")) = val(Me.TxtValue.text)
        .TextMatrix(LngNewRow, .ColIndex("Remark")) = Me.TxtRemark.text
       
        '.AutoSize 0, .Cols - 1, False
    End With

    Me.lbl(38).Caption = ModFgLib.GetItemsInFg(Fg, Fg.ColIndex("Account_Code"))

    Me.DcbAccount.text = ""
    Me.TxtRemark.text = ""
    Me.TxtValue.text = ""
DcbAccount.SetFocus
    
End Sub



'Private Sub Fg_Click()
'
 
'   On Error GoTo ErrTrap
'  '  FrmModels.FindRec
'   FrmSystemUnites.FindRec val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id")))
'ErrTrap:
'End Sub

Sub Retrive(Optional id As Integer = 0, Optional IDDet As Integer = 0)

 Dim RsDetails As ADODB.Recordset
 Dim StrSQL As String
 Dim i As Integer

   Set RsDetails = New ADODB.Recordset
StrSQL = "SELECT     dbo.TblDistriExpensItemDet3.ID, dbo.TblDistriExpensItemDet3.Ind, dbo.TblDistriExpensItemDet3.IDDet, dbo.TblDistriExpensItemDet3.TypeValue, "
StrSQL = StrSQL & "                      dbo.TblDistriExpensItemDet3.Vlue, dbo.TblDistriExpensItemDet3.Remark, REPLACE(REPLACE(dbo.TblDistriExpensItemDet3.Account_Code, CHAR(10), ''), CHAR(13),"
StrSQL = StrSQL & "                      '') AS Account_Code1, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Serial"
StrSQL = StrSQL & " FROM         dbo.TblDistriExpensItemDet3 LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.ACCOUNTS ON REPLACE(REPLACE(dbo.TblDistriExpensItemDet3.Account_Code, CHAR(10), ''), CHAR(13), '') = dbo.ACCOUNTS.Account_Code"
StrSQL = StrSQL & " Where (dbo.TblDistriExpensItemDet3.ind = " & id & ") And (dbo.TblDistriExpensItemDet3.IDDet = " & IDDet & ")"

   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails.RecordCount > 0 Then
   With Me.Fg
   .Rows = .Rows + RsDetails.RecordCount
   RsDetails.MoveFirst
   For i = 1 To .Rows - 1
   .TextMatrix(i, .ColIndex("Serial")) = i
   .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsDetails("Account_Code1").value), "", RsDetails("Account_Code1").value)
   .TextMatrix(i, .ColIndex("Account_Code1")) = IIf(IsNull(RsDetails("Account_Serial").value), "", RsDetails("Account_Serial").value)
   .TextMatrix(i, .ColIndex("Remark")) = IIf(IsNull(RsDetails("Remark").value), "", RsDetails("Remark").value)
   .TextMatrix(i, .ColIndex("Vlue")) = IIf(IsNull(RsDetails("Vlue").value), "", RsDetails("Vlue").value)
    .TextMatrix(i, .ColIndex("TypeValue")) = IIf(IsNull(RsDetails("TypeValue").value), "", RsDetails("TypeValue").value)
   
    If SystemOptions.UserInterface = EnglishInterface Then
     If val(RsDetails("TypeValue").value) = 0 Then
    .TextMatrix(i, .ColIndex("TypeValuename")) = "Rate"
    Else
    .TextMatrix(i, .ColIndex("TypeValuename")) = "Value"
    End If
     .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(RsDetails("Account_NameEng").value), "", RsDetails("Account_NameEng").value)
          Else
    .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(RsDetails("Account_Name").value), "", RsDetails("Account_Name").value)
        If val(RsDetails("TypeValue").value) = 0 Then
    .TextMatrix(i, .ColIndex("TypeValuename")) = "‰”»Â"
    Else
    .TextMatrix(i, .ColIndex("TypeValuename")) = "ﬁÌ„Â"
    End If
    End If
    RsDetails.MoveNext
   Next i
  
   End With
    End If
   Me.lbl(38).Caption = ModFgLib.GetItemsInFg(Fg, Fg.ColIndex("Account_Code"))
End Sub



Private Sub DcbAccount_Change()
TxtAccountCode.text = GetACCOUNTSCode(Me.DcbAccount.BoundText, 1)
End Sub

Private Sub DcbAccount_Click(Area As Integer)
DcbAccount_Change
End Sub

Private Sub DcbAccount_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Load Account_search
        Account_search.case_id = 667788
        Account_search.show vbModal
        
 End If
End Sub

Private Sub DcbTypevalue_Change()
If val(Me.DcbTypevalue.ListIndex) <> -1 Then
lbl(41).Caption = DcbTypevalue.text
End If
End Sub

Private Sub DcbTypevalue_Click()
If val(Me.DcbTypevalue.ListIndex) <> -1 Then
lbl(41).Caption = DcbTypevalue.text
End If
End Sub

Private Sub DcbTypevalue_LostFocus()
If val(Me.DcbTypevalue.ListIndex) <> -1 Then
lbl(41).Caption = DcbTypevalue.text
End If
End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
 
   Dcombos.GetAccountingCodes Me.DcbAccount
  If SystemOptions.UserInterface = EnglishInterface Then
DcbTypevalue.AddItem "Rate"
DcbTypevalue.AddItem "Value"
Else
DcbTypevalue.AddItem "‰”»Â"
DcbTypevalue.AddItem "ﬁÌ„Â"
End If

    Set DCboSearch = New clsDCboSearch
   
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
If val(FrmDistriExpensItems.XPTxtID.text) <> 0 And val(FrmDistriExpensItems.Fg.TextMatrix(FrmDistriExpensItems.LngRow, FrmDistriExpensItems.Fg.ColIndex("id"))) <> 0 Then
 Retrive val(FrmDistriExpensItems.XPTxtID.text), val(FrmDistriExpensItems.Fg.TextMatrix(FrmDistriExpensItems.LngRow, FrmDistriExpensItems.Fg.ColIndex("id")))
 End If
    Set GrdBack = New ClsBackGroundPic

'    With Me.Fg
'        Set .WallPaper = GrdBack.Picture
'        .AutoSize 0, .Cols - 1, False
'    End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
'

Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Save"
    Cmd(2).Caption = "Exit"
    
  Me.Caption = "Distribution Expenses on Items"
  
Label5.Caption = Me.Caption
'Me.LblClientName.Caption = "ClientName"
'lbl(4).Caption = "From"
'lbl(3).Caption = "To"
Cmd(24).Caption = "Add"
Cmd(25).Caption = "Delete"

lbl(1).Caption = "Account Code"
lbl(1).Caption = "Account Name"
lbl(51).Caption = "Type Value"
lbl(41).Caption = "Value  "
lbl(0).Caption = "Remarks  "
lbl(39).Caption = "Count"
'Me.lbreg.Caption = "Date Registration"

     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("Account_Name")) = "AccountName"
        .TextMatrix(0, .ColIndex("TypeValuename")) = "TypeVale"
         .TextMatrix(0, .ColIndex("Vlue")) = "Value  "
        .TextMatrix(0, .ColIndex("Remark")) = "Remarks  "
       '.TextMatrix(0, .ColIndex("PlateNo")) = "PlateNo"
    End With
  '
End Sub







Private Sub TxtAccountCode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        If TxtAccountCode.text = "" Then
            Me.DcbAccount.BoundText = ""
        Else
            Me.DcbAccount.BoundText = GetACCOUNTSCode(Trim$(Me.TxtAccountCode.text))
        End If
    End If

End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtValue.text, 1)
End Sub
