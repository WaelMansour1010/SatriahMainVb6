VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Frm_SandSearch 
   BackColor       =   &H009E7163&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ﬁÌÊœ «·ÌÊ„Ì…"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6510
   Icon            =   "Frm_SandSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra 
      BackColor       =   &H00DF967A&
      Height          =   2175
      Index           =   1
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   3390
      Width           =   6495
      Begin VB.TextBox TxtSer 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   210
         Width           =   1485
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1245
         Index           =   1
         Left            =   2520
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   180
         Width           =   1575
         _cx             =   2778
         _cy             =   2196
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
         Version         =   800
         BackColor       =   14653050
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "ﬁÌ„… «·”‰œ"
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
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
         Begin VB.TextBox TxtValueFrom 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   90
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Tag             =   "r"
            Top             =   360
            Width           =   1005
         End
         Begin VB.TextBox TxtValueTo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   90
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Tag             =   "r"
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00DF967A&
            Caption         =   "„‰"
            Height          =   285
            Index           =   2
            Left            =   1110
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   420
            Width           =   435
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00DF967A&
            Caption         =   "≈·Ï"
            Height          =   285
            Index           =   3
            Left            =   1110
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   780
            Width           =   435
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1245
         Index           =   0
         Left            =   1290
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
         _cx             =   2143
         _cy             =   2196
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
         Version         =   800
         BackColor       =   14653050
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "Õ«·… «·”‰œ"
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
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
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00DF967A&
            Caption         =   "„—Õ· "
            Height          =   285
            Index           =   0
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   330
            Width           =   1125
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00DF967A&
            Caption         =   "€Ì— „—Õ·"
            Height          =   285
            Index           =   1
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   630
            Width           =   1125
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00DF967A&
            Caption         =   "«·ﬂ·"
            Height          =   255
            Index           =   2
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   930
            Value           =   -1  'True
            Width           =   1125
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DF967A&
         Caption         =   "‰Ê⁄ «·”‰œ"
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1350
         Visible         =   0   'False
         Width           =   675
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00DF967A&
            Caption         =   "«·ﬂ·"
            Height          =   195
            Index           =   5
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   930
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00DF967A&
            Caption         =   "ﬁ»÷"
            Height          =   195
            Index           =   4
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   600
            Width           =   825
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00DF967A&
            Caption         =   "’—›"
            Height          =   195
            Index           =   3
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.CommandButton Cmd 
         BackColor       =   &H00DF967A&
         Cancel          =   -1  'True
         Caption         =   "Œ—ÊÃ"
         Height          =   345
         Index           =   2
         Left            =   90
         MouseIcon       =   "Frm_SandSearch.frx":27A2
         MousePointer    =   99  'Custom
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Cmd 
         BackColor       =   &H00DF967A&
         Caption         =   "„”Õ"
         Height          =   345
         Index           =   0
         Left            =   90
         MouseIcon       =   "Frm_SandSearch.frx":2AAC
         MousePointer    =   99  'Custom
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   210
         Width           =   1125
      End
      Begin VB.CommandButton Cmd 
         BackColor       =   &H00DF967A&
         Caption         =   "»ÕÀ"
         Default         =   -1  'True
         Height          =   345
         Index           =   1
         Left            =   90
         MouseIcon       =   "Frm_SandSearch.frx":2DB6
         MousePointer    =   99  'Custom
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   570
         Width           =   1125
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00DF967A&
         Caption         =   " «—ÌŒ  Õ—Ì— «·”‰œ ›Ï «·› —…"
         Height          =   1365
         Index           =   0
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   570
         Width           =   2295
         Begin NourAccounting.NourHijriCal DHijriTO 
            Height          =   315
            Left            =   90
            TabIndex        =   31
            Top             =   870
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
         End
         Begin NourAccounting.NourHijriCal DHijriFrom 
            Height          =   315
            Left            =   90
            TabIndex        =   30
            Top             =   300
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
         End
         Begin MSComCtl2.DTPicker DtpFrom 
            Height          =   345
            Left            =   60
            TabIndex        =   5
            Top             =   270
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   60817411
            CurrentDate     =   37954
         End
         Begin MSComCtl2.DTPicker DtpTo 
            Height          =   345
            Left            =   60
            TabIndex        =   6
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   60817411
            CurrentDate     =   37954
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00DF967A&
            Caption         =   "„‰"
            Height          =   285
            Index           =   0
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   270
            Width           =   315
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00DF967A&
            Caption         =   "≈·Ï"
            Height          =   285
            Index           =   1
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   870
            Width           =   315
         End
      End
      Begin MSDataListLib.DataCombo DcboEmp 
         Height          =   315
         Left            =   1020
         TabIndex        =   2
         Top             =   1470
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12648447
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboUsers 
         Height          =   315
         Left            =   1020
         TabIndex        =   29
         Top             =   1800
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12648447
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Õ—— «·ﬁÌœ"
         Height          =   255
         Index           =   6
         Left            =   3180
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1830
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00DF967A&
         Caption         =   "„”·”·"
         Height          =   285
         Index           =   5
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ÊŸ›"
         Height          =   255
         Index           =   4
         Left            =   3180
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1500
         Width           =   915
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid FgResult 
      Height          =   2745
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   6495
      _cx             =   11456
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   0
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"Frm_SandSearch.frx":30C0
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
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   4
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   6510
      _cx             =   11483
      _cy             =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   800
      BackColor       =   14653050
      ForeColor       =   16777215
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "Frm_SandSearch.frx":3207
      Caption         =   "«·»ÕÀ ⁄‰ ﬁÌÊœ «·ÌÊ„Ì…"
      Align           =   1
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
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
      PicturePos      =   1
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
   End
End
Attribute VB_Name = "Frm_SandSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TTP As New clstooltip
Dim OldIndex As Integer
Dim Dcombos As New ClsCombos
Private Sub Cmd_Click(Index As Integer)
Select Case Index
    Case 0
        Opt(5).Value = True
        Opt(2).Value = True
        Me.DtpFrom.Value = Null
        Me.DtpTo.Value = Null
        Me.DcboEmp.BoundText = ""
        Me.TxtValueFrom.Text = ""
        Me.TxtValueTo.Text = ""
    Case 1
        GetData
    Case 2
        Unload Me
End Select
End Sub

Private Sub Cmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If OldIndex = Index Then Exit Sub
Cmd(OldIndex).FontBold = False
Cmd(Index).FontBold = True
OldIndex = Index
End Sub

Private Sub FgResult_Click()
With FgResult
    If .TextMatrix(.Row, .ColIndex("Note_ID")) <> "" Then
        Frm_SandSarf.Retrive .TextMatrix(.Row, .ColIndex("Note_ID"))
    End If
End With
End Sub
Private Sub Form_Load()
Dim GrdBck As New ClsBackGroundPic
Dim My_SQL As String
Me.DtpFrom.Value = Date
Me.DtpFrom.Value = Null
Me.DtpTo.Value = Date
Me.DtpTo.Value = Null
If SystemOptions.SysDate = vbCalGreg Then
    Me.DHijriFrom.Visible = False
    Me.DHijriTo.Visible = False
Else
    Me.DHijriFrom.Visible = True
    Me.DHijriTo.Visible = True
End If
With Me.FgResult
    Set .WallPaper = GrdBck.NotesSearchWallpaper
    .AutoSize 0, .Cols - 1, False
End With
Dcombos.Get_Employees Me.DcboEmp
Dcombos.Get_Users Me.DcboUsers
If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
End If
CenterForm Me
AddTip
FormPostion Me, GetPostion
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Dcombos = Nothing
FormPostion Me, SavePostion

End Sub


Private Sub Fra_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Cmd(OldIndex).FontBold = False
End Sub
Private Sub GetData()
Dim StrSQL As String
Dim Msg As String
Dim Rs As New ADODB.Recordset
Dim I  As Integer
Dim StrPost As String
Dim StrUnPost As String
Dim StrType As String
StrSQL = BluidSQL
Rs.Open StrSQL, cn, adOpenStatic, adLockReadOnly, adCmdText
For I = 0 To Rs.Fields.Count - 1
    Debug.Print Rs(I).Name
Next I
For I = 0 To FgResult.Cols - 1
    Debug.Print FgResult.ColKey(I)
Next I
If SystemOptions.UserInterface = ArabicInterface Then
    StrPost = "„—Õ·"
    StrUnPost = "€Ì— „—Õ·"
    StrType = "ﬁÌœ ÌÊ„Ì…"
Else
    StrPost = "Posted"
    StrUnPost = "Not Posted"
    StrType = "Journal"
End If

If Not (Rs.BOF Or Rs.EOF) Then
    If Rs.RecordCount > 0 Then
        With FgResult
            .Rows = .FixedRows + Rs.RecordCount
            For I = 1 To Rs.RecordCount
                .TextMatrix(I, .ColIndex("Note_ID")) = IIf(IsNull(Rs("Note_ID").Value), _
                "", Rs("Note_ID").Value)
                .TextMatrix(I, .ColIndex("Emp")) = IIf(IsNull(Rs("Employee_Name").Value), _
                "", Rs("Employee_Name").Value)
                .TextMatrix(I, .ColIndex("Type")) = IIf(IsNull(Rs("Note_Type").Value), _
                "", IIf(Rs("Note_Type").Value = 20, StrType, StrType))
                .TextMatrix(I, .ColIndex("SandDate")) = IIf(IsNull(Rs("Note_Date").Value), _
                "", Format(Rs("Note_Date").Value, "yyyy/M/d"))
                .TextMatrix(I, .ColIndex("Value")) = IIf(IsNull(Rs("Value").Value), _
                "", Rs("Value").Value)
                .TextMatrix(I, .ColIndex("Serial")) = IIf(IsNull(Rs("Chique_Serial_No").Value), _
                "", Rs("Chique_Serial_No").Value)
                If Rs("NotePosted").Value = True Then
                    .TextMatrix(I, .ColIndex("State")) = StrPost
                    .Cell(flexcpForeColor, I, 1, I, .Cols - 1) = vbRed
                Else
                    .TextMatrix(I, .ColIndex("State")) = StrUnPost
                    .Cell(flexcpForeColor, I, 1, I, .Cols - 1) = vbBlack
                End If
                .TextMatrix(I, .ColIndex("Issued_by")) = IIf(IsNull(Rs("UserName").Value), _
                "", Rs("UserName").Value)
                
                Rs.MoveNext
            Next I
            .AutoSize 0, .Cols - 1, False
        End With
    End If
Else
    GetMsgs 77, vbExclamation
    FgResult.clear flexClearScrollable, flexClearEverything
    FgResult.Rows = FgResult.FixedRows + 1
End If
End Sub

Private Function BluidSQL() As String
Dim Begine As Boolean
Dim StrSQL As String
Dim StrWhere As String
Dim StrEmpFiled As String
If SystemOptions.UserInterface = ArabicInterface Then
    StrEmpFiled = "Employee_Name"
Else
    StrEmpFiled = "Employee_NameEng"
End If

If SystemOptions.UserShowDataEmployees = ShowArabicData Then
    StrEmpFiled = "Employee_Name"
ElseIf SystemOptions.UserShowDataEmployees = ShowEnglishData Then
    StrEmpFiled = "Employee_NameEng"
Else
    StrEmpFiled = StrEmpFiled
End If
StrSQL = "SELECT Notes.*, EMPLOYEES." & StrEmpFiled & " as Employee_Name, USERS.UserName" & _
    " FROM USERS INNER JOIN (EMPLOYEES INNER JOIN Notes ON  " & _
    " EMPLOYEES.Employee_Code = Notes.Employee_ID) ON " & _
    " USERS.User_ID = Notes.Issued_BY " & _
    " Where NOTES.Transaction_Header_ID is Null "
If Me.Opt(3).Value = True Then
    StrWhere = StrWhere & " And (Notes.Note_Type='20') "
ElseIf Opt(4).Value = True Then
    StrWhere = StrWhere & " And (Notes.Note_Type='20') "
ElseIf Opt(5).Value = True Then
    StrWhere = StrWhere & " And (Notes.Note_Type='20' Or Notes.Note_Type='20')"
End If
If Me.DcboEmp.BoundText <> "" Then
    StrWhere = StrWhere & " And (Notes.Employee_ID='" & Me.DcboEmp.BoundText & "')"
End If
If Me.DcboUsers.BoundText <> "" Then
    StrWhere = StrWhere & " And (Notes.Issued_BY=" & Me.DcboUsers.BoundText & ")"
End If
If Opt(0).Value = True Then
    StrWhere = StrWhere & " And(Notes.NotePosted=True)"
ElseIf Opt(1).Value = True Then
    StrWhere = StrWhere & " And(Notes.NotePosted=False)"
ElseIf Opt(2).Value = True Then
    StrWhere = StrWhere
End If
If Val(Me.TxtValueFrom.Text) > 0 Then
    StrWhere = StrWhere & " And(Notes.Value >=" & Val(TxtValueFrom.Text) & ")"
End If
If Val(Me.TxtValueTo.Text) > 0 Then
    StrWhere = StrWhere & " And(Notes.Value <=" & Val(TxtValueTo.Text) & ")"
End If
Select Case SystemOptions.SysDate
    Case VbCalendar.vbCalGreg
        If Not IsNull(Me.DtpFrom.Value) Then
            StrWhere = StrWhere & " And(Notes.Note_Date >=#" & SQLDate(Me.DtpFrom.Value) & "#)"
        End If
        If Not IsNull(Me.DtpTo.Value) Then
             StrWhere = StrWhere & " And(Notes.Note_Date <=#" & SQLDate(Me.DtpTo.Value) & "#)"
        End If
    Case VbCalendar.vbCalHijri
        StrWhere = StrWhere & " And(Notes.HijriDate >='" & Me.DHijriFrom.Value & "')"
        StrWhere = StrWhere & " And(Notes.HijriDate <='" & Me.DHijriTo.Value & "')"
End Select
StrSQL = StrSQL & StrWhere
BluidSQL = StrSQL
End Function

Private Sub TxtValueFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Exit Sub
If CBool(InStr(1, ".", Chr(KeyAscii))) And CBool(InStr(1, TxtValueFrom.Text, Chr(KeyAscii))) Then
    KeyAscii = 0
    Exit Sub
End If
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub


Private Sub TxtValueTo_KeyPress(KeyAscii As Integer)
If CBool(InStr(1, ".", Chr(KeyAscii))) And CBool(InStr(1, TxtValueTo.Text, Chr(KeyAscii))) Then
    KeyAscii = 0
    Exit Sub
End If
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Public Sub ChangeLang()
Me.Caption = "Journal Search"
EleHeader.Caption = Me.Caption
With Me.FgResult
    .TextMatrix(0, .ColIndex("Serial")) = "Voucher Serial"
    .TextMatrix(0, .ColIndex("Type")) = "Voucher Type"
    .TextMatrix(0, .ColIndex("Value")) = "Voucher Value"
    .TextMatrix(0, .ColIndex("State")) = "Posting State"
    .TextMatrix(0, .ColIndex("SandDate")) = "Issued Date"
    .TextMatrix(0, .ColIndex("Emp")) = "Employee Name"
    .TextMatrix(0, .ColIndex("Issued_by")) = "Issued By"
    .AutoSize 0, .Cols - 1, False
End With
Me.Lbl(0).Caption = "From"
Me.Lbl(1).Caption = "To"
Me.Lbl(2).Caption = "From"
Me.Lbl(3).Caption = "To"
Me.Lbl(4).Caption = "Employee"
Me.Lbl(5).Caption = "Serial"
Me.Lbl(6).Caption = "Issued By"

Me.Ele(0).Caption = "Posting State"
Opt(0).Caption = "Posted"
Opt(1).Caption = "Not Posted"
Opt(2).Caption = "All"
Me.Ele(1).Caption = "Voucher Value"
Cmd(0).Caption = "&Clear"
Cmd(1).Caption = "&Search"
Cmd(2).Caption = "E&xit"
Me.Fra(0).Caption = "With in interval"
End Sub

Private Sub AddTip()

Dim Wrap As String
Dim Msg As String
 
Wrap = Chr(13) + Chr(10)
If SystemOptions.UserInterface = ArabicInterface Then
    With TTP
        .Create Me.hwnd, "‰ «∆Ã «·»ÕÀ", 1, 15204351, -2147483630, True
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Â‰«  ⁄—÷ ‰ «∆Ã «·»ÕÀ..»«·÷€ÿ ⁄·Ï" & Wrap & _
                "«Ï Õ—ﬂ… Ì „ ≈” —Ã«⁄Â« ›Ï «·‘«‘…" & Wrap & _
                "«·—∆Ì”Ì…."
        .AddControl FgResult, Msg, True
    End With
    With TTP
        .Create Me.hwnd, "„”·”· «·”‰œ", 1, 15204351, -2147483630, True
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "≈–« √—œ  ≈” —Ã«⁄ Õ—ﬂ… „⁄Ì‰…" & Wrap & _
                "√œŒ· „”·”· Â–Â «·Õ—ﬂ… Â‰«"
        .AddControl TxtSer, Msg, True
    End With
    With TTP
        .Create Me.hwnd, " «—ÌŒ »œ«Ì… «·› —…", 1, 15204351, -2147483630, True
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "» ›⁄Ì· ⁄·«„… «·√Œ Ì«— Ì»œ√ «·»ÕÀ ⁄‰ «·”‰œ« " & Wrap & _
                " «· Ï Õ——  »œ«Ì… „‰ Â–« «· «—ÌŒ ›ﬁÿ ."
        .AddControl DtpFrom, Msg, True
    End With
    With TTP
        .Create Me.hwnd, " «—ÌŒ ‰Â«Ì… «·› —…", 1, 15204351, -2147483630, True
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "» ›⁄Ì· ⁄·«„… «·√Œ Ì«— Ì»œ√ «·»ÕÀ ⁄‰ «·”‰œ« " & Wrap & _
                " «· Ï Õ——  Õ Ï  Â–« «· «—ÌŒ  ›ﬁÿ ."
        .AddControl DtpTo, Msg, True
    End With
   
    With TTP
        .Create Me.hwnd, "»œ«Ì… ﬁÌ„… «·”‰œ", 1, 15204351, -2147483630, True
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "≈–« √—œ  «·»ÕÀ ⁄‰ «·”‰œ«  «· Ï  »œ√ „‰" & Wrap & _
                "ﬁÌ„… „«·Ì… „⁄Ì‰…...«œŒ· Â–Â «·ﬁÌ„… Â‰«" & Wrap & _
                "(«·”‰œ«   ﬁÌ„ Â« «ﬂ»— „‰ Â–Â «·ﬁÌ„…)."
        .AddControl TxtValueFrom, Msg, True
    End With
    With TTP
        .Create Me.hwnd, "‰Â«Ì… ﬁÌ„… «·”‰œ", 1, 15204351, -2147483630, True
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "≈–« √—œ  «·»ÕÀ ⁄‰ «·”‰œ«  «· Ï  ﬂÊ‰ √’€—" & Wrap & _
                "„‰ ﬁÌ„… „«·Ì… „⁄Ì‰…...«œŒ· Â–Â «·ﬁÌ„… Â‰«" & Wrap & _
                "(«·”‰œ«   ﬁÌ„ Â« √’€— „‰ Â–Â «·ﬁÌ„…)."
        .AddControl TxtValueTo, Msg, True
    End With
    With TTP
        .Create Me.hwnd, "Õ«·… «·”‰œ(„—Õ·)", 1, 15204351, -2147483630, True
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "» ›⁄Ì· Â–« «·ŒÌ«— Ì „ «·»ÕÀ ⁄‰" & Wrap & _
                "«·”‰œ«  «· Ï  „  —ÕÌ·Â« ›ﬁÿ."
        .AddControl Opt(0), Msg, True
    End With
    With TTP
        .Create Me.hwnd, "Õ«·… «·”‰œ(€Ì— „—Õ·)", 1, 15204351, -2147483630, True
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "» ›⁄Ì· Â–« «·ŒÌ«— Ì „ «·»ÕÀ ⁄‰" & Wrap & _
                "«·”‰œ«  «· Ï ·„ Ì „  —ÕÌ·Â« ›ﬁÿ."
        .AddControl Opt(1), Msg, True
    End With
    With TTP
        .Create Me.hwnd, "Õ«·… «·”‰œ(«·ﬂ·)", 1, 15204351, -2147483630, True
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "» ›⁄Ì· Â–« «·ŒÌ«— Ì „ «·»ÕÀ ⁄‰ Ã„Ì⁄" & Wrap & _
                "«·”‰œ«  ”Ê«¡ ﬂ«‰  Â–Â «·”‰œ«   „" & Wrap & _
                " —ÕÌ·Â« «„ ·„ Ì „  —ÕÌ·Â«"
        .AddControl Opt(2), Msg, True
    End With
    With TTP
        .Create Me.hwnd, "„”Õ", 1, 15204351, -2147483630, True
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "„”Õ Ã„Ì⁄ ‘—Êÿ «·»ÕÀ «·”«»ﬁ…" & Wrap & _
                "·»œ¡ ⁄„·Ì… «·»ÕÀ „‰ ÃœÌœ."
        .AddControl Cmd(0), Msg, True
    End With
    With TTP
        .Create Me.hwnd, "»ÕÀ", 1, 15204351, -2147483630, True
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "»œ¡ ⁄„·Ì… «·»ÕÀ Õ”» «·‘—Êÿ «·„Õœœ…"
        .AddControl Cmd(1), Msg, True
    End With
    With TTP
        .Create Me.hwnd, "Œ—ÊÃ", 1, 15204351, -2147483630, True
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ „‰ ‘«‘… «·»ÕÀ ⁄‰ ﬁÌÊœ «·ÌÊ„Ì…"
        .AddControl Cmd(2), Msg, True
    End With
    With TTP
        .Create Me.hwnd, "«”„ «·„ÊŸ›.", 1, 15204351, -2147483630, True
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "≈–« «—œ  «·»ÕÀ ⁄‰ «·”‰œ«  «· Ï  " & Wrap & _
                "Õ——  ·„ÊŸ› „⁄Ì‰ ...√Œ — Â–« «·„ÊŸ›"
        .AddControl DcboEmp, Msg, True
    End With
    With TTP
        .Create Me.hwnd, "«”„ «·„” Œœ„ «·„Õ——.", 1, 15204351, -2147483630, True
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "≈–« «—œ  «·»ÕÀ ⁄‰ «·”‰œ«  «· Ï Õ—— " & Wrap & _
                " »Ê«”ÿ… „” Œœ„ „⁄Ì‰ ...√Œ — Â–« «·„” Œœ„"
        .AddControl DcboUsers, Msg, True
    End With
Else
    With TTP
        .Create Me.hwnd, "Search Results", 1, 15204351, -2147483630, False
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Display search results. " & Wrap & _
                "click on any transaction" & Wrap & _
                "to retrive it."
        .AddControl FgResult, Msg, False
    End With
    With TTP
        .Create Me.hwnd, "Voucher Serail", 1, 15204351, -2147483630, False
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "If you want to retrive a specific" & Wrap & _
                "Voucher.!! enter it is serial here."
        .AddControl TxtSer, Msg, False
    End With
    With TTP
        .Create Me.hwnd, "Beginning Date", 1, 15204351, -2147483630, False
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "When Check is enabled...search will" & Wrap & _
                "start from this date."
        .AddControl DtpFrom, Msg, False
    End With
    With TTP
        .Create Me.hwnd, "End Date", 1, 15204351, -2147483630, False
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "When check is enabled...search will" & Wrap & _
                "end to this date."
        .AddControl DtpTo, Msg, False
    End With
    With TTP
        .Create Me.hwnd, "Voucher Value(Start)", 1, 15204351, -2147483630, False
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Enter the Value which you want to" & Wrap & _
                "start the search from it."
        .AddControl TxtValueFrom, Msg, False
    End With
    With TTP
        .Create Me.hwnd, "Voucher Value(End)", 1, 15204351, -2147483630, False
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Enter the Value which you want to" & Wrap & _
                "End the search to it."
        .AddControl TxtValueTo, Msg, False
    End With
    With TTP
        .Create Me.hwnd, "Voucher State(Posted)", 1, 15204351, -2147483630, False
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "This option when enabled... " & Wrap & _
                "the search will retrive the " & Wrap & _
                "posted voucheres only." & Wrap & _
                "(POSTED JOURNAL)."
        .AddControl Opt(0), Msg, False
    End With
    With TTP
        .Create Me.hwnd, "Voucher State(Not Posted)", 1, 15204351, -2147483630, False
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "This option when enabled... " & Wrap & _
                "the search will retrive the " & Wrap & _
                "not posted voucheres only." & Wrap & _
                "(NOT POSTED JOURNAL)."
        .AddControl Opt(1), Msg, False
    End With
    With TTP
        .Create Me.hwnd, "Voucher State(All)", 1, 15204351, -2147483630, False
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "This option when enabled... " & Wrap & _
                "the search will retrive the " & Wrap & _
                "All voucheres(ALL JOURNAL)."
        .AddControl Opt(2), Msg, False
    End With
    With TTP
        .Create Me.hwnd, "Clear", 1, 15204351, -2147483630, False
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Clear all search conditions" & Wrap & _
                "to start a new search ."
        .AddControl Cmd(0), Msg, False
    End With
    With TTP
        .Create Me.hwnd, "Search", 1, 15204351, -2147483630, False
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Start a search process."
        .AddControl Cmd(1), Msg, False
    End With
    With TTP
        .Create Me.hwnd, "Exit", 1, 15204351, -2147483630, False
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Exit from search journal screen."
        .AddControl Cmd(2), Msg, False
    End With
    With TTP
        .Create Me.hwnd, "Employee Name", 1, 15204351, -2147483630, False
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "If you want to retrive the vouchers  " & Wrap & _
                "which commissioned for this employee."
        .AddControl DcboEmp, Msg, False
    End With
    With TTP
        .Create Me.hwnd, "Issued By(Users)", 1, 15204351, -2147483630, False
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "If you want to retrive the vouchers" & Wrap & _
                "which issued by specific user. "
        .AddControl DcboUsers, Msg, False
    End With
End If
End Sub
