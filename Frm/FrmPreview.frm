VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmPreview 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„⁄«Ì‰…"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
   FillStyle       =   0  'Solid
   Icon            =   "FrmPreview.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   8505
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7725
      Top             =   3105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPreview.frx":038A
            Key             =   "Frist"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPreview.frx":0724
            Key             =   "Prev"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPreview.frx":0ABE
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPreview.frx":0E58
            Key             =   "Last"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPreview.frx":11F2
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPreview.frx":158C
            Key             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Tbr 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Œ—ÊÃ"
            ImageKey        =   "Exit"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "·ÿ»«⁄… «·›« Ê—… "
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Frist"
            Object.ToolTipText     =   "«·’›Õ… «·«Ê·Ì „‰ «·›« Ê—…"
            ImageKey        =   "Frist"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Prev"
            Object.ToolTipText     =   "«·’›Õ… «·”«»ﬁ… „‰ «·›« Ê—…"
            ImageKey        =   "Prev"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next"
            Object.ToolTipText     =   "«·’›Õ… «· «·Ì… „‰ «·›« Ê—…"
            ImageKey        =   "Next"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Last"
            Object.ToolTipText     =   "«·’›Õ… «·«ŒÌ—… „‰ «·›« Ê—…"
            ImageKey        =   "Last"
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   30
         Width           =   750
         Begin VB.Label LblTotPages 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   15
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   15
            Width           =   45
         End
      End
      Begin VB.TextBox TxtPageNum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   15
         Width           =   390
      End
   End
   Begin VB.PictureBox PictLogo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   210
      Left            =   90
      RightToLeft     =   -1  'True
      ScaleHeight     =   210
      ScaleWidth      =   1650
      TabIndex        =   5
      Top             =   4665
      Width           =   1650
      Begin VB.Label LabSite 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   1620
      End
   End
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5490
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   5490
      ScaleWidth      =   6840
      TabIndex        =   1
      Top             =   600
      Width           =   6840
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin ImpulseAniLabel.ISAniLabel GridColor 
         Height          =   315
         Left            =   -510
         TabIndex        =   2
         Top             =   135
         Visible         =   0   'False
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   556
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Andalus"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Andalus"
         FontSize        =   8.25
         Caption         =   "ISAniLabel2"
         ImageCount      =   0
      End
      Begin VSFlex8Ctl.VSFlexGrid Grd 
         Height          =   1140
         Index           =   0
         Left            =   990
         TabIndex        =   3
         Top             =   2445
         Visible         =   0   'False
         Width           =   2790
         _cx             =   4921
         _cy             =   2011
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   0
         BackColorSel    =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   8421631
         FloodColor      =   8421631
         SheetBorder     =   14737632
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   0
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmPreview.frx":1926
         ScrollTrack     =   0   'False
         ScrollBars      =   0
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
      Begin ImpulseAniLabel.ISAniLabel LblText 
         Height          =   330
         Index           =   0
         Left            =   2880
         TabIndex        =   4
         Top             =   780
         Visible         =   0   'False
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   582
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Arabic Transparent"
         FontSize        =   9.75
         MouseIcon       =   "FrmPreview.frx":199B
         BackColor       =   14737632
         Alignment       =   1
         Caption         =   "√œ«… ‰’"
         Interval        =   55
         StretchImages   =   0   'False
         RightToLeft     =   -1  'True
         ColorTextShadow =   16777215
         ColorShadow     =   16777215
         ImageCount      =   0
      End
      Begin VB.Image CoLog 
         Height          =   885
         Index           =   0
         Left            =   600
         Picture         =   "FrmPreview.frx":1AFD
         Stretch         =   -1  'True
         Top             =   7365
         Visible         =   0   'False
         Width           =   870
      End
   End
   Begin VB.ListBox LstCtrl 
      Height          =   2400
      Left            =   7530
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   495
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu op 
      Caption         =   "ŒÌ«—« "
      Visible         =   0   'False
      Begin VB.Menu prnt 
         Caption         =   "ÿ»«⁄…"
      End
      Begin VB.Menu sp 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFrist 
         Caption         =   "«·«Ê·"
      End
      Begin VB.Menu MnuNxt 
         Caption         =   "«· «·Ï"
      End
      Begin VB.Menu MnuPrev 
         Caption         =   "«·”«»ﬁ"
      End
      Begin VB.Menu MnuLast 
         Caption         =   "«·«ŒÌ—"
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Œ—ÊÃ"
      End
   End
End
Attribute VB_Name = "FrmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotPage As Integer
Dim RowsPerPage As Integer

Private Sub TotalPages()
    On Error GoTo ErrTrap
    Dim UpRow As Integer
    Dim i As Integer
    Dim LstRowVis As Integer
    Dim FrstRowHid As Integer
    'If Me.Grd(0).Visible = True Then
    LstRowVis = 0
    FrstRowHid = 0
    RowsPerPage = 0
    UpRow = Me.grd(0).TopRow
    TotPage = 1

    For i = 1 To Me.grd(0).Rows - 1

        'Debug.Print Me.Grd(0).TextMatrix(I, 0)
        If Me.grd(0).RowIsVisible(i) = False Then
            Me.grd(0).Height = i * Me.grd(0).RowHeight(i)
            FrstRowHid = i
            LstRowVis = i - 1
            RowsPerPage = LstRowVis
            
            Exit For
        End If

    Next

    If FrstRowHid > 0 And LstRowVis > 0 Then
        Me.Frame1.Visible = True
        Me.TBr.Buttons.Item("Frist").Visible = True
        Me.TBr.Buttons.Item("Last").Visible = True
        Me.TBr.Buttons.Item("Next").Visible = True
        Me.TBr.Buttons.Item("Prev").Visible = True
        Me.TxtPageNum.Visible = True
        TotPage = Int((grd(0).Rows - 1) / (LstRowVis))

        If (grd(0).Rows - 1) Mod (LstRowVis) > 0 Then
            TotPage = TotPage + 1
        End If

        grd(0).Rows = (TotPage * RowsPerPage) + 1
        
    End If

    Me.LblTotPages.Caption = TotPage
    Me.TxtPageNum.Text = 1
    'End If
ErrTrap:
End Sub

Private Sub FrmPrint()
    On Error GoTo ErrTrap

    'Dim Crep As ClsReportProp
    Dim Pr As Integer
    'Set Crep = New ClsReportProp
    'Crep.LoadFile "c:\Temp_Print\temp.drp", Frm_Print
    'Crep.InvoID = 5
    'Crep.PrintBill
    Dim CurPage As Integer
    CurPage = TxtPageNum.Text

    With Me
        .TBr.Visible = False
    
        For Pr = 1 To val(LblTotPages.Caption)
       
            .MovetoPage Pr
            .PrintForm

            DoEvents
        
        Next

        .TBr.Visible = True
    End With

    MovetoPage CurPage
    'Unload Frm_Print
    'Set Crep = Nothing
ErrTrap:
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrTrap

    'TotalPages
    'TotalPages
    'OtherData
    If FrmPreview.grd(0).Visible = True Then
    
        Me.Frame1.Visible = True
        Me.TBr.Buttons.Item("Frist").Visible = True
        Me.TBr.Buttons.Item("Last").Visible = True
        Me.TBr.Buttons.Item("Next").Visible = True
        Me.TBr.Buttons.Item("Prev").Visible = True
        Me.TxtPageNum.Visible = True
    Else
        Me.Frame1.Visible = False
        Me.TBr.Buttons.Item("Frist").Visible = False
        Me.TBr.Buttons.Item("Last").Visible = False
        Me.TBr.Buttons.Item("Next").Visible = False
        Me.TBr.Buttons.Item("Prev").Visible = False
        Me.TxtPageNum.Visible = False
  
    End If

    Me.Width = PicMain.Width + 100
    Me.Height = PicMain.Height + 630
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    'Set Crep = New ClsReportProp
    'LoadFile "c:\›« Ê—… »Ì⁄.drp", Me
    'Crep.InvoID = 5
    On Error GoTo ErrTrap
    Preview

    Me.PicMain.left = 0
    PicMain.top = 330
    Me.Width = PicMain.Width '+ 100
    Me.Height = PicMain.Height + 330

    Me.PicMain.Enabled = False
    Me.TxtPageNum.Visible = False
    Me.Frame1.Visible = False
    Me.TBr.Buttons.Item("Frist").Visible = False
    Me.TBr.Buttons.Item("Last").Visible = False
    Me.TBr.Buttons.Item("Next").Visible = False
    Me.TBr.Buttons.Item("Prev").Visible = False
    Me.PictLogo.left = Me.PicMain.left + 400
    Me.PictLogo.top = Me.PicMain.Height - Me.PictLogo.Height
    Me.PictLogo.ZOrder 0
    TotalPages
    OtherData
    Exit Sub
ErrTrap:

End Sub

Private Sub PicMain_Click()
    Unload Me
End Sub

Public Sub Preview()
    On Error GoTo ErrTrap
    '        Dim crep As ClsReportProp
    '        Set crep = New ClsReportProp
      
    crep.LoadFile crep.OpenFile, Me
    'Crep.InvoID = 5
    crep.ShowReport
        
    Dim My_SQL As String
    Dim RsFiled As ADODB.Recordset
    Dim Ctrl As Control
    Dim IX As Integer
    Dim TxtTag As String
    Set RsFiled = New ADODB.Recordset
    My_SQL = crep.SQLQuery(crep.InvoID)
    RsFiled.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic
    Dim II As Long

    If RsFiled.RecordCount > 0 Then
        grd(0).Rows = RsFiled.RecordCount + 1
        RsFiled.MoveFirst

        Do Until RsFiled.EOF
            II = II + 1

            For IX = 0 To grd(0).Cols - 1

                If Trim(grd(0).ColKey(IX)) <> "" Then
                    If Trim(grd(0).ColKey(IX)) = "Record_Number" Then
                        grd(0).TextMatrix(II, IX) = II
                    Else
                        grd(0).TextMatrix(II, IX) = IIf(IsNull(RsFiled.Fields(grd(0).ColKey(IX)).value), "", RsFiled.Fields(grd(0).ColKey(IX)).value)
                    End If
                End If

            Next

            RsFiled.MoveNext
        Loop

        RsFiled.MoveFirst

        For IX = 0 To LblText.count - 1

            If Trim(LblText(IX).Tag) <> "" Then
                If LblText(IX).Tag = "PrintDate" Then
                    LblText(IX).Caption = Date
                ElseIf LblText(IX).Tag = "PrintTime" Then
                    LblText(IX).Caption = Time
                Else

                    Select Case Trim(LblText(IX).Tag)

                        Case "PrintDate"
                            LblText(IX).Caption = Format(Date, "yyyy/mm/dd")

                        Case "PrintTime"
                            LblText(IX).Caption = Time

                        Case "PageNum", "PageOfTotal", "PageTotal"
                            GoTo EXT

                        Case Else
                            LblText(IX).Caption = IIf(IsNull(RsFiled.Fields(Trim(LblText(IX).Tag))), "", RsFiled.Fields(Trim(LblText(IX).Tag)))
                    End Select

                End If
            End If

EXT:
        Next

    End If

    PictLogo.top = Me.PicMain.Height - PictLogo.Height
    PictLogo.ZOrder 0
    Exit Sub
FillOthersCtrl:
ErrTrap:
End Sub

Private Sub prnt_Click()
    'Me.Grd(0).PrintGrid
    Printer.PrintQuality = vbPRPQMedium
    Me.PrintForm
End Sub

Private Sub Tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrTrap

    Select Case Button.Key

        Case "Frist"
            Me.TxtPageNum.Text = 1
            Me.grd(0).TopRow = 1

        Case "Last"
            Me.TxtPageNum.Text = Me.LblTotPages.Caption
        
        Case "Next"

            If val(Me.TxtPageNum.Text) >= val(Me.LblTotPages.Caption) Then
                Me.TxtPageNum.Text = val(Me.LblTotPages.Caption)
            Else
                Me.TxtPageNum.Text = val(Me.TxtPageNum.Text) + 1
            End If

        Case "Prev"

            If val(Me.TxtPageNum.Text) <= 1 Then
                Me.TxtPageNum.Text = 1
            Else
                Me.TxtPageNum.Text = val(Me.TxtPageNum.Text) - 1
            End If

        Case "Print"
            'FrmPrnt
            FrmPrint
            Exit Sub

        Case "Exit"
            Unload Me
            Exit Sub
    End Select

    MovetoPage val(Me.TxtPageNum.Text)
ErrTrap:
End Sub

Public Sub MovetoPage(Page As Integer)
    Dim TpRow As Integer
    TpRow = ((Page - 1) * RowsPerPage) + 1
    Me.grd(0).TopRow = TpRow
    OtherData
End Sub

Private Sub TxtPageNum_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrTrap
    Dim TxtAscii As String

    If KeyAscii = vbKeyReturn Then
        If val(Me.TxtPageNum.Text) = 0 Then Me.TxtPageNum.Text = 1
        If val(Me.TxtPageNum.Text) > val(Me.LblTotPages.Caption) Then Me.TxtPageNum.Text = val(Me.LblTotPages.Caption)
        MovetoPage val(Me.TxtPageNum.Text)
    End If

    TxtAscii = "1234567890"

    If KeyAscii <> vbKeyBack Then
        If InStr(1, TxtAscii, CHR(KeyAscii), vbTextCompare) Then
        Else
            KeyAscii = 0
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Public Sub OtherData()
    On Error GoTo ErrTrap
    Dim ctl As Control

    For Each ctl In Me.Controls

        If TypeOf ctl Is ISAniLabel Then
            Debug.Print ctl.Tag

            '        If Trim(Ctl.Tag) <> "" Then
            '            Stop
            '        End If
            If Trim(ctl.Tag) = "PageNum" Then
                ctl.Caption = Me.TxtPageNum.Text
            ElseIf Trim(ctl.Tag) = "PageOfTotal" Then
                ctl.Caption = Me.TxtPageNum.Text & "„‰" & Me.LblTotPages.Caption
            ElseIf Trim(ctl.Tag) = "PageTotal" Then
                ctl.Caption = Me.LblTotPages.Caption
            End If
        End If

    Next

    Exit Sub
ErrTrap:
End Sub

Private Sub FrmPrnt()
    On Error GoTo ErrTrap
    Dim Frm As New FrmPreview
    Frm.TBr.Visible = False
    Frm.PicMain.top = 0
    Frm.grd(0).TopRow = Me.grd(0).TopRow
    Frm.grd(0).Height = Me.grd(0).Height
    Dim Pr As Integer

    If Me.grd(0).Visible = False Then
        Me.LblTotPages.Caption = 1
    End If

    For Pr = 1 To val(LblTotPages.Caption)
        Frm.MovetoPage Pr
        Frm.show vbModal

        'Frm.PrintForm
        'MsgBox ""
        DoEvents
    Next

    Exit Sub
ErrTrap:
    'Unload Frm
End Sub

