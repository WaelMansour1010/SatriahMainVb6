VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmCustomerBalances1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " —”«∆· «·⁄„·«¡"
   ClientHeight    =   9690
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   8625
   HelpContextID   =   440
   Icon            =   "FrmCustomerBalances1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9690
   ScaleWidth      =   8625
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9675
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8655
      _cx             =   15266
      _cy             =   17066
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
         Caption         =   "„—›ﬁ«  «·«Ì„Ì·"
         Height          =   615
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1800
         Width           =   8055
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   255
            Left            =   2910
            TabIndex        =   25
            Top             =   315
            Width           =   255
         End
         Begin VB.TextBox txtAttach 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3375
            TabIndex        =   24
            Top             =   240
            Width           =   2355
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   255
            Left            =   9240
            TabIndex        =   22
            Top             =   840
            Width           =   255
         End
         Begin VB.TextBox xxxx 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6945
            TabIndex        =   21
            Top             =   765
            Width           =   2115
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   -300
            Top             =   90
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   315
            Left            =   1740
            TabIndex        =   27
            Top             =   180
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            Caption         =   "«” Ì—«œ «·„·› 2"
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
            ButtonImage     =   "FrmCustomerBalances1.frx":038A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   315
            Left            =   90
            TabIndex        =   28
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   180
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            Caption         =   "Õœœ «·„”«—"
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
            ButtonImage     =   "FrmCustomerBalances1.frx":6BEC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õœœ «·„—›ﬁ"
            Height          =   420
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Attachement"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   5970
            TabIndex        =   23
            Top             =   795
            Width           =   930
         End
      End
      Begin VB.TextBox TxtSearch 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4350
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   825
      End
      Begin VB.TextBox txtQuicSearch 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1530
         Width           =   4455
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1935
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   7710
         Width           =   8565
         _cx             =   15108
         _cy             =   3413
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
         Begin VB.TextBox txtMessage 
            Alignment       =   1  'Right Justify
            Height          =   1110
            Left            =   2415
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   420
            Width           =   5205
         End
         Begin VB.CheckBox ChkShow 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·«  ŸÂ— Â–Â «·‰«›–… ⁄‰œ  ‘€Ì· «·»—‰«„Ã"
            ForeColor       =   &H000000FF&
            Height          =   930
            Left            =   3465
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   2475
            Width           =   4875
         End
         Begin ImpulseButton.ISButton CmdExit 
            Cancel          =   -1  'True
            Height          =   690
            Left            =   105
            TabIndex        =   5
            Top             =   1215
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   1217
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
            ButtonImage     =   "FrmCustomerBalances1.frx":D44E
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
            Height          =   705
            Left            =   1950
            TabIndex        =   6
            Top             =   1860
            Visible         =   0   'False
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   1244
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
            ButtonImage     =   "FrmCustomerBalances1.frx":D7E8
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton SendMessage 
            Height          =   465
            Left            =   960
            TabIndex        =   8
            Top             =   240
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«—”«· SMS"
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   465
            Left            =   960
            TabIndex        =   19
            Top             =   720
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«—”«· Email"
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·—”«·…"
            Height          =   390
            Left            =   7620
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   420
            Width           =   765
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ì „  ÕœÌœ Â–Â «·»Ì«‰«  »‰«¡« ⁄·Ï «· «—ÌŒ «·Õ«·Ì ›Ì «·ÃÂ«“"
            ForeColor       =   &H000000FF&
            Height          =   675
            Index           =   0
            Left            =   4035
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   1605
            Visible         =   0   'False
            Width           =   4425
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   5250
         Left            =   30
         TabIndex        =   2
         Top             =   2505
         Width           =   8520
         _cx             =   15028
         _cy             =   9260
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmCustomerBalances1.frx":DB82
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
      Begin VB.CheckBox Check17 
         Alignment       =   1  'Right Justify
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   375
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   645
         Width           =   1425
      End
      Begin MSDataListLib.DataCombo DcbIqara 
         Height          =   315
         Left            =   750
         TabIndex        =   15
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
         Top             =   840
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcBranch 
         Height          =   315
         Left            =   750
         TabIndex        =   17
         Top             =   1140
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «·›—⁄"
         Height          =   195
         Index           =   32
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄ﬁ«—"
         Height          =   195
         Index           =   4
         Left            =   5310
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·»ÕÀ «·”—Ì⁄"
         Height          =   420
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1590
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "FrmCustomerBalances1.frx":DD31
         Top             =   30
         Width           =   480
      End
      Begin VB.Label LblCaption 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " —”«∆· «·⁄„·«¡"
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
         Height          =   630
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   30
         Width           =   8580
      End
   End
End
Attribute VB_Name = "FrmCustomerBalances1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Askinterval As String
Dim Askcount As Integer


Private Sub Command1_Click()
CD1.ShowOpen
TxtAttach.text = CD1.FileName
End Sub

Private Sub DcbIqara_Change()

DcbIqara_Click (0)
GetQuicSearch txtQuicSearch
End Sub

Private Sub DcbIqara_Click(Area As Integer)
      If val(DcbIqara.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , DcbIqara.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.text = EmpCode
    'dcsupplier.BoundText = ownerid
    'Calculte
    'DcbUnitType_Change
End Sub

Private Sub DcbIqara_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmAqarSearch
FrmAqarSearch.m_RetrunType = 1
FrmAqarSearch.show


End If


If KeyCode = vbKeyF5 Then
'ReloadCombos
End If

End Sub

Private Sub Dcbranch_Click(Area As Integer)
GetQuicSearch txtQuicSearch
End Sub

Private Sub ISButton1_Click()
    Dim Email As String
    Dim RowNum As Integer
    Dim opt As Integer
    Dim CurrentMessage As String
    Email = ""

    With FG

        For RowNum = .FixedRows To .rows - 1
    
            If .cell(flexcpChecked, RowNum, .ColIndex("Send")) = flexChecked Then

                '  MsgBox (.TextMatrix(RowNum, .ColIndex("Numbers")))
                If (.TextMatrix(RowNum, .ColIndex("email"))) <> "" Then
                    If Email = "" Then
                        Email = (.TextMatrix(RowNum, .ColIndex("email")))
                    Else
                    
                    
                    
                        Email = Email & "," & (.TextMatrix(RowNum, .ColIndex("email")))
                   End If
             
                End If
            End If
          
        Next RowNum
      
        CurrentMessage = txtMessage.text

        If Email = "" Then Exit Sub
        Dim RetVal As String
        
           RetVal = SendMail(Trim$(Email), _
        "", _
        "", _
        Trim$(CurrentMessage), _
        "", _
        0, _
        "", _
        "", _
        Trim$(TxtAttach.text), _
       False, True)
           MsgBox IIf(RetVal = "ok", "Message sent!", RetVal)
           
         'MsgBox RetVal
                                    
    End With


End Sub

Private Sub txtQuicSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
GetQuicSearch txtQuicSearch
End If
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Double
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.text, EmpID
        DcbIqara.BoundText = EmpID
        DcbIqara_Click (0)
    End If
End Sub

Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.FG
 
            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Send")) = True
            Next i

        End With

    Else

        With Me.FG

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Send")) = False
            Next i

        End With

    End If

End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()

    If DoPremis(Do_Print, Me.Name, True) = False Then
        Exit Sub
    End If
        
    On Error GoTo ErrTrap
    Dim Reports As ClsRepoerts
    Dim StrSQL As String

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_InstallmentMustPayed", True)
    Askcount = GetSetting(StrAppRegPath, "Setting", "count_InstallmentMustPayed", True)
    
    'StrSQL = "select * From QestNotReceipted where  DueDate<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
    ' StrSQL = StrSQL + " order by CusName,Transaction_ID,QeqtNum"

    StrSQL = "SELECT     TOP 100 PERCENT dbo.QryCust_Qest.QestID, dbo.QryCust_Qest.NoteID, dbo.QryCust_Qest.QeqtNum, dbo.QryCust_Qest.PartID, dbo.QryCust_Qest.[Value], "
    StrSQL = StrSQL + " dbo.QryCust_Qest.DueDate, dbo.QryCust_Qest.Receipt, dbo.QryCust_Qest.Summition, dbo.QryCust_Qest.CustID, dbo.QryCust_Qest.CusName,"
    StrSQL = StrSQL + "  dbo.QryCust_Qest.Transaction_ID , dbo.QryCust_Qest.Transaction_Date, dbo.Transactions.NoteSerial1"
    StrSQL = StrSQL + " FROM         dbo.QryCust_Qest LEFT OUTER JOIN"
    StrSQL = StrSQL + "  dbo.Transactions ON dbo.QryCust_Qest.Transaction_ID = dbo.Transactions.Transaction_ID"
    StrSQL = StrSQL + " WHERE     (dbo.QryCust_Qest.QestID NOT IN"
    StrSQL = StrSQL + " (SELECT     QestID"
    StrSQL = StrSQL + "  from InstallmentDet_Junc_Receipt"
    StrSQL = StrSQL + " WHERE     Status <> 1))"
    StrSQL = StrSQL + "  and DueDate <" & SQLDate(Date, True) & "'"
    StrSQL = StrSQL + "  order by CusName,QryCust_Qest.Transaction_ID,QeqtNum"
 
    Set Reports = New ClsRepoerts
    Reports.QestMustPayed StrSQL, , LblCaption.Caption
    Exit Sub
ErrTrap:
End Sub

Private Sub FG_BeforeEdit(ByVal row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)

    If Col <> FG.ColIndex("Send") Then
        Cancel = True
    End If

End Sub

Private Sub GetQuicSearch(ByVal TxtSearch As String)
Dim StrSQL As String
Dim RsTemp As New ADODB.Recordset

  Dim RowNum As Double
    Dim ReCount As Double
    
    Dim BGround As New ClsBackGroundPic
    Dim BolShowRequest As Boolean
         
        StrSQL = " SELECT distinct    CusName, CusNamee, CusID, Cus_Phone, Cus_mobile,E_mail"
        StrSQL = StrSQL + " FROM         dbo.TblCustemers where CusID>2 and (Type=1 or Type=56 or Type=57 or Type=55  or Type=20)"
               StrSQL = StrSQL + " And ( CusName like '%" & (TxtSearch) & "%'  or CusNamee like '%" & (TxtSearch) & "%' or FullCode  like '%" & (TxtSearch) & "%' or Cus_Phone  like '%" & (TxtSearch) & "%' or Cus_mobile  like '%" & (TxtSearch) & "%')  "
        If (DcbIqara.text) <> "" Then
            StrSQL = StrSQL + " And CusID In (Select CusID From TblContract Where Iqar = " & val(DcbIqara.BoundText) & " )"
        End If
        If (dcBranch.text) <> "" Then
            StrSQL = StrSQL + " And CusID In (Select CusID From TblContract Where Branch_NO = " & val(dcBranch.BoundText) & " )"
        End If


StrSQL = StrSQL + "  Union all  SELECT name,nameE,0,tel,tel,EMAIL  From TblCusCsh where 1=1 "
StrSQL = StrSQL + " And ( name like '%" & (TxtSearch) & "%'  or name like '%" & (TxtSearch) & "%'  or tel  like '%" & (TxtSearch) & "%'  )  "

 
        If SystemOptions.UserInterface = ArabicInterface Then
            StrSQL = StrSQL + " Order by CusName "
        Else
            StrSQL = StrSQL + " Order by CusNamee "
        End If

    

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
FG.rows = 0
        With FG
            .rows = .FixedRows

            For ReCount = 1 To RsTemp.RecordCount
                .rows = .rows + 1
                RowNum = .rows - 1
                   
                ', dbo.QryCust_Qest.CustID
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("CusName").value), "", RsTemp("CusName").value)
                Else
                    .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("CusNamee").value), "", RsTemp("CusNamee").value)

                End If
 
                .TextMatrix(RowNum, .ColIndex("Numbers")) = IIf(IsNull(RsTemp("Cus_mobile").value), "", RsTemp("Cus_mobile").value)
            .TextMatrix(RowNum, .ColIndex("email")) = IIf(IsNull(RsTemp("e_mail").value), "", RsTemp("e_mail").value)
            
                RsTemp.MoveNext
            Next ReCount

           ' .AutoSize 0, .Cols - 1, False
        End With
    Else
    
        FG.rows = FG.FixedRows
    End If



RsTemp.Close
Exit Sub
 

    FG.WallPaper = BGround.Picture
    
End Sub
Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim RowNum As Double
    Dim ReCount As Double
    Dim RsTemp As New ADODB.Recordset
    Dim BGround As New ClsBackGroundPic
    Dim BolShowRequest As Boolean


  Dim Dcombos As ClsDataCombos
 '   Dim My_SQL As String
  
 
    Set Dcombos = New ClsDataCombos
    
    Dcombos.GetIqar DcbIqara
   
    'Dcombos.GetIqarUnit 1, DcbUnitNo
    Dcombos.GetBranches dcBranch
    
    
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    'FormPostion Me, GetPostion
    LoadIcons

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        My_SQL = "Select * From QestNotReceipted where  DueDate <=#" & SQLDate(Date) & "#"
        My_SQL = My_SQL + " order by CusName,Transaction_ID,QeqtNum"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then

        '  Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_InstallmentMustPayed", True)
        '  Askcount = GetSetting(StrAppRegPath, "Setting", "count_InstallmentMustPayed", True)
        If Askinterval = "D" Then
            '            LblCaption.Caption = LblCaption.Caption & Askcount & "  ÌÊ„  "
        ElseIf Askinterval = "M" Then
            '            LblCaption.Caption = LblCaption.Caption & Askcount & "  ‘Â—  "
        ElseIf Askinterval = "Y" Then
            '            LblCaption.Caption = LblCaption.Caption & Askcount & "  ”‰…  "
        End If
    
        '    My_SQL = "Select * From QestNotReceipted where  DueDate <='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
        '    My_SQL = My_SQL + " order by CusName,Transaction_ID,QeqtNum"
        Dim StrSQL As String
        StrSQL = " SELECT distinct    CusName, CusNamee, CusID, Cus_Phone, Cus_mobile,E_mail"
        StrSQL = StrSQL + " FROM         dbo.TblCustemers where CusID>2 and Type=1 or Type=56 or Type=57 or Type=55  or Type=20"

StrSQL = StrSQL + "  Union all  SELECT name,nameE,0,tel,tel,EMAIL  From TblCusCsh"

        If SystemOptions.UserInterface = ArabicInterface Then
            StrSQL = StrSQL + " Order by CusName "
        Else
            StrSQL = StrSQL + " Order by CusNamee "
        End If

    End If

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        With FG
            .rows = .FixedRows

            For ReCount = 1 To RsTemp.RecordCount
                .rows = .rows + 1
                RowNum = .rows - 1
                   
                ', dbo.QryCust_Qest.CustID
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("CusName").value), "", RsTemp("CusName").value)
                Else
                    .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("CusNamee").value), "", RsTemp("CusNamee").value)

                End If
 
                .TextMatrix(RowNum, .ColIndex("Numbers")) = IIf(IsNull(RsTemp("Cus_mobile").value), "", RsTemp("Cus_mobile").value)
                .TextMatrix(RowNum, .ColIndex("email")) = IIf(IsNull(RsTemp("E_mail").value), "", RsTemp("E_mail").value)
                
            
                RsTemp.MoveNext
            Next ReCount

            .AutoSize 0, .Cols - 1, False
        End With

    End If



RsTemp.Close
Set RsTemp = Nothing

 StrSQL = " SELECT   distinct tel,Name from TblCusCsh"
         

        If SystemOptions.UserInterface = ArabicInterface Then
            StrSQL = StrSQL + " Order by Name "
        Else
            StrSQL = StrSQL + " Order by Namee "
        End If

      

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        With FG
           ' .Rows = .FixedRows

            For ReCount = 1 To RsTemp.RecordCount
                .rows = .rows + 1
                RowNum = .rows - 1
                   
                ', dbo.QryCust_Qest.CustID
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("Name").value), "", RsTemp("Name").value)
                Else
                    .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("Namee").value), "", RsTemp("Namee").value)

                End If
 
                .TextMatrix(RowNum, .ColIndex("Numbers")) = IIf(IsNull(RsTemp("tel").value), "", RsTemp("tel").value)
            
                RsTemp.MoveNext
            Next ReCount

            .AutoSize 0, .Cols - 1, False
        End With

    End If



    FG.WallPaper = BGround.Picture
    BolShowRequest = GetSetting(StrAppRegPath, "View_Type", "InstallmentMustPayed", True)

    If BolShowRequest = True Then
        ChkShow.value = Unchecked
    Else
        ChkShow.value = Checked
    End If

    'Resize_Form Me, ReportSize
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Me.Caption = "Installment Must Pay"
    LblCaption.Caption = Me.Caption
    ChkShow.Caption = "Dont Show at Start"
    'Label1.Caption = "Data Based in your System Date"
    Me.CmdExit.Caption = "Exit"
    Me.CmdPrint.Caption = "Print"

    With Me.FG
        .TextMatrix(0, .ColIndex("Name")) = "Customer Name"
        .TextMatrix(0, .ColIndex("BillIID")) = "BillI ID"
        .TextMatrix(0, .ColIndex("TransDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("QestNum")) = "installm. #"
        .TextMatrix(0, .ColIndex("DueDate")) = "DueDate"
        .TextMatrix(0, .ColIndex("value")) = "value"

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If ChkShow.value = Checked Then
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

    With FG
        .cell(flexcpPicture, 0, .ColIndex("Name")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
        .cell(flexcpPicture, 0, .ColIndex("BillIID")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .cell(flexcpPicture, 0, .ColIndex("TransDate")) = mdifrmmain.ImgLstTree.ListImages("qty").Picture
        .cell(flexcpPicture, 0, .ColIndex("QestNum")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .cell(flexcpPicture, 0, .ColIndex("DueDate")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .cell(flexcpPicture, 0, .ColIndex("Value")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
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

Function GetNumbers()

End Function

Private Sub SendMessage_Click()
'    Dim Numbers As String
'    Dim RowNum As Integer
'    Dim opt As Integer
'    Dim CurrentMessage As String
'    Numbers = ""
'
'    With Fg
'
'        For RowNum = .FixedRows To .rows - 1
'
'            If .cell(flexcpChecked, RowNum, .ColIndex("Send")) = flexChecked Then
'
'                '  MsgBox (.TextMatrix(RowNum, .ColIndex("Numbers")))
'                If (.TextMatrix(RowNum, .ColIndex("Numbers"))) <> "" Then
'                    If Numbers = "" Then
'                        Numbers = (.TextMatrix(RowNum, .ColIndex("Numbers")))
'                    Else
'
'
'
'                        Numbers = Numbers & "," & (.TextMatrix(RowNum, .ColIndex("Numbers")))
'                   End If
'
'                End If
'            End If
'
'        Next RowNum
'
'        CurrentMessage = txtMessage.text
'
'        If Numbers = "" Then Exit Sub
'sendMessageM "", "", CurrentMessage, "", Numbers
'        'SMSSeTTings.SendMessage CurrentMessage, Numbers
'        'SMSSeTTings.Hide
'
'    End With
    Dim Numbers As String
    Dim RowNum As Integer
    Dim opt As Integer
    Dim CurrentMessage As String
    Numbers = ""


    CurrentMessage = txtMessage.text
    Dim TxtPhone As String
    Dim mTelNo() As String
    ReDim Preserve mTelNo(FG.rows + 1)
    With FG

        For RowNum = .FixedRows To .rows - 1
    
            If .cell(flexcpChecked, RowNum, .ColIndex("Send")) = flexChecked Then

                '  MsgBox (.TextMatrix(RowNum, .ColIndex("Numbers")))
                If (.TextMatrix(RowNum, .ColIndex("Numbers"))) <> "" Then
                    If Numbers = "" Then
                        Numbers = (.TextMatrix(RowNum, .ColIndex("Numbers")))
                    Else
                    
                    
                    
                        Numbers = (.TextMatrix(RowNum, .ColIndex("Numbers")))
                   End If
                    TxtPhone = CheckPhoneNumber(Numbers, RowNum)
                    
                    mTelNo(RowNum) = TxtPhone
                    SendSms TxtPhone, CurrentMessage

                    
             
                End If
                
                
            End If
            
         
        Next RowNum
        
        'SENDSMSBulk CurrentMessage, mTelNo

        MsgBox " „ «·«—”«· »‰Ã«Õ"

        'SMSSeTTings.SendMessage CurrentMessage, Numbers
        'SMSSeTTings.Hide
                                    
    End With
End Sub

