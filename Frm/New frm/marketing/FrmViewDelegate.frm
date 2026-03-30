VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmViewDelegate 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ”ÃÌ·  «—ÌŒ  "
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11295
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Ê«⁄Ìœ «·„— »ÿÂ  »«·„‰«œÌ»"
      Height          =   1815
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2280
      Width           =   11295
      Begin VSFlex8Ctl.VSFlexGrid fg2 
         Height          =   1380
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   11145
         _cx             =   19659
         _cy             =   2434
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
         Rows            =   50
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmViewDelegate.frx":0000
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Ê«⁄Ìœ «·Œ«’Â »Â"
      Height          =   1815
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   11295
      Begin VSFlex8Ctl.VSFlexGrid fg 
         Height          =   1380
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   11145
         _cx             =   19659
         _cy             =   2434
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
         Rows            =   50
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmViewDelegate.frx":00C3
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
   Begin VB.TextBox TxtComment 
      Alignment       =   1  'Right Justify
      Height          =   975
      Left            =   30
      MaxLength       =   255
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   5190
      Width           =   4425
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   405
      Left            =   60
      TabIndex        =   4
      Top             =   4170
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "«€·«ﬁ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DtpDateFrom 
      Height          =   330
      Left            =   8280
      TabIndex        =   11
      Top             =   120
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   62783491
      CurrentDate     =   38887
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·“Ì«—« "
      Height          =   285
      Index           =   24
      Left            =   9840
      TabIndex        =   12
      Top             =   120
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   8
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   7
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4425
      X2              =   0
      Y1              =   4680
      Y2              =   4695
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   5
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3420
      Width           =   3645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   4
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   3
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   1545
   End
End
Attribute VB_Name = "FrmViewDelegate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub fillfg()

End Sub


Private Sub CmdCancel_Click()
    Unload Me
End Sub


Private Sub ChangeLang()
    CmdCancel.Caption = "CLose"
'CmdOk.Caption = "Save"
lbl(24).Caption = "DateVisit"
'lbl(9).Caption = "TimeEnter"
'Me.Caption = "Register Date & Time"
Me.Caption = "Register Date Visit "
With Fg
.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("fromtime")) = "From"
.TextMatrix(0, .ColIndex("totime")) = "To"
.TextMatrix(0, .ColIndex("CustomerName")) = "CustomerName"
.TextMatrix(0, .ColIndex("PersonConc")) = "Admin"
.TextMatrix(0, .ColIndex("Mobile")) = "Mobile"
.TextMatrix(0, .ColIndex("Tel")) = "Telephone"
.TextMatrix(0, .ColIndex("Email")) = "Email"
'.TextMatrix(0, .ColIndex("totime")) = "To"
End With
Frame2.Caption = "Appointments associated Mounadeb"
Frame1.Caption = "Appointments by"
With fg2
.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("fromtime")) = "From"
.TextMatrix(0, .ColIndex("totime")) = "To"
.TextMatrix(0, .ColIndex("name")) = "With"
.TextMatrix(0, .ColIndex("CustomerName")) = "CustomerName"

End With
End Sub

Private Sub Form_Load()
    CenterForm Me

    FormPostion Me, GetPostion

    'Me.CmdOk.ButtonStyle = impActive
    'Set CmdOk.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Save").Picture
    'CmdOk.ButtonPositionImage = impRightOfText

    Me.CmdCancel.ButtonStyle = impActive
    Set CmdCancel.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").Picture
    CmdCancel.ButtonPositionImage = impRightOfText
   ' XPDtbBill.value = Date
'Me.timeEnter.value = Time
If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

