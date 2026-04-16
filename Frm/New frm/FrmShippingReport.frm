VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmShippingReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
   Icon            =   "FrmShippingReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   10365
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
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   495
      Left            =   5880
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1095
      _cx             =   1931
      _cy             =   873
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
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   4725
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   10395
      Begin VB.TextBox TxtEmployeeID2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1800
         Width           =   825
      End
      Begin VB.TextBox TxtEmployeeID1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   480
         Width           =   825
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   2640
         Width           =   825
      End
      Begin VB.OptionButton ChHelper 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„”«⁄œÌ‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   32
         ToolTipText     =   "«’€— „‰"
         Top             =   2160
         Width           =   3555
      End
      Begin VB.OptionButton ChOpretor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„‘€·Ì‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   28
         ToolTipText     =   "«’€— „‰"
         Top             =   1320
         Width           =   3435
      End
      Begin VB.OptionButton ChDr 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—œÊœ «·”«∆ﬁÌ‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   27
         ToolTipText     =   "«’€— „‰"
         Top             =   120
         Width           =   4035
      End
      Begin VB.TextBox TxtValue 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4560
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   960
         Width           =   825
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   960
         Width           =   4065
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   24
            ToolTipText     =   "«’€— „‰"
            Top             =   0
            Width           =   555
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   23
            ToolTipText     =   "Ì”«ÊÏ"
            Top             =   0
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   22
            ToolTipText     =   "«ﬂ»— „‰"
            Top             =   0
            Width           =   465
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "=>"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   -120
            RightToLeft     =   -1  'True
            TabIndex        =   21
            ToolTipText     =   "«ﬂ»— „‰"
            Top             =   0
            Width           =   705
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "=<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   20
            ToolTipText     =   "«’€— „‰"
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "„‰ «·› —Â"
         Height          =   735
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   3240
         Width           =   4455
         Begin MSComCtl2.DTPicker DtpDateFrom 
            Height          =   330
            Left            =   2280
            TabIndex        =   12
            Top             =   270
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   91422723
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker DtpDateTo 
            Height          =   330
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   91422723
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   3
            Left            =   1710
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "„‰"
            Height          =   195
            Index           =   4
            Left            =   3690
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4575
         Left            =   6720
         TabIndex        =   5
         Top             =   120
         Width           =   3615
         Begin VB.Image Image1 
            Height          =   3675
            Left            =   120
            Picture         =   "FrmShippingReport.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3315
         End
         Begin VB.Label lblCompanyname 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·”« —Ì…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   1095
            Left            =   240
            TabIndex        =   6
            Top             =   3840
            Width           =   2895
         End
      End
      Begin MSDataListLib.DataCombo DcbDriver 
         Height          =   315
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbEp 
         Height          =   315
         Left            =   360
         TabIndex        =   29
         Top             =   1800
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCEmp2 
         Height          =   315
         Left            =   360
         TabIndex        =   33
         Top             =   2640
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„”«⁄œ"
         Height          =   285
         Index           =   2
         Left            =   5220
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„‘€·"
         Height          =   285
         Index           =   6
         Left            =   5220
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·—œ"
         Height          =   285
         Index           =   0
         Left            =   5340
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·”«∆ﬁ"
         Height          =   285
         Index           =   15
         Left            =   5220
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "‘«‘…  ﬁ«—Ì— —œÊœ «·”«∆ﬁÌ‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   420
         Index           =   25
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   4080
         Width           =   6255
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   495
         Left            =   120
         Top             =   4080
         Width           =   6375
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   5640
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   5640
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   873
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   31
      Top             =   5640
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "⁄—÷  Õ·Ì·Ì"
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ»ﬁ« ·„” √Ã— „Õœœ"
      Height          =   195
      Index           =   5
      Left            =   5400
      TabIndex        =   9
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ«—Ì— «·‘Õ‰"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmShippingReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecID As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Dim IDPr As Integer
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
   ' Set XPic = Me.btnFirst.ButtonImage
   ' Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
   ' Set Me.btnLast.ButtonImage = XPic
   ' Set XPic = Me.btnPrevious.ButtonImage
   ' Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
   ' Set Me.btnNext.ButtonImage = XPic
    Label5.Caption = "Shipping Reports"
    Lbl(25).Caption = Label5.Caption
    ChDr.Caption = "Responses Drivers"
    ChDr.RightToLeft = False
    Frame1.Caption = "Period"
    Lbl(3).Caption = "To"
    Lbl(4).Caption = "From"
Lbl(15).Caption = "Driver Name"
Lbl(6).Caption = "Operator Name"
ChOpretor.Caption = "Operators"
ChOpretor.RightToLeft = False
Lbl(0).Caption = "Replies value"
btnClear.Caption = "Clear"
Cmd(0).Caption = "Show Report"
Cmd(2).Caption = "Exit"
ChHelper.RightToLeft = False
ChHelper.Caption = "Helpers"
Cmd(1).Caption = "Print Analytic"
Lbl(2).Caption = "Helper"

End Sub
Private Sub btnClear_Click()
clear_all Me

DtpDateFrom.value = ""
DtpDateTo.value = ""
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
IDPr = 0
GetData
          
        Case 1
           IDPr = 1
GetData

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
            Case 3
'print_report
    End Select

End Sub




Private Sub DcbDriver_Change()
DcbDriver_Click (0)
End Sub

Private Sub DcbDriver_Click(Area As Integer)
 If val(DcbDriver.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String

    GetEmployeeIDFromCode , , DcbDriver.BoundText, EmpCode
    TxtEmployeeID1.text = EmpCode
End Sub

Private Sub DcbEp_Click(Area As Integer)
 If val(DcbEp.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String

    GetEmployeeIDFromCode , , DcbEp.BoundText, EmpCode
    TxtEmployeeID2.text = EmpCode
End Sub

Private Sub DCEmp2_Change()
DCEmp2_Click (0)
End Sub

Private Sub DCEmp2_Click(Area As Integer)
 If val(DCEmp2.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String

    GetEmployeeIDFromCode , , DCEmp2.BoundText, EmpCode
    Text2.text = EmpCode
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub




Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
  
  
    Set Dcombos = New ClsDataCombos
   Dcombos.GetEmployees Me.DcbDriver, , True
   Dcombos.GetEmployees Me.DcbEp
   Dcombos.GetEmployees Me.DCEmp2
DtpDateFrom.value = ""
DtpDateTo.value = ""
    Resize_Form Me
    If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
    End If

End Sub





Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim strSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    If ChDr.value = True Or Me.ChOpretor.value = True Or Me.ChHelper.value = True Then
If ChDr.value = True Then
strSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, "
strSQL = strSQL & "                        dbo.Transactions.DriverId, dbo.tblCarDrivers.DrivValue, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
strSQL = strSQL & "                        dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
strSQL = strSQL & "                        dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.Transactions.CusID,"
strSQL = strSQL & "                        dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS Expr1, dbo.Transactions.Transaction_HijriDate,"
strSQL = strSQL & "                        dbo.Transactions.TransactionComment, dbo.Transactions.CashCustomerComment, dbo.Transactions.CashCustomerAddress, dbo.Transactions.CashCustomerMobile,"
strSQL = strSQL & "                        dbo.Transactions.CashCustomerPhone, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.Transactions.Transporterdriver,"
strSQL = strSQL & "                        dbo.Transactions.DepartureDate, dbo.Transactions.DepartureTime, dbo.Transactions.Transporter, dbo.TblEmployee.Emp_ID,"
strSQL = strSQL & "                        dbo.GetDriverTrip(dbo.tblCarDrivers.EmpID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Date, 2) AS TotalTrip"
strSQL = strSQL & "  FROM         dbo.TblCustemers RIGHT OUTER JOIN"
strSQL = strSQL & "                        dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID LEFT OUTER JOIN"
strSQL = strSQL & "                        dbo.TblEmployee RIGHT OUTER JOIN"
strSQL = strSQL & "                        dbo.tblCarDrivers ON dbo.TblEmployee.Emp_ID = dbo.tblCarDrivers.EmpID ON dbo.Transactions.DriverId = dbo.tblCarDrivers.EmpID"
strSQL = strSQL & "   Where (dbo.Transactions.Transaction_Type = 55)and (NOT (dbo.TblEmployee.Emp_ID IS NULL))"

If Me.DcbDriver.text <> "" And val(DcbDriver.BoundText) <> 0 Then
    strSQL = strSQL & " AND   dbo.tblCarDrivers.EmpID = " & val(Me.DcbDriver.BoundText)

End If

If Me.TxtValue.text <> "" Then
If Opt(2).value = True Then
    strSQL = strSQL & " AND   dbo.tblCarDrivers.DrivValue >" & TxtValue.text & ""
 ElseIf Opt(1).value = True Then
    strSQL = strSQL & " AND   dbo.tblCarDrivers.DrivValue =" & TxtValue.text & ""
  ElseIf Opt(0).value = True Then
    strSQL = strSQL & " AND   dbo.tblCarDrivers.DrivValue <" & TxtValue.text & ""
   ElseIf Opt(4).value = True Then
    strSQL = strSQL & " AND   dbo.tblCarDrivers.DrivValue >=" & TxtValue.text & ""
  ElseIf Opt(3).value = True Then
    strSQL = strSQL & " AND   dbo.tblCarDrivers.DrivValue <=" & TxtValue.text & ""
    
  End If
End If
End If
''''//////////
If Me.ChOpretor.value = True Then
strSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, "
strSQL = strSQL & "                      dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
strSQL = strSQL & "                      dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2,"
strSQL = strSQL & "                      dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.Transaction_Details.Quantity, dbo.Transactions.Transaction_HijriDate,"
strSQL = strSQL & "                      dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
strSQL = strSQL & "                      dbo.TblCustemers.Fullcode AS Expr1, dbo.GetDriverTrip(dbo.Transactions.empID1, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Date, 0)"
strSQL = strSQL & "                      AS TotalTrip, dbo.Transactions.empID1"
strSQL = strSQL & " FROM         dbo.TblCustemers RIGHT OUTER JOIN"
strSQL = strSQL & "                      dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID LEFT OUTER JOIN"
strSQL = strSQL & "                      dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
strSQL = strSQL & "                      dbo.TblEmployee ON dbo.Transactions.empID1 = dbo.TblEmployee.Emp_ID"
strSQL = strSQL & " Where (dbo.Transactions.Transaction_Type = 61) And (Not (dbo.Transactions.empID1 Is Null))"
If Me.DcbEp.text <> "" And val(Me.DcbEp.BoundText) <> 0 Then
    strSQL = strSQL & " AND   dbo.Transactions.empID1 = " & val(Me.DcbEp.BoundText)

End If
End If
''////////////////
If Me.ChHelper.value = True Then
strSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, "
strSQL = strSQL & "                      dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
strSQL = strSQL & "                      dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2,"
strSQL = strSQL & "                      dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.Transaction_Details.Quantity, dbo.Transactions.empID2,"
strSQL = strSQL & "                      dbo.Transactions.Transaction_HijriDate, dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS Expr1,"
strSQL = strSQL & "                      dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, dbo.GetDriverTrip(dbo.Transactions.empID2, dbo.Transactions.Transaction_Date,"
strSQL = strSQL & "                      dbo.Transactions.Transaction_Date, 1) AS TotalTrip"
strSQL = strSQL & " FROM         dbo.TblCustemers RIGHT OUTER JOIN"
strSQL = strSQL & "                      dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID LEFT OUTER JOIN"
strSQL = strSQL & "                      dbo.TblEmployee ON dbo.Transactions.empID2 = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
strSQL = strSQL & "                      dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
strSQL = strSQL & " Where (dbo.Transactions.Transaction_Type = 61) And (Not (dbo.Transactions.empID2 Is Null))"
If Me.DCEmp2.text <> "" And val(Me.DCEmp2.BoundText) <> 0 Then
    strSQL = strSQL & " AND   dbo.Transactions.empID2 = " & val(Me.DCEmp2.BoundText)

End If
End If

 If Not IsNull(Me.DtpDateFrom.value) Then
                   strSQL = strSQL & " AND dbo.Transactions.Transaction_Date >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo.value) Then
                   strSQL = strSQL & " AND dbo.Transactions.Transaction_Date<=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If
 

    Set rs = New ADODB.Recordset
    rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
   '  Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
 rs.MoveFirst
' MsgBox rs("Granty").value
 print_report strSQL
'print_report StrSQL
       ' With Me.Fg
       '     .Clear flexClearScrollable, flexClearEverything
       '     .Rows = .FixedRows
       '     .Rows = rs.RecordCount + .FixedRows
'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
      End If
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— ‰Ê⁄ «· ﬁ—Ì—"
Else
MsgBox "Please Select Type of Reports"
End If
 Exit Sub
 

    End If
End Sub
Function print_report(Optional NoteSerial As String)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
      If IDPr = 0 Then
   If Me.ChDr.value = True Then
 
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepShippingDrivers.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepShippingDriversE.rpt"
       End If
      End If
       If Me.ChOpretor.value = True Then
       
             If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepShippingOpretor.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepShippingOpretorE.rpt"
            
       End If
       End If
      If Me.ChHelper.value = True Then
             If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepShippingOpretorHelper.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepShippingOpretorHelperE.rpt"
            
       End If
     End If
 Else
        If Me.ChDr.value = True Then
 
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepShippingDriversAna.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepShippingDriversAnaE.rpt"
            
       End If
       End If
        If Me.ChOpretor.value = True Then
             If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepShippingOpretorAnal.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepShippingOpretorAnalE.rpt"
            
       End If
      End If
           If Me.ChHelper.value = True Then
             If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepShippingOpretorHelperAna.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepShippingOpretorHelperAnaE.rpt"
            
       End If
     End If
End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
       'If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then xReport.ParameterFields(3).AddCurrentValue Format(Me.XPDtbFrom.value, "yyyy/M/d")
       ' If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(4).AddCurrentValue Format(Me.XPDtpTo.value, "yyyy/M/d")
       ' If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then xReport.ParameterFields(5).AddCurrentValue Me.Txt_DateHigriFrom.value
       ' If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(6).AddCurrentValue Me.Txt_DateHigriTO.value
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
  
    End If

   
  If DtpDateFrom.value <> "" And DtpDateTo.value <> "" Then
   xReport.ParameterFields(8).AddCurrentValue DtpDateFrom.value

    xReport.ParameterFields(10).AddCurrentValue DtpDateTo.value
  '  xReport.ParameterFields(11).AddCurrentValue DtpDateToH.value
    End If

  Dim Total As String
  Dim totl As Double


    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function



 
Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim empid As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text2.text, empid
        DCEmp2.BoundText = empid
    End If
End Sub



Private Sub TxtEmployeeID1_KeyPress(KeyAscii As Integer)
Dim empid As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtEmployeeID1.text, empid
        DcbDriver.BoundText = empid
    End If
End Sub

Private Sub TxtEmployeeID2_KeyPress(KeyAscii As Integer)
Dim empid As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtEmployeeID2.text, empid
        DcbEp.BoundText = empid
    End If
End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtValue.text, 0)
End Sub
