VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmShowTech 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "‘«‘… „Ê«œ «·⁄„·ÌÂ"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12495
   Icon            =   "FrmShowTech.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3990
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Ê«’›«  «·›‰Ì… ··„’«⁄œ"
      Height          =   1815
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1800
      Width           =   12540
      Begin VB.TextBox TxtWalls 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox TxtDorHight 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TxtDorWidth 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox DcbDOrFloor 
         Height          =   315
         ItemData        =   "FrmShowTech.frx":038A
         Left            =   7560
         List            =   "FrmShowTech.frx":038C
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   1320
         Width           =   3615
      End
      Begin VB.ComboBox DcbCabinDor 
         Height          =   315
         ItemData        =   "FrmShowTech.frx":038E
         Left            =   120
         List            =   "FrmShowTech.frx":0390
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox TxtCabinHigh 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox DcbOperatingMethod2 
         Height          =   315
         ItemData        =   "FrmShowTech.frx":0392
         Left            =   120
         List            =   "FrmShowTech.frx":0394
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox TxtLoadPerson 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtLoadKG 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9960
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtSpead 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox DcbHoist 
         Height          =   315
         ItemData        =   "FrmShowTech.frx":0396
         Left            =   120
         List            =   "FrmShowTech.frx":0398
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   4095
      End
      Begin VB.ComboBox DcbElectricMotor 
         Height          =   315
         ItemData        =   "FrmShowTech.frx":039A
         Left            =   7560
         List            =   "FrmShowTech.frx":039C
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   600
         Width           =   3615
      End
      Begin VB.ComboBox DcbOperatingMethod 
         Height          =   315
         ItemData        =   "FrmShowTech.frx":039E
         Left            =   3000
         List            =   "FrmShowTech.frx":03A0
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox TxtCabinWidth 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9960
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtCabinDepth 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ãœ—«‰ «·œ«Œ·ÌÂ ··ﬂ»Ì‰…"
         Height          =   375
         Index           =   27
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«— ›«⁄ «·»«»"
         Height          =   255
         Index           =   23
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄—÷ «·»«»"
         Height          =   255
         Index           =   19
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«»Ê«» «·ÿÊ«»ﬁ"
         Height          =   255
         Index           =   18
         Left            =   11400
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»«» «·ﬂ»Ì‰Â"
         Height          =   255
         Index           =   17
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·«— ›«⁄"
         Height          =   255
         Index           =   16
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ„Ê·… —«ﬂ»"
         Height          =   255
         Index           =   24
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ„Ê·… «·„’⁄œ ﬂÃ"
         Height          =   255
         Index           =   26
         Left            =   11280
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”—⁄… «·„’⁄œ"
         Height          =   255
         Index           =   25
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·… «·—›⁄"
         Height          =   255
         Index           =   22
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Õ—ﬂ «·ﬂÂ—»«∆Ì"
         Height          =   255
         Index           =   21
         Left            =   11280
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÿ—Ìﬁ… «· ‘€Ì·"
         Height          =   255
         Index           =   20
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ»Ì‰… «·„’⁄œ ⁄—÷"
         Height          =   255
         Index           =   15
         Left            =   11160
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄„ﬁ"
         Height          =   255
         Index           =   14
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame lblDataCli 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Ê«’›«  «·⁄«„…"
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   12540
      Begin VB.TextBox TxtParkingCount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7680
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtHighAllWell 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7680
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtDeptdrilled 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtWellHigh 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Txtspace 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtWellDepth 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7680
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtWellWidth 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox DcbEngineRoom 
         Height          =   315
         ItemData        =   "FrmShowTech.frx":03A2
         Left            =   120
         List            =   "FrmShowTech.frx":03A4
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   960
         Width           =   5895
      End
      Begin VB.ComboBox DcbBildWell 
         Height          =   315
         ItemData        =   "FrmShowTech.frx":03A6
         Left            =   4800
         List            =   "FrmShowTech.frx":03A8
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox DcbDorDirect 
         Height          =   315
         ItemData        =   "FrmShowTech.frx":03AA
         Left            =   120
         List            =   "FrmShowTech.frx":03AC
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox TxtDorNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtDoorCount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TXtElevatorCount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«— ›«⁄ «·»∆— «·ﬂ·Ì"
         Height          =   255
         Index           =   12
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄„ﬁ Õ›—… «·»∆—"
         Height          =   255
         Index           =   11
         Left            =   11280
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«— ›«⁄ ”ﬁ› «·»∆—"
         Height          =   255
         Index           =   10
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„”«›… «·—Õ·…"
         Height          =   255
         Index           =   9
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«»⁄«œ «·»∆— ⁄„ﬁ"
         Height          =   255
         Index           =   8
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«»⁄«œ «·»∆— ⁄—÷"
         Height          =   255
         Index           =   7
         Left            =   11280
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Êﬁ⁄ €—›… «·„Õ—ﬂ« "
         Height          =   255
         Index           =   6
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«‰‘«¡ «·»∆—"
         Height          =   255
         Index           =   5
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "≈ Ã«Â «·«»Ê«»"
         Height          =   255
         Index           =   4
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·«»Ê«»"
         Height          =   255
         Index           =   3
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·„Ê«ﬁ›"
         Height          =   255
         Index           =   2
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·«»Ê«»"
         Height          =   255
         Index           =   1
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·„’«⁄œ"
         Height          =   255
         Index           =   13
         Left            =   11640
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   0
      Top             =   3600
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
      TabIndex        =   1
      Top             =   3600
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
      TabIndex        =   2
      Top             =   3600
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·„’«⁄œ"
      Height          =   255
      Index           =   0
      Left            =   9600
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "«·„Ê«’›«  «·›‰Ì…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12465
   End
End
Attribute VB_Name = "FrmShowTech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Dim NewGrid As New ClsGrid


Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
    save
    Unload Me

       Case 24
     '  AddNewFgRowother
       Case 25
          '  DeleteFgRowAther
    End Select

End Sub
Sub save()
Dim str As String

str = ""

 str = str & Trim(TxtElevatorCount.text) & "#"
 str = str & Trim(TxtParkingCount.text) & "#"
 str = str & Trim(TxtDoorCount.text) & "#"
 str = str & Trim(DcbDorDirect.ListIndex) & "#"
 str = str & Trim(TxtDorNo.text) & "#"
 str = str & Trim(TxtWellWidth.text) & "#"
 str = str & Trim(TxtWellDepth.text) & "#"
 str = str & Trim(DcbBildWell.ListIndex) & "#"
 str = str & Trim(Txtspace.text) & "#"
 str = str & Trim(TxtWellHigh.text) & "#"
 str = str & Trim(TxtDeptdrilled.text) & "#"
 str = str & Trim(TxtHighAllWell.text) & "#"
 str = str & Trim(DcbEngineRoom.ListIndex) & "#"
 str = str & Trim(TxtLoadKG.text) & "#"
 str = str & Trim(TxtLoadPerson.text) & "#"
 str = str & Trim(TxtSpead.text) & "#"
 str = str & Trim(DcbHoist.ListIndex) & "#"
 str = str & Trim(DcbElectricMotor.ListIndex) & "#"
 str = str & Trim(DcbOperatingMethod.ListIndex) & "#"
 str = str & Trim(DcbOperatingMethod2.ListIndex) & "#"
 str = str & Trim(TxtCabinWidth.text) & "#"
 str = str & Trim(TxtCabinDepth.text) & "#"
 str = str & Trim(TxtCabinHigh.text) & "#"
 str = str & Trim(DcbCabinDor.ListIndex) & "#"
 
  str = str & Trim(DcbDOrFloor.ListIndex) & "#"
 str = str & Trim(TxtDorWidth.text) & "#"
 str = str & Trim(TxtDorHight.text) & "#"
 str = str & Trim(TxtWalls.text) & "#"

 str = str & Trim("@")
  str = str & Chr(13)
  str = Trim(str)


FrmQotation.UnitsGrid.TextMatrix(FrmQotation.LngRow, FrmQotation.UnitsGrid.ColIndex("DataTech")) = str


End Sub



Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub



    
Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
Dim Xpid As Integer
Dim rw As Integer
    Set Dcombos = New ClsDataCombos
'Dcombos.GetItemsNames Me.DCboItemsName
   ' Set NewGrid.DtpBillDate = Me.XPDtbBill
   If FrmQotation.TxtModFlg.text = "R" Then
   Cmd(0).Enabled = False
   Else
   Cmd(0).Enabled = True
   End If
   
DcbDorDirect.AddItem "« Ã«Â Ê«Õœ"
DcbDorDirect.AddItem "« Ã«ÂÌ‰"
DcbDorDirect.AddItem "90 œ—Ã…"
DcbBildWell.AddItem "„‰ «·Œ—”«‰…"
DcbBildWell.AddItem "ÕÃ— „⁄ «·Œ—”«‰…"
DcbEngineRoom.AddItem "»«·√⁄·Ï „»«‘—… ›Êﬁ »∆— «·„’⁄œ"
DcbEngineRoom.AddItem "»œÊ‰ €—›…"
DcbEngineRoom.AddItem "»Ã‰» Õ›—… «·»∆— „»«‘—…"
DcbElectricMotor.AddItem "„‰ «·‰Ê⁄ «·ÕÀÌ À·«ÀÌ «·√ÿÊ«— ÊÌ „ «· Õﬂ„ »”—⁄… «·„’⁄œ »Ê«”ÿ… ‰Ÿ«„  €ÌÌ— «·ÃÂœ Ê«·–»–»…VVVF"
DcbElectricMotor.AddItem "„‰ «·‰Ê⁄ «·ÕÀÌ À·«ÀÌ «·√ÿÊ«— ÊÌ „ «· Õﬂ„ »”—⁄… «·„’⁄œ »Ê«”ÿ… «·„·›«  «·œ«Œ·Ì… Ê⁄œœÂ« 2 ”—Ì⁄ »ÿÌ∆"
DcbElectricMotor.AddItem "„÷Œ… ÂÌœ—Ê·ÌﬂÌ…  ⁄„· »÷Œ 120 · — »«·œﬁÌﬁ… Êÿ»ﬁ« ··„Ê«’›«  «·«Ê—Ê»Ì… »ﬁÊ… 15 Õ’«‰"
DcbHoist.AddItem "„«ﬂÌ‰… Ã— „‰ «·‰Ê⁄ «·œÊœÌ"
DcbHoist.AddItem "„«ﬂÌ‰… Ã— »œÊ‰ ÃÌ—(Gearless)"
DcbHoist.AddItem "⁄«„Êœ ⁄Ìœ—Ê·Ìﬂ"
DcbOperatingMethod.AddItem "ﬂÊ‰ —Ê·: ›—œÌ SIMPLEX"
DcbOperatingMethod.AddItem "„“œÊÃ/DUPLEX"
DcbOperatingMethod2.AddItem "ÃÂ«“ «· Õﬂ„ »«·”—⁄… AC-VVVF"
DcbOperatingMethod2.AddItem "ÃÂ«“ «· Õﬂ„ »«·”—⁄…speed AC"
DcbCabinDor.AddItem " › Õ ›Ì « Ã«Â Ê«Õœ Ì”«—"
DcbCabinDor.AddItem " › Õ ›Ì « Ã«Â Ê«Õœ Ì„Ì‰"
DcbCabinDor.AddItem "„‰ «·Ê”ÿ"
DcbDOrFloor.AddItem " › Õ ›Ì « Ã«Â Ê«Õœ Ì”«—"
DcbDOrFloor.AddItem " › Õ ›Ì « Ã«Â Ê«Õœ Ì„Ì‰"
DcbDOrFloor.AddItem "„‰ «·Ê”ÿ"

    Set DCboSearch = New clsDCboSearch
    Set DCboSearch = New clsDCboSearch
   
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
   
  Xpid = val((FrmQotation.XPTxtID.text))
rw = val(FrmQotation.UnitsGrid.TextMatrix(FrmQotation.LngRow, FrmQotation.UnitsGrid.ColIndex("id")))
If rw <> 0 Then

Retrive Xpid, rw
End If


    Set GrdBack = New ClsBackGroundPic

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

 
  '
End Sub


Public Sub Retrive(Optional id As Integer = 0, Optional Row As Integer = 0)

    Dim rs As ADODB.Recordset
Dim str As String
     Set rs = New ADODB.Recordset
    str = " SELECT     dbo.TblQutationsTech.*"
str = str & " From dbo.TblQutationsTech"
 str = str & " Where (QutationsID = " & id & ") And (QutationsIDDet =" & Row & " )"
 rs.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs.RecordCount > 0 Then
    TxtElevatorCount.text = IIf(IsNull(rs("ElevatorCount").value), "", val(rs("ElevatorCount").value))
    TxtParkingCount.text = IIf(IsNull(rs("ParkingCount").value), "", rs("ParkingCount").value)
    TxtDoorCount.text = IIf(IsNull(rs("DoorCount").value), "", rs("DoorCount").value)
    DcbDorDirect.ListIndex = IIf(IsNull(rs("DorDirect").value), -1, rs("DorDirect").value)
    TxtDorNo.text = IIf(IsNull(rs("DorNo").value), "", rs("DorNo").value)
    Me.TxtWellWidth.text = IIf(IsNull(rs("WellWidth").value), "", rs("WellWidth").value)
    Me.TxtWellDepth.text = IIf(IsNull(rs("WellDepth").value), "", rs("WellDepth").value)
    Me.DcbBildWell.ListIndex = IIf(IsNull(rs("BilldWell").value), -1, rs("BilldWell").value)
    Me.Txtspace.text = IIf(IsNull(rs("space").value), "", rs("space").value)
    Me.TxtWellHigh.text = IIf(IsNull(rs("WellHigh").value), "", rs("WellHigh").value)
    Me.TxtDeptdrilled.text = IIf(IsNull(rs("Deptdrilled").value), "", rs("Deptdrilled").value)
    Me.TxtHighAllWell.text = val(IIf(IsNull(rs("HighAllWell").value), "", rs("HighAllWell").value))
    Me.DcbEngineRoom.ListIndex = val(IIf(IsNull(rs("EngineRoom").value), -1, rs("EngineRoom").value))
    Me.TxtLoadKG.text = val(IIf(IsNull(rs("LoadKG").value), "", rs("LoadKG").value))
    Me.TxtLoadPerson.text = val(IIf(IsNull(rs("LoadPerson").value), "", rs("LoadPerson").value))
    Me.TxtSpead.text = val(IIf(IsNull(rs("Spead").value), "", rs("Spead").value))
    Me.DcbHoist.ListIndex = val(IIf(IsNull(rs("Hoist").value), -1, rs("Hoist").value))
    Me.DcbElectricMotor.ListIndex = val(IIf(IsNull(rs("ElectricMotor").value), -1, rs("ElectricMotor").value))
    Me.DcbOperatingMethod.ListIndex = val(IIf(IsNull(rs("OperatingMethod").value), -1, rs("OperatingMethod").value))
    Me.DcbOperatingMethod2.ListIndex = val(IIf(IsNull(rs("OperatingMethod2").value), -1, rs("OperatingMethod2").value))
    Me.TxtCabinWidth.text = val(IIf(IsNull(rs("CabinWidth").value), "", rs("CabinWidth").value))
    Me.TxtCabinDepth.text = val(IIf(IsNull(rs("CabinDepth").value), "", rs("CabinDepth").value))
    Me.TxtCabinHigh.text = val(IIf(IsNull(rs("CabinHigh").value), "", rs("CabinHigh").value))
    Me.DcbCabinDor.ListIndex = val(IIf(IsNull(rs("CabinDor").value), -1, rs("CabinDor").value))
    Me.DcbDOrFloor.ListIndex = val(IIf(IsNull(rs("DOrFloor").value), -1, rs("DOrFloor").value))
    Me.TxtDorWidth.text = val(IIf(IsNull(rs("DorWidth").value), "", rs("DorWidth").value))
    Me.TxtDorHight.text = val(IIf(IsNull(rs("DorHight").value), "", rs("DorHight").value))
    Me.TxtWalls.text = IIf(IsNull(rs("Walls").value), "", rs("Walls").value)
   End If
End Sub

