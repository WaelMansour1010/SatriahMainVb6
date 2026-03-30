VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmExpensesType 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "√‰Ê«⁄ «·„’—Ê›« "
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7260
   Icon            =   "FrmExpensesType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   7260
   Begin VB.CheckBox chkTransportation 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„’«—Ì› ‰ﬁ·Ì« "
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   4050
      Width           =   1695
   End
   Begin VB.CheckBox chkComposeExpenses 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„’«—Ì›  ”ÊÌﬁ"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   3690
      Width           =   1695
   End
   Begin VB.TextBox TxtAccount_Serial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   2160
      Width           =   1245
   End
   Begin VB.CheckBox chkIndirectCosts 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ﬂ«·Ì› €Ì— „»«‘—…"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3690
      Width           =   1695
   End
   Begin VB.CheckBox ChkTypicalProduction 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ì »⁄ «·«‰ «Ã «·‰„ÿÌ"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   3690
      Width           =   1695
   End
   Begin VB.TextBox XPTxtBankNamee 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      MaxLength       =   1000
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1440
      Width           =   5685
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   660
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox XPTxtBankID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4950
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   885
   End
   Begin VB.TextBox XPTxtBankName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      MaxLength       =   1000
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1095
      Width           =   5685
   End
   Begin VB.TextBox XPMTxtRemark 
      Alignment       =   1  'Right Justify
      Height          =   435
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3090
      Width           =   5685
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   675
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   7155
      _cx             =   12621
      _cy             =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   20.25
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
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "√‰Ê«⁄ «·„’—Ê›« "
      Align           =   0
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
      Begin VB.CheckBox ChkManual 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " ⁄—Ì› ÌœÊÌ"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   4
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmExpensesType.frx":038A
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   2
         Left            =   90
         TabIndex        =   5
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmExpensesType.frx":0724
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   6
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmExpensesType.frx":0ABE
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         AlignmentVertical=   1
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   3
         Left            =   615
         TabIndex        =   7
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmExpensesType.frx":0E58
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   2160
         Picture         =   "FrmExpensesType.frx":11F2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   5430
      TabIndex        =   8
      Top             =   4485
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÃœÌœ"
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
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   9
      Top             =   4485
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   " ⁄œÌ·"
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
      Index           =   2
      Left            =   3795
      TabIndex        =   10
      Top             =   4485
      Width           =   675
      _ExtentX        =   1191
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
      Index           =   3
      Left            =   3045
      TabIndex        =   11
      Top             =   4485
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   " —«Ã⁄"
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
      Index           =   4
      Left            =   2280
      TabIndex        =   12
      Top             =   4485
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   661
      ButtonPositionImage=   1
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
      Left            =   450
      TabIndex        =   13
      Top             =   4485
      Width           =   675
      _ExtentX        =   1191
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   23
      Top             =   4485
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
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
   Begin MSDataListLib.DataCombo DboParentAccount 
      Height          =   315
      Left            =   120
      TabIndex        =   28
      Top             =   1800
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DboAcc 
      Height          =   315
      Left            =   120
      TabIndex        =   30
      Top             =   2160
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo cmbDataTypeExchange 
      Height          =   315
      Left            =   120
      TabIndex        =   35
      Top             =   2580
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483624
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„’—Ê›"
      Height          =   285
      Index           =   5
      Left            =   5880
      TabIndex        =   36
      Top             =   2550
      Width           =   1140
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Õ”«»  "
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   5
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   2160
      Width           =   1125
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ «‰Ã·Ì“Ì"
      Height          =   315
      Index           =   4
      Left            =   5970
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Õ”«» «·—∆Ì”Ì"
      Height          =   315
      Index           =   3
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1800
      Width           =   1125
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬂÊœ"
      Height          =   315
      Index           =   2
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   720
      Width           =   885
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ ⁄—»Ì"
      Height          =   315
      Index           =   1
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1095
      Width           =   885
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   315
      Index           =   0
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3000
      Width           =   885
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   4080
      Width           =   1245
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   4080
      Width           =   1365
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   4080
      Width           =   705
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   2970
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   4080
      Width           =   825
   End
End
Attribute VB_Name = "FrmExpensesType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim ScreenNameArabic As String

Dim ScreenNameEnglish As String


Private Sub ChkManual_Click()
If ChkManual = vbChecked Then
DboParentAccount.Enabled = False
DboAcc.Enabled = True
XPTxtBankName.Enabled = False
Else
DboParentAccount.Enabled = True
DboAcc.Enabled = False
XPTxtBankName.Enabled = True
End If

End Sub

Private Sub Cmd_Click(index As Integer)
     On Error GoTo ErrTrap

    Select Case index

        Case 0
            TxtModFlg.text = "N"
            clear_all Me
            XPTxtBankID.text = CStr(new_id("ExpensesType", "ID", "", True))
            '        XPTxtBankName.SetFocus
        
            Dim Account_Code_dynamic As String
            Account_Code_dynamic = get_account_code_branch(33, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                      Else
                      MsgBox "Define Branch Firstly", vbCritical
                      End If
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» ··„’—Ê›«    ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                 Else
                 MsgBox "Expense Account Not Defined", vbCritical
                 End If
       
                End If
            End If
        
            DboParentAccount.BoundText = Account_Code_dynamic
        
        Case 1
            TxtModFlg.text = "E"
            DboParentAccount.Enabled = False
            CuurentLogdata

        Case 2
            SaveData

        Case 3
            Call Undo

        Case 4
            Del_ExpensesType

        Case 5
            FrmExpensesSearch.show
            FrmExpensesSearch.RetrunType = 0
            FrmExpensesSearch.Indx = 0

        Case 6
            Unload Me
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub DboAcc_Change()
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
DboParentAccount.BoundText = Get_Account_Parent_code(DboAcc.BoundText)
XPTxtBankName.text = DboAcc.text
End If
End Sub

Private Sub DboAcc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 161115
    End If
End Sub

Private Sub DboParentAccount_KeyUp(KeyCode As Integer, _
                                   Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 55
    End If

End Sub

Private Sub Form_Activate()
    XPTxtBankID.SetFocus
End Sub
 
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ " & XPTxtBankID.text & CHR(13) & "   «·«”„ " & XPTxtBankName & CHR(13) & "   «·Õ”«» «·—∆Ì”Ì " & DboParentAccount & CHR(13) & "   „·«ÕŸ«  " & XPMTxtRemark

    If ChkTypicalProduction.value = Checked Then
        LogTextA = LogTextA & CHR(13) & "   Ì »⁄ «·«‰ «Ã «·‰„ÿÌ  "
    End If
         
    If chkIndirectCosts.value = Checked Then
        LogTextA = LogTextA & CHR(13) & "   ﬂ«·Ì› €Ì— „»«‘—… "
    End If
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Code  " & XPTxtBankID.text & CHR(13) & "   Name " & XPTxtBankNamee & CHR(13) & "Parent Account  " & DboParentAccount & CHR(13) & "   Remarks " & XPMTxtRemark

    If ChkTypicalProduction.value = Checked Then
        LogTexte = LogTexte & CHR(13) & "    According To Typical Production"
    End If
         
    If chkIndirectCosts.value = Checked Then
        LogTexte = LogTexte & CHR(13) & " According To  InDirect Cost "
    End If
                     
         If ChkManual.value = Checked Then
        LogTexte = LogTexte & CHR(13) & " According To Manual Entry"
    End If
    
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.text = "R" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
            XPBtnMove_Click (2)
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
            XPBtnMove_Click (1)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
            XPBtnMove_Click (3)
        ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
            XPBtnMove_Click (0)
        End If
    End If

    If KeyCode = vbKeyF12 Then
        If Cmd(0).Enabled = False Then Exit Sub
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If Cmd(1).Enabled = False Then Exit Sub
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If Cmd(2).Enabled = False Then Exit Sub
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If Cmd(3).Enabled = False Then Exit Sub
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    ScreenNameArabic = " √‰Ê«⁄ «·„’—Ê›« "
    ScreenNameEnglish = "Expenses Types"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    Dim Dcombos As New ClsDataCombos

    Dcombos.GetAccountingCodes Me.DboParentAccount, False, True, 3
    Dcombos.GetAccountingCodes Me.DboAcc, True, False
     
    
    
       Dim StrSQL As String
        StrSQL = "SELECT id,name From TblDataTypeExchange "
        fill_combo cmbDataTypeExchange, StrSQL
    
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
    rs.Open "ExpensesType", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    Me.TxtModFlg.text = "R"
    Retrive

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

            Case "N"
    
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & CHR(13)
                    StrMSG = StrMSG & " the new data  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & CHR(13)
                    StrMSG = StrMSG & " the Modifications  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
                
                End If

        End Select

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

        Select Case IntResult

            Case vbYes
                Cancel = True
       
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '   Me.Caption = "√‰Ê«⁄ «·„’—Ê›« "
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            DboParentAccount.Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        ChkManual.Enabled = False
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            Me.XPTxtBankID.locked = True
            Me.XPTxtBankName.locked = True
            Me.XPMTxtRemark.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

        Case "N"
            '      Me.Caption = "√‰Ê«⁄ «·„’—Ê›« ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        ChkManual.Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
        
            Me.XPTxtBankID.locked = True
            Me.XPTxtBankName.locked = False
            Me.XPMTxtRemark.locked = False
            DboParentAccount.Enabled = True

        Case "E"
            '      Me.Caption = "√‰Ê«⁄ «·„’—Ê›« (  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            DboParentAccount.Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        ChkManual.Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            Me.XPTxtBankID.locked = True
            Me.XPTxtBankName.locked = False
            Me.XPMTxtRemark.locked = False
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    On Error GoTo ErrTrap
    Dim i As Integer

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount

            If rs("ID").value = Lngid Then
                GoTo ll
            End If

            rs.MoveNext
        Next i

        Exit Sub
    End If

ll:
    XPTxtBankID.text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    XPTxtBankName.text = IIf(IsNull(rs("Name").value), "", Trim(rs("Name").value))
    XPTxtBankNamee.text = IIf(IsNull(rs("Namee").value), "", Trim(rs("Namee").value))

    XPMTxtRemark.text = IIf(IsNull(rs("Remarks").value), "", Trim(rs("Remarks").value))

    'DboParentAccount.Text = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
    DboParentAccount.BoundText = Get_Account_Parent_code(IIf(IsNull(rs("Account_Code").value), "", Trim(rs("Account_Code").value)))
    cmbDataTypeExchange.BoundText = IIf(IsNull(rs("DataTypeExchangeCode").value), "", Trim(rs("DataTypeExchangeCode").value))
    
DboAcc.BoundText = IIf(IsNull(rs("Account_Code").value), "", Trim(rs("Account_Code").value))
tXTAccount_Serial.text = getAccountSerial_Code("Account_Serial", "Account_Code", IIf(IsNull(rs("Account_Code").value), "", Trim(rs("Account_Code").value)))
    If rs("TypicalProduction").value = vbTrue Then
        Me.ChkTypicalProduction.value = vbChecked
    Else
        Me.ChkTypicalProduction.value = vbUnchecked
    End If


    If rs("Transportation").value = vbTrue Then
        Me.chkTransportation.value = vbChecked
    Else
        Me.chkTransportation.value = vbUnchecked
    End If



    If rs("IndirectCosts").value = vbTrue Then
        Me.chkIndirectCosts.value = vbChecked
    Else
        Me.chkIndirectCosts.value = vbUnchecked
    End If
    
        If rs("ComposeExpenses").value = 1 Then
        Me.chkComposeExpenses.value = vbChecked
    Else
        Me.chkComposeExpenses.value = vbUnchecked
    End If
    

    If rs("ManualEntrty").value = 1 Then
        Me.ChkManual.value = vbChecked
        DboParentAccount.Enabled = False
DboAcc.Enabled = True
    Else
        Me.ChkManual.value = vbUnchecked
        DboParentAccount.Enabled = True
DboAcc.Enabled = False
    End If
 
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnMove_Click(index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case index

        Case 0

            If Not (rs.EOF Or rs.BOF) Then
                rs.MovePrevious

                If rs.BOF Then rs.MoveFirst
            End If

        Case 1

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
            End If

        Case 2

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If

        Case 3

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveNext

                If rs.EOF Then rs.MoveLast
            End If

    End Select

    Retrive
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
     On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If XPTxtBankName.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· ‰Ê⁄ «·„’—Ê›«  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         Else
         MsgBox "Enter Expenses Name ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
         
            XPTxtBankName.SetFocus
            Exit Sub
        End If
 
        If DboParentAccount.BoundText = "" And ChkManual.value = vbUnchecked Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "„‰ ›÷·ﬂ «Œ — «·Õ”«» «·—∆Ì”Ì ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                 Else
                 MsgBox "Select Parent Account ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                 End If
                    DboParentAccount.SetFocus
                    Sendkeys ("{F4}")
                    Exit Sub
        End If
    
        Select Case Me.TxtModFlg.text

            Case "N"
                StrSQL = "select * From  ExpensesType where Name='" & Trim(XPTxtBankName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "Â‰«ﬂ ‰Ê⁄ „’—Ê›«  „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ ‰Ê⁄ «·„’—Ê›«  «·„Õœœ"
                          Else
                          Msg = "Expenses With same name Exist " & CHR(13)
                            Msg = Msg + " Change the name  " & CHR(13)
                        
                            
                          End If
                            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTxtBankName.SetFocus
                    Exit Sub
                End If

            Case "E"
                StrSQL = "select * From  ExpensesType where Name='" & Trim(XPTxtBankName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("ID").value <> val(XPTxtBankID.text) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "Â‰«ﬂ ‰Ê⁄ „’—Ê›«  „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ ‰Ê⁄ «·„’—Ê›«  «·„Õœœ"
                     Else
                     Msg = "Expenses With same name Exist " & CHR(13)
                            Msg = Msg + " Change the name  " & CHR(13)
                        
                     End If
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTxtBankName.SetFocus
                        Exit Sub
                    End If
                End If

        End Select

        Cn.BeginTrans
        BeginTrans = True

        Select Case Me.TxtModFlg.text

            Case "N"
        
                rs.AddNew
                rs("ID").value = val(XPTxtBankID.text)

            Case "E"
        End Select

        rs("Name").value = Trim(XPTxtBankName.text)
        rs("Namee").value = Trim(XPTxtBankNamee.text)
     
        rs("Remarks").value = IIf(XPMTxtRemark.text = "", Null, Trim(XPMTxtRemark.text))
        rs("parent_account").value = IIf(DboParentAccount.BoundText = "", Null, (DboParentAccount.text))
        rs("DataTypeExchangeCode").value = IIf(cmbDataTypeExchange.BoundText = "", Null, val(cmbDataTypeExchange.BoundText))
        
        If Me.ChkTypicalProduction.value = vbChecked Then
            rs("TypicalProduction").value = 1
        ElseIf Me.ChkTypicalProduction.value = vbUnchecked Then
            rs("TypicalProduction").value = 0
        End If


        If Me.chkTransportation.value = vbChecked Then
            rs("Transportation").value = 1
        ElseIf Me.ChkTypicalProduction.value = vbUnchecked Then
            rs("Transportation").value = 0
        End If


        If Me.chkIndirectCosts.value = vbChecked Then
            rs("IndirectCosts").value = 1
        ElseIf Me.ChkTypicalProduction.value = vbUnchecked Then
            rs("IndirectCosts").value = 0
        End If
        
             If Me.chkComposeExpenses.value = vbChecked Then
            rs("ComposeExpenses").value = 1
        ElseIf Me.ChkTypicalProduction.value = vbUnchecked Then
            rs("ComposeExpenses").value = 0
        End If
        
        If Me.ChkManual.value = vbChecked Then
            rs("ManualEntrty").value = 1
        ElseIf Me.ChkTypicalProduction.value = vbUnchecked Then
            rs("ManualEntrty").value = 0
        End If




        If ChkManual.value = vbUnchecked Then
                        If Me.TxtModFlg.text = "N" Then
                            rs("Account_Code").value = ModAccounts.AddNewAccount(DboParentAccount.BoundText, Trim$(Me.XPTxtBankName.text), True, False, XPTxtBankNamee.text, , , , , , , , , , 2, 3, 0, 0, 0)
                            '   Rs("Account_Code").value = ModAccounts.AddNewAccount("a3a1", Trim$(Me.XPTxtBankName.text), True, False)
                        Else
            
                                    If Not IsNull(rs("Account_Code").value) Then
                                        ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtBankName.text, Me.XPTxtBankNamee.text, , , , , , , , , 2, 3, 0, 0, 0, , , , True
                                    End If
                        End If
          Else
          
          rs("Account_Code").value = DboAcc.BoundText
        End If

        rs.update
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        CuurentLogdata

        Select Case Me.TxtModFlg.text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·‰Ê⁄" & CHR(13)
                                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                             Else
                             Msg = " saved " & CHR(13)
                                Msg = Msg + "Ener Another record yes/no? "
                             
                             End If
                Else
        
                    Msg = "Saved" & CHR(13)
                    Msg = Msg + "Do you want enter another One"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If

        End Select

        TxtModFlg.text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "ID=" & val(XPTxtBankID.text) & "", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_ExpensesType()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset

    On Error GoTo ErrTrap

    If XPTxtBankID.text <> "" Then
        StrSQL = "select * From Notes where ExpensesID=" & Trim(XPTxtBankID.text)
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·‰Ê⁄" & CHR(13)
            Msg = Msg + "· ﬂ«„· «·»Ì«‰« "
        Else
        Msg = "Can't Delete This Type" & CHR(13)
            Msg = Msg + "data integration"
        End If
            
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·‰Ê⁄ —ﬁ„ " & CHR(13)
        Msg = Msg + (XPTxtBankID.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
Msg = "Confirm Delete this type" & CHR(13)
        Msg = Msg + (XPTxtBankID.text) & CHR(13)
        Msg = Msg + " yes / no  ?"

End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                Dim StrAccountCode As String
                StrAccountCode = rs("Account_Code").value

                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
                    CuurentLogdata ("D")
                    rs.delete
                Else
                    Exit Sub
                End If

                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then

        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        Msg = "No record"
        End If
        
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        '
                                   If SystemOptions.UserInterface = ArabicInterface Then
                          Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·‰Ê⁄ "
                             Else
                         Msg = "Cant Delete this Record " & CHR(13) & "Data integration "
                             End If
 
 
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ‰Ê⁄ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–« «·‰Ê⁄" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·‰Ê⁄ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·‰Ê⁄" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Me.Caption = "Expenses Type"
    Me.Ele.Caption = Me.Caption
    lbl(3).Caption = "Parent Acc"
    Cmd(5).Caption = "Search"
    Me.lbl(0).Caption = "Comment"
 ChkManual.Caption = "Manual"
 lbl(5).Caption = "Account"
 chkComposeExpenses.Caption = "Marketing EXP."
    Me.lbl(1).Caption = "Ar Name"
    Me.lbl(4).Caption = "En Name"
chkIndirectCosts.Caption = "Indirect Costs"
    ChkTypicalProduction.Caption = "Typical Production"
    Me.lbl(2).Caption = "Code"
    Me.lbl(7).Caption = "Current Record:"
    Me.lbl(6).Caption = "Records NO:"
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(6).Caption = "Exit"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    'XPBtnMove(2).RightToLeft = False
End Sub

Private Sub XPTxtBankName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtBankNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
