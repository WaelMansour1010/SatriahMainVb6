VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»ÕÀ"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   Icon            =   "FrmSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2835
   ScaleWidth      =   4650
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   2835
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4695
      _cx             =   8281
      _cy             =   5001
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
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈ Ã«Â «·»ÕÀ"
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
         Height          =   1035
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   510
         Width           =   3945
         Begin VB.OptionButton OptForward 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "··√”›·(„‰ «·’›Õ… «·Õ«·Ì… ≈·Ï «·’›Õ«  «· «·Ì…)"
            Height          =   345
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   570
            Value           =   -1  'True
            Width           =   3645
         End
         Begin VB.OptionButton OptBackward 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "··√⁄·Ï( „‰ «·’›Õ… «·Õ«·Ì… ≈·Ï «·’›Õ«  «·”«»ﬁ…)"
            Height          =   345
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   240
            Width           =   3645
         End
      End
      Begin VB.TextBox TxtSearch 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   690
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   150
         Width           =   3075
      End
      Begin ImpulseButton.ISButton CmdSearch 
         Height          =   375
         Left            =   1050
         TabIndex        =   2
         Top             =   2400
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
         ButtonStyle     =   1
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
         ButtonImage     =   "FrmSearch.frx":038A
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton CmdExit 
         Height          =   405
         Left            =   60
         TabIndex        =   3
         Top             =   2370
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   714
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
         ButtonImage     =   "FrmSearch.frx":0724
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin VB.Label LblCurrentPageNumber 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1560
         Width           =   465
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·’›Õ… «·Õ«·Ì…:"
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
         Height          =   255
         Index           =   1
         Left            =   510
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1590
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·»ÕÀ ⁄‰"
         Height          =   195
         Index           =   0
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   795
      End
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Viewer As CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer

Private m_Report As CRAXDRT.Report

Dim m_PageGen  As CRAXDRT.PageGenerator
Dim m_PageEngine As CRAXDRT.PageEngine

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSearch_Click()
    Dim StrVal As String
    Dim BolResult As Boolean
    Dim x As Variant
    Dim x_Report As CRAXDRT.Report
    Dim LngStartPage As Long
    Dim Msg As String
    Static LngOldPage As Long

    StrVal = TxtSearch.text
    x = Array()

    Set m_PageEngine = Me.Report.PageEngine
    Set m_PageGen = m_PageEngine.CreatePageGenerator(x)
    LngStartPage = val(Me.LblCurrentPageNumber.Caption)

    If Me.OptForward.value = True Then
        BolResult = m_PageGen.FindText(StrVal, crForward, LngStartPage)
    Else
        BolResult = m_PageGen.FindText(StrVal, crBackward, LngStartPage)
    End If
 
    If BolResult = True Then
        Me.Viewer.ShowNthPage LngStartPage
        Me.Viewer.SearchForText StrVal
        Me.LblCurrentPageNumber.Caption = LngStartPage
    Else
        Msg = "·„ Ì „ «·⁄ÀÊ— ⁄·Ï «Ï »Ì«‰ „ÿ«»ﬁ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If

End Sub

Public Property Get Viewer() As CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer
    Set Viewer = m_Viewer
End Property

Public Property Set Viewer(ByVal vNewValue As CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer)
    Set m_Viewer = vNewValue
    Me.LblCurrentPageNumber.Caption = m_Viewer.GetCurrentPageNumber()
End Property

Public Property Get Report() As CRAXDRT.Report
    Set Report = m_Report
End Property

Public Property Set Report(ByVal vNewValue As CRAXDRT.Report)
    Set m_Report = vNewValue
End Property

