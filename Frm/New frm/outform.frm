VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form outform 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   7260
   Begin VB.Frame Frame16 
      Height          =   4455
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   7215
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   4215
         Left            =   0
         TabIndex        =   19
         Top             =   120
         Width           =   6735
         _cx             =   11880
         _cy             =   7435
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
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "‘Â—Ì|«”»Ê⁄Ì|ÌÊ„Ì"
         Align           =   0
         CurrTab         =   2
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   0   'False
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   3840
            Left            =   -7290
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   330
            Width           =   6645
            _cx             =   11721
            _cy             =   6773
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
            Begin VB.TextBox Text3 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   480
               TabIndex        =   21
               Top             =   480
               Width           =   1815
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   315
               Left            =   480
               TabIndex        =   22
               Top             =   840
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               _Version        =   393216
               Format          =   98304001
               CurrentDate     =   41640
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "Ì»œ« „‰  «—ÌŒ"
               Height          =   720
               Index           =   7
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   885
               Width           =   2370
            End
            Begin VB.Label Label62 
               Caption         =   "Õœœ ⁄œœ «·«”«»Ì⁄"
               Height          =   735
               Index           =   0
               Left            =   3000
               TabIndex        =   23
               Top             =   480
               Width           =   1695
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   3840
            Left            =   -7590
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   330
            Width           =   6645
            _cx             =   11721
            _cy             =   6773
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
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   480
               TabIndex        =   26
               Top             =   360
               Width           =   1815
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   480
               TabIndex        =   27
               Top             =   720
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               _Version        =   393216
               Format          =   98304001
               CurrentDate     =   41640
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "Ì»œ« „‰  «—ÌŒ"
               Height          =   240
               Index           =   3
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   765
               Width           =   1890
            End
            Begin VB.Label Label61 
               Caption         =   "Õœœ ⁄œœ «·‘ÂÊ—"
               Height          =   735
               Left            =   3000
               TabIndex        =   28
               Top             =   360
               Width           =   1695
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   3840
            Left            =   45
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   330
            Width           =   6645
            _cx             =   11721
            _cy             =   6773
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
            Begin VB.CheckBox chkDays 
               Height          =   255
               Index           =   0
               Left            =   3120
               TabIndex        =   38
               Top             =   480
               Width           =   375
            End
            Begin VB.CheckBox chkDays 
               Height          =   255
               Index           =   1
               Left            =   3120
               TabIndex        =   37
               Top             =   840
               Width           =   375
            End
            Begin VB.CheckBox chkDays 
               Height          =   255
               Index           =   2
               Left            =   3120
               TabIndex        =   36
               Top             =   1200
               Width           =   375
            End
            Begin VB.CheckBox chkDays 
               Height          =   255
               Index           =   3
               Left            =   3120
               TabIndex        =   35
               Top             =   1560
               Width           =   375
            End
            Begin VB.CheckBox chkDays 
               Height          =   255
               Index           =   4
               Left            =   3120
               TabIndex        =   34
               Top             =   1920
               Width           =   375
            End
            Begin VB.CheckBox chkDays 
               Height          =   255
               Index           =   5
               Left            =   3120
               TabIndex        =   33
               Top             =   2280
               Width           =   375
            End
            Begin VB.CheckBox chkDays 
               Height          =   255
               Index           =   6
               Left            =   3120
               TabIndex        =   32
               Top             =   2640
               Width           =   375
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·Œ„Ì”"
               Height          =   315
               Index           =   8
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   2385
               Width           =   1545
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·À·«À«¡"
               Height          =   330
               Index           =   6
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   1620
               Width           =   1545
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Õœ"
               Height          =   330
               Index           =   1
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   855
               Width           =   1545
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·Ã„⁄…"
               Height          =   180
               Index           =   9
               Left            =   1815
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   2685
               Width           =   1530
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«—»⁄«¡"
               Height          =   300
               Index           =   7
               Left            =   1815
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   2025
               Width           =   1530
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«À‰Ì‰"
               Height          =   285
               Index           =   4
               Left            =   1815
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   1260
               Width           =   1530
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·”» "
               Height          =   300
               Index           =   2
               Left            =   1815
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   480
               Width           =   1530
            End
         End
      End
      Begin VB.Label Label62 
         Caption         =   "Õ›Ÿ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   6720
         TabIndex        =   30
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox TxtEmp_Code 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check7 
      Alignment       =   1  'Right Justify
      Caption         =   "Check7"
      Height          =   255
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "›Ì Õ«·…  «·Œ—ÊÃ Ê«·⁄Êœ…"
      Height          =   1455
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Right Justify
         Caption         =   " „ «” ·«„ „” Õﬁ«  «·«Ã«“…"
         Height          =   255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   " „   ”·Ì„ «·⁄Âœ «· Ì ·œÌ…"
         Height          =   255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   600
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   " „   ”œ«œ «·”·›"
         Height          =   255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   960
         Value           =   1  'Checked
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " ÕœÌœ «·„œ…"
      Height          =   615
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   360
      Width           =   2415
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„œ…"
         Height          =   255
         Index           =   0
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ÌÊ„"
         Height          =   255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÿ»«⁄Â"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "outform.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "›Ì Õ«·… «·Œ—ÊÃ «·‰Â«∆Ì"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Width           =   2535
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   " „   ”œ«œ «·”·›"
         Height          =   255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   " „   ”·Ì„ «·⁄Âœ «· Ì ·œÌ…"
         Height          =   255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   " „  ’›Ì… „” Õﬁ« …"
         Height          =   255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "Œ—ÊÃ ‰Â«∆Ì"
      Height          =   195
      Index           =   2
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "⁄œÂ ”›—Ì« "
      Height          =   195
      Index           =   1
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "”›—… Ê«Õœ…"
      Height          =   195
      Index           =   0
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   240
      Value           =   -1  'True
      Width           =   1815
   End
End
Attribute VB_Name = "outform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim EmpReport As ClsEmployeeReport
Dim xReport As New CRAXDRT.Report

Function SHOWPIC(PICNAME As String)
    Dim xLogo As CRAXDRT.OLEObject
    '    On Error Resume Next
    StrFileName = App.path & "\Images\" & PICNAME & ".JPG"
If Dir(StrFileName) = "" Then Exit Function
    Set xLogo = xReport.Areas(3).Sections(1).AddPictureObject(StrFileName, 4000, 300)
    xLogo.Width = 1700
    xLogo.Height = 1700
    xLogo.backcolor = vbWhite
    xLogo.BorderColor = 255
    xLogo.CloseAtPageBreak = True
    '  xLogo.HyperlinkText = "BYTE"
    '  xLogo.HyperlinkType = crHyperlinkWebsite
    '  rep.Areas(1).Sections(1).SuppressIfBlank = True
    '  rep.Areas(1).Sections(1).Height = xLogo.Height + 250
 
End Function

Private Sub ALLButton1_Click()
    Dim xApp As New CRAXDRT.Application

    Dim rs As New ADODB.Recordset

'     If Frame2.Enabled = True And Frame3.Enabled = True Then
If Option1(0).value = True Or Option1(1).value = True Then
        If Not IsNumeric(Text1.Text) Then MsgBox "·«»œ „‰  ÕœÌœ «·„œ… »«·«Ì«„", vbCritical: Text1.SetFocus:   Exit Sub
        If Check4.value = 0 Or Check5.value = 0 Or Check6.value = 0 Then MsgBox "·« Ì„ﬂ‰ ·Â–« «·„ÊŸ› «·”›—   «·« »⁄œ «” ﬂ„«· „« ⁄·ÌÂ", vbCritical: Exit Sub
Else
        If Check1.value = 0 Or Check2.value = 0 Or Check3.value = 0 Then MsgBox "·« Ì„ﬂ‰ ·Â–« «·„ÊŸ› «·Œ—ÊÃ ‰Â«∆Ì« «·« »⁄œ «” ﬂ„«· „« ⁄·ÌÂ", vbCritical: Exit Sub

End If

        ' Emp_ID=" & Me.XPTxtEmpID.text
        
        sql = "SELECT * from emp_all_details WHERE Emp_ID=" & val(FrmEmployee.XPTxtEmpID.Text)
        rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText
    If Option1(0).value = True Then
   '     FrmReport.TxtPath = system_path & "\reports\emp\REPORT5.rpt"
         Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT5.rpt")
         ElseIf Option1(1).value = True Then
   '     FrmReport.TxtPath = system_path & "\reports\emp\REPORT6.rpt"
         Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT6.rpt")
        ElseIf Option1(2).value = True Then
   '     FrmReport.TxtPath = system_path & "\reports\emp\REPORT7.rpt"
         Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT7.rpt")
        End If
        
       
        xReport.Database.SetDataSource rs
 
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
   If Option1(0).value = True Then
        FrmReport.txtPath = system_path & "\reports\emp\REPORT5.rpt"
     '    Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT5.rpt")
         ElseIf Option1(1).value = True Then
        FrmReport.txtPath = system_path & "\reports\emp\REPORT6.rpt"
     '    Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT6.rpt")
        ElseIf Option1(2).value = True Then
        FrmReport.txtPath = system_path & "\reports\emp\REPORT7.rpt"
     '    Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT7.rpt")
        End If
        
'      xReport.ParameterFields(1).AddCurrentValue FrmEmployee.DcboJobsType3.Text
        FrmReport.CRViewer.viewReport
        
        FrmReport.show
        Screen.MousePointer = vbDefault
        xReport.reporttitle = Text1.Text
        
        
        SendKeys "{RIGHT}"

        If FrmEmployee.Check1.value = 1 Then
            SHOWPIC (Me.TxtEmp_Code.Text)
        End If
    
        '        Form3.Show
        '        Form3.case_id = 5
        '     Form3.noofmonth = Text1.text
        '      Form3.SHOWPICTURE.Caption = Me.Check7.value
        '       Form3.TxtEmp_Code = Me.TxtEmp_Code.text
Exit Sub
'        GoTo ll

'    End If

    If Frame2.Enabled = False And Frame3.Enabled = True Then
    
        If Check4.value = 0 Or Check5.value = 0 Or Check6.value = 0 Then MsgBox "·« Ì„ﬂ‰ ·Â–« «·„ÊŸ› «·”›—   «·« »⁄œ «” ﬂ„«· „« ⁄·ÌÂ", vbCritical: Exit Sub
  If Not IsNumeric(Text1.Text) Then MsgBox "·«»œ „‰  ÕœÌœ «·„œ… »«·«Ì«„", vbCritical: Exit Sub
  
        sql = "SELECT * from emp_all_details WHERE Emp_ID=" & val(FrmEmployee.XPTxtEmpID.Text)
        rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

        Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT6.rpt")
        xReport.Database.SetDataSource rs
 
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
  
        FrmReport.CRViewer.viewReport
        FrmReport.show
        FrmReport.txtPath = system_path & "\reports\emp\REPORT6.rpt"
        Screen.MousePointer = vbDefault
               xReport.reporttitle = Text1.Text
        SendKeys "{RIGHT}"

        If FrmEmployee.Check1.value = 1 Then
            'SHOWPIC (FrmEmployee.TxtEmp_Code.text)
            SHOWPIC (Me.TxtEmp_Code.Text)
        End If
  
        'Form3.Show
        'Form3.case_id = 6
        GoTo ll
    End If

    If Frame1.Enabled = True Then

        If Check1.value = 0 Or Check2.value = 0 Or Check3.value = 0 Then
            MsgBox "·« Ì„ﬂ‰ ·Â–« «·„ÊŸ› «·Œ—ÊÃ ‰Â«∆Ì« «·« »⁄œ «” ﬂ„«· „« ⁄·ÌÂ", vbCritical: Exit Sub
 
        Else
            sql = "SELECT * from emp_all_details WHERE Emp_ID=" & val(FrmEmployee.XPTxtEmpID.Text)
            rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

            Set xReport = xApp.OpenReport(system_path & "\reports\emp\REPORT7.rpt")
            xReport.Database.SetDataSource rs
 
            Set FrmReport = New FrmReportViewer
            FrmReport.CRViewer.ReportSource = xReport
  
            FrmReport.CRViewer.viewReport
            FrmReport.show
            Screen.MousePointer = vbDefault
            ' xReport.ReportTitle = X
            SendKeys "{RIGHT}"

            If FrmEmployee.Check1.value = 1 Then
                'SHOWPIC (FrmEmployee.TxtEmp_Code.text)
                SHOWPIC (Me.TxtEmp_Code.Text)
            End If

            'Form3.Show
            'Form3.case_id = 7
            GoTo ll
        End If
    End If

ll:
    Unload Me

End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
End Sub

Private Sub Label62_Click(Index As Integer)
Unload Me
End Sub

Private Sub Option1_Click(Index As Integer)

    Select Case Index

        Case 0
            Frame2.Enabled = True
            Frame1.Enabled = False
            Frame3.Enabled = True

        Case 1
            Frame2.Enabled = True
            Frame1.Enabled = False
            Frame3.Enabled = True

        Case 2
            Frame2.Enabled = False
            Frame1.Enabled = True
            Frame3.Enabled = False

    End Select

End Sub
