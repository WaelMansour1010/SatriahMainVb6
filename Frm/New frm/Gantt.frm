VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F871B372-BD4B-4283-A10E-0AB1C61FA941}#1.0#0"; "JanChart.ocx"
Begin VB.Form Gantt 
   Caption         =   "ÇáÌÇäÊ"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18960
   Icon            =   "Gantt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   18960
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
      Height          =   9315
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   18960
      _cx             =   33443
      _cy             =   16431
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
      Align           =   5
      AutoSizeChildren=   1
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
      Begin C1SizerLibCtl.C1Elastic Header 
         Height          =   1155
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   18960
         _cx             =   33443
         _cy             =   2037
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
         Align           =   1
         AutoSizeChildren=   3
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
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   " ÇáÌÇäÊ åæ äæÚ ãä ÇáÊÎØíØ íæÖÍ ÇáÌÏæá ÇáÒãäí ááÊäÝíÐ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Tag             =   $"Gantt.frx":000C
            Top             =   615
            Width           =   18780
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "ÇáÌÇäÊ      GANTT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   465
            Left            =   90
            TabIndex        =   2
            Top             =   90
            Width           =   18780
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   1155
         Left            =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   8160
         Width           =   18960
         _cx             =   33443
         _cy             =   2037
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
         Align           =   2
         AutoSizeChildren=   3
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
         Begin VB.Frame Frame1 
            Caption         =   "ÏáÇáÇÊ ÇáÇáæÇä"
            Height          =   975
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   90
            Width           =   18780
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H000000FF&
               Caption         =   "ÍÑÌ"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   375
               Left            =   17160
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   240
               Width           =   1215
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Main 
         Height          =   7005
         Left            =   0
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1155
         Width           =   18960
         _cx             =   33443
         _cy             =   12356
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
         Align           =   5
         AutoSizeChildren=   3
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
         Begin prjJanChart.JanChart JanChart1 
            Height          =   6825
            Left            =   90
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   90
            Width           =   18780
            _ExtentX        =   33126
            _ExtentY        =   12039
            BackColor       =   -2147483633
            ChartWidth      =   18525
            ChartHeight     =   6570
            CaptionNumber   =   1
            CaptionBackColor=   8438015
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            WorksheetNumber =   1
            WorksheetGroupNumber=   5
            WorksheetWidth  =   400
            BeginProperty DataFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "Gantt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************************************************************
' Hi,
' This is VB project for JanChart demo
' The chart will have:
'   2 column for caption (IO Number, Date In)
'   4 column for data (1, 2, 3, 4)
'   20 row for data (1, 2, ..., 20)
'
' If you need some help or have some question or bug report
' please feel free to contact me
' My email address is: jimmi_kembaren@sancerta.com
'
' Btw, I live in Bandung - Java
'
' Chrs,
' Jimmi A. Kembaren
'**********************************************************************

Private Sub cmdAbout_Click(Index As Integer)
     
End Sub

Private Sub CmdExit_Click(Index As Integer)
    End
End Sub

Private Sub Form_Load()
  
    If SystemOptions.UserInterface = EnglishInterface Then
        ChangeLang
    End If
  
End Sub

Private Function ChangeLang()
    SetInterface Me
    Me.Caption = "GANTT"

    Label9.Caption = Me.Caption
    Frame1.Caption = "Color Map"
    Label1.Caption = "Critical"
End Function

Public Sub Init_Chart(weeknos As Integer)
On Error Resume Next
    'init gantt chart
    Dim i As Integer
    JanChart1.CaptionNumber = 10                   'set caption column 2
    JanChart1.Set_CaptionWidth 1, 800              'set 1st caption width
    JanChart1.Set_CaptionWidth 2, 2000              'set 2nd caption width
If weeknos > 60 Then weeknos = 60
    For i = 3 To 10
        JanChart1.Set_CaptionWidth i, 1000
    Next i
    
    JanChart1.Set_CaptionName 1, "id"        'set 1st caption label
    JanChart1.Set_CaptionName 2, "Brief"          'set 2nd caption label
    JanChart1.Set_CaptionName 3, "Duration"          'set 2nd caption label
    JanChart1.Set_CaptionName 4, "Based On"          'set 2nd caption label
    JanChart1.Set_CaptionName 5, "Early Start"          'set 2nd caption label
    JanChart1.Set_CaptionName 6, "Start "          'set 2nd caption label
    JanChart1.Set_CaptionName 7, "Early End"          'set 2nd caption label
    JanChart1.Set_CaptionName 8, "End "          'set 2nd caption label
    JanChart1.Set_CaptionName 9, "Crash Time "          'set 2nd caption label
    JanChart1.Set_CaptionName 10, "Critical "          'set 2nd caption label
    JanChart1.WorksheetNumber = weeknos                 'set worksheet column number
    JanChart1.WorksheetWidth = 1300                 'set worksheet column width
    
    JanChart1.WorksheetGroupNumber = weeknos              'set worksheet column group
    JanChart1.Set_WorksheetGroupLabel 1, getoprTitle    'set worksheet column group label
   
    '   JanChart1.Set_WorksheetLabel 1, "1"             'set 1st worksheet column label
    For i = 1 To weeknos
        JanChart1.Set_WorksheetLabel i, "" & i & ""             'set 1st worksheet column label
    Next i

    '   JanChart1.Set_WorksheetLabel 2, "2"             'set 2st worksheet column label
    '   JanChart1.Set_WorksheetLabel 3, "3"             'set 3st worksheet column label
    '   JanChart1.Set_WorksheetLabel 4, "4"             'set 4st worksheet column label
    
    JanChart1.Refresh
End Sub 'init_chart

Public Sub Draw_Data(current_terms As String)

    'draw chart data
    If current_terms = "" Then
        Exit Sub
    End If

    Dim i As Integer
    Dim intNumData As Integer
    Dim strIONo As String, strDateIn As String
    Dim dblBarWidth1 As Double, dblBarStart1 As Double
    Dim dblBarWidth2 As Double, dblBarStart2 As Double
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim StrSQL As String
    'StrSQL = "select * from terms_operations Where term_fullcode ='" & current_terms & "'" ' Val(Me.txt_project_id.text) & "AND item_id=" & current_terms"
    
    StrSQL = "SELECT     *, dbo.TblProcessDEF.ProcessNameE, dbo.TblProcessDEF.ProcessName"
StrSQL = StrSQL & " FROM         dbo.terms_operations INNER JOIN"
StrSQL = StrSQL & "                       dbo.TblProcessDEF ON dbo.terms_operations.id = dbo.TblProcessDEF.TblProcessDEFID"
           StrSQL = StrSQL & " Where term_fullcode ='" & current_terms & "'" ' Val(Me.txt_project_id.text) & "AND item_id=" & current_terms"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.RecordCount > 0 Then
        'set number of data (number of row)
        intNumData = rs.RecordCount
        JanChart1.DataNumber = intNumData
    
        'draw every rows
        For i = 1 To rs.RecordCount
       
            JanChart1.Set_DataCaptionLabel i, 1, IIf(IsNull(rs("fullcode").value), "", rs("fullcode").value)       'set io number for this row
            JanChart1.Set_DataCaptionLabel i, 2, IIf(IsNull(rs("ProcessNameE").value), "", rs("ProcessNameE").value)       'set io number for this row
            JanChart1.Set_DataCaptionLabel i, 3, IIf(IsNull(rs("period").value), 0, rs("period").value)       'set io number for this row
            JanChart1.Set_DataCaptionLabel i, 4, IIf(IsNull(rs("Pre").value), "", rs("Pre").value)       'set io number for this row
            JanChart1.Set_DataCaptionLabel i, 5, IIf(IsNull(rs("EarlyStartWeek").value), 0, rs("EarlyStartWeek").value)       'set io number for this row
            JanChart1.Set_DataCaptionLabel i, 6, IIf(IsNull(rs("StartWeek").value), 0, rs("StartWeek").value)       'set io number for this row
            JanChart1.Set_DataCaptionLabel i, 7, IIf(IsNull(rs("EarlyEndWeek").value), 0, rs("EarlyEndWeek").value)       'set io number for this row
            JanChart1.Set_DataCaptionLabel i, 8, IIf(IsNull(rs("EndWeek").value), 0, rs("EndWeek").value)       'set io number for this row
            JanChart1.Set_DataCaptionLabel i, 9, IIf(IsNull(rs("Period1").value), 0, rs("Period1").value)       'set io number for this row
            JanChart1.Set_DataCaptionLabel i, 10, IIf(IsNull(rs("Critical").value), 0, rs("Critical").value)       'set io number for this row
        
            dblBarStart1 = IIf(IsNull(rs("EarlyStartWeek").value), 0, rs("EarlyStartWeek").value)
            dblBarWidth1 = dblBarStart1 + IIf(IsNull(rs("Period").value), 0, rs("Period").value)
        
            If (dblBarStart1 < JanChart1.WorksheetNumber) Then
                'draw bar
                JanChart1.Set_DataWorksheet i, dblBarStart1, dblBarWidth1, &H80FF80  ' &HFF8080
            End If
        
            dblBarStart2 = IIf(IsNull(rs("StartWeek").value), 0, rs("StartWeek").value)
            dblBarWidth2 = dblBarStart2 + IIf(IsNull(rs("Period").value), 0, rs("Period").value)           'create randomize data

            '
            If (dblBarStart2 < JanChart1.WorksheetNumber) Then
                'draw bar
                JanChart1.Set_DataWorksheet i, dblBarStart2, dblBarWidth2, &HFF&
            End If

            rs.MoveNext
        Next i

    End If

    JanChart1.Refresh
        
End Sub 'draw_data
 
