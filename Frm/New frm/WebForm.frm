VERSION 5.00
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form WebForm 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   ClientHeight    =   16200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   26070
   Enabled         =   0   'False
   Icon            =   "WebForm.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   16200
   ScaleWidth      =   26070
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
   Begin VB.Timer info1Timer 
      Interval        =   500
      Left            =   3120
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Interval        =   500
      Left            =   9360
      Top             =   3240
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   120
      Top             =   5160
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3960
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1440
      Top             =   2760
   End
   Begin MSDataListLib.DataCombo DcSearch 
      Height          =   315
      Left            =   12960
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
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
   Begin ImpulseAniLabel.ISAniLabel LblLink 
      Height          =   315
      Left            =   960
      TabIndex        =   4
      ToolTipText     =   "«÷€ÿ ·› Õ „Êﬁ⁄ «·‘—ﬂ…"
      Top             =   3840
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   556
      ActiveUnderline =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      ForeColor       =   16777152
      MousePointer    =   99
      MouseIcon       =   "WebForm.frx":000C
      BackColor       =   14871017
      Caption         =   "Web Site : "
      ColorHover      =   16777152
      RightToLeft     =   -1  'True
      ImageCount      =   0
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   300
      Left            =   12960
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   " ‰»ÌÂ«  «·ÌÊ„"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   255
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "WebForm.frx":016E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblLabel1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lblLabel6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   1
      Left            =   21480
      TabIndex        =   13
      Top             =   9480
      Width           =   3375
   End
   Begin VB.Label lblLabel1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   21480
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   9000
      Width           =   3375
   End
   Begin VB.Label lblLabel6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "hiiiiiiiiii"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Tech."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   24960
      TabIndex        =   12
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Riyadh KSA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   855
      Left            =   18840
      TabIndex        =   6
      Top             =   7560
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Dynamic Byte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   1095
      Left            =   15480
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   6615
      Left            =   9600
      TabIndex        =   3
      Top             =   360
      Width           =   7335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "«·Ê’Ê· «·”—Ì⁄"
      Height          =   255
      Left            =   19080
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      Height          =   255
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   17160
      Top             =   15000
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "WebForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ALLButton1_Click()

          If checkApility("System_alarms") = False Then
                Exit Sub
            End If

            System_alarms.show
            
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DcSearch_GotFocus()

    If DcSearch.Text = "" Then
        ' SendKeys "{F4}"
    End If

End Sub

Private Sub DcSearch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)
    On Error GoTo errStrap

    'SendKeys ("{F4}")
    If KeyCode = vbKeyReturn Then
        ShowForm (CStr(Me.DcSearch.BoundText))
        DcSearch.Text = ""
        'Me.WindowState = 2
    End If

    Exit Sub
errStrap:

    If SystemOptions.UserInterface = ArabicInterface Then
        'MsgBox " ·«  ÊÃœ ‘«‘… »Â–« «·«”„ ", vbInformation
    Else
        'MsgBox "Form Not found", vbInformation
    End If

End Sub

Private Sub Form_Activate()
 
    If mdifrmmain.ActiveForm Is Nothing Then

        Exit Sub
    
    Else
        '   Msg = "ÌÃ» €·ﬁ «Ï ‘«‘… „‰ ‘«‘«  «·»—‰«„Ã ﬁ»·"
        '   Msg = Msg & Chr(13) & "«‰  ” Œœ„ Â–« «·‘«‘…....!!!!"
        '   MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If


 
End Sub
Private Sub ChangeLang()

    Label7.Caption = "Quick Access"
  ALLButton1.Caption = "Show Alrams"
End Sub
Sub SetText1(StrText As String)


    lblLabel6(0) = StrText & Space(10)
    lblLabel6(1) = lblLabel6(0)
    lblLabel6(0).Left = 0
    lblLabel6(1).Left = lblLabel6(0).Width
End Sub

Sub SetText(StrText As String)

    lblLabel1(0) = StrText & Space(10)
    lblLabel1(1) = lblLabel1(0)
    lblLabel1(0).Left = 0
    lblLabel1(1).Left = lblLabel1(0).Width
End Sub
'Here is where we do all the work
Public Sub ScrollText()
 Static i As Integer
 Dim k As Integer
 k = i Xor 1 'other label
 'move the label left by one pixel
 lblLabel1(i).Left = lblLabel1(i).Left - 30
 'other label follows like a train
 lblLabel1(k).Left = lblLabel1(i).Left + lblLabel1(i).Width
 'if engine is off screen, then make it caboose
 If lblLabel1(k).Left = 0 Then i = k: lblLabel1(k).Left = Me.Width
 
End Sub

 

Private Sub Form_GotFocus()
'    OldWindowProc = SetWindowLong( _
        hwnd, GWL_WNDPROC, _
        AddressOf NewWindowProc)
End Sub

Private Sub Form_Load()
On Error Resume Next
    'SkinFramework1.ApplyWindow Me.hWnd
    ' SkinFramework1.LoadSkin App.path & "\style\Vista.cjstyles", "Normalblack.ini"
Image1.Left = 0
Image1.Top = 0

           Me.Width = (mdifrmmain.Width)
   Me.Height = (mdifrmmain.Height) - 1300
    Image1.Width = Me.Width
   Image1.Height = Me.Height
   
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
Image1.Left = 0
Image1.Top = 0
    End If
   
   
    If Dir(BackGroundImag) <> "" Then
      Image1.Picture = LoadPicture(BackGroundImag)
        'Set Me.PopMenu1.BackgroundPicture = Me.Picture
        Me.Picture = Image1.Picture
        
Else

                If Dir(App.path & "\Garphics\wallpaper_Main.jpg") <> "" Then
                  '    Image1.Picture = LoadPicture(App.path & "\Garphics\wallpaper_Main.jpg")
                  '    Me.Picture = Image1.Picture
                 End If
 
    End If


   

   
'       Me.Height = 1400
  
  
 '    OldWindowProc = SetWindowLong( _
        hwnd, GWL_WNDPROC, _
        AddressOf NewWindowProc)
        
    Shape1.Width = Me.Width
    Shape1.Left = 0
    'Me.Height = 1100
lblLabel1(0).Width = Me.Width
    Dim StrSQL As String
    Dim cOptions As ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " select ScreenName,ScreenCaption from Screens order by ScreenCaption"
      '  Label7.Visible = True
    Else
        StrSQL = " select ScreenName,ScreenTitleEng from Screens order by ScreenTitleEng"
        'Label1.Visible = True
 
    End If

    Set cOptions = New ClsCompanyInfo
    
    If bigUser = True Then
user_name = "Power Admin"
End If

    If SystemOptions.UserInterface = ArabicInterface Then
    Label8.Caption = cOptions.ArabCompanyName & CHR(13) & CurrentBranchName & CHR(13) & " «·„” Œœ„   " & CHR(13) & user_name
Else
Label8.Caption = cOptions.EngCompanyName & CHR(13) & CurrentBranchNameE & CHR(13) & "User  " & CHR(13) & user_name
End If



    fill_combo Me.DcSearch, StrSQL

    lblLabel1(0).AutoSize = True
    Load lblLabel1(1)
 '   lblLabel1(1).Visible = True
lblLabel1(1).Width = Me.Width
lblLabel1(1).Left = Me.Width
lblLabel1(1).BackStyle = 0
    
showmessage
 


End Sub
Public Function showmessage(Optional speed1 As Double = 0, Optional fontcolor1 As Double = 0 _
, Optional fontsize1 As Double = 0, Optional backcolor1 As Double = 0)
Dim Message As String
Dim speed As Double
Dim fontsize As Double
Dim fontcolor As Double
Dim backcolor As Double
Dim show As Boolean
On Error Resume Next
 getInfoMessage 0, Message, speed, fontsize, fontcolor, backcolor, show
    If show = True Then
    Timer2.Enabled = True
         SetText Message
 
 If speed1 > 0 Then
  Timer2.interval = speed1
  Else
  Timer2.interval = speed
  End If
  If fontsize1 > 0 Then
  fontsize = fontsize1
  End If
  
  If fontcolor1 > 0 Then
  fontcolor = fontcolor1
  End If
  
  If backcolor1 > 0 Then
  backcolor = backcolor1
  End If
  
  lblLabel1(1).fontsize = fontsize
 lblLabel1(1).ForeColor = fontcolor
'  lblLabel1(1).backcolor = backcolor
  If backcolor = 0 Then
'    lblLabel1(1).BackStyle = 0
  Else
'    lblLabel1(1).BackStyle = 1
  End If
    Else
    Timer2.Enabled = False
    End If

'temp Info
getInfoMessage1 0, Message, speed, fontsize, fontcolor, backcolor, show
    If show = True Then
    info1Timer.Enabled = True
         SetText1 Message
 
  lblLabel6(1).fontsize = fontsize
 lblLabel6(1).ForeColor = fontcolor
' lblLabel6(1).backcolor = backcolor
  If backcolor = 0 Then
'    lblLabel6(1).BackStyle = 0
  Else
'    lblLabel6(1).BackStyle = 1
  End If
    Else
    info1Timer.Enabled = False
    End If










End Function

Private Sub Form_LostFocus()
'    OldWindowProc = SetWindowLong( _
'        hwnd, GWL_WNDPROC, _
'        AddressOf NewWindowProc)
End Sub

Private Sub info1Timer_Timer()
ScrollText
If lblLabel6(1).Left + lblLabel6(1).Width <= 0 Then
lblLabel6(1).Left = Me.Width
End If
lblLabel6(1).Left = lblLabel6(1).Left - 100
End Sub

Private Sub Label8_Click()
    'SkinFramework1.ApplyWindow Me.hWnd
    ' SkinFramework1.LoadSkin App.path & "\style\Vista.cjstyles", "Normalblack.ini"

End Sub

Private Sub LblLink_Click()
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE http://www.sattaryah.com"
End Sub

Private Sub Timer1_Timer()

    If Me.Top <> 0 And Me.Left <> 0 Then
        Me.Top = 0
        Me.Left = 0
    End If

End Sub

Private Sub Timer2_Timer()
ScrollText
If lblLabel1(0).Left + lblLabel1(0).Width <= 0 Then
lblLabel1(0).Left = Me.Width
End If
lblLabel1(0).Left = lblLabel1(0).Left - 100

End Sub

Private Sub Timer3_Timer()
'  CurrentTime = Format(Time, "hh:mm")
'    If CurrentTime = Text1.text Then
'        Beep
'        MsgBox (Text2.text), , "Personal Alarm"
'        Timer1.Enabled = False
'        Form1.WindowState = 0 'Restore form
'    End If
End Sub

Private Sub Timer4_Timer()
If ALLButton1.ForeColor = &HFF& Then
 ALLButton1.ForeColor = &HFFFFFF
Else
 ALLButton1.ForeColor = &HFF&
End If

End Sub
