VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmPaneHelp 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   1830
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   9735
      Left            =   2400
      TabIndex        =   6
      Top             =   -120
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   600
         Top             =   8160
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·—⁄«Â"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   1785
         Left            =   120
         Stretch         =   -1  'True
         Top             =   5280
         Width           =   1860
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "⁄„·«∆‰«"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Image Image10 
         Height          =   1785
         Left            =   120
         Stretch         =   -1  'True
         Top             =   720
         Visible         =   0   'False
         Width           =   1860
      End
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   1800
      Top             =   1800
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1260
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneHelp.frx":0000
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneHelp.frx":077A
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneHelp.frx":0EF4
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneHelp.frx":166E
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneHelp.frx":1DE8
            Key             =   "Home"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneHelp.frx":2562
            Key             =   "History"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   480
      Left            =   150
      TabIndex        =   1
      Top             =   -1140
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   847
      ButtonWidth     =   820
      ButtonHeight    =   794
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.ComboBox CboHistoryList 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   60
         Width           =   3495
      End
   End
   Begin VB.PictureBox WbHelp1 
      Height          =   3945
      Left            =   4260
      ScaleHeight     =   3885
      ScaleWidth      =   1905
      TabIndex        =   0
      Top             =   2310
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.PictureBox PicTip 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   5070
      ScaleHeight     =   375
      ScaleWidth      =   4425
      TabIndex        =   4
      Top             =   -180
      Width           =   4455
      Begin VB.Label LblTip 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tip"
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
         Height          =   225
         Left            =   330
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   60
         Visible         =   0   'False
         Width           =   3915
      End
   End
   Begin XtremeSuiteControls.WebBrowser WbHelp 
      Height          =   9675
      Left            =   1800
      TabIndex        =   3
      Top             =   0
      Width           =   1965
      _Version        =   786432
      _ExtentX        =   3466
      _ExtentY        =   17066
      _StockProps     =   173
      BackColor       =   -2147483643
      ScrollBarStyle  =   2
   End
End
Attribute VB_Name = "FrmPaneHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim websiteurl As String
Dim firstrun As Integer
    Public Enum Exec
        OLECMDID_OPTICAL_ZOOM = 63
    End Enum
    Private Enum execOpt
        OLECMDEXECOPT_DODEFAULT = 0
        OLECMDEXECOPT_PROMPTUSER = 1
        OLECMDEXECOPT_DONTPROMPTUSER = 2
        OLECMDEXECOPT_SHOWHELP = 3
    End Enum
    Dim InitialZoom As Integer
    Dim imageCounter As Integer
    Dim imageCounter1 As Integer
    Dim CURRENTPATH As String
    
 


Public Sub PerformZoom(ByVal value As Integer)
 On Error Resume Next
       '  value = 50
          WbHelp.ExecWB Exec.OLECMDID_OPTICAL_ZOOM, execOpt.OLECMDEXECOPT_PROMPTUSER, value, 0
   
    End Sub
 
Private Sub Form_Load()
On Error Resume Next
CURRENTPATH = App.path
If mId(CURRENTPATH, Len(App.path), 1) = "/" Or mId(CURRENTPATH, Len(App.path), 1) = "\" Then
CURRENTPATH = mId(CURRENTPATH, 1, Len(CURRENTPATH) - 1)

End If
InitialZoom = 100
PerformZoom (-50)
'waelComment    LoadHomePage
    firstrun = 0
 '   CreateToolBar
 ' InternetGetConnectedState(0, 0)
 'InternetGetConnectedState(0, 1) = 1
'      Frame1.Visible = False
'     If 1 = 1 Then
'    LoadHomePage
'             WbHelp.Visible = True
'             WbHelp1.Visible = True
'
'         Timer5.Enabled = False
'         Timer1.Enabled = False
'         Frame1.Visible = False
'             Else
'             Frame1.Visible = True
'         WbHelp.Visible = False
'         Image10.Visible = True
'         Timer5.Enabled = True
'         Timer1.Enabled = True
'
'   Timer1.Enabled = True
'   WbHelp1.Visible = False
'    End If
    
    
     '/WbHelp.ExecWB 50, 2, 35&
    If SystemOptions.UserInterface = ArabicInterface Then
'        Me.LblTip.Caption = "· ﬂ»Ì— «·Œÿ «Ê  ’€Ì—Â ... ﬁ„ »·› «·⁄Ã·… «·œÊ«—… ›Ï «·„«Ê” Ê«‰  ÷«€ÿ ⁄·Ï “— Ctrl"
    Else
        SetInterface Me
'        Me.LblTip.Caption = "To Resize the Font ... Scroll the Mouse Wheel While you Pree the {Ctrl} Key"
    End If

End Sub

Private Sub Form_Resize()

    Dim SngTop As Single
    Dim SngWidth As Single
    Dim SngTemp As Single
    On Error Resume Next

    If Me.WindowState = vbNormal Then
       ' Me.Toolbar1.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth
        Me.WbHelp.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight + 50
Exit Sub
End If
    If Me.WindowState = vbNormal Then
        Me.Toolbar1.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth
        Me.WbHelp.Move Me.ScaleLeft, Me.ScaleTop + Me.Toolbar1.Height, Me.ScaleWidth, Me.ScaleHeight - (Me.Toolbar1.Height + Me.PicTip.Height + 50)

        With Me.PicTip
            .left = Me.ScaleLeft
            .top = Me.WbHelp.top + Me.WbHelp.Height + 25
            .Width = Me.ScaleWidth
        End With

        With Me.LblTip
            .top = Me.PicTip.ScaleTop
            .left = Me.PicTip.ScaleLeft
            .Width = Me.PicTip.ScaleWidth
            .Height = Me.PicTip.ScaleHeight
        End With

        With Me.Toolbar1.Buttons("CboHistoryList")
            SngTop = (Me.Toolbar1.Height - Me.CboHistoryList.Height) / 2
            SngWidth = Me.Toolbar1.Width - Me.Toolbar1.Buttons("CboHistoryList").left
            Me.Toolbar1.Buttons("CboHistoryList").Width = SngWidth
            SngTemp = (Me.Toolbar1.Buttons("CboHistoryList").Width - 100)

            If SngTemp > 0 Then
                Me.CboHistoryList.Move .left, SngTop, (Me.Toolbar1.Buttons("CboHistoryList").Width - 100)
            End If

        End With

    End If

End Sub

Private Sub CreateToolBar()
    Dim xButton As MSComctlLib.Button

    With Me.Toolbar1
        .Buttons.Clear
        Set .ImageList = Me.ImageList1
        .Appearance = ccFlat
        .Style = tbrFlat
        .AllowCustomize = False
        .BorderStyle = ccNone
        '.ShowTips = True
        .TextAlignment = tbrTextAlignBottom
        .AllowCustomize = False
        .Wrappable = False
        Set xButton = .Buttons.Add(, "Back", "", tbrDefault, "Back")
        xButton.ToolTipText = "⁄Êœ… ··’›Õ… «·”«»ﬁ…"
        Set xButton = .Buttons.Add(, "Forward", "", tbrDefault, "Forward")
        xButton.ToolTipText = "«·’›Õ… «· «·Ì…"
        .Buttons.Add , , , tbrSeparator
        Set xButton = .Buttons.Add(, "Stop", "", tbrDefault, "Stop")
        xButton.ToolTipText = "≈Ìﬁ«›  Õ„Ì· «·’›Õ… «·Õ«·Ì…"
    
        Set xButton = .Buttons.Add(, "Refresh", "", tbrDefault, "Refresh")
        xButton.ToolTipText = "≈⁄«œ…  Õ„Ì· «·’›Õ… «·Õ«·Ì…"
     
        Set xButton = .Buttons.Add(, "Home", "", tbrDefault, "Home")
        xButton.ToolTipText = "⁄—÷ «·’›Õ… «·—∆Ì”Ì…"
        '.Buttons.Add , , , tbrSeparator
        Set xButton = .Buttons.Add(, "History", , tbrCheck, "History")
        xButton.ToolTipText = "⁄—÷ ﬁ«∆„… »«·’›Õ«  «·”«»ﬁ ﬁ—«∆ Â«"
    
        Set xButton = .Buttons.Add(, "CboHistoryList", , tbrPlaceholder)
        xButton.Width = .Width - xButton.left
        xButton.ToolTipText = "«Œ— 20 ’›Õ…  „  ﬁ—«∆ Â„"
        CboHistoryList.Move xButton.left, CboHistoryList.top, xButton.Width
    End With

End Sub

'Private Function AddButton(Controls As CommandBarControls, _
'                           ControlType As XTPControlType, _
'                           ID As Long, _
'                           Optional BeginGroup As Boolean = False, _
'                           Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic) As CommandBarControl
'    Dim Control As CommandBarControl
'    Set Control = Controls.Add(ControlType, ID, "")
'
'    Control.BeginGroup = BeginGroup
'    Control.Style = ButtonStyle
'    Set AddButton = Control
'End Function

Private Sub Timer1_Timer()
If imageCounter = 0 Then imageCounter = 1
If imageCounter = 10 Then imageCounter = 1
On Error Resume Next
'Image1.Picture = LoadPicture(App.path & "\Images\localwebCustomer\" & imageCounter & ".jpg")
' imageCounter = imageCounter + 1
 
 
End Sub

Private Sub Timer5_Timer()
On Error Resume Next
If imageCounter1 = 0 Then imageCounter1 = 1
If imageCounter1 = 10 Then imageCounter1 = 1
On Error Resume Next
'Image10.Picture = LoadPicture(App.path & "\Images\localwebAdv\" & imageCounter1 & ".jpg")
 imageCounter1 = imageCounter1 + 1
 
 
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo hErr

    Select Case Button.Key

        Case "Home"
            LoadHomePage

        Case "Back"
            Me.WbHelp.GoBack

        Case "Forward"
            Me.WbHelp.GoForward

        Case "Refresh"
            Me.WbHelp.Refresh

        Case "Stop"
            Me.WbHelp.Stop
    End Select

    Exit Sub
hErr:
End Sub
Function checkInternetConnection() As Boolean

checkInternetConnection = False
Exit Function
'code to check for internet connection
'by Daniel Isoje
On Error Resume Next
' checkInternetConnection = False
' Dim objSvrHTTP As ServerXMLHTTP
' Dim varProjectID, varCatID, strt As String
' Set objSvrHTTP = New ServerXMLHTTP
' objSvrHTTP.Open "GET", "http://www.google.com"
' objSvrHTTP.setRequestHeader "Accept", "application/xml"
' objSvrHTTP.setRequestHeader "Content-Type", "application/xml"
' objSvrHTTP.send strt
 If Err = 0 Then
 checkInternetConnection = True
  
 End If
End Function
Private Sub LoadHomePage()
On Error Resume Next

  '  If Dir(App.path & "\DynamicHelp\index.htm") <> "" Then
  '      Me.WbHelp.Navigate App.path & "\DynamicHelp\index.htm"
  '  End If
'Me.WbHelp.Navigate "http://www.alsattaryahgroupservice.somee.com/home/test"

If SystemOptions.WebAdv = "" Then
websiteurl = "http://sattaryahadv.xyz/MainAdvertisement/index"
Else
websiteurl = SystemOptions.WebAdv
End If
If SystemOptions.UserInterface = ArabicInterface Then
websiteurl = "http://sattaryahadv.xyz/MainAdvertisement/ViewAds"
Else
websiteurl = "http://sattaryahadv.xyz/MainAdvertisement/ViewAdsE"
End If
'Me.WbHelp.Navigate "\static\staticpage.html"
'Exit Sub
If checkInternetConnection = True Then ' connected
Me.WbHelp.Navigate websiteurl
Else
        If SystemOptions.UserInterface = ArabicInterface Then
                Me.WbHelp.Navigate CURRENTPATH & "\static\staticpage.html"
        Else
                Me.WbHelp.Navigate CURRENTPATH & "\static\staticpagee.html"
        End If
End If

End Sub

Private Sub WbHelp_BeforeNavigate2(ByVal pDisp As Object, _
                                   URL As Variant, _
                                   Flags As Variant, _
                                   TargetFrameName As Variant, _
                                   PostData As Variant, _
                                   Headers As Variant, _
                                   Cancel As Boolean)
On Error Resume Next
If firstrun = 0 Then firstrun = 1: Exit Sub
'    If InStr(1, URL, "70.htm", vbTextCompare) <> 0 Then
        'OpenScreen EmployeesScreen
'    End If

'If Mid(CStr(URL), Len(CStr(URL)), 1) = "/" Then

'URL = Mid(CStr(URL), 1, Len(CStr(URL)) - 1)

'End If
'If LCase(CStr(URL)) = LCase(CURRENTPATH & "static\staticpage.html") Then Exit Sub
 
If CStr(URL) <> websiteurl Then

     
    OpenWebSite CStr(URL)
   
    
    
End If
'If InStr(1, URL, "http://drabdulrahmanalmishari.com.sa/", vbTextCompare) <> 0 Then
'    OpenWebSite "http://drabdulrahmanalmishari.com.sa/"
'
'
    'End If
'
End Sub

