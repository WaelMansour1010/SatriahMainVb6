VERSION 5.00
Object = "{D95CB779-00CB-4B49-97B9-9F0B61CAB3C1}#4.0#0"; "biokey.ocx"
Begin VB.Form FPFRM 
   Caption         =   "ÇáÈÕãå"
   ClientHeight    =   7635
   ClientLeft      =   420
   ClientTop       =   1050
   ClientWidth     =   12390
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   826
   Begin VB.CommandButton cmdSaveImage 
      Caption         =   "ÍÝÙ"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   25
      Top             =   5520
      Width           =   3135
   End
   Begin VB.CommandButton cmdInit 
      Caption         =   "ÇÏÑÇÌ ÇáÈÕãå"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   24
      Top             =   4440
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Template"
      Height          =   495
      Left            =   -4080
      TabIndex        =   18
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "From File"
      Height          =   2415
      Left            =   -8880
      TabIndex        =   16
      Top             =   4200
      Width           =   7455
      Begin VB.CommandButton Command4 
         Caption         =   "product name"
         Height          =   495
         Left            =   360
         TabIndex        =   22
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Vendor"
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label7 
         Height          =   495
         Left            =   2160
         TabIndex        =   23
         Top             =   1680
         Width           =   4575
      End
      Begin VB.Label Label6 
         Height          =   495
         Left            =   1920
         TabIndex        =   21
         Top             =   1080
         Width           =   5415
      End
   End
   Begin VB.Frame FrameCommands 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin ZKFPEngXControl.ZKFPEngX ZKFPEngX1 
         Left            =   5760
         Top             =   840
         EnrollCount     =   3
         SensorIndex     =   0
         Threshold       =   10
         VerTplFileName  =   ""
         RegTplFileName  =   ""
         OneToOneThreshold=   10
         Active          =   0   'False
         IsRegister      =   0   'False
         EnrollIndex     =   0
         SensorSN        =   ""
         FPEngineVersion =   "9"
         ImageWidth      =   0
         ImageHeight     =   0
         SensorCount     =   0
         TemplateLen     =   1152
         EngineValid     =   0   'False
         ForceSecondMatch=   0   'False
         IsReturnNoLic   =   -1  'True
         LowestQuality   =   30
         FakeFunOn       =   1
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close  Sensor"
         Height          =   375
         Left            =   1920
         TabIndex        =   19
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdEnroll 
         Caption         =   "Register"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "Verify(1:1)"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton cmdIdentify 
         Caption         =   "identify(1:N)"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox TextFingerName 
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2640
         TabIndex        =   7
         Text            =   "fingerprint1"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox TextSensorCount 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TextSensorIndex 
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3600
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TextSensorSN 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         Caption         =   "Image Format"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   1
         Top             =   1560
         Width           =   2415
         Begin VB.OptionButton OptionBmp 
            Caption         =   "BMP"
            BeginProperty Font 
               Name            =   "SimSun"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton OptionJpg 
            Caption         =   "JPEG"
            BeginProperty Font 
               Name            =   "SimSun"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Label LBLID 
         Height          =   495
         Left            =   960
         TabIndex        =   26
         Top             =   3720
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Sensor Cnt"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Current"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Serial Number"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Label StatusBar 
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   6720
      Visible         =   0   'False
      Width           =   5295
   End
End
Attribute VB_Name = "FPFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FTempLen As Integer
Dim FRegTemplate As String
Dim FRegTemp As Variant
Dim FingerCount As Long
Dim fpcHandle As Long
Dim FFingerNames() As String
Dim FMatchType As Integer

Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Private Sub cmdEnroll_Click()
  If TextFingerName.Text = "" Then
     MessageBox 0, "Please enter ID", "Error", 0
     Exit Sub
  End If
  ZKFPEngX1.BeginEnroll
  StatusBar.Caption = "Begin Register"
End Sub

Private Sub cmdIdentify_Click()
    If ZKFPEngX1.IsRegister Then
        ZKFPEngX1.CancelEnroll
    End If
    StatusBar.Caption = "Start Identification(1:N)"
    FMatchType = 2
End Sub

Private Sub cmdInit_Click()
  ZKFPEngX1.SensorIndex = 0
  If ZKFPEngX1.InitEngine = 0 Then
  '  MessageBox 0, ZKFPEngX1.ProduceName, ZKFPEngX1.Vendor, 0
  '  MessageBox 0, "init success", "Hint", 0
     
     StatusBar.Caption = "Êã ÇáÊÔÛíá"
     TextSensorCount.Text = ZKFPEngX1.SensorCount & ""
     TextSensorIndex.Text = ZKFPEngX1.SensorIndex & ""
     TextSensorSN.Text = ZKFPEngX1.SensorSN
     
     cmdInit.Enabled = False
     FMatchType = 0
  End If
End Sub

Private Sub cmdReadMemory_Click()

End Sub

Private Sub cmdSaveMemory_Click()
End Sub

Private Sub cmdSaveImage_Click()
    Dim sFileName As String
    sFileName = App.Path & "\IMAGES\FP\1"
    If OptionBmp.Value Then
        ZKFPEngX1.SaveBitmap sFileName & ".bmp"
    Else
        ZKFPEngX1.SaveJPG sFileName + ".jpg"
    End If
    MsgBox "Êã ÇáÍÝÙ"
End Sub

Private Sub cmdVerify_Click()
    If ZKFPEngX1.IsRegister Then
        ZKFPEngX1.CancelEnroll
    End If
    StatusBar.Caption = "Start Verification(1:1)"
    FMatchType = 1
End Sub

Private Sub Command1_Click()
Dim RegTemplateStr As String
Dim VerTemplateStr As String
Dim RegChanged As Boolean
ZKFPEngX1.GenRegTemplateAsStringFromFile ".\test.bmp", 500, RegTemplateStr
ZKFPEngX1.GenVerTemplateAsStringFromFile ".\test.bmp", 500, VerTemplateStr
If ZKFPEngX1.VerFingerFromStr(RegTemplateStr, VerTemplateStr, False, RegChanged) Then
          MessageBox 0, "Verify success", "information", 0
       Else
          MessageBox 0, "Verify Failed", "information", 0
       End If




End Sub

Private Sub Command2_Click()
    ZKFPEngX1.EndEngine
    cmdInit.Enabled = True
    
End Sub

Private Sub Command3_Click()
Label6.Caption = ZKFPEngX1.Vendor
End Sub

Private Sub Command4_Click()
Label7.Caption = ZKFPEngX1.ProduceName
End Sub

Private Sub Form_Load()
    FingerCount = 0
    fpcHandle = ZKFPEngX1.CreateFPCacheDB
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ZKFPEngX1.FreeFPCacheDB (fpcHandle)
End Sub

Private Sub ZKFPEngX1_OnCapture(ByVal ActionResult As Boolean, ByVal atemplate As Variant)
    Dim fi As Long, i As Long
    Dim lentgh  As Long
    Dim Score As Long, ProcessNum As Long
    Dim RegChanged As Boolean
    Dim sTemp1 As String
    Dim sTemp As Variant
    Dim AtempFinger
    sTemp1 = ZKFPEngX1.GetTemplateAsString()
    sTemp = ZKFPEngX1.GetTemplate()
  
  AtempFinger = atemplate
    StatusBar.Caption = "Acqired Template"
   If ZKFPEngX1.SaveTemplate("d:\fingerprint.tpl", AtempFinger) Then
   MsgBox "SaveTemplate d:\fingerprint.tpl success"
   Else
   MsgBox "SaveTemplate fail"
   End If
   
    If FMatchType = 1 Then  '1:1
       If ZKFPEngX1.VerFingerFromStr(FRegTemplate, sTemp1, False, RegChanged) Then
          MessageBox 0, "Verify success", "information", 0
       Else
          MessageBox 0, "Verify Failed", "information", 0
       End If
     
    ElseIf FMatchType = 2 Then '1:N
       Score = 8
    
        fi = ZKFPEngX1.IdentificationInFPCacheDB(fpcHandle, sTemp, Score, ProcessNum)
       If fi = -1 Then
          MessageBox 0, "Identification failed£¡", "information", 0
       Else
          MessageBox 0, "Identification Success Name=" & FFingerNames(fi) & " Score = " & Score & " Processed Number = " & ProcessNum, "information", 0
       End If
    End If
    
   
End Sub

Private Sub ZKFPEngX1_OnEnroll(ByVal ActionResult As Boolean, ByVal atemplate As Variant)
  Dim i As Long
  
  If Not ActionResult Then
    MessageBox 0, "Register failed", "Warning", 0
  Else
    MessageBox 0, "Regsiter success", "Information", 0
    
  
    FRegTemplate = ZKFPEngX1.GetTemplateAsString()
    FRegTemp = ZKFPEngX1.GetTemplate()
     
               

     ZKFPEngX1.AddRegTemplateToFPCacheDB fpcHandle, FingerCount, FRegTemp
        ReDim Preserve FFingerNames(FingerCount + 1)
    FFingerNames(FingerCount) = TextFingerName.Text
    FingerCount = FingerCount + 1
  End If
End Sub

Private Sub ZKFPEngX1_OnFeatureInfo(ByVal AQuality As Long)
  Dim sTemp As String
  
  sTemp = ""
  If ZKFPEngX1.IsRegister Then
     If ZKFPEngX1.EnrollIndex - 1 > 0 Then
     sTemp = "Register status: still press finger " & ZKFPEngX1.EnrollIndex - 1 & " times"
     Else
     sTemp = ""
     End If
  End If
  sTemp = sTemp & " Fingerprint quality"
  If AQuality <> 0 Then
     sTemp = sTemp & " no good " & AQuality
  Else
     sTemp = sTemp & " good"
  End If
  StatusBar.Caption = sTemp
End Sub

Private Sub ZKFPEngX1_OnImageReceived(AImageValid As Boolean)
  ZKFPEngX1.PrintImageAt hDC, FrameCommands.Width + 6, FrameCommands.Top, ZKFPEngX1.ImageWidth, ZKFPEngX1.ImageHeight
  End Sub
