VERSION 5.00
Object = "{D95CB779-00CB-4B49-97B9-9F0B61CAB3C1}#4.0#0"; "biokey.ocx"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form FPFRM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«œŒ«· ﬂÊœ «·»’„Â"
   ClientHeight    =   9075
   ClientLeft      =   345
   ClientTop       =   975
   ClientWidth     =   12915
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormX1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   605
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   861
   Begin VB.Frame FramerETERIVE 
      Height          =   9135
      Left            =   0
      TabIndex        =   33
      Top             =   -120
      Visible         =   0   'False
      Width           =   13095
      Begin VB.PictureBox Picture3 
         Height          =   5175
         Left            =   240
         ScaleHeight     =   5115
         ScaleWidth      =   4275
         TabIndex        =   34
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   5175
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   7215
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   5415
      Left            =   14400
      ScaleHeight     =   5355
      ScaleWidth      =   4755
      TabIndex        =   31
      Top             =   480
      Width           =   4815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "„”Õ «·’Ê—… „‰ «·”ﬂ«‰—"
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
      Left            =   5640
      TabIndex        =   30
      Top             =   7680
      Width           =   4335
   End
   Begin VB.PictureBox Picture1 
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5115
      ScaleWidth      =   4275
      TabIndex        =   28
      Top             =   360
      Width           =   4335
   End
   Begin VB.TextBox TxtId 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   120
      TabIndex        =   27
      Top             =   5880
      Width           =   4335
   End
   Begin VB.CommandButton cmdSaveImage 
      Caption         =   "Õ›Ÿ"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   25
      Top             =   7560
      Width           =   4335
   End
   Begin VB.CommandButton cmdInit 
      Caption         =   "«œ—«Ã «·»’„Â"
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
      Left            =   13320
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
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
      Height          =   6495
      Left            =   11160
      TabIndex        =   0
      Top             =   -4440
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
      Begin VB.CommandButton Command5 
         Caption         =   "Õ›Ÿ"
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
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   4335
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
   Begin DBPIXLib.DBPix20 DBPix201 
      Height          =   5295
      Left            =   5520
      TabIndex        =   32
      Top             =   2520
      Width           =   5295
      _Version        =   131072
      _ExtentX        =   9340
      _ExtentY        =   9340
      _StockProps     =   1
      _Image          =   "FormX1.frx":000C
      ImageResampleWidth=   100
      ImageResampleHeight=   100
      ImageResampleMode=   1
      ImageSaveFormat =   0
      JPEGQuality     =   75
      JPEGEncoding    =   0
      JPEGColorMode   =   0
      JPEGNoRecompress=   -1  'True
      JPEGRotateWarning=   0
      PNGColorDepth   =   0
      PNGCompression  =   0
      PNGFilter       =   0
      PNGInterlace    =   1
      ImageDitherMethod=   3
      ImagePaletteMethod=   4
      ImagePreviewMode=   0   'False
      ImageKeepMetaData=   -1  'True
      UseAmbientBackcolor=   -1  'True
      ViewAsyncDecoding=   -1  'True
      ViewEnableMouseZoom=   -1  'True
      ViewInitialZoom =   0
      ViewHAlign      =   1
      ViewVAlign      =   1
      ViewMenuMode    =   0
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
      Left            =   0
      TabIndex        =   17
      Top             =   9000
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

  

Private Sub cmdInit_Click()
  ZKFPEngX1.SensorIndex = 0
  If ZKFPEngX1.InitEngine = 0 Then
   
     Me.Caption = " „ «· ‘€Ì·"
     TextSensorCount.Text = ZKFPEngX1.SensorCount & ""
     TextSensorIndex.Text = ZKFPEngX1.SensorIndex & ""
     TextSensorSN.Text = ZKFPEngX1.SensorSN
     
     cmdInit.Enabled = False
     FMatchType = 0
  End If
End Sub
 

Private Sub cmdSaveImage_Click()
On Error Resume Next
    Dim sFileName As String
    Dim sFileName1 As String
    
    sFileName = "C:\\" & "x.bmp"
    sFileName1 = "C:\\" & "y.bmp"
    ZKFPEngX1.SaveBitmap sFileName
    Picture1.Picture = LoadPicture(sFileName)
    
    DBPix201.ImageSaveFile (sFileName1)
    Picture2.Picture = LoadPicture(sFileName1)
    
    saveimage (Val(Me.TxtId))


End Sub

Public Function retimage(FCusID As Double)

On Error Resume Next
'Me.Caption = Month(Date) * 500 + 3
Dim StrSQL  As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
 
If Val(FCusID) = 0 Then Exit Function
 
StrSQL = "select * From tblCustomerFingers  where FCusID =" & FCusID & "  "
rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
If rs.RecordCount > 0 Then
     If Not IsNull(rs("ItemPhoto").Value) Then
        If LenB(rs("ItemPhoto")) Then
            LoadPictureFromDB Me.Picture3, rs, "ItemPhoto"
        Else
            Set Picture3.Picture = Nothing
        End If

    Else
        Set Picture3.Picture = Nothing
    End If



     If Not IsNull(rs("ItemPhoto1").Value) Then
        If LenB(rs("ItemPhoto1")) Then
            LoadPictureFromDB Me.Picture2, rs, "ItemPhoto1"
        Else
            Set Picture2.Picture = Nothing
        End If

    Else
        Set Picture2.Picture = Nothing
    End If
    Image1.Picture = Picture2.Picture
DBPix201.ImageLoadBlob (Image1.Picture)
    
    

Else
MsgBox "·« ÌÊÃœ »Ì«‰«  »Â–« «·—ﬁ„", vbCritical, ""

End If
Exit Function
ErrTrap:
MsgBox "Œÿ√ ›Ì «·»Ì«‰« ", vbCritical, ""
End Function


Function saveimage(FCusID As Double)

On Error Resume Next
'Me.Caption = Month(Date) * 500 + 3
Dim StrSQL  As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
'StrSQL = "select * From TblCustemers CusID =" & CusID & "  "
If Val(FCusID) = 0 Then Exit Function
Cn.Execute "DELETE  tblCustomerFingers WHERE  FCusID = " & FCusID
Cn.Execute "insert into  tblCustomerFingers (FCusID)   values (" & FCusID & ")"

StrSQL = "select * From tblCustomerFingers  where FCusID =" & FCusID & "  "
rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
If rs.RecordCount > 0 Then
         If Picture1.Picture = 0 Then
             
        Else

            If SavePictureToDB(Picture1, rs, "ItemPhoto") = False Then
                GoTo ErrTrap
            End If
        End If
        
        
        
        If Picture2.Picture = 0 Then
             
        Else

            If SavePictureToDB(Picture2, rs, "ItemPhoto1") = False Then
                GoTo ErrTrap
            End If
        End If
        
          
          
        rs.Update
MsgBox " „ «·Õ›Ÿ", vbInformation, ""
Else
MsgBox "·« ÌÊÃœ „«·ﬂ »Â–« «·—ﬁ„", vbCritical, ""

End If
Exit Function
ErrTrap:
MsgBox "Œÿ√ ›Ì «·»Ì«‰« ", vbCritical, ""
End Function

 
 

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

 

Private Sub Command6_Click()

     DBPix201.TWAINAcquire
        MsgBox " „ „”Õ ÷Ê∆Ì  ··’Ê—…", vbInformation, ""

                DoEvents
                
End Sub

Private Sub Command7_Click()
retimage (Val(Me.TxtId))
End Sub

Private Sub Form_Load()
    FingerCount = 0
    fpcHandle = ZKFPEngX1.CreateFPCacheDB
    cmdInit_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ZKFPEngX1.FreeFPCacheDB (fpcHandle)
End Sub

  
Private Sub ZKFPEngX1_OnImageReceived(AImageValid As Boolean)
 Dim sFileName As String
  ZKFPEngX1.PrintImageAt hDC, DBPix201.Left + 50, DBPix201.Top, ZKFPEngX1.ImageWidth, ZKFPEngX1.ImageHeight
       sFileName = "C:\\" & "x.bmp"
    
    ZKFPEngX1.SaveBitmap sFileName
    Picture1.Picture = LoadPicture(sFileName)
         
         
         
  
  End Sub
