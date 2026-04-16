VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form FrmCustomerDisplay 
   ClientHeight    =   9855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15225
   Icon            =   "FrmCustomerDisplay.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9855
   ScaleWidth      =   15225
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
      Height          =   9855
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15225
      _cx             =   26855
      _cy             =   17383
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
      Appearance      =   6
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   -1  'True
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
      Begin VB.Label LblInformation3 
         Caption         =   "Label3"
         Height          =   795
         Left            =   3150
         TabIndex        =   7
         Top             =   2190
         Width           =   2595
      End
      Begin VB.Label LblInformation2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   3285
         Left            =   240
         TabIndex        =   6
         Top             =   5520
         Width           =   5985
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00400040&
         X1              =   7080
         X2              =   7080
         Y1              =   1800
         Y2              =   8880
      End
      Begin VB.Image Image1 
         Height          =   3285
         Left            =   7440
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   5985
      End
      Begin VB.Label lblItemID 
         Caption         =   "Label3"
         Height          =   345
         Left            =   2970
         TabIndex        =   5
         Top             =   165
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Visit Us Again"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   645
         Left            =   60
         TabIndex        =   4
         Top             =   8910
         Width           =   14805
      End
      Begin VB.Label lblcompanyname 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "company name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   945
         Left            =   7380
         TabIndex        =   3
         Top             =   165
         Visible         =   0   'False
         Width           =   7380
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to your Site"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   510
         Left            =   7965
         TabIndex        =   2
         Top             =   165
         Visible         =   0   'False
         Width           =   4830
      End
      Begin VB.Image Image9 
         Height          =   1305
         Left            =   1110
         Picture         =   "FrmCustomerDisplay.frx":000C
         Stretch         =   -1  'True
         Top             =   60
         Width           =   12660
      End
      Begin VB.Label LblInformation 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   3285
         Left            =   7440
         TabIndex        =   1
         Top             =   5520
         Width           =   5985
      End
   End
End
Attribute VB_Name = "FrmCustomerDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function loadLogo()
    
    
    Image9.Visible = False
    Dim BackGroundImag As String
 BackGroundImag = App.path & "\Images\POS2\logo.jpg"
    
     If Dir(BackGroundImag) <> "" Then
     Image9.Picture = LoadPicture(BackGroundImag)
     Image9.Visible = True
     Else
     Image9.Visible = False
     
    End If
    
    
    
    
    
    
    
    
Exit Function
    Dim rs As ADODB.Recordset
    Dim BolShowLogo As Boolean
    Dim xLogo As CRAXDRT.OLEObject
    Dim StrFileName As String
    Dim MsgErr As String

     

    Set rs = New ADODB.Recordset
    rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

    If rs.BOF Or rs.EOF Then
       
        Exit Function
    End If

   

   If Not (IsNull(rs("CompanyLogo").value)) Then
    
     LoadPictureFromDB Image9, rs, "CompanyLogo"
        rs.Close
        Set rs = Nothing
     Else
     'Image9.Picture = Nothing
    End If
    
    
End Function
Public Function maxformz()
Me.WindowState = 2
End Function
Private Sub Form_Load()
 loadLogo
    
End Sub

Private Sub lblItemID_Change()
  'load pictur
Image1.Visible = False
  Dim StrSQL As String
    Dim rs As ADODB.Recordset
  
  StrSQL = " Select * from TblItems where ItemID=" & val(Me.lblItemID.Caption)
     

    Set rs = New ADODB.Recordset
    'rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdText
rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.BOF Or rs.EOF Then
       
        Exit Sub
    End If

   

   If Not (IsNull(rs("ItemPhoto").value)) Then
    Image1.Visible = True
     LoadPictureFromDB Image1, rs, "ItemPhoto"
        rs.Close
        Set rs = Nothing
     Else
     Image1.Visible = False
    End If
    
    
    



End Sub

