VERSION 5.00
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmItemTreeTip 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.PictureBox PicContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   8175
      Index           =   1
      Left            =   0
      ScaleHeight     =   8145
      ScaleWidth      =   5265
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.Timer TmrFade 
         Left            =   480
         Top             =   1290
      End
      Begin ImpulseButton.ISButton CmdClose 
         Height          =   345
         Left            =   4920
         TabIndex        =   20
         Top             =   30
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   609
         ButtonStyle     =   1
         Caption         =   ""
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmItemTreeTip.frx":0000
         ColorButton     =   -2147483624
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseAniLabel.ISAniLabel LblTimer 
         Height          =   285
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   503
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Arial"
         FontSize        =   14.25
         ForeColor       =   192
         BackColor       =   -2147483624
         Alignment       =   2
         Caption         =   "15"
         ColorHover      =   192
         ImageCount      =   0
      End
      Begin VB.Timer TmrClose 
         Interval        =   1000
         Left            =   60
         Top             =   1290
      End
      Begin ImpulseAniLabel.ISAniLabel LblReport 
         Height          =   225
         Index           =   0
         Left            =   2670
         TabIndex        =   18
         Top             =   7530
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   0
         MousePointer    =   99
         MouseIcon       =   "FrmItemTreeTip.frx":039A
         BackColor       =   -2147483624
         Alignment       =   1
         Caption         =   "ÚÑÖ ÔÇÔÉ ÊÞÇÑíÑ ÇáÕäÝ"
         ColorHover      =   16711680
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblReport 
         Height          =   285
         Index           =   1
         Left            =   1500
         TabIndex        =   19
         Top             =   7800
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   503
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   0
         MousePointer    =   99
         MouseIcon       =   "FrmItemTreeTip.frx":04FC
         BackColor       =   -2147483624
         Alignment       =   1
         Caption         =   "ÚÑÖ ÈíÇäÇÊ ÇáÕäÝ Ýì ÔÇÔÉ ÇáÃÕäÇÝ"
         ColorHover      =   16711680
         ImageCount      =   0
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÑÞã ÇáÕäÝ:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   0
         Left            =   3900
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   60
         Width           =   885
      End
      Begin VB.Label LblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   30
         Width           =   1035
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   1
         Left            =   4950
         Picture         =   "FrmItemTreeTip.frx":065E
         Top             =   7800
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   0
         Left            =   4950
         Picture         =   "FrmItemTreeTip.frx":09E8
         Top             =   7500
         Width           =   240
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   330
         Width           =   3495
      End
      Begin VB.Label LblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   690
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   30
         Width           =   975
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇÓã ÇáãÌãæÚÉ:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Index           =   3
         Left            =   3390
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   1425
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇÓã ÇáÕäÝ:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Index           =   2
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   330
         Width           =   1425
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ßæÏ ÇáÕäÝ:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Index           =   1
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   30
         Width           =   1005
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÓÚÑ ÇáÈíÚ ááãÓÊåáß:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Index           =   4
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   900
         Width           =   1755
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÓÚÑ ÇáÈíÚ ááÚãíá:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Index           =   5
         Left            =   3390
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1170
         Width           =   1425
      End
      Begin VB.Label LblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1050
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   870
         Width           =   1995
      End
      Begin VB.Label LblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   1050
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1170
         Width           =   1995
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   210
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   210
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   210
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   990
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   2820
         Width           =   1635
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÓÚÑ ÇáÈíÚ ááÏíáÑ-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Index           =   6
         Left            =   3390
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1470
         Width           =   1425
      End
      Begin VB.Label LblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1050
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1470
         Width           =   1995
      End
   End
End
Attribute VB_Name = "FrmItemTreeTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Enum dlgShowActions
    fadeNone = 0
    fadeIn = 1
    fadeOut = 2
    fadeInOut = 3
End Enum

Private unloadAction As dlgShowActions

Private fadeMode As dlgShowActions

Private winstyle As Long

Private Const GWL_EXSTYLE As Long = (-20)

Private Const WS_EX_RIGHT As Long = &H1000

Private Const WS_EX_LEFTSCROLLBAR As Long = &H4000

Private Const WS_EX_LAYERED As Long = &H80000

Private Const WS_EX_TRANSPARENT = &H20&

Private Const LWA_COLORKEY As Long = &H1

Private Const LWA_ALPHA As Long = &H2

Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long
   
Private Declare Function SetLayeredWindowAttributes _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal crKey As Long, _
                              ByVal bAlpha As Long, _
                              ByVal dwFlags As Long) As Long

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim XFont As IFontDisp
    Set XFont = Me.Font
    XFont.name = "Tahoma"
    XFont.Size = 8
    XFont.Bold = True
    XFont.Charset = Me.Font.Charset
    Set Me.LblReport(0).Font = XFont
    Set Me.LblReport(1).Font = XFont

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Me.PicContainer(1).Move Me.ScaleLeft, Me.ScaleTop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    'IF the user presses the close button
    'and the 'out' mode is fade, cancel
    'the close and instead invoke the
    'timer to cause the fade.
    '
    'The timer code changes the unloadAction
    'value to prevent this check From executing
    'again when the timer code issues the
    'Unload command.
    If ((UnloadMode = vbFormControlMenu) Or (UnloadMode = vbFormCode)) And (unloadAction = fadeOut) Or (unloadAction = fadeInOut) Then

        Cancel = True

        fadeMode = fadeOut
        TmrFade.interval = 20
        TmrFade.Enabled = True
    End If
   
    If TmrClose.Enabled = True Then
        TmrClose.Enabled = False
    End If

End Sub

Private Sub LblReport_Click(Index As Integer)

    Select Case Index

        Case 0
            OpenScreen PopUpShowItemCardScreen, val(Me.LblData(0).Caption)

        Case 1
            OpenScreen ItemsDataScreen, val(Me.LblData(0).Caption)
    End Select

End Sub

Private Sub TmrClose_Timer()
    Static i As Integer

    'PutFormOnTop Me.hwnd
    If IsMouseOverMe = False Then
        i = i + 1
        Me.LblTimer.Caption = 15 - i

        If i >= 15 Then
            TmrClose.Enabled = False
            Unload Me
        End If
    End If

End Sub

Public Function LoadData(LngItemID As Long)

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer, J As Integer
    Dim XFont As IFontDisp
    '-----------------------------------

    'StrSQL = "SELECT TblItems.ItemID, TblItems.ItemCode, TblItems.ItemName," & _
    '"Groups.GroupName, TblItems.SallingPrice, TblItems.CustomerPrice, TblItems.DealerPrice "
    'StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID"
    'StrSQL = StrSQL + " Where TblItems.ItemID=" & LngItemID
    'Set Rs = New ADODB.Recordset
    'Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If Not (Rs.BOF Or Rs.EOF) Then
    '    Me.LblData(0).Caption = LngItemID
    '    Me.LblData(1).Caption = IIf(IsNull(Rs("ItemCode").Value), "", Rs("ItemCode").Value)
    '    Me.LblData(2).Caption = IIf(IsNull(Rs("ItemName").Value), "", Rs("ItemName").Value)
    '    Me.LblData(3).Caption = IIf(IsNull(Rs("GroupName").Value), "", Rs("GroupName").Value)
    '    Me.LblData(4).Caption = IIf(IsNull(Rs("SallingPrice").Value), "", Rs("SallingPrice").Value)
    '    Me.LblData(5).Caption = IIf(IsNull(Rs("CustomerPrice").Value), "", Rs("CustomerPrice").Value)
    '    Me.LblData(6).Caption = IIf(IsNull(Rs("DealerPrice").Value), "", Rs("DealerPrice").Value)
    'End If
    ''-----------------------------------
    'If SystemOptions.SysDataBaseType = SQLServerDataBase Then
    '    StrSQL = "Select ItemID,ItemCode,ItemName,Sum(Total) as Totals,Sum(Quantity) as TotalQty, " & _
    '                "MOnth(DrivTable.Transaction_Date) as MonthNumber  "
    '    StrSQL = StrSQL + "From (" & _
    '                "SELECT TblItems.Item" & _
    '                "ID,TblItems.ItemCode, TblItems.ItemName,Transactions.Transaction_Date,'Total'=" & _
    '                "Case   When ItemDiscountType=1 Or ItemDiscountType=0 Then Transaction_Details.Qua" & _
    '                "ntity*Transaction_Details.Price    When ItemDiscountType=2 Then ((Transaction_Deta" & _
    '                "ils.Quantity*Transaction_Details.Price)-ItemDiscount)  When ItemDiscountType=3 T" & _
    '                "hen (Transaction_Details.Quantity*Transaction_Details.Price) *( 1- (ItemDiscount" & _
    '                "/100))     Else  0             End     ,Transaction_Details.Quantity " & _
    '                "FROM dbo.TblItems INNER JOIN  dbo.Transaction_Details ON dbo.TblItems.ItemID = " & _
    '                "dbo.Transaction_Details.Item_ID INNER JOIN dbo.Transactions ON dbo.Transaction_D" & _
    '                "etails.Transaction_ID = dbo.Transactions.Transaction_ID "
    '    StrSQL = StrSQL + " WHERE (Transactions.Transaction_Type=2  OR Transactions.Transaction_Type=0)"
    '    StrSQL = StrSQL + " AND Year(Transactions.Transaction_Date)=" & Year(Date)
    '    StrSQL = StrSQL + " AND (Transaction_Details.Item_ID)=" & LngItemID
    '    StrSQL = StrSQL + " )DrivTable  "
    '    StrSQL = StrSQL + " Group By ItemID,ItemCode,ItemName,MOnth(DrivTable.Transaction_Date)"
    '    StrSQL = StrSQL + "Order By ItemID,MOnth(DrivTable.Transaction_Date)"
    'ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
    '    StrSQL = "Select ItemID,ItemCode,ItemName,Sum(Total) as Totals,Sum(Quantity) as TotalQty,M" & _
    '    "Onth(DrivTable.Transaction_Date) as MonthNumber  From ( SELECT TblItems.ItemID, " & _
    '    "TblItems.ItemCode, TblItems.ItemName ,Transaction_Details.Quantity, Transactions" & _
    '    ".Transaction_Date, (IIf(ItemDiscountType=1 Or ItemDiscountType=0,(Transaction_De" & _
    '    "tails.Quantity*Transaction_Details.Price),IIf(ItemDiscountType=2,((Transaction_D" & _
    '    "etails.Quantity*Transaction_Details.Price)-ItemDiscount),IIf(ItemDiscountType=3," & _
    '    "(Transaction_Details.Quantity*Transaction_Details.Price)-((Transaction_Details.Q" & _
    '    "uantity*Transaction_Details.Price)*(ItemDiscount/100)),0)))) AS Total FROM Trans" & _
    '    "actions INNER JOIN (TblItems INNER JOIN Transaction_Details ON TblItems.ItemID =" & _
    '    " Transaction_Details.Item_ID) ON Transactions.Transaction_ID = Transaction_Detai" & _
    '    "ls.Transaction_ID WHERE (Transactions.Transaction_Type=2  OR Transactions.Transa" & _
    '    "ction_Type=0) AND Year(Transactions.Transaction_Date)=" & Year(Date) & " AND (Transaction_Deta" & _
    '    "ils.Item_ID)=" & LngItemID & " ) as  DrivTable  Group By  ItemID,ItemCode,ItemName,MOnth(DrivTab" & _
    '    "le.Transaction_Date)  Order By ItemID,MOnth(DrivTable.Transaction_Date)"
    'End If
    'Set Rs = New ADODB.Recordset
    'Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If SystemOptions.UserInterface = ArabicInterface Then
    '    ItemChart.SetMessageText "NoData", "áÇÊæÌÏ ÈíÇäÇÊ ááÚÑÖ"
    '    ItemChartPrice.SetMessageText "NoData", "áÇÊæÌÏ ÈíÇäÇÊ ááÚÑÖ"
    'End If
    'If Not (Rs.BOF Or Rs.EOF) Then
    '    ItemChart.Gallery = Gallery_Bar
    '    ItemChart.PointLabels = True
    ''    ItemChart.OpenData COD_Values, 2, Rs.RecordCount
    ''    I = 0: j = 0
    ''    For j = 0 To Rs.RecordCount - 1
    ''        ItemChart.Value(0, j) = Rs("Totals").Value
    ''        ItemChart.Value(1, j) = Rs("TotalQty").Value
    ''        Rs.MoveNext
    ''    Next j
    ''    ItemChart.CloseData COD_Values
    '    ItemChart.DataType.Item(0) = DataType_NotUsed
    '    ItemChart.DataType.Item(1) = DataType_NotUsed
    '    ItemChart.DataType.Item(2) = DataType_NotUsed
    '    ItemChart.DataType.Item(3) = DataType_Value
    '    ItemChart.DataType.Item(4) = DataType_Value
    '    ItemChart.DataType.Item(5) = DataType_Label
    '    ItemChart.DataSource = Rs
    '    'ItemChart.Series(0).PointLabelAlign=
    '    ItemChart.BackColor = &HE2E9E9
    '    ItemChart.Titles(0).Font.name = "Tahoma"
    '    ItemChart.Titles(0).Font.Charset = 178
    '    ItemChart.Titles(0).Font.Bold = True
    '    ItemChart.Titles(0).Font.Size = 8
    '    ItemChart.Titles(0).Alignment = StringAlignment_Center
    '    ItemChart.Titles(0).TextColor = &H80&
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        ItemChart.Titles(0).text = "ÊÍáíá ãÈíÚÇÊ ÇáÕäÝ Ýì ÇáÚÇã ÇáÍÇáì " & Year(Date)
    '        ItemChart.AxisX.Title.text = "ÔåæÑ ÇáÚÇã " & Year(Date)
    '    Else
    '        ItemChart.Titles(0).text = "Analysis of Item Sales in the Year " & Year(Date)
    '        ItemChart.AxisX.Title.text = "Months Of Year " & Year(Date)
    '    End If
    '
    '    ItemChart.Series(0).PointLabelColor = vbBlack
    '    ItemChart.WallWidth = 2
    '    ItemChart.Border = True
    '    ItemChart.BorderEffect = BorderEffect_Raised
    '    '--------------------------------------------
    '    Dim series0 As Cfx62ClientServerCtl.SeriesAttributes
    '    Set series0 = ItemChart.Series(0)
    '    series0.Gallery = Gallery_Bar
    '    series0.LineWidth = 2
    '    series0.MarkerShape = MarkerShape_None
    '    series0.YAxis = YAxis_Main
    '    ItemChart.AxisY.Pane = 0
    '
    '    Dim Series1 As Cfx62ClientServerCtl.SeriesAttributes
    '
    '    Set Series1 = ItemChart.Series(1)
    '    Series1.Border = False
    '    Series1.YAxis = YAxis_Secondary
    '    ItemChart.AxisY2.Pane = 1
    '
    '    Dim Axis0 As Cfx62ClientServerCtl.Axis
    '    Set Axis0 = ItemChart.AxisY
    '    'Axis0.Style=
    '    Axis0.ForceZero = False
    '    Axis0.LabelsFormat.Format = AxisFormat_Number
    '    Set Axis0 = ItemChart.Axis(YAxis_Secondary)
    '    Axis0.LabelsFormat.Decimals = 0
    '    ItemChart.RecalcScale
    '    ItemChart.AxisY2.Position = AxisPosition_Near
    '
    '    Dim Pane1 As Cfx62ClientServerCtl.Pane
    '    Set Pane1 = ItemChart.Panes(0)
    '    Pane1.Proportion = 15
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        Pane1.Title.text = "ÅÌãÇáì ÞíãÉ ÇáãÈíÚÇÊ"
    '    Else
    '        Pane1.Title.text = "Sales Value"
    '    End If
    '    Pane1.Title.Alignment = StringAlignment_Center
    '
    '    Dim pane2 As Cfx62ClientServerCtl.Pane
    '    Set pane2 = ItemChart.Panes(1)
    '    pane2.Proportion = 8
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        pane2.Title.text = "ÅÌãÇáì ÇáßãíÇÊ"
    '    Else
    '        pane2.Title.text = "Quantity Volume"
    '    End If
    '    pane2.Title.Alignment = StringAlignment_Center
    '
    '    pane2.Title.BackColor = ItemChart.BackColor
    '    Pane1.Title.BackColor = ItemChart.BackColor
    '    '------------------------------------------
    'Else
    '    ItemChart.ClearData ClearDataFlag_AllData
    'End If
    'Rs.Close
    'Set Rs = Nothing
    ''----------------------------
    'If SystemOptions.SysDataBaseType = SQLServerDataBase Then
    '    StrSQL = "SELECT Transactions.Transaction_Date,Transaction_Details.Price  " & _
    '    "FROM dbo.TblItems INNER JOIN  dbo.Transaction_Details ON dbo.TblItems.ItemID =" & _
    '    "dbo.Transaction_Details.Item_ID INNER JOIN dbo.Transactions ON  " & _
    '    "dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID "
    '    StrSQL = StrSQL + " WHERE (Transactions.Transaction_Type=2  OR Transactions.Transaction_Type=0)" & _
    '    " AND  Year(Transactions.Transaction_Date)=" & Year(Date) & _
    '    " AND (Transaction_Details.Item_ID)=" & LngItemID & _
    '    " Order By ItemID,Transaction_Date"
    'ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
    '    StrSQL = "SELECT Transactions.Transaction_Date,Transaction_Details.Price  " & _
    '    " FROM Transactions INNER JOIN Transaction_Details ON Transactions.Transaction_ID " & _
    '    "= Transaction_Details.Transaction_ID"
    '    StrSQL = StrSQL + " WHERE (Transactions.Transaction_Type=2  OR Transactions.Transaction_Type=0)" & _
    '    " AND  Year(Transactions.Transaction_Date)=" & Year(Date) & _
    '    " AND (Transaction_Details.Item_ID)=" & LngItemID & _
    '    " Order By Item_ID,Transaction_Date"
    'End If
    'Set Rs = New ADODB.Recordset
    'Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If Rs.BOF Or Rs.EOF Then
    '    ItemChartPrice.ClearData ClearDataFlag_AllData
    'Else
    '    'Setting the Chart series
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        ItemChartPrice.Titles(0).text = "ÊÍáíá ÃÓÚÇÑ ÇáÕäÝ Ýì ÇáÚÇã ÇáÍÇáì " & Year(Date)
    '    Else
    '        ItemChartPrice.Titles(0).text = "Analysis of Item Prices on the Year " & Year(Date)
    '    End If
    '    ItemChartPrice.Titles(0).TextColor = SysMaronColor
    '    ItemChartPrice.Titles(0).Font.name = "Tahoma"
    '    ItemChartPrice.Titles(0).Font.Size = 8
    '    ItemChartPrice.Titles(0).Font.Bold = True
    '    ItemChartPrice.Gallery = Gallery_Lines
    '    ItemChartPrice.DataType.Item(0) = DataType_Label
    '    ItemChartPrice.DataType.Item(1) = DataType_Value
    '    ItemChartPrice.AxisX.LabelsFormat.Format = AxisFormat_Date
    '    ItemChartPrice.DataSource = Rs
    '    'Dim series0 As Object
    '    Set series0 = ItemChartPrice.Series(0)
    '    series0.Gallery = Gallery_Lines
    '    series0.LineWidth = 2
    '    series0.MarkerShape = MarkerShape_None
    '    series0.YAxis = YAxis_Main
    '
    '    ItemChartPrice.AxisY.Pane = 0
    'End If
    '----------------------------
End Function

Private Function IsMouseOverMe() As Boolean
    Dim t As POINTAPI
    Dim LngX As Single
    Dim LngY As Single
    GetCursorPos t
    LngX = t.x * Screen.TwipsPerPixelX
    LngY = t.Y * Screen.TwipsPerPixelY

    If (LngX) > (Me.left + Me.Width) Or LngX < Me.left Then
        IsMouseOverMe = False
        Exit Function
    ElseIf (LngY) > (Me.top + Me.Height) Or LngY < Me.top Then
        IsMouseOverMe = False
        Exit Function
    End If

    IsMouseOverMe = True
End Function

Private Sub TmrFade_Timer()

    Static fadeValue As Long
    Dim alpha As Long
   
    Select Case fadeMode
   
        Case fadeOut:
      
            'prevents the form's QueryUnload sub
            'From stopping the unloading of the
            'form via code here
            unloadAction = 0
      
            If (fadeValue + (256 * 0.05)) >= 256 Then
             
                'done, so reset the fadeValue to
                'allow for fading out if required
                TmrFade.Enabled = False
                fadeValue = 0
                Unload Me
                Exit Sub
            End If
         
            fadeValue = fadeValue + (256 * 0.05)
            alpha = (256 - fadeValue)
      
        Case fadeIn:
      
            If (fadeValue + (256 * 0.05)) >= 256 Then
                'done, but one more call to
                'SetLayeredWindowAttributes is
                'required to set the final opacity to 255
                TmrFade.Enabled = False
                fadeValue = 0
                alpha = 255
            Else
                fadeValue = fadeValue + (256 * 0.05)
                alpha = fadeValue
            End If
      
        Case Else
    End Select

    SetLayeredWindowAttributes Me.hWnd, 0&, alpha, LWA_ALPHA
   
End Sub

Public Sub DialogAction(dlgEffectsMethod As dlgShowActions)
   
    Dim alpha As Long

    'alpha=0: window transparent
    'alpha=255: window opaque
 
    Select Case dlgEffectsMethod
   
        Case fadeNone  'show 'normally'
            'nothing to do, so exit and let
            'the calling routine's Show command
            'control the display
            Exit Sub
      
        Case fadeOut 'show normally but prepare for a fade out
      
            'this requires changing the window style
            'and calling SetLayeredWindowAttributes once
            'specifying a value of opaque (255). To
            'cause the form to fade out, an 'unloadAction'
            'flag is set
            unloadAction = dlgEffectsMethod
            Call AdjustWindowStyle
            alpha = 2
            SetLayeredWindowAttributes Me.hWnd, 0&, alpha, LWA_ALPHA
      
        Case fadeIn, fadeInOut 'show form by fading in
      
            'just adjust the window style and
            'use a timer to fade the window in
            Call AdjustWindowStyle
            fadeMode = fadeIn
            TmrFade.interval = 20
            TmrFade.Enabled = True
         
            'but ... if the effect mode is
            'to fade in/out, set the unloadAction
            'flag
            If dlgEffectsMethod = fadeInOut Then unloadAction = fadeOut
         
    End Select
   
End Sub

Private Function AdjustWindowStyle()

    Dim Style As Long

    'in order to have transparent windows, the
    'WS_EX_LAYERED window style must be applied
    'to the form
    Style = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
   
    If Not (Style And WS_EX_LAYERED = WS_EX_LAYERED) Then
      
        Style = Style Or WS_EX_LAYERED
        SetWindowLong Me.hWnd, GWL_EXSTYLE, Style
      
    End If
    
End Function

Private Sub ChangeLang()
    Dim i As Integer
    Me.lbl(0).Caption = "Item ID:"
    Me.lbl(1).Caption = "Item Code:"
    Me.lbl(2).Caption = "Item Name:"
    Me.lbl(3).Caption = "Group Name:"
    Me.lbl(4).Caption = "User Price:"
    Me.lbl(5).Caption = "Customer Price:"
    Me.lbl(6).Caption = "Dealer Price:"

    For i = lbl.LBound To lbl.UBound
        Me.lbl(i).FontBold = True
    Next i

    LblReport(0).Caption = "Show Item Reports Screen"
    LblReport(1).Caption = "Show Item Data Screen"
End Sub
