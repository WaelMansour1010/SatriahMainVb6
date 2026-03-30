VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form FrmItemsOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«⁄œ«œ  ﬂÊÌœ ·«’‰«›"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   Icon            =   "FrmItemsOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   5445
   Begin VB.OptionButton Opt_ItemcodeGroupOnly 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ «·’‰› ÌÕ ÊÏ   „Ã„Ê⁄Â «·’‰›"
      Height          =   255
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   720
      Width           =   4095
   End
   Begin VB.OptionButton opt_ItemcodeGroupandParentGroup 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ «·’‰› ÌÕ ÊÏ   „Ã„Ê⁄Â «·’‰› Ê «·„Ã„Ê⁄Â «·—∆Ì”Ì…"
      Height          =   255
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1200
      Width           =   4095
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Õ›Ÿ"
      Height          =   315
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame Frame20 
      Caption         =   "«· ⁄«„· „⁄ «·«’‰«›"
      Height          =   2655
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1680
      Width           =   5505
      Begin VB.TextBox TXTCodeDigits 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox CboSep 
         Height          =   315
         Index           =   1
         ItemData        =   "FrmItemsOptions.frx":000C
         Left            =   3480
         List            =   "FrmItemsOptions.frx":0025
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1800
         Width           =   735
      End
      Begin VB.ComboBox CboSep 
         Height          =   315
         Index           =   0
         ItemData        =   "FrmItemsOptions.frx":0045
         Left            =   3600
         List            =   "FrmItemsOptions.frx":005E
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox CmbCode 
         Height          =   315
         Index           =   0
         ItemData        =   "FrmItemsOptions.frx":007E
         Left            =   2760
         List            =   "FrmItemsOptions.frx":008B
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox CmbCode 
         Height          =   315
         Index           =   1
         ItemData        =   "FrmItemsOptions.frx":00B1
         Left            =   2760
         List            =   "FrmItemsOptions.frx":00BE
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox CmbCode 
         Height          =   315
         Index           =   2
         ItemData        =   "FrmItemsOptions.frx":00E4
         Left            =   2760
         List            =   "FrmItemsOptions.frx":00F1
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox TXTCodeDigits 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TXTCodeDigits 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄œœ «·Œ«‰« "
         Height          =   255
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "«·›«’·"
         Height          =   255
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "«·›«’·"
         Height          =   255
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "‘ﬂ· «·ﬂÊœ «· ›’Ì·Ï ··’‰›"
         Height          =   255
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«Ê·"
         Height          =   255
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Ê”ÿ"
         Height          =   255
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«Œ—"
         Height          =   255
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄œœ «·Œ«‰« "
         Height          =   255
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄œœ «·Œ«‰« "
         Height          =   255
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5475
      _cx             =   9657
      _cy             =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "«⁄œ«œ  ﬂÊÌœ ·«’‰«›"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   2
      ChildSpacing    =   1
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
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Visible         =   0   'False
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmItemsOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmbCode_Click(Index As Integer)

    If CmbCode(Index).ListIndex = 0 Then
        TXTCodeDigits(Index).Visible = False
        TXTCodeDigits(Index).text = 0
    Else
        TXTCodeDigits(Index).Visible = True
        TXTCodeDigits(Index).text = 1

    End If

End Sub

Private Sub CmdSave_Click()
    Dim sql As String
    sql = "Update tbloptions "
    sql = sql & " Set  itemcodePart1 =" & CmbCode(0).ListIndex & ","
    sql = sql & "    itemcodePart2 =" & CmbCode(1).ListIndex & ","
    sql = sql & "    itemcodePart3 =" & CmbCode(2).ListIndex & ","
    sql = sql & "    itemcodeSeperator1 ='" & CboSep(0).text & "',"
    sql = sql & "    itemcodeSeperator2 ='" & CboSep(1).text & "',"

    sql = sql & "    itemcodePart1NoOFDigit =" & val(TXTCodeDigits(0).text) & ","
    sql = sql & "    itemcodePart2NoOFDigit =" & val(TXTCodeDigits(1).text) & ","
    sql = sql & "    itemcodePart3NoOFDigit =" & val(TXTCodeDigits(2).text) & ","
    sql = sql & "    ItemcodeGroupOnly =" & IIf(Opt_ItemcodeGroupOnly.value = True, 1, 0) & ","
    sql = sql & "    ItemcodeGroupandParentGroup =" & IIf(opt_ItemcodeGroupandParentGroup.value = True, 1, 0)
 
    Cn.Execute sql
    MsgBox " „ «·Õ›Ÿ"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Resize_Form Me

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        '   ChangeLang
    End If

    Dim i As Integer

    If SystemOptions.UserInterface = EnglishInterface Then

        For i = 1 To 3

            With Me.CmbCode(i)
                .Clear
                .AddItem "Basic Item Code"
                .AddItem "Size"
                .AddItem "Color"
         
            End With

        Next i

    Else

    End If

    Set rs = New ADODB.Recordset
    rs.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rs.EOF Or rs.BOF) Then

        CmbCode(0).ListIndex = IIf(IsNull(rs("itemcodePart1").value), -1, rs("itemcodePart1").value)
        CmbCode(1).ListIndex = IIf(IsNull(rs("itemcodePart2").value), -1, rs("itemcodePart2").value)
        CmbCode(2).ListIndex = IIf(IsNull(rs("itemcodePart3").value), -1, rs("itemcodePart3").value)

        CboSep(0).text = IIf(IsNull(rs("itemcodeSeperator1").value), "", rs("itemcodeSeperator1").value)
        CboSep(1).text = IIf(IsNull(rs("itemcodeSeperator2").value), "", rs("itemcodeSeperator2").value)
        TXTCodeDigits(0).text = IIf(IsNull(rs("itemcodePart1NoOFDigit").value), "", rs("itemcodePart1NoOFDigit").value)
        TXTCodeDigits(1).text = IIf(IsNull(rs("itemcodePart2NoOFDigit").value), "", rs("itemcodePart2NoOFDigit").value)
        TXTCodeDigits(2).text = IIf(IsNull(rs("itemcodePart3NoOFDigit").value), "", rs("itemcodePart3NoOFDigit").value)
   
        If rs("ItemcodeGroupOnly").value = True Then
            Opt_ItemcodeGroupOnly.value = True
          
        End If

        If rs("ItemcodeGroupandParentGroup").value = True Then
            Opt_IItemcodeGroupandParentGroup.value = True
          
        End If
        
    End If

ErrTrap:
End Sub
