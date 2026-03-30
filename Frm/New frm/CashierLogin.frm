VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form CashierLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘… œŒÊ· «·‘Ì› "
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9720
   Icon            =   "CashierLogin.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   9720
   Begin VB.Frame Frame3 
      Caption         =   "»Ì«‰«  «·ﬂ«‘Ì—"
      Height          =   2895
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3720
      Visible         =   0   'False
      Width           =   6015
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1080
         Width           =   6015
      End
      Begin VB.TextBox XPTxtPass 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   600
         Width           =   3465
      End
      Begin VB.TextBox TxtBalance 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1080
         Width           =   975
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   360
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16761024
         ForeColor       =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ALLButtonS.ALLButton CMDLogin 
         Height          =   375
         Left            =   2160
         TabIndex        =   32
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   " ”ÃÌ· «·œŒÊ·"
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
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "CashierLogin.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton CMDCancel 
         Height          =   375
         Left            =   480
         TabIndex        =   33
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "«·€«¡"
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
         BCOL            =   192
         BCOLO           =   192
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "CashierLogin.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄Âœ… ”«»ﬁ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   3720
         TabIndex        =   37
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label LblPetty 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1560
         Width           =   1400
      End
      Begin VB.Label LBLBalance 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1080
         Width           =   1400
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "—’Ìœ ”«»ﬁ"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   29
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄Âœ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   28
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·ﬂ«‘Ì—"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   4320
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ·„… «·„—Ê—"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   4320
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   3840
         Picture         =   "CashierLogin.frx":0044
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Õœœ «·‰ﬁÿ…"
      Height          =   615
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1680
      Width           =   6135
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dcpoint 
         Height          =   360
         Left            =   240
         TabIndex        =   19
         Top             =   120
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16761024
         ForeColor       =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Œ — «·‰ﬁÿ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   4440
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "»Ì«‰«  «·„‘—›"
      Height          =   1575
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2280
      Width           =   6015
      Begin VB.TextBox XPTxtPass1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Text            =   "1234"
         Top             =   600
         Width           =   3465
      End
      Begin MSDataListLib.DataCombo DCboUserName1 
         Height          =   360
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16761024
         ForeColor       =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ALLButtonS.ALLButton ALLButton1 
         Default         =   -1  'True
         Height          =   375
         Left            =   2160
         TabIndex        =   31
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   " ”ÃÌ· «·œŒÊ·"
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
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "CashierLogin.frx":0721
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·„‘—›"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   4200
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ·„… «·„—Ê—"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   4200
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   495
         Left            =   3840
         Picture         =   "CashierLogin.frx":073D
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
   End
   Begin MSComCtl2.DTPicker ShfitFrom 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "'Time: 'hh:mm tt"
      Format          =   161218563
      UpDown          =   -1  'True
      CurrentDate     =   39240
   End
   Begin MSComCtl2.DTPicker ShfitTo 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "'Time: 'hh:mm tt"
      Format          =   161218563
      UpDown          =   -1  'True
      CurrentDate     =   39240
   End
   Begin MSDataListLib.DataCombo dcShift 
      Height          =   360
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16761024
      ForeColor       =   0
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DcboDebitSide 
      Height          =   360
      Left            =   2520
      TabIndex        =   34
      Top             =   8880
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16761024
      ForeColor       =   0
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DcboCreditSide 
      Height          =   360
      Left            =   0
      TabIndex        =   35
      Top             =   9360
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16761024
      ForeColor       =   0
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Labelx 
      Alignment       =   1  'Right Justify
      Caption         =   "«·œŒÊ· »«·»’„…"
      DataField       =   "’„…"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   13
      Left            =   8160
      TabIndex        =   38
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Labelx 
      Alignment       =   1  'Right Justify
      Caption         =   "Õœœ «·‘Ì› "
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   4440
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1725
      Left            =   6120
      Picture         =   "CashierLogin.frx":0E1A
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1620
   End
   Begin VB.Label LBLShiftID 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   600
      Width           =   255
   End
   Begin VB.Label LBLPOSName 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   9360
      Width           =   2175
   End
   Begin VB.Label LBLPOSCode 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Labelx 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Labelx 
      Alignment       =   1  'Right Justify
      Caption         =   "„‰"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Labelx 
      Alignment       =   1  'Right Justify
      Caption         =   "»Ì«‰«  «·‘Ì› "
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   11040
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Labelx 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·‰ﬁÿ…"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   9960
      TabIndex        =   2
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label LblHeader 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "  ”ÃÌ· œŒÊ· «·‘Ì›  Ê› Õ «·‰ﬁÿ…   "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   585
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   9675
   End
   Begin VB.Label Labelx 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ «·‰ﬁÿ…"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   19
      Left            =   11520
      TabIndex        =   0
      Top             =   4440
      Width           =   975
   End
End
Attribute VB_Name = "CashierLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PointID As Integer
Dim Pointname As String
Dim Balance As Double
Dim pettyBalance As Double

Private Sub ALLButton1_Click()
On Error Resume Next
    'On Error GoTo ErrTrap
    If DCboUserName1.text = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ» «œŒ«· «”„ «·„‘—›"
    Else
    Msg = "Enter Admin User Name"
    End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboUserName.SetFocus
        Exit Sub
    End If
 
     If DcShift.text = "" Then
     If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ» «œŒ«· «·‘Ì›   "
     Else
     Msg = "Select Shift"
     End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcShift.SetFocus
        Exit Sub
    End If
    

    StrSQL = "Select * From cachierData Where id=" & Me.DCboUserName1.BoundText & " AND password='" & Trim(Me.XPTxtPass1.text) & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
         
         


            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = " √ﬂœ „‰ ’Õ… «”„ «·„‘—› " & CHR(13)
                Msg = Msg + "Êﬂ·„… «·„—Ê— Ê√⁄œ «·„Õ«Ê·…"
            Else
            
            Msg = "User Name Or Password Incorrect " & CHR(13)
            End If



        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboUserName1.SetFocus
        Exit Sub
    End If

    ' user_name = rs("UserName").value
    ' user_id = rs("UserID").value
    ' User_Password = rs("PassWord").value
 
    AddToLogFile CInt(user_id), 0, Date, Time, "  ”ÃÌ· «·œŒÊ· ·‰›ÿ… «·»Ì⁄ »√”„ «·„‘—›  " & DCboUserName1.text, " System Login", Me.Name, "L", "", ""
  '  AddSessonData val(LBLPOSCode), val(LBLShiftID.Caption), val(DCboUserName.BoundText), ShfitFrom.value, ShfitTo.value, val(LBLBalance.Caption), val(TxtBalance.text), Now
   Frame2.Visible = True
 Frame3.Visible = True
CMDLogin.Default = True
 
    Exit Sub
ErrTrap:


End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CMDLogin_Click()
On Error Resume Next
    'On Error GoTo ErrTrap
    If DCboUserName.text = "" Then
     If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ» «œŒ«· «”„ «·„” Œœ„"
     Else
     Msg = "Select User First"
     End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboUserName.SetFocus
        Exit Sub
    End If


    If dcpoint.text = "" Then
         If SystemOptions.UserInterface = ArabicInterface Then
Msg = "ÌÃ»   ÕœÌœ ‰ﬁÿ…    "
     Else
     Msg = "Select Pos First"
     End If
     
        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        dcpoint.SetFocus
        Exit Sub
    End If
    
    If XPTxtPass.text = "" Then
        '    Msg = "„‰ ›÷·ﬂ √œŒ· ﬂ·„… «·„—Ê—"
        '    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '    XPTxtPass.SetFocus
        '    Exit Sub
    End If

    StrSQL = "Select * From cachierData Where id=" & Me.DCboUserName.BoundText & " AND password='" & Trim(Me.XPTxtPass.text) & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
     
        
               If SystemOptions.UserInterface = ArabicInterface Then
   Msg = " √ﬂœ „‰ ’Õ… «”„ «·„” Œœ„ " & CHR(13)
        Msg = Msg + "Êﬂ·„… «·„—Ê— Ê√⁄œ «·„Õ«Ê·…"
        
     Else
     Msg = " wrong  user name or password"
     End If
       
       
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboUserName.SetFocus
        Exit Sub
    End If

    ' user_name = rs("UserName").value
    ' user_id = rs("UserID").value
    ' User_Password = rs("PassWord").value
 If CheckAcconts = False Then
   Exit Sub
   End If
'    AddToLogFile CInt(user_id), 0, Date, Time, "  ”ÃÌ· «·œŒÊ· ·‰›ÿ… «·»Ì⁄ »√”„  " & DCboUserName.Text, " System Login", Me.Name, "L", "", ""
AddSessonData val(LBLPOSCode), val(LBLShiftID.Caption), val(DCboUserName.BoundText), ShfitFrom.value, ShfitTo.value, val(LBLBalance.Caption), val(TxtBalance.text), Now, Null
If val(TxtBalance.text) > 0 Then
' createVoucher
    End If
    
   
     PPointID = val(dcpoint.BoundText)
   CurrentCashireID = val(DCboUserName.BoundText)
   
If SystemOptions.posshape2 = True Then
       frmsalebill3.show vbModal
     frmsalebill3.LblSessionID = SessionD
     Unload Me
 GoTo endme
 End If
 
   
 If SystemOptions.TradingPOS = True Then
   frmsalebill2.show
   frmsalebill2.LblSessionID = SessionD
' FrmCustomerDisplay.show
'
'FrmCustomerDisplay.Left = Screen.Width
'
' FrmCustomerDisplay.Top = 0
'
'FrmCustomerDisplay.maxformz

   
Else
        frmsalebill1.show
     frmsalebill1.LblSessionID = SessionD
  
  End If
  
    
    'FRMPOS.Show
    'FRMPOS.LblUserName.Caption = DCboUserName.text
 'mdifrmmain.Visible = False
' mdifrmmain.Enabled = False
endme:
    Unload Me
 
    Exit Sub
ErrTrap:

End Sub

Private Sub createVoucher()
Dim des As String
Dim DebitAccount As String
Dim CreditAccount As String
DebitAccount = DcboDebitSide.BoundText
CreditAccount = DcboCreditSide.BoundText
des = "   ”‰œ  ÕÊÌ· ⁄Âœ… »‰«¡ ⁄·Ï œŒÊ· ··ﬂ«‘Ì—  " & DCboUserName1.text & "   ··‰ﬁÿ… " & dcpoint.text & "   ··‘Ì›   " & DcShift.text & "   »√‘—«› " & DCboUserName.text & "   »—’Ìœ ”«»ﬁ  " & LBLBalance.Caption
Dim NoteID As Long
 CreateNotes NoteID, Date, Current_branch, 63, val(TxtBalance.text)
         
       CREATE_VOUCHER_GE NoteID, Current_branch, user_id, val(TxtBalance.text), DebitAccount, CreditAccount, des, Date


End Sub

Private Sub DCboUserName_Change()
 Dim PettyId As Long
   Exit Sub
    If val(DCboUserName.BoundText) = 0 Then Me.DcboDebitSide.BoundText = "":   Exit Sub
    getCashireData val(DCboUserName.BoundText), PointID, Pointname, Balance, PettyId, pettyBalance
Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", PettyId)
    Me.LBLPOSCode = PointID
    Me.LBLPOSName = Pointname
    Me.LBLBalance.Caption = Balance
   LblPetty.Caption = pettyBalance
    
    
    
    
    
End Sub

Private Sub DCboUserName_Click(Area As Integer)

DCboUserName_Change
End Sub

Function CheckAcconts() As Boolean
CheckAcconts = True
Exit Function
If Me.DcboDebitSide.BoundText = "" Then MsgBox "Õ”«» ⁄Âœ… «·ﬂ«‘Ì— €Ì— „Õœœ", vbCritical: CheckAcconts = False
If Me.DcboCreditSide.BoundText = "" Then MsgBox "Õ”«» ⁄Âœ… «·„‘—› €Ì— „Õœœ", vbCritical: CheckAcconts = False

 
End Function
Private Sub DCboUserName1_Change()
Dim My_SQL As String

 
    
    
    
 Dim PettyId As Long
    If val(DCboUserName1.BoundText) = 0 Then Me.DcboCreditSide.BoundText = "":    Exit Sub
    getCashireData val(DCboUserName1.BoundText), , 0, , PettyId
Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", PettyId)

End Sub

Private Sub DCboUserName1_Click(Area As Integer)
DCboUserName1_Change
End Sub

Private Sub dcShift_Click(Area As Integer)
   Dim ShiftFrom As Date
    Dim ShiftTo As Date
   GetShiftData val(DcShift.BoundText), ShiftFrom, ShiftTo

    LBLShiftID.Caption = val(DcShift.BoundText)
    ShfitFrom.value = ShiftFrom
    ShfitTo.value = ShiftTo
End Sub



Private Sub ChangeLang()
Me.Caption = "Cashier Login"
LblHeader.Caption = Me.Caption
    Labelx(r).Caption = "Shift"
    Labelx(2).Caption = "From"
    Labelx(3).Caption = "To"
    Frame1.Caption = "Supervisor"
     Frame2.Caption = "Point"
      Frame3.Caption = "Cashier"
      Labelx(13).Caption = "Finger Print"
          Labelx(8).Caption = " Shift"
    Labelx(10).Caption = " Name"
    Labelx(9).Caption = "Password"
ALLButton1.Caption = "LogIn"
 
    Labelx(6).Caption = " Name"
    Labelx(7).Caption = "Password"
CMDLogin.Caption = "LogIn"
    Labelx(11).Caption = " Point"
        Labelx(4).Caption = "Balance"
              Labelx(12).Caption = "Pettycash Blance"
              
     Labelx(5).Caption = "Pettycash"
     cmdCancel.Caption = "Cancel"
      
End Sub
Private Sub Form_Load()
On Error Resume Next
    Resize_Form Me
    Dim My_SQL As String
    Dim Shiftcode As String
    Dim shiftname As String
    Dim FromDate1 As Date
    Dim ToDate1 As Date

If SystemOptions.HideInfroCasher = True Then
 
Frame4.Visible = True
 

End If


Dim Dcombos As New ClsDataCombos
Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    
    
        If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "select SeftCode,SheftName From  TbLSheft "
    Else
        My_SQL = "select SeftCode,SheftNamee From TbLSheft "
    End If

   fill_combo DcShift, My_SQL

DcShift.BoundText = 1

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    'My_SQL = "select id,name From cachierData where ctype=0 "
If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "SELECT     dbo.cachierData.id, dbo.cachierData.name"
    My_SQL = My_SQL & " FROM         dbo.cachierData LEFT OUTER JOIN"
    My_SQL = My_SQL & " dbo.TblShiftWorker ON dbo.cachierData.EmpID = dbo.TblShiftWorker.EmpID"
 
 
 Else
 
   My_SQL = "SELECT     dbo.cachierData.id, dbo.cachierData.namee"
    My_SQL = My_SQL & " FROM         dbo.cachierData LEFT OUTER JOIN"
    My_SQL = My_SQL & " dbo.TblShiftWorker ON dbo.cachierData.EmpID = dbo.TblShiftWorker.EmpID"
 End If
 
    My_SQL = My_SQL & " WHERE     (dbo.cachierData.Ctype = 0)"
'    My_SQL = My_SQL & " and  PointId in (SELECT     BoxID  From dbo.Tblposdata  Where (BranchId = " & Current_branch & ")) "
  
'  My_SQL = My_SQL & " and  PointId in (SELECT     BoxID  From dbo.Tblposdata  Where (BranchId = " & Current_branch & ")) "
    My_SQL = My_SQL & "  and  cachierData.EmpID IN ( SELECT    DISTINCT dbo.cachierData.EmpID "
    My_SQL = My_SQL & " FROM         dbo.cachierData INNER JOIN"
   My_SQL = My_SQL & "                   dbo.TblUsers ON dbo.cachierData.EmpID = dbo.TblUsers.Empid"
My_SQL = My_SQL & " Where (dbo.TblUsers.UserID = " & user_id & "))"
    
    
    My_SQL = My_SQL & "       and (isCachDeactivated is null  or isCachDeactivated=0)"
   
   'My_SQL = My_SQL & "      and (isCachDeactivated is null  or isCachDeactivated=0)"
    fill_combo DCboUserName, My_SQL
    
  If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "SELECT     dbo.cachierData.id, dbo.cachierData.name"
    My_SQL = My_SQL & " FROM         dbo.cachierData  "
   My_SQL = My_SQL & " WHERE     (dbo.cachierData.Ctype = 1)"
   
    'My_SQL = My_SQL & " and  PointId in (SELECT     BoxID  From dbo.Tblposdata  Where (BranchId = " & Current_branch & ")) "
    Else
    
            My_SQL = "SELECT     dbo.cachierData.id, dbo.cachierData.namee"
    My_SQL = My_SQL & " FROM         dbo.cachierData  "
   My_SQL = My_SQL & " WHERE     (dbo.cachierData.Ctype = 1)"
   

    End If
  My_SQL = My_SQL & " and  PointId in (SELECT     BoxID  From dbo.Tblposdata  Where (BranchId = " & Current_branch & ")) "
    My_SQL = My_SQL & "       and (isCachDeactivated is null  or isCachDeactivated=0)"
        fill_combo DCboUserName1, My_SQL
        
        
            If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "select BoxID,BoxName From Tblposdata  where BranchId =" & Current_branch
    Else
        My_SQL = "select BoxID,BoxNamee From Tblposdata    where BranchId =" & Current_branch
    End If

    My_SQL = My_SQL & "  and  BoxID IN ( SELECT    DISTINCT dbo.cachierData.pointid "
    My_SQL = My_SQL & " FROM         dbo.cachierData INNER JOIN"
   My_SQL = My_SQL & "                   dbo.TblUsers ON dbo.cachierData.EmpID = dbo.TblUsers.Empid"
My_SQL = My_SQL & " Where (dbo.TblUsers.UserID = " & user_id & "))"

    fill_combo dcpoint, My_SQL


DcShift.BoundText = 1
DCboUserName1.BoundText = 3
dcpoint.BoundText = 1
DCboUserName.BoundText = user_id
End Sub

Private Sub Image2_Click()
Call Shell("OSK.exe")
End Sub

Private Sub TxtBalance_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtBalance.text, 0)
End Sub

Private Sub XPTxtPass1_GotFocus()
XPTxtPass1.text = ""
End Sub
