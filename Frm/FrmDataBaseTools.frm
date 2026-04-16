VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FrmDataBaseTools 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄„· ‰”Œ…  «Õ Ì«ÿÌ… „‰ ﬁ«⁄œ… «·»Ì«‰« "
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   Icon            =   "FrmDataBaseTools.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   7215
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
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„”«— «·„·› «·„—«œ ≈” ⁄«œ Â"
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
      Height          =   585
      Index           =   2
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   7185
      Begin ImpulseButton.ISButton Cmd 
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   150
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   609
         ButtonPositionImage=   1
         Caption         =   "..."
         BackColor       =   14871017
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmDataBaseTools.frx":000C
         ColorButton     =   14871017
      End
      Begin VB.Label LblPath 
         BackColor       =   &H00E2E9E9&
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   11
         Top             =   210
         Width           =   6405
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„”«— Õ›Ÿ «·„·› «·Õ«·Ï"
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
      Height          =   585
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3060
      Width           =   7185
      Begin ImpulseButton.ISButton Cmd 
         Height          =   345
         Index           =   4
         Left            =   90
         TabIndex        =   7
         Top             =   150
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   609
         ButtonPositionImage=   1
         Caption         =   "..."
         BackColor       =   14871017
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmDataBaseTools.frx":05A6
         ColorButton     =   14871017
      End
      Begin VB.Label LblPath 
         BackColor       =   &H00E2E9E9&
         Height          =   315
         Index           =   1
         Left            =   720
         TabIndex        =   8
         Top             =   690
         Width           =   6405
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "≈Œ — «·⁄„·Ì… «· Ï  —ÌœÂ«"
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
      Height          =   645
      Index           =   1
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   630
      Width           =   5505
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "÷€Ÿ ﬁ«⁄œ… «·»Ì«‰« "
         Height          =   345
         Index           =   2
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   930
         Width           =   4935
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈” ⁄«œ… ‰”Œ… ≈Õ Ì«ÿÌ… ”«»ﬁ… „‰ ﬁ«⁄œ… «·»Ì«‰« "
         Height          =   345
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   930
         Width           =   4935
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ›Ÿ ‰”Œ… ≈Õ Ì«ÿÌ… „‰ ﬁ«⁄œ… «·»Ì«‰« "
         Height          =   195
         Index           =   0
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   4935
      End
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   435
      Index           =   1
      Left            =   1140
      TabIndex        =   12
      Top             =   2580
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   767
      ButtonPositionImage=   1
      Caption         =   "«÷€ÿ ·⁄„· «·‰”Œ…"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmDataBaseTools.frx":0B40
      ColorButton     =   14871017
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   435
      Index           =   2
      Left            =   60
      TabIndex        =   13
      Top             =   2580
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      ButtonPositionImage=   1
      Caption         =   "Œ—ÊÃ"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmDataBaseTools.frx":0EDA
      ColorButton     =   14871017
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   435
      Index           =   3
      Left            =   3720
      TabIndex        =   14
      Top             =   3930
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      ButtonPositionImage=   1
      Caption         =   "„”«⁄œ…"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmDataBaseTools.frx":1274
      ColorButton     =   14871017
   End
   Begin MSComDlg.CommonDialog Cdg 
      Left            =   60
      Top             =   780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   435
      Index           =   5
      Left            =   2940
      TabIndex        =   15
      Top             =   2580
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   767
      ButtonPositionImage=   1
      Caption         =   "Trunacate Log File"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmDataBaseTools.frx":160E
      ColorButton     =   14871017
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -30
      X2              =   7215
      Y1              =   2460
      Y2              =   2475
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "⁄„· ‰”Œ…  «Õ Ì«ÿÌ… „‰ ﬁ«⁄œ… «·»Ì«‰« "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   555
      Index           =   0
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   7155
   End
End
Attribute VB_Name = "FrmDataBaseTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_Click(index As Integer)

    Select Case index

        Case 0

            With Me.cdg
                .CancelError = False
                .DialogTitle = "Õ›Ÿ «·‰”Œ… «·≈Õ Ì«ÿÌ… „‰ ﬁ«⁄œ… «·»Ì«‰« "
                .InitDir = App.path
'                If SystemOptions.SysDataBaseType = SQLServerDataBase Then
'                    .filter = "SQL Server Backup|*.bak"
'                Else
'                    .filter = "Microsoft Access|*.bak"
'                End If
                If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                    .filter = "All Backup Types (*.bak;*.zip;*.rar)|*.bak;*.zip;*.rar|SQL Server Backup (*.bak)|*.bak|ZIP Archive (*.zip)|*.zip|RAR Archive (*.rar)|*.rar|All Files (*.*)|*.*"
                Else
                    .filter = "All Backup Types (*.bak;*.zip;*.rar)|*.bak;*.zip;*.rar|Microsoft Access (*.bak)|*.bak|ZIP Archive (*.zip)|*.zip|RAR Archive (*.rar)|*.rar|All Files (*.*)|*.*"
                End If


                .ShowSave

                If .FileName <> "" Then
                    Me.LblPath(0).Caption = .FileName
                Else
                    Exit Sub
                End If

            End With

        Case 1

            If Me.opt(0).value = True Then

                DoBackup
                
            ElseIf Me.opt(1).value = True Then
            
            ElseIf Me.opt(2).value = True Then
            
            End If

        Case 2
            Unload Me
        Case 5
            Dim rslogFile As New ADODB.Recordset
            Dim sql8      As String
            sql8 = ""
            sql8 = "SELECT  name , DB_NAME(database_id) AS dbname  "
            sql8 = sql8 & "FROM sys.master_files "
            sql8 = sql8 & "WHERE database_id = db_id() "
            sql8 = sql8 & "  AND type = 1"
            rslogFile.Open sql8, Cn, adOpenForwardOnly, adLockReadOnly
            Dim dbname As String
            Dim dblogFile As String
            If Not rslogFile.EOF Then
                dbname = rslogFile!dbname & ""
                dblogFile = rslogFile!Name & ""
            End If
            rslogFile.Close
            If dbname <> "" Then
                sql8 = "ALTER DATABASE " & dbname & " SET RECOVERY SIMPLE;"
                sql8 = sql8 & "DBCC SHRINKFILE(" & dblogFile & ", 1);"
                sql8 = sql8 & "ALTER DATABASE " & dbname & "  SET RECOVERY FULL;"
 
                Cn.Execute sql8
                MsgBox "Done"
            End If
            
    End Select

End Sub

Private Sub Form_Load()
    CenterForm Me

    FormPostion Me, GetPostion
    opt(0).value = True
    Opt_Click 0
    LblPath(0).Caption = App.path
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub Opt_Click(index As Integer)

    If Me.opt(0).value = True Then
        Me.Fra(0).Visible = False
        Me.Fra(2).Caption = "„”«— Õ›Ÿ «·‰”Œ… «·≈Õ Ì«ÿÌ…"
    ElseIf Me.opt(1).value = True Then
        Me.Fra(0).Visible = True
        Me.Fra(2).Visible = True
        Me.Fra(2).Caption = "„”«— «·„·› «·„—«œ ≈” ⁄«œ Â"
    ElseIf Me.opt(2).value = True Then
        Me.Fra(0).Visible = False
        Me.Fra(2).Visible = False
    End If

End Sub

Private Sub DoBackup()
On Error GoTo ErrTrap
    If Trim$(Me.LblPath(0).Caption) = "" Then
        Msg = "ÌÃ»  ÕœÌœ „”«— Õ›Ÿ «·‰”Œ… «·≈Õ Ì«ÿÌ… „‰ ﬁ«⁄œ… «·»Ì«‰«  ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.Cmd(0).SetFocus
        Exit Sub
    End If

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        If Not Cn Is Nothing Then
            If Cn.State = adStateOpen Then
                Cn.Execute "BackUp DataBase " & SystemOptions.SysSQLServerDataBaseName & " To Disk = '" & Trim$(Me.LblPath(0).Caption) & "'"
MsgBox " „  ⁄„· «·‰”Œ…"

            End If
        End If

    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
    
    End If
    
    Exit Sub
ErrTrap:
Msg = " not allowed to access this location ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
