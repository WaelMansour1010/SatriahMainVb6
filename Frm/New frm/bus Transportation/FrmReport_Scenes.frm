VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReport_Scenes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "       ﬁ‹‹‹‹«—Ì— «·‰ﬁ·Ì‹‹«    "
   ClientHeight    =   6975
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   10320
   Icon            =   "FrmReport_Scenes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   6972
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10332
      Begin VB.OptionButton OPtDue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì— «·«” Õﬁ«ﬁ«  «· Ì ·Ì” ·Â« ”‰œ ’—›"
         Height          =   372
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   2520
         Width           =   3855
      End
      Begin VB.OptionButton OPTReq 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì— ÿ·»«  «·’—› ··„ ⁄ÂœÌ‰"
         Height          =   372
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   732
         Left            =   120
         TabIndex        =   43
         Top             =   6000
         Width           =   10092
         Begin VB.CommandButton Cmd 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ»«⁄…"
            Height          =   492
            Index           =   0
            Left            =   6240
            TabIndex        =   47
            Top             =   120
            Width           =   1452
         End
         Begin VB.CommandButton Cmd 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ»«⁄…  ﬁ—Ì—  Õ·Ì·Ï"
            Height          =   492
            Index           =   4
            Left            =   4800
            TabIndex        =   46
            Top             =   120
            Width           =   1452
         End
         Begin VB.CommandButton Cmd 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”Õ"
            Height          =   492
            Index           =   1
            Left            =   3360
            TabIndex        =   45
            Top             =   120
            Width           =   1452
         End
         Begin VB.CommandButton Cmd 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ—ÊÃ"
            Height          =   492
            Index           =   2
            Left            =   1920
            TabIndex        =   44
            Top             =   120
            Width           =   1452
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Height          =   3132
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   10092
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2E9E9&
            Height          =   492
            Left            =   360
            TabIndex        =   48
            Top             =   600
            Width           =   9612
            Begin VB.TextBox txtCount 
               Alignment       =   2  'Center
               Height          =   288
               Left            =   4920
               TabIndex        =   52
               Top             =   120
               Width           =   3252
            End
            Begin VB.ComboBox cbDay 
               Height          =   288
               ItemData        =   "FrmReport_Scenes.frx":000C
               Left            =   0
               List            =   "FrmReport_Scenes.frx":0025
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   120
               Width           =   3372
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ê· «·‘Â—"
               Height          =   252
               Index           =   12
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   120
               Width           =   972
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·«Ì«„"
               Height          =   252
               Index           =   10
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   120
               Width           =   972
            End
         End
         Begin VB.TextBox txtfullcode 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   5316
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   1200
            Width           =   3240
         End
         Begin MSDataListLib.DataCombo dcMangerialArea 
            Height          =   288
            Left            =   360
            TabIndex        =   22
            Top             =   1560
            Width           =   3408
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcDuration 
            Height          =   288
            Left            =   5316
            TabIndex        =   23
            Top             =   240
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCustomer 
            Height          =   315
            Left            =   360
            TabIndex        =   24
            Top             =   1200
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCar 
            Height          =   288
            Left            =   360
            TabIndex        =   25
            Top             =   1920
            Width           =   3408
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   288
            Left            =   5316
            TabIndex        =   26
            Top             =   1560
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcMonth 
            Height          =   288
            Left            =   360
            TabIndex        =   27
            Top             =   240
            Width           =   3408
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcViolationType 
            Height          =   288
            Left            =   5316
            TabIndex        =   28
            Top             =   1920
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcSchool 
            Height          =   288
            Left            =   360
            TabIndex        =   29
            Top             =   2280
            Width           =   3408
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcMContract 
            Height          =   315
            Left            =   5310
            TabIndex        =   30
            Top             =   2280
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcMinistry 
            Height          =   288
            Left            =   5316
            TabIndex        =   31
            Top             =   2640
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—… «· ⁄·Ì„Ì…"
            Height          =   312
            Index           =   9
            Left            =   3804
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   1560
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   252
            Index           =   1
            Left            =   8808
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   972
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ ⁄Âœ"
            Height          =   312
            Index           =   6
            Left            =   4548
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   1200
            Width           =   468
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”Ì«—…"
            Height          =   312
            Index           =   2
            Left            =   3996
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   1920
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   312
            Index           =   0
            Left            =   9300
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   1560
            Width           =   468
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   312
            Index           =   3
            Left            =   3804
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   240
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„Œ«·›…"
            Height          =   312
            Index           =   4
            Left            =   8700
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   1920
            Width           =   1068
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œ—”…"
            Height          =   312
            Index           =   5
            Left            =   3996
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   2280
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄ﬁœ «”‰«œ"
            Height          =   312
            Index           =   7
            Left            =   8700
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   2280
            Width           =   1068
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„ ⁄Âœ"
            Height          =   312
            Index           =   22
            Left            =   8796
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   1200
            Width           =   972
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄ﬁœ Ê“«—Ï"
            Height          =   312
            Index           =   8
            Left            =   8700
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   2640
            Width           =   1068
         End
      End
      Begin VB.OptionButton opt_Paid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì— «·„œ›Ê⁄«  ··„ ⁄ÂœÌ‰"
         Height          =   372
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1680
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.OptionButton opt_Scene 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì— «·„‘«Âœ"
         Height          =   372
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   840
         Value           =   -1  'True
         Width           =   1452
      End
      Begin VB.OptionButton opt_Attribution 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì— ⁄ﬁÊœ «·«”‰«œ"
         Height          =   372
         Left            =   8160
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   1812
      End
      Begin VB.OptionButton opt_Violation 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì— «·„Œ«·›« "
         Height          =   372
         Left            =   8160
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1920
         Width           =   1812
      End
      Begin VB.OptionButton opt_MC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì— œ›⁄«  «·„ ⁄ÂœÌ‰"
         Height          =   372
         Left            =   8160
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   2280
         Width           =   1812
      End
      Begin VB.OptionButton opt_AttribDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì—  Õ·Ì·Ï ⁄ﬁÊœ «·«”‰«œ"
         Height          =   372
         Left            =   8160
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1560
         Width           =   1812
      End
      Begin VB.OptionButton opt_VehicleAllocation 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì—  Œ’Ì’ «·Õ«›·« "
         Height          =   372
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1920
         Width           =   2175
      End
      Begin VB.OptionButton opt_stat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì— »Ì«‰ „Êﬁ› «·‰ﬁ·"
         Height          =   372
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   1812
      End
      Begin VB.OptionButton opt_oper 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì— «·Œÿ… «· ‘€Ì·Ì…"
         Height          =   372
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1560
         Width           =   1812
      End
      Begin VB.OptionButton opt_1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ·»«  «·’—›"
         Height          =   372
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   2280
         Width           =   1812
      End
      Begin VB.OptionButton opt_2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì— «·„‘«Âœ »œÊ‰ ÿ·«»"
         Height          =   372
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   1932
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   588
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   10296
         _cx             =   18150
         _cy             =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   24
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
         Caption         =   "       ﬁ‹‹‹‹«—Ì— «·‰ﬁ·Ì‹‹«    "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   6
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   1
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
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   312
         Left            =   2304
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1464
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   89587715
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   312
         Left            =   2304
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   960
         Width           =   1464
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   89587715
         CurrentDate     =   37140
      End
      Begin Dynamic_Byte.NourHijriCal FromDateH 
         Height          =   312
         Left            =   720
         TabIndex        =   15
         Top             =   960
         Width           =   1596
         _ExtentX        =   2805
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal ToDateH 
         Height          =   312
         Left            =   720
         TabIndex        =   16
         Top             =   1320
         Width           =   1596
         _ExtentX        =   2805
         _ExtentY        =   556
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   3600
         TabIndex        =   18
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ï  «—ÌŒ"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   0
         Left            =   3744
         TabIndex        =   17
         Top             =   1320
         Width           =   1176
      End
   End
End
Attribute VB_Name = "frmReport_Scenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public calltype As Integer
Public SendForm As String
Dim rs_hol As ADODB.Recordset
Dim rs_Dur As ADODB.Recordset
Dim LastDate As String
Private Sub Cmd_Click(Index As Integer)
' On Error Resume Next
    Select Case Index
    
        Case 0
                 If opt_Scene.value = True Then
                        If IsValid = True Then
                            AddSch
                            Me.print_report
                        End If
                 ElseIf opt_Attribution = True Then
                        print_AttributionContractHeader
                ElseIf opt_Violation.value = True Then
                        print_Violation
                 ElseIf opt_MC.value = True Then
                            print_MC
                 ElseIf opt_Paid.value = True Then
                            print_Paid
                ElseIf opt_AttribDetail.value = True Then
                
                
                        print_AttributionContractHeader2
                  ElseIf OPTReq.value = True Then
               print_AttributionContractHeader3
               ElseIf OPtDue.value = True Then
               print_Due
               
               
                ElseIf opt_stat.value = True Then
                        print_Statement
                        
                ElseIf opt_oper.value = True Then
                        print_oper
                        
               ElseIf opt_1.value = True Then
                    print_allReq
                     
               ElseIf opt_2.value = True Then
                        
                          If IsValid = True Then
                            AddSch
                            print_report "1"
                        End If
                 End If
          
        Case 1
            clear_all Me
            Reset_Date
        Case 2
            Unload Me
           Case 3
           GetData
           
           Case 4
                 If opt_Scene.value = True Then
                         Me.print_report
                 ElseIf opt_Attribution = True Then
                        print_AttributionContractHeader 2
                        
                    ElseIf opt_VehicleAllocation = True Then
                            'print_VehicleAllocationH
                            print_VA
                 End If
           
           
    End Select

End Sub


Private Sub dcCustomer_Change()
Dim val1, val2, recordno As String, Fullcode As String
If dcCustomer.BoundText = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and cusid = " & val(dcCustomer.BoundText)
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
          Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     End If
       TxtFullcode.Text = Fullcode
End Sub

Private Sub dcCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
        Unload FrmCompanySearch
        FrmCompanySearch.lblSearchtype = "2020"
        FrmCompanySearch.show vbModal
End If
End Sub

Private Sub dcDuration_Click(Area As Integer)
Dim i As Integer, j As Integer, str As String
    i = val(dcDuration.BoundText)
    
    If i > 0 Then
        str = "  select id , Name  from TblDurations_Details where did =   " & i
        fill_combo dcMonth, str
    Else
        str = "  select id , Name  from TblDurations_Details where did =   " & -1
        fill_combo dcMonth, str
    End If
End Sub

Private Sub dcDuration_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF3 Then
  '      FrmSearch_Duration.SendForm = "scene"
  '      FrmSearch_Duration.show
End If

End Sub

Private Sub dcMangerialArea_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF3 Then
            Unload FrmSearch_BasicData
            FrmSearch_BasicData.SendForm = "scene"
            FrmSearch_BasicData.show
            
    End If

End Sub



Private Sub dcMContract_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
    Unload FrmSearch_MinistryContract
            FrmSearch_MinistryContract.SendForm = "ReportScene_Attribuation"
            FrmSearch_MinistryContract.show
           
    End If
End Sub




Private Sub dcMinistry_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF3 Then
            Unload FrmSearch_MinistryContract
            FrmSearch_MinistryContract.SendForm = "ReportScene_MinistryContract"
            FrmSearch_MinistryContract.show
End If

End Sub

Private Sub dcSchool_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
            Unload FrmSearch_BasicData
            FrmSearch_BasicData.SendForm = "ReportScene_School"
            FrmSearch_BasicData.show
End If
End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    'Dcombos.getCountriesGovernments dcCity
    Dcombos.GetCustomersSuppliers 2, dcCustomer
    Dcombos.GetBranches Dcbranch
    Resize_Form Me
    Dim str As String
    If SystemOptions.UserInterface = ArabicInterface Then
    str = " Select ID , Name   from TblManagerialArea "
    Else
    str = " Select ID , NameE   from TblManagerialArea "
    End If
    fill_combo dcMangerialArea, str
    str = "select id , name  from TblDurations "
    fill_combo dcDuration, str
 
    
     str = " select ID , BoardNo  from tblvendorcars "
    fill_combo DCCar, str
 
    str = " select ID , NAME  from TblViolationTypes "
    fill_combo dcViolationType, str
    
    str = " select id , Name from TblSchooleFile  "
    fill_combo dcSchool, str
    
    str = " select IDAC , IDAC from TblAttributionContract    "
    fill_combo dcMContract, str
    
    str = " select IDMC , MinistryContractNo from  TblMinistryContract   "
    fill_combo dcMinistry, str
    
    Set DCboSearch = New clsDCboSearch
    Set GrdBack = New ClsBackGroundPic
    'With Me.Fg
    '    Set .WallPaper = GrdBack.Picture
    '    .AutoSize 0, .Cols - 1, False
    'End With
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    
    Reset_Date
    
End Sub

Private Sub Reset_Date()

FromDate.value = Date
ToDate.value = Date

FromDate.value = Null
ToDate.value = Null


End Sub


Private Sub Form_Unload(Cancel As Integer)

   ' FormPostion Me, SavePostion
   ' Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer


  StrSQL = StrSQL & "   SELECT dbo.TblVehicleAllocation_Details.Type, dbo.TblManagerialArea.Name, dbo.TblVehicleAllocation_Details.SchoolFileID, dbo.TblVehicleAllocation_Details.SchoolFile,"
  StrSQL = StrSQL & "                   dbo.TblVehicleAllocation_Details.SchoolMinistryno, dbo.TblVehicleAllocation_Details.BoardNo, dbo.TblVehicleAllocation_Details.DriverTel,"
  StrSQL = StrSQL & "                  dbo.TblVehicleAllocation_Details.Driver, dbo.TblAttributionContract.DurationID, dbo.TblAttributionContract.AdditionalType, dbo.TblAttributionContract.VendorID,"
  StrSQL = StrSQL & "                  dbo.TblVehicleAllocation_Details.CarID, dbo.TblAttributionContract.EndContractDateh, dbo.TblAttributionContract.MangerialAreaID,"
  StrSQL = StrSQL & "                   dbo.TblVehicleAllocation_Details.StudentCustom, dbo.TblVehicleAllocation_Details.StudentCount, dbo.TblAttributionContract.Depend, dbo.TblAttributionContract.CityID,"
  StrSQL = StrSQL & "                  dbo.TblCustemers.CusName, dbo.TblDurations.FromDate, dbo.TblDurations.FromDateH, dbo.TblDurations.ToDate, dbo.TblDurations.TODateH, dbo.TblSchooleFile.phase,"
  StrSQL = StrSQL & "                 dbo.TblSchooleFile.Telephone, dbo.TblAttributionContract.BranchID, dbo.TblVacationSchedule.DDID, dbo.TblVacationSchedule.VacationTypeID,"
  StrSQL = StrSQL & "                  dbo.TblVacationSchedule.color, dbo.TblVacationSchedule.Day, dbo.TblVacationSchedule.DateH, dbo.TblVacationSchedule.ISVac, dbo.TblVacationSchedule.Date,"
  StrSQL = StrSQL & "                  dbo.TblDurations_Details.FromDate AS Expr1, dbo.TblDurations_Details.FromDateH AS Expr2, dbo.TblDurations_Details.Name AS MonthName"
  StrSQL = StrSQL & "  , dbo.TblDurations_Details.ToDate AS Expr3, dbo.TblDurations_Details.TODateH AS Expr4 "
  StrSQL = StrSQL & "    FROM     dbo.TblVehicleAllocation_Details INNER JOIN"
  StrSQL = StrSQL & "                 dbo.TblAttributionContract ON dbo.TblVehicleAllocation_Details.IDVA = dbo.TblAttributionContract.IDAC INNER JOIN"
  StrSQL = StrSQL & "                 dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID INNER JOIN"
  StrSQL = StrSQL & "                   dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                  dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID INNER JOIN"
  StrSQL = StrSQL & "                dbo.TblSchooleFile ON dbo.TblVehicleAllocation_Details.SchoolFileID = dbo.TblSchooleFile.ID INNER JOIN"
  StrSQL = StrSQL & "               dbo.TblDurations_Details ON dbo.TblDurations.ID = dbo.TblDurations_Details.DID INNER JOIN"
  StrSQL = StrSQL & "                dbo.TblVacationSchedule ON dbo.TblDurations_Details.ID = dbo.TblVacationSchedule.DDID"
  StrSQL = StrSQL & "   Where (dbo.TblVehicleAllocation_Details.Type = 3)"
   

   

     If Me.dcDuration.BoundText <> "" Then
            StrSQL = StrSQL & "   and  DurationID =  " & val(Me.dcDuration.BoundText)
    End If
     
     If Me.dcMangerialArea.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblAttributionContract.MangerialAreaID =  " & val(Me.dcMangerialArea.BoundText)
    End If
     
     If Me.DCCar.BoundText <> "" Then
            StrSQL = StrSQL & "   and  Carid =  " & val(Me.DCCar.BoundText)
    End If
     
    If Me.dcCustomer.BoundText <> "" Then
            StrSQL = StrSQL & "   and  vendorid =  " & val(Me.dcCustomer.BoundText)
    End If
     
         If Me.Dcbranch.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblAttributionContract.BranchID =  " & val(Me.Dcbranch.BoundText)
    End If
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By IDMC "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
          '  Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.Lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

         
    End If

End Sub

Private Sub ChangeLang()
 
End Sub

Function print_Scene()
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

     MySQL = "  SELECT dbo.TblVehicleAllocation_Details.Type, dbo.TblManagerialArea.Name, dbo.TblVehicleAllocation_Details.SchoolFileID, dbo.TblVehicleAllocation_Details.SchoolFile, "
     MySQL = MySQL & "   dbo.TblVehicleAllocation_Details.SchoolMinistryno, dbo.TblVehicleAllocation_Details.BoardNo, dbo.TblVehicleAllocation_Details.DriverTel,"
     MySQL = MySQL & "               dbo.TblVehicleAllocation_Details.Driver, dbo.TblAttributionContract.DurationID, dbo.TblAttributionContract.AdditionalType, dbo.TblAttributionContract.VendorID,"
     MySQL = MySQL & "                dbo.TblVehicleAllocation_Details.CarID, dbo.TblAttributionContract.EndContractDateh, dbo.TblAttributionContract.MangerialAreaID,"
     MySQL = MySQL & "               dbo.TblVehicleAllocation_Details.StudentCustom, dbo.TblVehicleAllocation_Details.StudentCount, dbo.TblAttributionContract.Depend, dbo.TblAttributionContract.CityID,"
     MySQL = MySQL & "               dbo.TblCustemers.CusName , dbo.TblDurations.FromDate, dbo.TblDurations.FromDateH, dbo.TblDurations.ToDate, dbo.TblDurations.ToDateH"
     MySQL = MySQL & " , dbo.TblSchooleFile.phase,  dbo.TblSchooleFile.Telephone "
     MySQL = MySQL & "  FROM     dbo.TblVehicleAllocation_Details INNER JOIN"
     MySQL = MySQL & "                dbo.TblAttributionContract ON dbo.TblVehicleAllocation_Details.IDVA = dbo.TblAttributionContract.IDAC INNER JOIN"
     MySQL = MySQL & "              dbo.TblCustemers ON dbo.TblAttributionContract.vendorid = dbo.TblCustemers.CusID INNER JOIN"
     MySQL = MySQL & "               dbo.TblDurations ON dbo.TblAttributionContract.DurationID  = dbo.TblDurations.ID LEFT OUTER JOIN"
     MySQL = MySQL & "               dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID"
     MySQL = MySQL & "  INNER JOIN  dbo.TblSchooleFile ON dbo.TblVehicleAllocation_Details.SchoolFileID = dbo.TblSchooleFile.ID"
     MySQL = MySQL & "  WHERE  (dbo.TblVehicleAllocation_Details.Type = 3)"
   
   

     If Me.dcDuration.BoundText <> "" Then
            MySQL = MySQL & "   and  DurationID =  " & val(Me.dcDuration.BoundText)
    End If
 
    'If Me.dcCity.BoundText <> "" Then
    '        MySQL = MySQL & "   and   dbo.TblAttributionContract.CityID =  " & val(Me.dcCity.BoundText)
    'End If
    
     If Me.dcMangerialArea.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.MangerialAreaID =  " & val(Me.dcMangerialArea.BoundText)
    End If
     
     If Me.DCCar.BoundText <> "" Then
            MySQL = MySQL & "   and  Carid =  " & val(Me.DCCar.BoundText)
    End If
     
    
    


    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_Scenes.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_Scenes.rpt"
    End If

    If Dir(StrFileName) = "" Then

        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If RsData.BOF Or RsData.EOF Then

    '    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    RsData.Close
    '    Set RsData = Nothing
    '    Screen.MousePointer = vbDefault
    '    Exit Function
    'End If
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    'xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    'If SystemOptions.UserInterface = ArabicInterface Then
    '    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    '
    '    StrReportTitle = "" '& StrAccountName
    ' Else
    '
    '    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    '
    '     xReport.ParameterFields(4).AddCurrentValue get_branch_name(Val(my_branch))
    '    StrReportTitle = ""
    ' End If
    'xReport.ParameterFields(3).AddCurrentValue user_name
    'xReport.ReportTitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    'xReport.ApplicationName = App.Title
    'xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault



End Function

Function print_report(Optional NoteSerial As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

   

  MySQL = MySQL & "  SELECT  dbo.TblVehicleAllocation_Details.Type, dbo.TblManagerialArea.Name, dbo.TblVehicleAllocation_Details.SchoolFileID,"
  MySQL = MySQL & "                  dbo.TblVehicleAllocation_Details.SchoolFile, dbo.TblVehicleAllocation_Details.SchoolMinistryno, dbo.TblVehicleAllocation_Details.BoardNo,"
  MySQL = MySQL & "                  dbo.TblVehicleAllocation_Details.DriverTel, dbo.TblVehicleAllocation_Details.Driver, dbo.TblAttributionContract.DurationID, dbo.TblAttributionContract.AdditionalType,"
  MySQL = MySQL & "                  dbo.TblAttributionContract.VendorID, dbo.TblVehicleAllocation_Details.CarID, dbo.TblAttributionContract.EndContractDateh, dbo.TblAttributionContract.MangerialAreaID,"
  MySQL = MySQL & "                 dbo.TblVehicleAllocation_Details.StudentCustom, dbo.TblVehicleAllocation_Details.StudentCount, dbo.TblAttributionContract.Depend, dbo.TblAttributionContract.CityID,"
  MySQL = MySQL & "                  dbo.TblCustemers.CusName, dbo.TblDurations.FromDate, dbo.TblDurations.FromDateH, dbo.TblDurations.ToDate, dbo.TblDurations.TODateH, dbo.TblSchooleFile.phase,"
  MySQL = MySQL & "                  dbo.TblSchooleFile.Telephone, dbo.TblAttributionContract.BranchID, dbo.TblVacationSchedule2.DDID, dbo.TblVacationSchedule2.VacationTypeID,"
  MySQL = MySQL & "                  dbo.TblVacationSchedule2.color, dbo.TblVacationSchedule2.Day, dbo.TblVacationSchedule2.DateH, dbo.TblVacationSchedule2.ISVac, dbo.TblVacationSchedule2.Date,"
  MySQL = MySQL & "                  dbo.TblDurations_Details.FromDate AS Expr1, dbo.TblDurations_Details.FromDateH AS Expr2, dbo.TblDurations_Details.Name AS MonthName,"
  MySQL = MySQL & "                  dbo.TblDurations_Details.ToDate AS Expr3, dbo.TblDurations_Details.TODateH AS Expr4, dbo.TblAttributionContract.IDAC"
  MySQL = MySQL & "  FROM     dbo.TblVehicleAllocation_Details INNER JOIN"
  MySQL = MySQL & "                  dbo.TblAttributionContract ON dbo.TblVehicleAllocation_Details.IDVA = dbo.TblAttributionContract.IDAC INNER JOIN"
  MySQL = MySQL & "                   dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID INNER JOIN"
  MySQL = MySQL & "                  dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                  dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID INNER JOIN"
  MySQL = MySQL & "                 dbo.TblDurations_Details ON dbo.TblDurations.ID = dbo.TblDurations_Details.DID INNER JOIN"
  MySQL = MySQL & "                 dbo.TblVacationSchedule2 ON dbo.TblDurations_Details.ID = dbo.TblVacationSchedule2.DDID LEFT OUTER JOIN"
  MySQL = MySQL & "                  dbo.TblSchooleFile ON dbo.TblVehicleAllocation_Details.SchoolFileID = dbo.TblSchooleFile.ID"
  MySQL = MySQL & "  Where (dbo.TblVehicleAllocation_Details.Type = 3)"
   

     If Me.dcDuration.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.DurationID =  " & val(Me.dcDuration.BoundText)
    End If
   
    If Me.dcMangerialArea.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.MangerialAreaID =  " & val(Me.dcMangerialArea.BoundText)
    End If
     
     If Me.DCCar.BoundText <> "" Then
            MySQL = MySQL & "   and  Carid =  " & val(Me.DCCar.BoundText)
    End If
   
    If Me.Dcbranch.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.BranchID =  " & val(Me.Dcbranch.BoundText)
    End If
  
     If Me.dcCustomer.BoundText <> "" Then
           MySQL = MySQL & "   and  vendorid =  " & val(Me.dcCustomer.BoundText)
    End If
    
    If Me.dcMonth.BoundText <> "" Then
           MySQL = MySQL & "   and  TblDurations_Details.id =  " & val(Me.dcMonth.BoundText)
    End If
    
    
     If Me.dcMContract.BoundText <> "" Then
           MySQL = MySQL & "   and  TblAttributionContract.IDAC  =  " & val(Me.dcMContract.BoundText)
    End If
    
      If Me.dcSchool.BoundText <> "" Then
           MySQL = MySQL & "   and  TblSchooleFile.id =  " & val(Me.dcSchool.BoundText)
    End If
    
    
    If Not IsNull(FromDate.value) Then
            MySQL = MySQL & "  and TblAttributionContract.FromDate >= " & SQLDate(FromDate.value, True)
    End If
    
    If Not IsNull(ToDate.value) Then
            MySQL = MySQL & "  and TblAttributionContract.ToDate <= " & SQLDate(ToDate.value, True)
    End If
  

   If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Scenes.rpt"
    Else
          StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Scenes.rpt"
    End If
    

    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
            xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    Else
            xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    End If
    
    If NoteSerial = "1" Then
            xReport.ParameterFields(2).AddCurrentValue "1"
    Else
            xReport.ParameterFields(2).AddCurrentValue ""
    End If
  xReport.ParameterFields(3).AddCurrentValue MonthStart
     xReport.ParameterFields(4).AddCurrentValue LastDate
     
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
End Function

Function print_AttributionContractHeader(Optional opt As Integer = 1)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

      If opt = 1 Then
                    MySQL = MySQL & "      SELECT dbo.TblAttributionContract.IDAC, dbo.TblAttributionContract.Name, dbo.TblAttributionContract.FromDate, dbo.TblAttributionContract.FromDateH,"
                    MySQL = MySQL & "      dbo.TblAttributionContract.ToDate, dbo.TblAttributionContract.ToDateH, dbo.TblAttributionContract.VendorID, dbo.TblAttributionContract.CityID,"
                    MySQL = MySQL & "      dbo.TblAttributionContract.StudentCount, dbo.TblAttributionContract.StudentCustom, dbo.TblAttributionContract.DisCount, dbo.TblAttributionContract.MinistryContractNo,"
                    MySQL = MySQL & "      dbo.TblAttributionContract.MangerialAreaID, dbo.TblManagerialArea.Name AS TblManagerialAreaName, dbo.TblManagerialArea.Namee AS TblManagerialAreaNameE,"
                    MySQL = MySQL & "      dbo.TblDurations.Name AS DurationName, dbo.TblAttributionContract.BranchID, dbo.TblBranchesData.branch_namee, dbo.TblBranchesData.branch_name,"
                    MySQL = MySQL & "      dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.EndContractDateh, dbo.TblAttributionContract.StartContractDate,"
                    MySQL = MySQL & "      dbo.TblAttributionContract.StartContractDateh, dbo.TblAttributionContract.NetValue, dbo.TblAttributionContract.ContractType, dbo.TblAttributionContract.ActualDayValue,"
                    MySQL = MySQL & "      dbo.TblAttributionContract.DayValue, dbo.TblAttributionContract.DaysCount, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.CusID,"
                    MySQL = MySQL & "      dbo.TblCountriesGovernments.GovernmentName   ,  dbo.TblCustemers.Fullcode, dbo.TblCustemers.RecordNo"
                    MySQL = MySQL & "        FROM     dbo.TblAttributionContract INNER JOIN"
                    MySQL = MySQL & "      dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID INNER JOIN"
                    MySQL = MySQL & "      dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID INNER JOIN"
                    MySQL = MySQL & "      dbo.TblBranchesData ON dbo.TblAttributionContract.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
                    MySQL = MySQL & "      dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID INNER JOIN"
                    MySQL = MySQL & "       dbo.TblCountriesGovernments ON dbo.TblAttributionContract.CityID = dbo.TblCountriesGovernments.GovernmentID  "
                    MySQL = MySQL & "      WHERE  (dbo.TblCustemers.Type = 2)  "
      ElseIf opt = 2 Then
                    MySQL = MySQL & "  SELECT dbo.TblAttributionContract.IDAC, dbo.TblAttributionContract.Name, dbo.TblAttributionContract.FromDate, dbo.TblAttributionContract.FromDateH,"
                    MySQL = MySQL & "                     dbo.TblAttributionContract.ToDate, dbo.TblAttributionContract.ToDateH, dbo.TblAttributionContract.VendorID, dbo.TblAttributionContract.CityID,"
                    MySQL = MySQL & "                     dbo.TblAttributionContract.StudentCount, dbo.TblAttributionContract.StudentCustom, dbo.TblAttributionContract.DisCount, dbo.TblAttributionContract.MinistryContractNo,"
                    MySQL = MySQL & "                     dbo.TblVehicleAllocation_Details.CarID, dbo.TblVehicleAllocation_Details.Driver, dbo.TblVehicleAllocation_Details.SchoolFileID, dbo.TblVehicleAllocation_Details.SchoolFile,"
                    MySQL = MySQL & "                     dbo.TblVehicleAllocation_Details.SchoolMinistryno, dbo.TblVehicleAllocation_Details.SchoolStudentCount, dbo.TblVehicleAllocation_Details.SchoolStudentAvailable,"
                    MySQL = MySQL & "                     dbo.TblVehicleAllocation_Details.Custom, dbo.TblVehicleAllocation_Details.DayRate, dbo.TblVehicleAllocation_Details.MaxCap,"
                    MySQL = MySQL & "                     dbo.TblVehicleAllocation_Details.VehicleAvailableSite, dbo.TblVehicleAllocation_Details.VehicleSiteCount, dbo.TblVehicleAllocation_Details.Capecity,"
                    MySQL = MySQL & "                     dbo.TblVehicleAllocation_Details.DriverID, dbo.TblVehicleAllocation_Details.BoardNo, dbo.TblAttributionContract.MangerialAreaID,"
                    MySQL = MySQL & "                     dbo.TblManagerialArea.Name AS TblManagerialAreaName, dbo.TblManagerialArea.Namee AS TblManagerialAreaNameE,"
                    MySQL = MySQL & "                     dbo.TblVehicleAllocation_Details.StudentCount AS StudentCountD, dbo.TblVehicleAllocation_Details.StudentCustom AS StudentCustomD,"
                    MySQL = MySQL & "                     dbo.TblDurations.Name AS DurationName, dbo.TblAttributionContract.BranchID, dbo.TblBranchesData.branch_namee, dbo.TblBranchesData.branch_name,"
                    MySQL = MySQL & "                     dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.EndContractDateh, dbo.TblAttributionContract.StartContractDate,"
                    MySQL = MySQL & "                     dbo.TblAttributionContract.StartContractDateh, dbo.TblAttributionContract.NetValue, dbo.TblAttributionContract.ContractType, dbo.TblAttributionContract.ActualDayValue,"
                    MySQL = MySQL & "                     dbo.TblAttributionContract.DayValue, dbo.TblAttributionContract.DaysCount, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.CusID,"
                    MySQL = MySQL & "                     dbo.TblCountriesGovernments.GovernmentName  , dbo.TblVehicleAllocation_Details.ID AS IDD ,  dbo.TblCustemers.Fullcode, dbo.TblCustemers.RecordNo"
                    MySQL = MySQL & "   FROM     dbo.TblAttributionContract INNER JOIN"
                    MySQL = MySQL & "                     dbo.TblVehicleAllocation_Details ON dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA INNER JOIN"
                    MySQL = MySQL & "                     dbo.TblVendorCars ON dbo.TblVehicleAllocation_Details.CarID = dbo.TblVendorCars.ID INNER JOIN"
                    MySQL = MySQL & "                     dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID INNER JOIN"
                    MySQL = MySQL & "                     dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID INNER JOIN"
                    MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TblAttributionContract.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
                    MySQL = MySQL & "                     dbo.TblCustemers ON dbo.TblAttributionContract.VendorID  = dbo.TblCustemers.CusID INNER JOIN"
                    MySQL = MySQL & "                     dbo.TblCountriesGovernments ON dbo.TblAttributionContract.CityID = dbo.TblCountriesGovernments.GovernmentID"
                    MySQL = MySQL & "   Where  (dbo.TblVehicleAllocation_Details.Type = 3) AND (dbo.TblCustemers.Type = 2)  "
                    
                    
                    If Me.dcSchool.BoundText <> "" Then
                                 MySQL = MySQL & "   and  TblVehicleAllocation_Details.SchoolFileID =  " & val(Me.dcSchool.BoundText)
                    End If
                    
                    
      End If
     
   

    If Me.dcDuration.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.DurationID =  " & val(Me.dcDuration.BoundText)
    End If
 
     
   If Me.dcMangerialArea.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.MangerialAreaID =  " & val(Me.dcMangerialArea.BoundText)
    End If
     
     If Me.DCCar.BoundText <> "" Then
            MySQL = MySQL & "   and  Carid =  " & val(Me.DCCar.BoundText)
    End If
   
    If Me.Dcbranch.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.BranchID =  " & val(Me.Dcbranch.BoundText)
    End If
  
    If dcCustomer.BoundText <> "" Then
              MySQL = MySQL & "   and  dbo.TblAttributionContract.VendorID =  " & val(Me.dcCustomer.BoundText)
    End If
  
  '  If Me.dcMonth.BoundText <> "" Then
  '         MySQL = MySQL & "   and  TblDurations_Details.id =  " & val(Me.dcMonth.BoundText)
  '  End If
  
    If Me.dcMContract.BoundText <> "" Then
           MySQL = MySQL & "   and  TblAttributionContract.IDAC  =  " & val(Me.dcMContract.BoundText)
    End If
    
      If Not IsNull(FromDate.value) Then
            MySQL = MySQL & "  and TblAttributionContract.FromDate >= " & SQLDate(FromDate.value, True)
    End If
    
    If Not IsNull(ToDate.value) Then
            MySQL = MySQL & "  and TblAttributionContract.ToDate <= " & SQLDate(ToDate.value, True)
    End If
    
      If dcMinistry.BoundText <> "" Then
            MySQL = MySQL & "  and TblAttributionContract.IDMC = " & val(dcMinistry.BoundText)
    End If
                            
            
  If opt = 1 Then
        MySQL = MySQL & "    order by IDAC "
  ElseIf opt = 2 Then
          MySQL = MySQL & "  ORDER BY dbo.TblAttributionContract.IDAC, IDD    "
  End If
   'MySQL = MySQL & "   group by TblAttributionContract.IDAC , TblAttributionContract.VendorID "
   If SystemOptions.UserInterface = ArabicInterface Then
            If opt = 2 Then
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractHeader.rpt"
            ElseIf opt = 1 Then
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractDetail.rpt"
            End If
    Else
             If opt = 2 Then
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractHeader.rpt"
            ElseIf opt = 1 Then
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractDetail.rpt"
            End If
            
    End If


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
 
End Function

Function print_VehicleAllocation(Optional opt As Integer = 1)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

      If opt = 1 Then
                    MySQL = MySQL & "      SELECT dbo.TblAttributionContract.IDAC, dbo.TblAttributionContract.Name, dbo.TblAttributionContract.FromDate, dbo.TblAttributionContract.FromDateH,"
                    MySQL = MySQL & "      dbo.TblAttributionContract.ToDate, dbo.TblAttributionContract.ToDateH, dbo.TblAttributionContract.VendorID, dbo.TblAttributionContract.CityID,"
                    MySQL = MySQL & "      dbo.TblAttributionContract.StudentCount, dbo.TblAttributionContract.StudentCustom, dbo.TblAttributionContract.DisCount, dbo.TblAttributionContract.MinistryContractNo,"
                    MySQL = MySQL & "      dbo.TblAttributionContract.MangerialAreaID, dbo.TblManagerialArea.Name AS TblManagerialAreaName, dbo.TblManagerialArea.Namee AS TblManagerialAreaNameE,"
                    MySQL = MySQL & "      dbo.TblDurations.Name AS DurationName, dbo.TblAttributionContract.BranchID, dbo.TblBranchesData.branch_namee, dbo.TblBranchesData.branch_name,"
                    MySQL = MySQL & "      dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.EndContractDateh, dbo.TblAttributionContract.StartContractDate,"
                    MySQL = MySQL & "      dbo.TblAttributionContract.StartContractDateh, dbo.TblAttributionContract.NetValue, dbo.TblAttributionContract.ContractType, dbo.TblAttributionContract.ActualDayValue,"
                    MySQL = MySQL & "      dbo.TblAttributionContract.DayValue, dbo.TblAttributionContract.DaysCount, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.CusID,"
                    MySQL = MySQL & "      dbo.TblCountriesGovernments.GovernmentName   ,  dbo.TblCustemers.Fullcode, dbo.TblCustemers.RecordNo"
                    MySQL = MySQL & "        FROM     dbo.TblAttributionContract INNER JOIN"
                    MySQL = MySQL & "      dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID INNER JOIN"
                    MySQL = MySQL & "      dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID INNER JOIN"
                    MySQL = MySQL & "      dbo.TblBranchesData ON dbo.TblAttributionContract.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
                    MySQL = MySQL & "      dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID INNER JOIN"
                    MySQL = MySQL & "       dbo.TblCountriesGovernments ON dbo.TblAttributionContract.CityID = dbo.TblCountriesGovernments.GovernmentID  "
                    MySQL = MySQL & "      WHERE  (dbo.TblCustemers.Type = 2)  "
      ElseIf opt = 2 Then
                   MySQL = MySQL & "  SELECT dbo.TblCarsData.BoardNO, dbo.TblCarsData.id, dbo.TblVehicleAllocation.DurationID, dbo.TblVehicleAllocation.MinistryNo, dbo.TblVehicleAllocation.StudentAlloc, "
                   MySQL = MySQL & "  dbo.TblVehicleAllocation.StudentAttrib, dbo.TblVehicleAllocation.ToDateH, dbo.TblVehicleAllocation.FromDateH, dbo.TblVehicleAllocation.ToDate,"
                   MySQL = MySQL & "     dbo.TblVehicleAllocation.FromDate, dbo.TblVehicleAllocation.StudentCount, dbo.TblVehicleAllocation.IDMC, dbo.TblVehicleAllocation.SchoolFileID,"
                   MySQL = MySQL & "      dbo.TblVehicleAllocation.ProcessNo, dbo.TblVehicleAllocation.IDVA, dbo.TblSchooleFile.Name AS SchoolName, dbo.TblSchooleFile.StudentCount AS SchoolStudentCount,"
            MySQL = MySQL & "        dbo.TblSchooleFile.ministerNo AS SchoolministerNo, dbo.TblDurations.Name AS DurationName, dbo.TblDurations.FromDate AS DurFromDate,"
             MySQL = MySQL & "       dbo.TblDurations.FromDateH AS DurFromDateH, dbo.TblDurations.ToDate AS DurToDate, dbo.TblDurations.TODateH AS DurToDateH, dbo.TblEmployee.Emp_Name,"
              MySQL = MySQL & "      dbo.TblEmployee.Emp_Name1 , dbo.TblEmployee.fullcode, dbo.TblVehicleAllocation_Details.Type"
MySQL = MySQL & "  FROM     dbo.TblEmployee RIGHT OUTER JOIN"
      MySQL = MySQL & "              dbo.TblDurations INNER JOIN"
          MySQL = MySQL & "         dbo.TblVehicleAllocation ON dbo.TblDurations.ID = dbo.TblVehicleAllocation.DurationID INNER JOIN"
         MySQL = MySQL & "           dbo.TblSchooleFile ON dbo.TblVehicleAllocation.SchoolFileID = dbo.TblSchooleFile.ID INNER JOIN"
            MySQL = MySQL & "        dbo.TblVehicleAllocation_Details ON dbo.TblVehicleAllocation.IDVA = dbo.TblVehicleAllocation_Details.IDVA ON"
            MySQL = MySQL & "        dbo.TblEmployee.Emp_ID = dbo.TblVehicleAllocation_Details.DriverID LEFT OUTER JOIN"
           MySQL = MySQL & "         dbo.TblCarsData ON dbo.TblVehicleAllocation_Details.CarID = dbo.TblCarsData.id"
MySQL = MySQL & "  WHERE  (dbo.TblVehicleAllocation_Details.Type = 2)  "
                    
                    
                    If Me.dcSchool.BoundText <> "" Then
                                 MySQL = MySQL & "   and  TblVehicleAllocation_Details.SchoolFileID =  " & val(Me.dcSchool.BoundText)
                    End If
      End If
     
   

    If Me.dcDuration.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.DurationID =  " & val(Me.dcDuration.BoundText)
    End If
 
     
   If Me.dcMangerialArea.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.MangerialAreaID =  " & val(Me.dcMangerialArea.BoundText)
    End If
     
     If Me.DCCar.BoundText <> "" Then
            MySQL = MySQL & "   and  Carid =  " & val(Me.DCCar.BoundText)
    End If
   
    If Me.Dcbranch.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.BranchID =  " & val(Me.Dcbranch.BoundText)
    End If
  
    If dcCustomer.BoundText <> "" Then
              MySQL = MySQL & "   and  dbo.TblAttributionContract.VendorID =  " & val(Me.dcCustomer.BoundText)
    End If
  
  '  If Me.dcMonth.BoundText <> "" Then
  '         MySQL = MySQL & "   and  TblDurations_Details.id =  " & val(Me.dcMonth.BoundText)
  '  End If
  
    If Me.dcMContract.BoundText <> "" Then
           MySQL = MySQL & "   and  TblAttributionContract.IDAC  =  " & val(Me.dcMContract.BoundText)
    End If
    
      If Not IsNull(FromDate.value) Then
            MySQL = MySQL & "  and TblAttributionContract.FromDate >= " & SQLDate(FromDate.value, True)
    End If
    
    If Not IsNull(ToDate.value) Then
            MySQL = MySQL & "  and TblAttributionContract.ToDate <= " & SQLDate(ToDate.value, True)
    End If
    
      If dcMinistry.BoundText <> "" Then
            MySQL = MySQL & "  and TblAttributionContract.IDMC = " & val(dcMinistry.BoundText)
    End If
                            
            
  If opt = 1 Then
        MySQL = MySQL & "    order by IDAC "
  ElseIf opt = 2 Then
          MySQL = MySQL & "  ORDER BY dbo.TblAttributionContract.IDAC, IDD "
  End If
  
   If SystemOptions.UserInterface = ArabicInterface Then
            If opt = 2 Then
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractHeader.rpt"
            ElseIf opt = 1 Then
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractDetail.rpt"
            End If
    Else
             If opt = 2 Then
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractHeader.rpt"
            ElseIf opt = 1 Then
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractDetail.rpt"
            End If
            
    End If


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
 
End Function



Function print_Violation(Optional opt As Integer = 1)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    
    MySQL = ""

MySQL = MySQL & "  SELECT DISTINCT  TblConfirmViolation.id , dbo.TblVehicleAllocation_Details.BoardNo, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblConfirmViolation.VendorID AS Customerid,"
MySQL = MySQL & "  dbo.TblViolationTypes.Type, dbo.TblViolationTypes.Name, dbo.TblViolationTypes.absence, dbo.TblConfirmViolation.AbsenceCount, dbo.TblConfirmViolation.Value,"
MySQL = MySQL & "  dbo.TblConfirmViolation.Date, dbo.TblConfirmViolation.DateH, dbo.TblConfirmViolation.ViolationID, dbo.TblConfirmViolation.DurationID, dbo.TblDurations.Name AS DurName,"
MySQL = MySQL & "  dbo.TblConfirmViolation.MonthID, dbo.TblConfirmViolation.CarID, dbo.TblConfirmViolation.MinistryContractID, dbo.TblConfirmViolation.ViolationType,"
MySQL = MySQL & "  dbo.TblConfirmViolation.MinistryContractValue, dbo.TblConfirmViolation.ID, dbo.TblDurations_Details.Name AS MonthName, dbo.TblCustemers.Fullcode,"
MySQL = MySQL & "  dbo.TblCustemers.RecordNo, dbo.TblVehicleAllocation_Details.CarID AS Expr1"
MySQL = MySQL & "  FROM     dbo.TblConfirmViolation INNER JOIN"
MySQL = MySQL & "  dbo.TblViolationTypes ON dbo.TblConfirmViolation.ViolationID = dbo.TblViolationTypes.ID INNER JOIN"
MySQL = MySQL & "  dbo.TblCustemers ON dbo.TblConfirmViolation.VendorID = dbo.TblCustemers.CusID INNER JOIN"
MySQL = MySQL & "  dbo.TblDurations ON dbo.TblConfirmViolation.DurationID = dbo.TblDurations.ID INNER JOIN"
MySQL = MySQL & "  dbo.TblDurations_Details ON dbo.TblConfirmViolation.MonthID = dbo.TblDurations_Details.ID INNER JOIN"
MySQL = MySQL & "  dbo.TblAttributionContract ON dbo.TblConfirmViolation.MinistryContractID = dbo.TblAttributionContract.IDAC INNER JOIN"
MySQL = MySQL & "  dbo.TblVehicleAllocation_Details ON dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA AND"
MySQL = MySQL & "  dbo.TblConfirmViolation.CarID = dbo.TblVehicleAllocation_Details.CarID"
MySQL = MySQL & "  Where (dbo.TblVehicleAllocation_Details.Type = 3)"
   
   

     If Me.dcDuration.BoundText <> "" Then
            MySQL = MySQL & "   and  TblConfirmViolation.DurationID =  " & val(Me.dcDuration.BoundText)
    End If
 
     
     If Me.DCCar.BoundText <> "" Then
            MySQL = MySQL & "   and  TblConfirmViolation.CarID =  " & val(Me.DCCar.BoundText)
    End If
   
  
    If dcCustomer.BoundText <> "" Then
              MySQL = MySQL & "   and  TblConfirmViolation.VendorID =  " & val(Me.dcCustomer.BoundText)
    End If
  
    If dcMonth.BoundText <> "" Then
              MySQL = MySQL & "   and  TblConfirmViolation.MonthID =  " & val(Me.dcMonth.BoundText)
    End If
    
    If dcViolationType.BoundText <> "" Then
              MySQL = MySQL & "   and  TblConfirmViolation.ViolationID =  " & val(Me.dcViolationType.BoundText)
    End If
    
    
   If Me.dcMContract.BoundText <> "" Then
           MySQL = MySQL & "   and  TblConfirmViolation.MinistryContractID  =  " & val(Me.dcMContract.BoundText)
    End If
    
      
       If Not IsNull(FromDate.value) Then
            MySQL = MySQL & "  and TblConfirmViolation.Date >= " & SQLDate(FromDate.value, True)
    End If
    
    If Not IsNull(ToDate.value) Then
            MySQL = MySQL & "  and TblConfirmViolation.Date <= " & SQLDate(ToDate.value, True)
    End If
      
      
      MySQL = MySQL & "   order by TblConfirmViolation.id  "
    
    If SystemOptions.UserInterface = ArabicInterface Then
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_Violation.rpt"
    Else
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_Violation.rpt"
    End If


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

       
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
 
End Function


Function print_MC(Optional opt As Integer = 1)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    
   
  MySQL = MySQL & " SELECT dbo.TblMinistryContract_Installment.InstallmentNo, dbo.TblMinistryContract_Installment.Value, dbo.TblAttributionContract.DurationID,"
  MySQL = MySQL & " dbo.TblDurations.Name AS DurationName, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblMinistryContract_Installment.Type,"
  MySQL = MySQL & " dbo.TblMinistryContract_Installment.Due_DateH, dbo.TblMinistryContract_Installment.Due_Date, dbo.TblCustemers.Fullcode, dbo.TblMinistryContract_Installment.ID,"
  MySQL = MySQL & " dbo.TblCustemers.CusID, dbo.TblMinistryContract_Installment.MonthID, dbo.TblMinistryContract_Installment.IDMC, dbo.TblAttributionContract.StartContractDate,"
  MySQL = MySQL & " dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.FromDate, dbo.TblDurations.FromDate AS DurFromDate, dbo.TblDurations.ToDate AS DurToDate,"
  MySQL = MySQL & " dbo.ACCOUNTS.Account_Code , dbo.ACCOUNTS.account_serial"
  MySQL = MySQL & " , dbo.TblAttributionContract.MangerialAreaID, dbo.TblMinistryContract_Installment.CarID,   dbo.TblAttributionContract.BranchId , dbo.TblAttributionContract.VendorID"
  MySQL = MySQL & "  FROM     dbo.TblAttributionContract INNER JOIN"
  MySQL = MySQL & "  dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID INNER JOIN"
  MySQL = MySQL & "  dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
  MySQL = MySQL & "  dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID LEFT OUTER JOIN"
  MySQL = MySQL & "  dbo.TblMinistryContract_Installment ON dbo.TblAttributionContract.IDAC = dbo.TblMinistryContract_Installment.IDMC"
   
  MySQL = MySQL & "  where   TblMinistryContract_Installment.Type = 2"
    If Me.dcDuration.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.DurationID =  " & val(Me.dcDuration.BoundText)
    End If
      
    If Me.dcMangerialArea.BoundText <> "" Then
            MySQL = MySQL & "   and  dbo.TblAttributionContract.MangerialAreaID =  " & val(Me.dcMangerialArea.BoundText)
    End If
     
     If Me.DCCar.BoundText <> "" Then
            MySQL = MySQL & "   and  dbo.TblMinistryContract_Installment.CarID =  " & val(Me.DCCar.BoundText)
    End If
   
    If Me.Dcbranch.BoundText <> "" Then
            MySQL = MySQL & "   and  dbo.TblAttributionContract.BranchId =  " & val(Me.Dcbranch.BoundText)
    End If
  
    If dcCustomer.BoundText <> "" Then
              MySQL = MySQL & "   and  dbo.TblAttributionContract.VendorID =  " & val(Me.dcCustomer.BoundText)
    End If
  
    If Me.dcMonth.BoundText <> "" Then
           MySQL = MySQL & "   and  dbo.TblMinistryContract_Installment.MonthID  =  " & val(Me.dcMonth.BoundText)
    End If
  
  
  
   If Me.dcMContract.BoundText <> "" Then
           MySQL = MySQL & "   and  TblAttributionContract.IDAC  =  " & val(Me.dcMContract.BoundText)
    End If
    
  If Not IsNull(FromDate.value) Then
            MySQL = MySQL & "  and TblAttributionContract.FromDate >= " & SQLDate(FromDate.value, True)
    End If
    
    If Not IsNull(ToDate.value) Then
            MySQL = MySQL & "  and TblAttributionContract.ToDate <= " & SQLDate(ToDate.value, True)
    End If
    
     If SystemOptions.UserInterface = ArabicInterface Then
              StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VendorRequest.rpt"
    Else
              StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VendorRequest.rpt"
    End If

'MySQL = " order by TblAttributionContract.idac "
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

   
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
 
End Function

Function print_Paid(Optional opt As Integer = 1)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    
   
  MySQL = MySQL & "  SELECT dbo.TblVendorReceipt.ID, dbo.TblVendorReceipt.BranchID, dbo.TblVendorReceipt.Depend, dbo.TblVendorReceipt.Code, dbo.TblVendorReceipt.ExchangeType, "
  MySQL = MySQL & "                  dbo.TblVendorReceipt.DurationID, dbo.TblVendorReceipt.Month, dbo.TblVendorReceipt.DurationName, dbo.TblVendorReceipt.Date, dbo.TblVendorReceipt.DateH,"
  MySQL = MySQL & "                  dbo.TblVendorReceipt.DependID, dbo.TblVendorReceipt_Details.CusID, dbo.TblDurations.Name AS DurName, dbo.TblDurations_Details.Name AS MonthName,"
  MySQL = MySQL & "                  dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblVendorCars.BoardNo, dbo.TblCustemers.Fullcode, dbo.TblCustemers.RecordNo, dbo.TblCustemers.IBAN,"
  MySQL = MySQL & "                  dbo.TblAttributionContract.IDAC , dbo.TblVendorReceipt_Details.net"
  MySQL = MySQL & "    FROM     dbo.TblVendorReceipt INNER JOIN"
  MySQL = MySQL & "                  dbo.TblVendorReceipt_Details ON dbo.TblVendorReceipt.ID = dbo.TblVendorReceipt_Details.HID INNER JOIN"
  MySQL = MySQL & "                  dbo.TblDurations ON dbo.TblVendorReceipt.DurationID = dbo.TblDurations.ID INNER JOIN"
  MySQL = MySQL & "                  dbo.TblCustemers ON dbo.TblVendorReceipt_Details.CusID = dbo.TblCustemers.CusID INNER JOIN"
  MySQL = MySQL & "                  dbo.TblDurations_Details ON dbo.TblVendorReceipt.Month = dbo.TblDurations_Details.ID INNER JOIN"
  MySQL = MySQL & "                  dbo.TblVendorCars ON dbo.TblVendorReceipt_Details.carid = dbo.TblVendorCars.ID INNER JOIN"
  MySQL = MySQL & "                  dbo.TblAttributionContract ON dbo.TblVendorReceipt_Details.IDAC = dbo.TblAttributionContract.IDAC"



    If Me.dcDuration.BoundText <> "" Then
            MySQL = MySQL & "   and  TblVendorReceipt.DurationID =  " & val(Me.dcDuration.BoundText)
    End If
   
     
     If Me.DCCar.BoundText <> "" Then
            MySQL = MySQL & "   and  dbo.TblVendorCars.ID =  " & val(Me.DCCar.BoundText)
    End If
   
    If Me.Dcbranch.BoundText <> "" Then
            MySQL = MySQL & "   and  dbo.TblVendorReceipt.BranchID  =  " & val(Me.Dcbranch.BoundText)
    End If
  
    If dcCustomer.BoundText <> "" Then
              MySQL = MySQL & "   and  dbo.TblVendorReceipt_Details.CusID=  " & val(Me.dcCustomer.BoundText)
    End If
  
    If Me.dcMonth.BoundText <> "" Then
           MySQL = MySQL & "   and  dbo.TblVendorReceipt.Month   =  " & val(Me.dcMonth.BoundText)
    End If
  
  
  
   If Me.dcMContract.BoundText <> "" Then
           MySQL = MySQL & "   and  dbo.TblAttributionContract.IDAC  =  " & val(Me.dcMContract.BoundText)
    End If
    
    
      If Not IsNull(FromDate.value) Then
            MySQL = MySQL & "  and TblVendorReceipt.Date >= " & SQLDate(FromDate.value, True)
    End If
    
    If Not IsNull(ToDate.value) Then
            MySQL = MySQL & "  and TblVendorReceipt.Date <= " & SQLDate(ToDate.value, True)
    End If
    
    
     If SystemOptions.UserInterface = ArabicInterface Then
              StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VendorPaid.rpt"
    Else
              StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VendorPaid.rpt"
    End If

'MySQL = " order by TblAttributionContract.idac "
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

   
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
 
End Function


Private Sub FromDate_Change()
On Error Resume Next
  FromDateH.value = ToHijriDate(FromDate.value)
End Sub


Private Sub Fromdateh_LostFocus()
VBA.Calendar = vbCalGreg
          FromDate.value = ToGregorianDate(FromDateH.value)
End Sub




Private Sub Option1_Click()

End Sub

Private Sub opt_1_Click()
Frame4.Enabled = False
End Sub

Private Sub opt_2_Click()
Frame4.Enabled = True
End Sub

Private Sub opt_AttribDetail_Click()
Frame4.Enabled = False
End Sub

Private Sub opt_Attribution_Click()
Frame4.Enabled = False
End Sub

Private Sub opt_MC_Click()
Frame4.Enabled = False
End Sub

Private Sub opt_oper_Click()
Frame4.Enabled = False
End Sub

Private Sub opt_Scene_Click()
Frame4.Enabled = True

End Sub

Private Sub opt_stat_Click()
Frame4.Enabled = False
End Sub

Private Sub opt_VehicleAllocation_Click()
Frame4.Enabled = False
End Sub

Private Sub opt_Violation_Click()
Frame4.Enabled = False
End Sub

Private Sub ToDate_Change()
On Error Resume Next
ToDateH.value = ToHijriDate(ToDate.value)
End Sub

Private Sub todateH_LostFocus()
VBA.Calendar = vbCalGreg
          ToDate.value = ToGregorianDate(ToDateH.value)
End Sub

Private Sub txtfullcode_Change()
Dim val1, val2
If TxtFullcode.Text = "" Then Exit Sub
Dim str As String, recordno As String, CusID As String
recordno = ""
CusID = ""

    str = " select * From TblCustemers where Type=2  and fullcode = '" & TxtFullcode & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
     Else
             dcCustomer.BoundText = ""
    End If
        dcCustomer.BoundText = CusID
End Sub



Function print_AttributionContractHeader2(Optional opt As Integer = 1)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

        MySQL = MySQL & "   SELECT dbo.TblAttributionContract.IDAC, dbo.TblAttributionContract.Name, dbo.TblAttributionContract.FromDate, dbo.TblAttributionContract.FromDateH,"
        MySQL = MySQL & "                     dbo.TblAttributionContract.ToDate, dbo.TblAttributionContract.ToDateH, dbo.TblAttributionContract.VendorID, dbo.TblAttributionContract.CityID,"
        MySQL = MySQL & "                     dbo.TblAttributionContract.StudentCount, dbo.TblAttributionContract.StudentCustom, dbo.TblAttributionContract.DisCount, dbo.TblAttributionContract.MinistryContractNo,"
        MySQL = MySQL & "                     dbo.TblVehicleAllocation_Details.CarID, dbo.TblVehicleAllocation_Details.Driver, dbo.TblVehicleAllocation_Details.SchoolFileID, dbo.TblVehicleAllocation_Details.SchoolFile,"
        MySQL = MySQL & "                     dbo.TblVehicleAllocation_Details.SchoolMinistryno, dbo.TblVehicleAllocation_Details.SchoolStudentCount, dbo.TblVehicleAllocation_Details.SchoolStudentAvailable,"
        MySQL = MySQL & "                     dbo.TblVehicleAllocation_Details.Custom, dbo.TblVehicleAllocation_Details.DayRate, dbo.TblVehicleAllocation_Details.MaxCap,"
        MySQL = MySQL & "                     dbo.TblVehicleAllocation_Details.VehicleAvailableSite, dbo.TblVehicleAllocation_Details.VehicleSiteCount, dbo.TblVehicleAllocation_Details.Capecity,"
        MySQL = MySQL & "                     dbo.TblVehicleAllocation_Details.DriverID, dbo.TblVehicleAllocation_Details.BoardNo, dbo.TblAttributionContract.MangerialAreaID,"
        MySQL = MySQL & "                     dbo.TblManagerialArea.Name AS TblManagerialAreaName, dbo.TblManagerialArea.Namee AS TblManagerialAreaNameE,"
        MySQL = MySQL & "                     dbo.TblVehicleAllocation_Details.StudentCount AS StudentCountD, dbo.TblVehicleAllocation_Details.StudentCustom AS StudentCustomD,"
        MySQL = MySQL & "                     dbo.TblDurations.Name AS DurationName, dbo.TblAttributionContract.BranchID, dbo.TblBranchesData.branch_namee, dbo.TblBranchesData.branch_name,"
        MySQL = MySQL & "                     dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.EndContractDateh, dbo.TblAttributionContract.StartContractDate,"
        MySQL = MySQL & "                     dbo.TblAttributionContract.StartContractDateh, dbo.TblAttributionContract.NetValue, dbo.TblAttributionContract.ContractType, dbo.TblAttributionContract.ActualDayValue,"
        MySQL = MySQL & "                     dbo.TblAttributionContract.DayValue, dbo.TblAttributionContract.DaysCount, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.CusID,"
        MySQL = MySQL & "                     dbo.TblCountriesGovernments.GovernmentName  , dbo.TblVehicleAllocation_Details.ID AS IDD ,  dbo.TblCustemers.Fullcode, dbo.TblCustemers.RecordNo"
        MySQL = MySQL & "   FROM     dbo.TblAttributionContract INNER JOIN"
        MySQL = MySQL & "                     dbo.TblVehicleAllocation_Details ON dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA INNER JOIN"
        MySQL = MySQL & "                     dbo.TblVendorCars ON dbo.TblVehicleAllocation_Details.CarID = dbo.TblVendorCars.ID INNER JOIN"
        MySQL = MySQL & "                     dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID INNER JOIN"
        MySQL = MySQL & "                     dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID INNER JOIN"
        MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TblAttributionContract.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
        MySQL = MySQL & "                     dbo.TblCustemers ON dbo.TblAttributionContract.VendorID  = dbo.TblCustemers.CusID INNER JOIN"
        MySQL = MySQL & "                     dbo.TblCountriesGovernments ON dbo.TblAttributionContract.CityID = dbo.TblCountriesGovernments.GovernmentID"
        MySQL = MySQL & "   Where  (dbo.TblVehicleAllocation_Details.Type = 3) AND (dbo.TblCustemers.Type = 2)  "
        
                    
        If Me.dcSchool.BoundText <> "" Then
                     MySQL = MySQL & "   and  TblVehicleAllocation_Details.SchoolFileID =  " & val(Me.dcSchool.BoundText)
        End If
 
   

    If Me.dcDuration.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.DurationID =  " & val(Me.dcDuration.BoundText)
    End If
 
     
   If Me.dcMangerialArea.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.MangerialAreaID =  " & val(Me.dcMangerialArea.BoundText)
    End If
     
     If Me.DCCar.BoundText <> "" Then
            MySQL = MySQL & "   and  Carid =  " & val(Me.DCCar.BoundText)
    End If
   
    If Me.Dcbranch.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.BranchID =  " & val(Me.Dcbranch.BoundText)
    End If
  
    If dcCustomer.BoundText <> "" Then
              MySQL = MySQL & "   and  dbo.TblAttributionContract.VendorID =  " & val(Me.dcCustomer.BoundText)
    End If
  
  '  If Me.dcMonth.BoundText <> "" Then
  '         MySQL = MySQL & "   and  TblDurations_Details.id =  " & val(Me.dcMonth.BoundText)
  '  End If
  
    If Me.dcMContract.BoundText <> "" Then
           MySQL = MySQL & "   and  TblAttributionContract.IDAC  =  " & val(Me.dcMContract.BoundText)
    End If
    
      If Not IsNull(FromDate.value) Then
            MySQL = MySQL & "  and TblAttributionContract.FromDate >= " & SQLDate(FromDate.value, True)
    End If
    
    If Not IsNull(ToDate.value) Then
            MySQL = MySQL & "  and TblAttributionContract.ToDate <= " & SQLDate(ToDate.value, True)
    End If
     
  If opt = 1 Then
        MySQL = MySQL & "    order by IDAC "
  ElseIf opt = 2 Then
          MySQL = MySQL & "  ORDER BY dbo.TblAttributionContract.IDAC, IDD "
  End If
  
   If SystemOptions.UserInterface = ArabicInterface Then
           
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractHeader2.rpt"
          
    Else
           
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractHeader2.rpt"
           
    End If


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
 
End Function
Function print_AttributionContractHeader3(Optional opt As Integer = 1)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

                     
 MySQL = "SELECT     dbo.TblExchangeRequest.Remarks, dbo.TblExchangeReques_Detailst.remark, dbo.TblExchangeReques_Detailst.BankAccount, dbo.TblDurations.Name AS DurName, "
MySQL = MySQL & "                       dbo.TblExchangeReques_Detailst.HID, dbo.TblExchangeReques_Detailst.CusID, dbo.TblExchangeReques_Detailst.InsID, dbo.TblExchangeReques_Detailst.InsNo,"
MySQL = MySQL & "                       dbo.TblExchangeReques_Detailst.CusName, dbo.TblExchangeReques_Detailst.FullCode, dbo.TblExchangeReques_Detailst.[Value],"
MySQL = MySQL & "                       dbo.TblExchangeReques_Detailst.Total_Deduct, dbo.TblExchangeReques_Detailst.net, dbo.TblExchangeReques_Detailst.wokdays,"
MySQL = MySQL & "                       dbo.TblExchangeReques_Detailst.stopdays, dbo.TblExchangeReques_Detailst.stopvalue, dbo.TblExchangeReques_Detailst.Account_Serial,"
MySQL = MySQL & "                       dbo.TblExchangeReques_Detailst.Account_Code, dbo.TblExchangeReques_Detailst.carid, dbo.TblExchangeReques_Detailst.boardno,"
MySQL = MySQL & "                       dbo.TblExchangeReques_Detailst.dayvalue, dbo.TblExchangeReques_Detailst.absenceDays, dbo.TblExchangeReques_Detailst.absenceValue,"
MySQL = MySQL & "                       dbo.TblExchangeReques_Detailst.ContractBDate, dbo.TblExchangeReques_Detailst.ContractEDate, dbo.TblExchangeReques_Detailst.ContractDate,"
MySQL = MySQL & "                       dbo.TblExchangeRequest.DurationID, dbo.TblExchangeRequest.[Month], dbo.TblDurations_Details.Name AS MonthName, dbo.TblExchangeRequest.[Date],"
MySQL = MySQL & "                       dbo.TblExchangeRequest.DateH, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblExchangeRequest.ID AS MainTblID,"
MySQL = MySQL & "                       dbo.TblExchangeReques_Detailst.IDAC, dbo.TblExchangeReques_Detailst.ID AS DID, dbo.TblAttributionInstallmentDivided.RE_Paid,"
MySQL = MySQL & "                       dbo.TblAttributionInstallmentDivided.REID, dbo.TblExchangeRequest.EntryCreated, dbo.TblAttributionInstallmentDivided.TotalValue,"
MySQL = MySQL & "                       dbo.TblAttributionInstallmentDivided.PayMentPayed , dbo.TblAttributionInstallmentDivided.VReceipt_Paid, dbo.TblAttributionInstallmentDivided.VReceiptID"
MySQL = MySQL & " FROM         dbo.TblDurations_Details INNER JOIN"
MySQL = MySQL & "                       dbo.TblDurations INNER JOIN"
MySQL = MySQL & "                       dbo.TblExchangeRequest INNER JOIN"
MySQL = MySQL & "                       dbo.TblExchangeReques_Detailst ON dbo.TblExchangeRequest.ID = dbo.TblExchangeReques_Detailst.HID ON"
MySQL = MySQL & "                       dbo.TblDurations.ID = dbo.TblExchangeRequest.DurationID ON dbo.TblDurations_Details.ID = dbo.TblExchangeRequest.[Month] INNER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.TblExchangeRequest.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
MySQL = MySQL & "                       dbo.TblAttributionInstallmentDivided ON dbo.TblExchangeReques_Detailst.InsID = dbo.TblAttributionInstallmentDivided.ID"
  '
  '      If Me.dcSchool.BoundText <> "" Then
  '                   MySQL = MySQL & "   and  TblVehicleAllocation_Details.SchoolFileID =  " & val(Me.dcSchool.BoundText)
  '      End If
 '
   MySQL = "  SELECT     dbo.TblAttributionInstallmentDivided.REID, dbo.TblExchangeRequest.Remarks, dbo.TblExchangeReques_Detailst.remark, "
       MySQL = MySQL & "                        dbo.TblExchangeReques_Detailst.BankAccount, dbo.TblDurations.Name AS DurName, dbo.TblExchangeReques_Detailst.HID,"
       MySQL = MySQL & "                        dbo.TblExchangeReques_Detailst.CusID, dbo.TblExchangeReques_Detailst.InsID, dbo.TblExchangeReques_Detailst.InsNo,"
       MySQL = MySQL & "                        dbo.TblExchangeReques_Detailst.CusName, dbo.TblExchangeReques_Detailst.FullCode, dbo.TblExchangeReques_Detailst.[Value],"
       MySQL = MySQL & "                        dbo.TblExchangeReques_Detailst.Total_Deduct, dbo.TblExchangeReques_Detailst.net, dbo.TblExchangeReques_Detailst.wokdays,"
       MySQL = MySQL & "                        dbo.TblExchangeReques_Detailst.stopdays, dbo.TblExchangeReques_Detailst.stopvalue, dbo.TblExchangeReques_Detailst.Account_Serial,"
       MySQL = MySQL & "                        dbo.TblExchangeReques_Detailst.Account_Code, dbo.TblExchangeReques_Detailst.carid, dbo.TblExchangeReques_Detailst.boardno,"
       MySQL = MySQL & "                        dbo.TblExchangeReques_Detailst.dayvalue, dbo.TblExchangeReques_Detailst.absenceDays, dbo.TblExchangeReques_Detailst.absenceValue,"
       MySQL = MySQL & "                        dbo.TblExchangeReques_Detailst.ContractBDate, dbo.TblExchangeReques_Detailst.ContractEDate, dbo.TblExchangeReques_Detailst.ContractDate,"
       MySQL = MySQL & "                        dbo.TblExchangeRequest.DurationID, dbo.TblExchangeRequest.[Month], dbo.TblDurations_Details.Name AS MonthName, dbo.TblExchangeRequest.[Date],"
       MySQL = MySQL & "                        dbo.TblExchangeRequest.DateH, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblExchangeRequest.ID AS MainTblID,"
       MySQL = MySQL & "                        dbo.TblExchangeReques_Detailst.IDAC, dbo.TblExchangeReques_Detailst.ID AS DID, dbo.TblAttributionInstallmentDivided.RE_Paid,"
       MySQL = MySQL & "                        dbo.TblExchangeRequest.EntryCreated, dbo.TblAttributionInstallmentDivided.TotalValue, dbo.TblAttributionInstallmentDivided.PayMentPayed,"
       MySQL = MySQL & "                        dbo.TblAttributionInstallmentDivided.VReceipt_Paid , dbo.TblAttributionInstallmentDivided.VReceiptID"
       MySQL = MySQL & "  FROM         dbo.TblDurations_Details INNER JOIN"
       MySQL = MySQL & "                        dbo.TblDurations INNER JOIN"
       MySQL = MySQL & "                        dbo.TblExchangeRequest INNER JOIN"
       MySQL = MySQL & "                        dbo.TblExchangeReques_Detailst ON dbo.TblExchangeRequest.ID = dbo.TblExchangeReques_Detailst.HID ON"
       MySQL = MySQL & "                        dbo.TblDurations.ID = dbo.TblExchangeRequest.DurationID ON dbo.TblDurations_Details.ID = dbo.TblExchangeRequest.[Month] INNER JOIN"
       MySQL = MySQL & "                        dbo.TblBranchesData ON dbo.TblExchangeRequest.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
       MySQL = MySQL & "                        dbo.TblAttributionInstallmentDivided ON dbo.TblExchangeReques_Detailst.InsID = dbo.TblAttributionInstallmentDivided.DetailsID AND"
       MySQL = MySQL & "                        dbo.TblExchangeReques_Detailst.HID = dbo.TblAttributionInstallmentDivided.REID"
                      
MySQL = MySQL & "         where 1=1"

    If Me.dcDuration.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionInstallmentDivided.DurationID =  " & val(Me.dcDuration.BoundText)
    End If
 
     
   If Me.dcMangerialArea.BoundText <> "" Then
            MySQL = MySQL & "   and  TblExchangeRequest .MangerialAreaID =  " & val(Me.dcMangerialArea.BoundText)
    End If
     
     If Me.DCCar.BoundText <> "" Then
            MySQL = MySQL & "   and TblExchangeReques_Detailst.boardno ='" & (Me.DCCar.Text) & "'"
    End If
   
    If Me.Dcbranch.BoundText <> "" Then
            MySQL = MySQL & "   and  TblExchangeRequest.BranchID =  " & val(Me.Dcbranch.BoundText)
    End If
  
    If dcCustomer.BoundText <> "" Then
              MySQL = MySQL & "   and  dbo.TblExchangeReques_Detailst.CusID =  " & val(Me.dcCustomer.BoundText)
    End If
  
    If Me.dcMonth.BoundText <> "" Then
           MySQL = MySQL & "   and  TblExchangeRequest.month =  " & val(Me.dcMonth.BoundText)
    End If
  
    If Me.dcMContract.BoundText <> "" Then
           MySQL = MySQL & "   and  TblExchangeReques_Detailst.IDAC  =  " & val(Me.dcMContract.BoundText)
    End If
    
      If Not IsNull(FromDate.value) Then
            MySQL = MySQL & "  and TblExchangeRequest.Date >= " & SQLDate(FromDate.value, True)
    End If
    
    If Not IsNull(ToDate.value) Then
            MySQL = MySQL & "  and TblExchangeRequest.Date <= " & SQLDate(ToDate.value, True)
    End If
     
 ' If Opt = 1 Then
 '       MySQL = MySQL & "    order by IDAC "
 ' ElseIf Opt = 2 Then
 '         MySQL = MySQL & "  ORDER BY dbo.TblAttributionContract.IDAC, IDD "
 ' End If
  
   If SystemOptions.UserInterface = ArabicInterface Then
           
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractHeader3.rpt"
          
    Else
           
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractHeader3.rpt"
           
    End If


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
 
End Function
Function print_Due(Optional opt As Integer = 1)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

                     
 MySQL = " SELECT DISTINCT "
MySQL = MySQL & "                          dbo.TblVehicleAllocation_Details.DayRate * dbo.MonthActualWorkDays(dbo.TblDurations_Details.ID) AS Totalvalue,"
MySQL = MySQL & "                         dbo.TblDurations_Details.ID AS MonthID, dbo.TblDurations.Name AS DurationName, dbo.TblMinistryContract_Installment.VRID,"
MySQL = MySQL & "                         dbo.TblVehicleAllocation_Details.DayRate, dbo.TblAttributionContract.IDAC, dbo.TblVehicleAllocation_Details.BoardNo, dbo.TblCustemers.RecordNo,"
MySQL = MySQL & "                         dbo.TblCustemers.Fullcode, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblAttributionContract.StartContractDate,"
MySQL = MySQL & "                         dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.FromDate, dbo.TblDurations.FromDate AS DurFromDate,"
MySQL = MySQL & "                         dbo.TblDurations.FromDateH AS DurFromDateH, dbo.TblDurations.ToDate AS DurToDate, dbo.TblVehicleAllocation_Details.StudentCount,"
MySQL = MySQL & "                         dbo.TblVehicleAllocation_Details.Custom, dbo.TblVehicleAllocation_Details.StudentCustom, dbo.TblVehicleAllocation_Details.CarID, dbo.TblDurations_Details.Name,"
MySQL = MySQL & "                         dbo.TblVehicleAllocation_Details.ID, dbo.TblVehicleAllocation_Details.SchoolFileID, dbo.TblAttributionContract.StopDeal, dbo.TblAttributionContract.StopDate,"
MySQL = MySQL & "                         dbo.TblAttributionContract.StopDateH, dbo.TblAttributionContract.BranchID, dbo.TblCustemers.Account_Code, dbo.ACCOUNTS.Account_Serial, dbo.TblCustemers.IBAN,"
MySQL = MySQL & "                         dbo.TblCustemers.BankAccount, dbo.TblDurations.Type AS DurationType, dbo.TblAttributionInstallmentDivided.ID AS DividID,"
MySQL = MySQL & "                         dbo.TblAttributionContract.MangerialAreaID, dbo.TblManagerialArea.Name AS MAName, dbo.TblAttributionContract.CityID, dbo.TblAttributionContract.VendorID,"
MySQL = MySQL & "                         dbo.TblSchooleFile.Name AS schoolName, dbo.TblSchooleFile.Namee AS schoolNamee"
MySQL = MySQL & "  , dbo.TblMinistryContract_Installment.Due_Date  FROM         dbo.TblCustemers LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblDurations_Details INNER JOIN"
MySQL = MySQL & "                         dbo.TblAttributionContract INNER JOIN"
MySQL = MySQL & "                         dbo.TblVehicleAllocation_Details ON dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA INNER JOIN"
MySQL = MySQL & "                         dbo.TblAttributionInstallmentDivided ON dbo.TblVehicleAllocation_Details.ID = dbo.TblAttributionInstallmentDivided.DetailsID ON"
MySQL = MySQL & "                         dbo.TblDurations_Details.ID = dbo.TblAttributionInstallmentDivided.MonthID INNER JOIN"
MySQL = MySQL & "                         dbo.TblSchooleFile ON dbo.TblVehicleAllocation_Details.SchoolFileID = dbo.TblSchooleFile.ID INNER JOIN"
MySQL = MySQL & "                         dbo.TblMinistryContract_Installment ON dbo.TblAttributionContract.IDAC = dbo.TblMinistryContract_Installment.IDMC AND"
MySQL = MySQL & "                         dbo.TblAttributionInstallmentDivided.MonthID = dbo.TblMinistryContract_Installment.MonthID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID ON"
MySQL = MySQL & "                         dbo.TblCustemers.CusID = dbo.TblAttributionContract.VendorID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblVendorCars ON dbo.TblVehicleAllocation_Details.CarID = dbo.TblVendorCars.ID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID"
MySQL = MySQL & "   WHERE     (dbo.TblVehicleAllocation_Details.Type = 3) AND (dbo.TblAttributionInstallmentDivided.RE_Paid IS NULL OR"
 MySQL = MySQL & "                        dbo.TblAttributionInstallmentDivided.RE_Paid = 0) AND (dbo.TblMinistryContract_Installment.Type = 2) AND (NOT (dbo.TblMinistryContract_Installment.VRID IS NULL))"

 

    If Me.dcDuration.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionInstallmentDivided.DurationID =  " & val(Me.dcDuration.BoundText)
    End If
 
     
   If Me.dcMangerialArea.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract .MangerialAreaID =  " & val(Me.dcMangerialArea.BoundText)
    End If
     
     If Me.DCCar.BoundText <> "" Then
            MySQL = MySQL & "   and TblVehicleAllocation_Details.boardno ='" & (Me.DCCar.Text) & "'"
    End If
   
    If Me.Dcbranch.BoundText <> "" Then
            MySQL = MySQL & "   and  TblAttributionContract.BranchID =  " & val(Me.Dcbranch.BoundText)
    End If
  
    If dcCustomer.BoundText <> "" Then
              MySQL = MySQL & "   and  dbo.TblCustemers.CusID =  " & val(Me.dcCustomer.BoundText)
    End If
  
    If Me.dcMonth.BoundText <> "" Then
           MySQL = MySQL & "   and  TblDurations_Details.id =  " & val(Me.dcMonth.BoundText)
    End If
  
    If Me.dcMContract.BoundText <> "" Then
           MySQL = MySQL & "   and  TblAttributionContract.IDAC  =  " & val(Me.dcMContract.BoundText)
    End If
    
      If Not IsNull(FromDate.value) Then
            MySQL = MySQL & "  and TblMinistryContract_Installmen.Due_Date >= " & SQLDate(FromDate.value, True)
    End If
    
    If Not IsNull(ToDate.value) Then
            MySQL = MySQL & "  and TblMinistryContract_Installmen.Due_Date <= " & SQLDate(ToDate.value, True)
    End If
     
 ' If Opt = 1 Then
 '       MySQL = MySQL & "    order by IDAC "
 ' ElseIf Opt = 2 Then
 '         MySQL = MySQL & "  ORDER BY dbo.TblAttributionContract.IDAC, IDD "
 ' End If
  
   If SystemOptions.UserInterface = ArabicInterface Then
           
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractDUE.rpt"
          
    Else
           
                    StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractDUE.rpt"
           
    End If


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
 
End Function



Function print_VehicleAllocationH()
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String


MySQL = MySQL & "   SELECT dbo.TblCarsData.BoardNO, dbo.TblVehicleAllocation.DurationID, dbo.TblVehicleAllocation.MinistryNo, dbo.TblVehicleAllocation.StudentAlloc,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.StudentAttrib, dbo.TblVehicleAllocation.ToDateH, dbo.TblVehicleAllocation.FromDateH, dbo.TblVehicleAllocation.ToDate,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.FromDate, dbo.TblVehicleAllocation.StudentCount, dbo.TblVehicleAllocation.IDMC, dbo.TblVehicleAllocation.SchoolFileID,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.ProcessNo, dbo.TblVehicleAllocation.IDVA, dbo.TblSchooleFile.Name AS SchoolName, dbo.TblSchooleFile.StudentCount AS SchoolStudentCount,"
MySQL = MySQL & "   dbo.TblSchooleFile.ministerNo AS SchoolministerNo, dbo.TblDurations.Name AS DurationName, dbo.TblDurations.FromDate AS DurFromDate,"
MySQL = MySQL & "   dbo.TblDurations.FromDateH AS DurFromDateH, dbo.TblDurations.ToDate AS DurToDate, dbo.TblDurations.TODateH AS DurToDateH, dbo.TblEmployee.Emp_Name,"
MySQL = MySQL & "   dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Fullcode, dbo.TblVehicleAllocation_Details.Type, dbo.TblMinistryContract.MinistryContractNo,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.ProcessDate, dbo.TblVehicleAllocation_Details.Capecity, dbo.TblVehicleAllocation_Details.MaxCap, dbo.TblCarsData.code,"
MySQL = MySQL & "   dbo.TblVehicleAllocation_Details.StudentCount AS StudentCount1, dbo.TblVehicleAllocation_Details.VehicleAvailableSite, dbo.TblEmployee.NumEkama,"
MySQL = MySQL & "   dbo.TblEmployee.Emp_Phone , dbo.TblEmployee.Emp_mobile, dbo.TblVehicleAllocation_Details.CarID , dbo.TblVehicleAllocation_Details.VehicleSiteCount"
MySQL = MySQL & "   FROM     dbo.TblMinistryContract INNER JOIN"
MySQL = MySQL & "   dbo.TblDurations INNER JOIN"
MySQL = MySQL & "   dbo.TblVehicleAllocation ON dbo.TblDurations.ID = dbo.TblVehicleAllocation.DurationID INNER JOIN"
MySQL = MySQL & "   dbo.TblSchooleFile ON dbo.TblVehicleAllocation.SchoolFileID = dbo.TblSchooleFile.ID INNER JOIN"
MySQL = MySQL & "   dbo.TblVehicleAllocation_Details ON dbo.TblVehicleAllocation.IDVA = dbo.TblVehicleAllocation_Details.IDVA ON"
MySQL = MySQL & "   dbo.TblMinistryContract.IDMC = dbo.TblVehicleAllocation.IDMC INNER JOIN"
MySQL = MySQL & "   dbo.TblCarsData ON dbo.TblVehicleAllocation_Details.CarID = dbo.TblCarsData.id left outer JOIN"
MySQL = MySQL & "   dbo.TblEmployee ON dbo.TblVehicleAllocation_Details.DriverID = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & "   WHERE   dbo.TblVehicleAllocation_Details.Type = 2 "




        If Me.dcSchool.BoundText <> "" Then
                     MySQL = MySQL & "   and  TblVehicleAllocation.SchoolFileID =  " & val(Me.dcSchool.BoundText)
        End If
 
   

    If Me.dcDuration.BoundText <> "" Then
            MySQL = MySQL & "   and  TblVehicleAllocation.DurationID =  " & val(Me.dcDuration.BoundText)
    End If
 
  
    If Me.dcMinistry.BoundText <> "" Then
           MySQL = MySQL & "   and  TblVehicleAllocation.IDMC  =  " & val(Me.dcMContract.BoundText)
    End If
    
      If Not IsNull(FromDate.value) Then
            MySQL = MySQL & "  and TblVehicleAllocation.FromDate >= " & SQLDate(FromDate.value, True)
    End If
    
    If Not IsNull(ToDate.value) Then
            MySQL = MySQL & "  and TblVehicleAllocation.ToDate <= " & SQLDate(ToDate.value, True)
    End If
     
        MySQL = MySQL & "    order by TblVehicleAllocation.IDVA "


    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VehicleAllocationH.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VehicleAllocationH.rpt"
    End If

    If Dir(StrFileName) = "" Then

        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.EnableParameterPrompting = False
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault



End Function

Function print_Statement(Optional opt As Integer = 1)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   
   
  If dcDuration.BoundText = "" Then
    MsgBox ("«Œ — «·› —… «·œ—«”Ì… «Ê·« ")
    Exit Function
  End If
   
  
 ' If dcMinistry.BoundText = "" Then
 '   MsgBox ("«Œ — «·› —… «·⁄ﬁœ «·Ê“«—Ï «Ê·« ")
 '   Exit Function
 ' End If
  
   
  MySQL = MySQL & "  select id , name , studentcount , vehiclealloc ,atrributionalloc , ( coalesce(vehiclealloc,0) +  coalesce(atrributionalloc,0)) tot "
 MySQL = MySQL & "  ,  (coalesce ( studentcount , 0 ) -  ( coalesce(vehiclealloc,0) +  coalesce(atrributionalloc,0)) ) diff"
 MySQL = MySQL & "  From"
 MySQL = MySQL & "  ("
 MySQL = MySQL & "  SELECT dbo.TblSchooleFile.id, dbo.TblSchooleFile.Name, dbo.TblSchooleFile.StudentCount,"
 MySQL = MySQL & "  (select  sum (StudentAlloc )  from  TblVehicleAllocation"
 MySQL = MySQL & "  where TblVehicleAllocation.SchoolFileID = TblSchooleFile.id and durationid = " & val(dcDuration.BoundText) & " "
  
 If dcMinistry.BoundText <> "" Then
 MySQL = MySQL & " and IDMC = " & val(dcMinistry.BoundText) & ""
 End If
 
 MySQL = MySQL & " ) VehicleAlloc ,"
 MySQL = MySQL & "  (select  sum (D.StudentCount )  from  TblAttributionContract H , TblVehicleAllocation_Details D"
 MySQL = MySQL & "  where h.IDAC = d.IDVA and h.durationid = " & val(dcDuration.BoundText) & ""
 
 If dcMinistry.BoundText <> "" Then
        MySQL = MySQL & "  and h.IDMC = " & val(dcMinistry.BoundText) & ""
 End If
 
 MySQL = MySQL & "  and   d.SchoolFileID = TblSchooleFile.id   ) AtrributionAlloc"
 MySQL = MySQL & "  FROM     dbo.TblSchooleFile  ) tbl1"

   
   
  '  If Me.dcBranch.BoundText <> "" Then
  '          MySQL = MySQL & "   and  dbo.TblVendorReceipt.BranchID  =  " & val(Me.dcBranch.BoundText)
  '  End If
  
    If dcSchool.BoundText <> "" Then
              MySQL = MySQL & "   and  dbo.TblSchooleFile.id=  " & val(Me.dcSchool.BoundText)
    End If
  
  '  If Me.dcMonth.BoundText <> "" Then
  '         MySQL = MySQL & "   and  dbo.TblVendorReceipt.Month   =  " & val(Me.dcMonth.BoundText)
  '  End If
  
  
  '
  ' If Me.dcMContract.BoundText <> "" Then
  '         MySQL = MySQL & "   and  dbo.TblAttributionContract.IDAC  =  " & val(Me.dcMContract.BoundText)
  '  End If
  '
    
  '  If Not IsNull(FromDate.value) Then
  '          MySQL = MySQL & "  and TblVendorReceipt.Date >= " & SQLDate(FromDate.value, True)
  '  End If
    
  '  If Not IsNull(ToDate.value) Then
  '          MySQL = MySQL & "  and TblVendorReceipt.Date <= " & SQLDate(ToDate.value, True)
  '  End If
    
    
     If SystemOptions.UserInterface = ArabicInterface Then
              StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_SchoolStatements.rpt"
    Else
              StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_SchoolStatements.rpt"
    End If

'MySQL = " order by TblAttributionContract.idac "
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

   
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
 
End Function


Function print_VA(Optional opt As Integer = 1)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   

MySQL = MySQL & "   SELECT dbo.TblCarsData.BoardNO, dbo.TblVehicleAllocation.DurationID, dbo.TblVehicleAllocation.MinistryNo, dbo.TblVehicleAllocation.StudentAlloc,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.StudentAttrib, dbo.TblVehicleAllocation.ToDateH, dbo.TblVehicleAllocation.FromDateH, dbo.TblVehicleAllocation.ToDate,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.FromDate, dbo.TblVehicleAllocation.StudentCount, dbo.TblVehicleAllocation.IDMC, dbo.TblVehicleAllocation.SchoolFileID,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.ProcessNo, dbo.TblVehicleAllocation.IDVA, dbo.TblSchooleFile.Name AS SchoolName, dbo.TblSchooleFile.StudentCount AS SchoolStudentCount,"
MySQL = MySQL & "   dbo.TblSchooleFile.ministerNo AS SchoolministerNo, dbo.TblDurations.Name AS DurationName, dbo.TblDurations.FromDate AS DurFromDate,"
MySQL = MySQL & "   dbo.TblDurations.FromDateH AS DurFromDateH, dbo.TblDurations.ToDate AS DurToDate, dbo.TblDurations.TODateH AS DurToDateH, dbo.TblEmployee.Emp_Name,"
MySQL = MySQL & "   dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Fullcode, dbo.TblVehicleAllocation_Details.Type, dbo.TblMinistryContract.MinistryContractNo,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.ProcessDate, dbo.TblVehicleAllocation_Details.Capecity, dbo.TblVehicleAllocation_Details.MaxCap, dbo.TblCarsData.code,"
MySQL = MySQL & "   dbo.TblVehicleAllocation_Details.StudentCount AS StudentCount1, dbo.TblVehicleAllocation_Details.VehicleAvailableSite, dbo.TblEmployee.NumEkama,"
MySQL = MySQL & "   dbo.TblEmployee.Emp_Phone , dbo.TblEmployee.Emp_mobile, dbo.TblVehicleAllocation_Details.CarID , dbo.TblVehicleAllocation_Details.VehicleSiteCount"
MySQL = MySQL & "   FROM     dbo.TblMinistryContract INNER JOIN"
MySQL = MySQL & "   dbo.TblDurations INNER JOIN"
MySQL = MySQL & "   dbo.TblVehicleAllocation ON dbo.TblDurations.ID = dbo.TblVehicleAllocation.DurationID INNER JOIN"
MySQL = MySQL & "   dbo.TblSchooleFile ON dbo.TblVehicleAllocation.SchoolFileID = dbo.TblSchooleFile.ID INNER JOIN"
MySQL = MySQL & "   dbo.TblVehicleAllocation_Details ON dbo.TblVehicleAllocation.IDVA = dbo.TblVehicleAllocation_Details.IDVA ON"
MySQL = MySQL & "   dbo.TblMinistryContract.IDMC = dbo.TblVehicleAllocation.IDMC INNER JOIN"
MySQL = MySQL & "   dbo.TblCarsData ON dbo.TblVehicleAllocation_Details.CarID = dbo.TblCarsData.id left outer JOIN"
MySQL = MySQL & "   dbo.TblEmployee ON dbo.TblVehicleAllocation_Details.DriverID = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & "   WHERE   dbo.TblVehicleAllocation_Details.Type = 2 "




        If Me.dcSchool.BoundText <> "" Then
                     MySQL = MySQL & "   and  TblVehicleAllocation.SchoolFileID =  " & val(Me.dcSchool.BoundText)
        End If
 
   

    If Me.dcDuration.BoundText <> "" Then
            MySQL = MySQL & "   and  TblVehicleAllocation.DurationID =  " & val(Me.dcDuration.BoundText)
    End If
 
  
    If Me.dcMinistry.BoundText <> "" Then
           MySQL = MySQL & "   and  TblVehicleAllocation.IDMC  =  " & val(Me.dcMContract.BoundText)
    End If
    
      If Not IsNull(FromDate.value) Then
            MySQL = MySQL & "  and TblVehicleAllocation.FromDate >= " & SQLDate(FromDate.value, True)
    End If
    
    If Not IsNull(ToDate.value) Then
            MySQL = MySQL & "  and TblVehicleAllocation.ToDate <= " & SQLDate(ToDate.value, True)
    End If
     
        MySQL = MySQL & "    order by TblVehicleAllocation.IDVA "


    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VehicleAllocationH.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VehicleAllocationH.rpt"
    End If





'MySQL = " order by TblAttributionContract.idac "
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

   
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
 
End Function

Function print_oper(Optional opt As Integer = 1)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   
   
MySQL = MySQL & " SELECT * FROM ( "
MySQL = MySQL & "  SELECT   Branch_NO as BranchID ,dbo.TblCountriesGovernments.GovernmentName, dbo.TblManagerialArea.Name AS MangerialAreaName, dbo.TblSchooleFile.Name AS SchoolName,"
MySQL = MySQL & "  dbo.TblSchooleFile.ministerNo, dbo.TblSchooleFile.StudentCount, dbo.TblVehicleAllocation.StudentAlloc AS custom, '—∆Ì”Ï' AS trans_type,"
MySQL = MySQL & "  dbo.TblEmployee.Emp_Name AS Driver, dbo.TblEmployee.Emp_Phone AS DriverTel, dbo.TblCarsData.OperatorN, dbo.TblCarsData.BoardNO, dbo.TblCarsData.SetCount,"
MySQL = MySQL & "  '' AS BrandName, dbo.TblCarsData.Model, dbo.TblSchooleFile.phase, dbo.TblVehicleAllocation_Details.Type"
MySQL = MySQL & " , dbo.TblVehicleAllocation.SchoolFileID ,tblschoolefile.sextype "
MySQL = MySQL & "  FROM     dbo.TblSchooleFile INNER JOIN"
MySQL = MySQL & "  dbo.TblVehicleAllocation ON dbo.TblSchooleFile.ID = dbo.TblVehicleAllocation.SchoolFileID INNER JOIN"
MySQL = MySQL & "  dbo.TblVehicleAllocation_Details ON dbo.TblVehicleAllocation.IDVA = dbo.TblVehicleAllocation_Details.IDVA LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblEmployee ON dbo.TblVehicleAllocation_Details.DriverID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblCarsData ON dbo.TblVehicleAllocation_Details.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblCountriesGovernments ON dbo.TblSchooleFile.CityID = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblManagerialArea ON dbo.TblSchooleFile.ID = dbo.TblManagerialArea.ID"
MySQL = MySQL & "  Where (dbo.TblVehicleAllocation_Details.Type = 2)"

MySQL = MySQL & "  Union"

MySQL = MySQL & "  SELECT  BranchID,dbo.TblCountriesGovernments.GovernmentName, dbo.TblManagerialArea.Name AS MangerialAreaName, dbo.TblSchooleFile.Name AS SchoolName,"
MySQL = MySQL & "  dbo.TblSchooleFile.ministerNo, dbo.TblSchooleFile.StudentCount, dbo.TblVehicleAllocation_Details.StudentCount AS custom, '»«ÿ‰' AS trans_type,"
MySQL = MySQL & "  dbo.TblVehicleAllocation_Details.Driver, dbo.TblVehicleAllocation_Details.DriverTel, '' AS OperatorN, dbo.TblVendorCars.BoardNo, dbo.TblVendorCars.Count AS SetCount,"
MySQL = MySQL & "  dbo.TBLCarTypes.name AS BrandName, dbo.TblVendorCars.ModelID + 1900 AS Model, dbo.TblSchooleFile.phase, '' AS type"
MySQL = MySQL & " ,dbo.TblVehicleAllocation_Details.SchoolFileID , tblschoolefile.sextype"
MySQL = MySQL & "  FROM     dbo.TblAttributionContract INNER JOIN"
MySQL = MySQL & "  dbo.TblVehicleAllocation_Details ON dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA INNER JOIN"
MySQL = MySQL & "  dbo.TblVendorCars ON dbo.TblVehicleAllocation_Details.CarID = dbo.TblVendorCars.ID INNER JOIN"
MySQL = MySQL & "  dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID INNER JOIN"
MySQL = MySQL & "  dbo.TblCountriesGovernments ON dbo.TblAttributionContract.CityID = dbo.TblCountriesGovernments.GovernmentID INNER JOIN"
MySQL = MySQL & "  dbo.TblSchooleFile ON dbo.TblVehicleAllocation_Details.SchoolFileID = dbo.TblSchooleFile.ID LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TBLCarTypes ON dbo.TblVendorCars.BrandID = dbo.TBLCarTypes.id"

MySQL = MySQL & " ) AS TB1 "
  
    If dcSchool.BoundText <> "" Then
              MySQL = MySQL & "   WHERE   SchoolFileID  =  " & val(Me.dcSchool.BoundText)
    End If
    
        If Dcbranch.BoundText <> "" Then
              MySQL = MySQL & "   WHERE   BranchID  =  " & val(Me.Dcbranch.BoundText)
    End If
    
     If SystemOptions.UserInterface = ArabicInterface Then
              StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_OperationPlan.rpt"
    Else
              StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_OperationPlan.rpt"
    End If

'MySQL = " order by TblAttributionContract.idac "
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

   
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
 
End Function

Function print_allReq(Optional opt As Integer = 1)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   



MySQL = MySQL & "  SELECT ID, Date, DateH, NoteSerial, SUM(net) AS net, Name, PayMentPayed + SUM(net) - 1 AS tot  FROM     "
MySQL = MySQL & "  (SELECT dbo.TblExchangeRequest.ID, dbo.TblExchangeRequest.Date, dbo.TblExchangeRequest.DateH, dbo.TblExchangeRequest.NoteSerial, dbo.TblExchangeReques_Detailst.net,"
MySQL = MySQL & "  dbo.TblExchangeReques_Detailst.boardno, dbo.TblAttributionInstallmentDivided.PayMentPayed, dbo.TblBranchesData.branch_name name,"
MySQL = MySQL & "  dbo.TblExchangeRequest.BranchID"
MySQL = MySQL & "  FROM     dbo.TblExchangeRequest INNER JOIN"
MySQL = MySQL & "  dbo.TblExchangeReques_Detailst ON dbo.TblExchangeRequest.ID = dbo.TblExchangeReques_Detailst.HID INNER JOIN"
MySQL = MySQL & "  dbo.TblBranchesData ON dbo.TblExchangeRequest.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblAttributionInstallmentDivided ON dbo.TblExchangeReques_Detailst.HID = dbo.TblAttributionInstallmentDivided.REID"
MySQL = MySQL & "  GROUP BY dbo.TblExchangeRequest.ID, dbo.TblExchangeRequest.Date, dbo.TblExchangeRequest.DateH, dbo.TblExchangeRequest.NoteSerial,"
MySQL = MySQL & "  dbo.TblExchangeReques_Detailst.boardno, dbo.TblExchangeReques_Detailst.net, dbo.TblAttributionInstallmentDivided.PayMentPayed, dbo.TblBranchesData.branch_name  ,"
MySQL = MySQL & "  dbo.TblExchangeRequest.BranchID) AS tb1"
MySQL = MySQL & "  GROUP BY ID, Date, DateH, NoteSerial, Name, PayMentPayed"


     If SystemOptions.UserInterface = ArabicInterface Then
              StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VendorExchangeCollection.rpt"
    Else
              StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VendorExchangeCollection.rpt"
    End If

'MySQL = " order by TblAttributionContract.idac "
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

   
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
 
End Function
Private Sub AddSch()

   Dim StrSQL As String
   StrSQL = "delete From TblVacationSchedule2 "
   Cn.Execute StrSQL, , adExecuteNoRecords
    Add_ScheduleH MonthStart, val(dcDuration.BoundText), val(dcMonth.BoundText), val(txtCount.Text)
    Schudle_vacation
    
End Sub


Private Sub Add_ScheduleH(FromDate As String, dur As Integer, DDID As Integer, daycount As Integer)
  
  Dim str As String, str1 As String, day As String, i As Integer, DI As Integer, year As String, Month As String, DM As String, j   As Integer
  FromDate = Format(FromDate, "yyyy/MM/dd")
'  ToDate = Format(ToDate, "yyyy/MM/dd")
   
  Dim d As Integer
  d = cbDay.ListIndex + 1
      
      
   ' j = 1 '
   year = Format(FromDate, "yyyy")
   Month = Format(FromDate, "MM")
   
   Dim ddd As Integer
   ddd = j & Int(Format(FromDate, "dd"))
   Do While j < daycount
        VBA.Calendar = vbCalHijri
        LastDate = FromDate
        
        If j < 10 Then
                
                DM = "0" & ddd
        Else
                DM = ddd
        End If
        
        FromDate = year & "/" & Month & "/" & DM
        VBA.Calendar = vbCalHijri
        day = WeekdayName(d, False, vbSaturday)
        VBA.Calendar = vbCalGreg
        
        If IsHoliday("" & d & "", dur) Then
                 AddRowToSchedule dur, ToGregorianDate(FromDate), FromDate, True, day, DDID
        Else
                 AddRowToSchedule dur, ToGregorianDate(FromDate), FromDate, False, day, DDID
        End If
         
         VBA.Calendar = vbCalHijri
         FromDate = DateAdd("d", 1, FromDate)
        VBA.Calendar = vbCalGreg
             
        FromDate = Format(FromDate, "yyyy/MM/dd")
        
        d = d + 1
        If d = 8 Then
            d = 1
        End If
        j = j + 1
        ddd = ddd + 1
        If ddd > 30 Then
            Exit Sub
        End If
   Loop

End Sub


Private Sub AddRowToSchedule(dur As Integer, dt As Date, dth As String, isvac As Boolean, day As String, DDID As Integer)
        
       Set rs_Dur = New ADODB.Recordset
       rs_Dur.Open " TblVacationschedule2", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       
       rs_Dur.AddNew
       rs_Dur("ID") = CStr(new_id("TblVacationschedule2", "ID", "", True))
       rs_Dur("DurationID") = dur
       rs_Dur("Date") = dt
       rs_Dur("DateH") = Format(dth, "yyyy/MM/dd")
       rs_Dur("isvac") = isvac
       rs_Dur("day") = day
       rs_Dur("DDID") = DDID
       If isvac = True Then
            rs_Dur("color") = "255"
       End If
       rs_Dur.update

End Sub


Private Function IsHoliday(day As String, dur As Integer) As Boolean
    Dim str As String
    str = " select * from  tblholidays  where DurationID =" & dur
    Set rs_hol = New ADODB.Recordset
    rs_hol.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs_hol.RecordCount > 0 Then
            If rs_hol("sa").value = True And day = "1" Then
                IsHoliday = True
            ElseIf rs_hol("su").value = True And day = "2" Then
                IsHoliday = True
            ElseIf rs_hol("Mo").value = True And day = "3" Then
                IsHoliday = True
            ElseIf rs_hol("Tu").value = True And day = "4" Then
                 IsHoliday = True
            ElseIf rs_hol("We").value = True And day = "5" Then
                IsHoliday = True
            ElseIf rs_hol("Th").value = True And day = "6" Then
                IsHoliday = True
            ElseIf rs_hol("Fr").value = True And day = "7" Then
                IsHoliday = True
            End If
  
    End If

End Function

Private Sub Schudle_vacation()
            
       Dim str As String, sql   As String, i As Integer, cnt As Integer
       str = " select * from TblVacationSchedule where DurationID = " & val(dcDuration.BoundText) & " and DDID = " & val(dcMonth.BoundText) & " and ISVac = 1 and   VacationTypeID is not NULL "
       Set rs = New ADODB.Recordset
       rs.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
       
       cnt = rs.RecordCount
       If cnt > 0 Then
            For i = 0 To cnt - 1
                sql = " update TblVacationSchedule2 set isvac = 1 , VacationTypeID = " & IIf(IsNull(rs("VacationTypeID").value), 0, rs("VacationTypeID").value) & _
                 " where   dateh =  '" & IIf(IsNull(rs("dateh").value), "", rs("dateh").value) & "'  and DurationID = " & val(dcDuration.BoundText) & " and DDID =   " & val(dcMonth.BoundText) & ""  ' Date = " & IIf(IsNull(rs("Date").value), Date, rs("Date").value) & " and
                 Cn.Execute sql, , adExecuteNoRecords
                 rs.MoveNext
            Next
       End If
End Sub

Private Function IsValid() As Boolean
 Dim valid As Boolean
 valid = True
 
    If dcDuration.BoundText = "" Then
         MsgBox ("«Œ — «·”‰… «·œ—«”Ì… «Ê·«")
         valid = False
         Exit Function
    End If
 
    If dcMonth.BoundText = "" Then
         MsgBox ("«Œ — «·› —… «Ê·«")
         valid = False
         Exit Function
    End If
    
    If cbDay.ListIndex = -1 Then
         MsgBox ("«Œ — ÌÊ„ «Ê· «·‘Â— ")
         valid = False
         Exit Function
    End If
    
   If val(txtCount.Text) > 30 Then
        MsgBox (" ·«Ì„ﬂ‰ «‰ Ì Ã«Ê“ ⁄œœ «·«Ì«„ 30 ÌÊ„ ")
        valid = False
         Exit Function
   End If
    
    IsValid = valid

End Function

Private Function MonthStart() As String
            
       Dim str As String, sql   As String, i As Integer, cnt As Integer
       str = " select * from TblDurations_Details  where DID =  " & val(dcDuration.BoundText) & " and id  = " & val(dcMonth.BoundText) & "  "
       Set rs = New ADODB.Recordset
       rs.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
       cnt = rs.RecordCount
       If cnt > 0 Then
            MonthStart = rs("FromdateH").value
       End If
End Function




