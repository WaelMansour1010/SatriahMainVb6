VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmArchReports 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   9525
   ClientLeft      =   3720
   ClientTop       =   2295
   ClientWidth     =   15915
   Icon            =   "FrmArchReports.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   15915
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame111 
      Caption         =   "Frame1"
      Height          =   8895
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   15975
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   9135
         Left            =   0
         TabIndex        =   2
         Top             =   -360
         Width           =   15960
         _cx             =   28152
         _cy             =   16113
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
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   " ﬁ«—Ì— "
         Align           =   0
         CurrTab         =   2
         FirstTab        =   0
         Style           =   3
         Position        =   1
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   0   'False
         DogEars         =   0   'False
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   1
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Picture(0)      =   "FrmArchReports.frx":038A
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Height          =   8700
            Index           =   1
            Left            =   30
            TabIndex        =   3
            Top             =   30
            Width           =   15885
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Height          =   1695
               Left            =   7800
               TabIndex        =   82
               Top             =   720
               Width           =   7965
               Begin VB.ListBox ListDeptSelect 
                  Height          =   1230
                  ItemData        =   "FrmArchReports.frx":0724
                  Left            =   120
                  List            =   "FrmArchReports.frx":072B
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   330
                  Width           =   3705
               End
               Begin VB.ListBox ListDeptAll 
                  Height          =   1230
                  ItemData        =   "FrmArchReports.frx":073F
                  Left            =   4485
                  List            =   "FrmArchReports.frx":0746
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   330
                  Width           =   3345
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·ﬁ”„"
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Left            =   3360
                  TabIndex        =   89
                  Top             =   120
                  Width           =   1470
               End
               Begin VB.Label Label10 
                  Alignment       =   2  'Center
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   300
                  Width           =   495
               End
               Begin VB.Label Label9 
                  Alignment       =   2  'Center
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   240
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   630
                  Width           =   495
               End
               Begin VB.Label Label8 
                  Alignment       =   2  'Center
                  Caption         =   "<<"
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
                  Height          =   360
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   1200
                  Width           =   495
               End
               Begin VB.Label Label7 
                  Alignment       =   2  'Center
                  Caption         =   "<"
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
                  Height          =   345
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   855
                  Width           =   495
               End
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00E2E9E9&
               Height          =   1695
               Left            =   -120
               TabIndex        =   74
               Top             =   720
               Width           =   7965
               Begin VB.ListBox ListAllUsers 
                  Height          =   1230
                  ItemData        =   "FrmArchReports.frx":0758
                  Left            =   4605
                  List            =   "FrmArchReports.frx":075F
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   330
                  Width           =   3225
               End
               Begin VB.ListBox ListUserSelect 
                  Height          =   1230
                  ItemData        =   "FrmArchReports.frx":0771
                  Left            =   120
                  List            =   "FrmArchReports.frx":0778
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   330
                  Width           =   3705
               End
               Begin VB.Label Label19 
                  Alignment       =   2  'Center
                  Caption         =   "<"
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
                  Height          =   345
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   855
                  Width           =   495
               End
               Begin VB.Label Label20 
                  Alignment       =   2  'Center
                  Caption         =   "<<"
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
                  Height          =   360
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   1185
                  Width           =   495
               End
               Begin VB.Label Label21 
                  Alignment       =   2  'Center
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   240
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   600
                  Width           =   495
               End
               Begin VB.Label Label22 
                  Alignment       =   2  'Center
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   300
                  Width           =   495
               End
               Begin VB.Label Label23 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·„ÊŸ›Ì‰ «·„—”· «·ÌÂ„"
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Left            =   3330
                  TabIndex        =   77
                  Top             =   120
                  Width           =   1950
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Height          =   1695
               Left            =   7800
               TabIndex        =   66
               Top             =   2280
               Width           =   7965
               Begin VB.ListBox ListAllArche 
                  Height          =   1230
                  ItemData        =   "FrmArchReports.frx":078C
                  Left            =   4485
                  List            =   "FrmArchReports.frx":0793
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   330
                  Width           =   3345
               End
               Begin VB.ListBox ListSelectArche 
                  Height          =   1230
                  ItemData        =   "FrmArchReports.frx":07A5
                  Left            =   120
                  List            =   "FrmArchReports.frx":07AC
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   330
                  Width           =   3705
               End
               Begin VB.Label Label16 
                  Alignment       =   2  'Center
                  Caption         =   "<"
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
                  Height          =   345
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   960
                  Width           =   495
               End
               Begin VB.Label Label15 
                  Alignment       =   2  'Center
                  Caption         =   "<<"
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
                  Height          =   360
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   1185
                  Width           =   495
               End
               Begin VB.Label Label14 
                  Alignment       =   2  'Center
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   360
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   600
                  Width           =   495
               End
               Begin VB.Label Label13 
                  Alignment       =   2  'Center
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   300
                  Width           =   495
               End
               Begin VB.Label Label12 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·«—‘Ì›"
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Left            =   3330
                  TabIndex        =   69
                  Top             =   120
                  Width           =   1470
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Height          =   1695
               Left            =   -120
               TabIndex        =   58
               Top             =   2280
               Width           =   7965
               Begin VB.ListBox ListSelectRoom 
                  Height          =   1230
                  ItemData        =   "FrmArchReports.frx":07C0
                  Left            =   120
                  List            =   "FrmArchReports.frx":07C7
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   330
                  Width           =   3705
               End
               Begin VB.ListBox ListAllRoom 
                  Height          =   1230
                  ItemData        =   "FrmArchReports.frx":07DB
                  Left            =   4605
                  List            =   "FrmArchReports.frx":07E2
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   330
                  Width           =   3225
               End
               Begin VB.Label Label6 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·€—›…"
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Left            =   3330
                  TabIndex        =   65
                  Top             =   120
                  Width           =   1470
               End
               Begin VB.Label Label4 
                  Alignment       =   2  'Center
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   300
                  Width           =   495
               End
               Begin VB.Label Label3 
                  Alignment       =   2  'Center
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   360
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   630
                  Width           =   495
               End
               Begin VB.Label Label2 
                  Alignment       =   2  'Center
                  Caption         =   "<<"
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
                  Height          =   360
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   1305
                  Width           =   495
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  Caption         =   "<"
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
                  Height          =   345
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   975
                  Width           =   495
               End
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00E2E9E9&
               Height          =   1695
               Left            =   7800
               TabIndex        =   50
               Top             =   3840
               Width           =   7965
               Begin VB.ListBox ListAllBox 
                  Height          =   1230
                  ItemData        =   "FrmArchReports.frx":07F4
                  Left            =   4485
                  List            =   "FrmArchReports.frx":07FB
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   330
                  Width           =   3345
               End
               Begin VB.ListBox ListSelectBox 
                  Height          =   1230
                  ItemData        =   "FrmArchReports.frx":080D
                  Left            =   120
                  List            =   "FrmArchReports.frx":0814
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   360
                  Width           =   3735
               End
               Begin VB.Label Label31 
                  Alignment       =   2  'Center
                  Caption         =   "<"
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
                  Height          =   345
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   855
                  Width           =   495
               End
               Begin VB.Label Label30 
                  Alignment       =   2  'Center
                  Caption         =   "<<"
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
                  Height          =   360
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   1185
                  Width           =   495
               End
               Begin VB.Label Label29 
                  Alignment       =   2  'Center
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   240
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   630
                  Width           =   495
               End
               Begin VB.Label Label28 
                  Alignment       =   2  'Center
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Left            =   3915
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   300
                  Width           =   495
               End
               Begin VB.Label Label27 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·’‰œÊﬁ"
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Left            =   3330
                  TabIndex        =   53
                  Top             =   120
                  Width           =   1470
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E2E9E9&
               Height          =   1695
               Left            =   0
               TabIndex        =   42
               Top             =   3840
               Width           =   7845
               Begin VB.ListBox ListSelectRaf 
                  Height          =   1230
                  ItemData        =   "FrmArchReports.frx":0828
                  Left            =   120
                  List            =   "FrmArchReports.frx":082F
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   330
                  Width           =   3585
               End
               Begin VB.ListBox ListAllRaf 
                  Height          =   1230
                  ItemData        =   "FrmArchReports.frx":0843
                  Left            =   4485
                  List            =   "FrmArchReports.frx":084A
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   330
                  Width           =   3225
               End
               Begin VB.Label Label26 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·—›"
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Left            =   3330
                  TabIndex        =   49
                  Top             =   120
                  Width           =   1470
               End
               Begin VB.Label Label25 
                  Alignment       =   2  'Center
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Left            =   3795
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   300
                  Width           =   495
               End
               Begin VB.Label Label24 
                  Alignment       =   2  'Center
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   360
                  Left            =   3795
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   630
                  Width           =   495
               End
               Begin VB.Label Label18 
                  Alignment       =   2  'Center
                  Caption         =   "<<"
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
                  Height          =   360
                  Left            =   3795
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   1305
                  Width           =   495
               End
               Begin VB.Label Label17 
                  Alignment       =   2  'Center
                  Caption         =   "<"
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
                  Height          =   345
                  Left            =   3795
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   960
                  Width           =   495
               End
            End
            Begin VB.Frame Frame10 
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·Õ—ﬂ…"
               Height          =   1575
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   7080
               Width           =   4455
               Begin MSComCtl2.DTPicker FrmRecordDate 
                  Height          =   330
                  Left            =   1920
                  TabIndex        =   36
                  Top             =   510
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   66125827
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker ToRecordDate 
                  Height          =   330
                  Left            =   1920
                  TabIndex        =   37
                  Top             =   960
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   66125827
                  CurrentDate     =   41640
               End
               Begin Dynamic_Byte.NourHijriCal FrmRecordDateH 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   38
                  Top             =   480
                  Width           =   1635
                  _ExtentX        =   2884
                  _ExtentY        =   556
               End
               Begin Dynamic_Byte.NourHijriCal ToRecordDateH 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   39
                  Top             =   990
                  Width           =   1635
                  _ExtentX        =   2884
                  _ExtentY        =   556
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   195
                  Index           =   13
                  Left            =   3630
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   1080
                  Width           =   480
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   15
                  Left            =   3690
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   600
                  Width           =   420
               End
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00E2E9E9&
               Height          =   1695
               Left            =   0
               TabIndex        =   27
               Top             =   5400
               Width           =   7845
               Begin VB.ListBox ListSelectProcess 
                  Height          =   1230
                  ItemData        =   "FrmArchReports.frx":085C
                  Left            =   120
                  List            =   "FrmArchReports.frx":0863
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   330
                  Width           =   3585
               End
               Begin VB.ListBox ListAllProcess 
                  Height          =   1230
                  ItemData        =   "FrmArchReports.frx":0877
                  Left            =   4485
                  List            =   "FrmArchReports.frx":087E
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   330
                  Width           =   3225
               End
               Begin VB.Label Label41 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·⁄„·Ì…"
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Left            =   3330
                  TabIndex        =   34
                  Top             =   120
                  Width           =   1470
               End
               Begin VB.Label Label40 
                  Alignment       =   2  'Center
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Left            =   3795
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   300
                  Width           =   495
               End
               Begin VB.Label Label39 
                  Alignment       =   2  'Center
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   240
                  Left            =   3795
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   630
                  Width           =   495
               End
               Begin VB.Label Label38 
                  Alignment       =   2  'Center
                  Caption         =   "<<"
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
                  Height          =   360
                  Left            =   3795
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   1185
                  Width           =   495
               End
               Begin VB.Label Label37 
                  Alignment       =   2  'Center
                  Caption         =   "<"
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
                  Height          =   345
                  Left            =   3795
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   855
                  Width           =   495
               End
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00E2E9E9&
               Height          =   3255
               Left            =   7800
               TabIndex        =   5
               Top             =   5400
               Width           =   7965
               Begin VB.TextBox Summary 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   1200
                  Width           =   6705
               End
               Begin VB.ComboBox DcbImportExport 
                  Height          =   315
                  ItemData        =   "FrmArchReports.frx":0890
                  Left            =   3120
                  List            =   "FrmArchReports.frx":0892
                  RightToLeft     =   -1  'True
                  TabIndex        =   23
                  Top             =   840
                  Width           =   3705
               End
               Begin VB.Frame Frame12 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "› —… «·Œ—ÊÃ "
                  Height          =   735
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   2400
                  Width           =   7815
                  Begin MSComCtl2.DTPicker FrmExitDate 
                     Height          =   330
                     Left            =   3600
                     TabIndex        =   19
                     Top             =   270
                     Width           =   2295
                     _ExtentX        =   4048
                     _ExtentY        =   582
                     _Version        =   393216
                     CheckBox        =   -1  'True
                     CustomFormat    =   "dd/MM/yyyy,hh:mm:tt"
                     Format          =   66125827
                     CurrentDate     =   41640
                  End
                  Begin MSComCtl2.DTPicker ToExitDate 
                     Height          =   330
                     Left            =   120
                     TabIndex        =   20
                     Top             =   270
                     Width           =   2295
                     _ExtentX        =   4048
                     _ExtentY        =   582
                     _Version        =   393216
                     CheckBox        =   -1  'True
                     CustomFormat    =   "dd/MM/yyyy,hh:mm:tt"
                     Format          =   66125827
                     CurrentDate     =   41640
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "≈·Ï"
                     Height          =   195
                     Index           =   6
                     Left            =   2550
                     RightToLeft     =   -1  'True
                     TabIndex        =   22
                     Top             =   240
                     Width           =   480
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„‰"
                     Height          =   195
                     Index           =   5
                     Left            =   5970
                     RightToLeft     =   -1  'True
                     TabIndex        =   21
                     Top             =   240
                     Width           =   540
                  End
               End
               Begin VB.Frame Frame11 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "› —… «·œŒÊ·"
                  Height          =   735
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   13
                  Top             =   1680
                  Width           =   7815
                  Begin MSComCtl2.DTPicker FrmEnterDate 
                     Height          =   330
                     Left            =   3600
                     TabIndex        =   14
                     Top             =   270
                     Width           =   2295
                     _ExtentX        =   4048
                     _ExtentY        =   582
                     _Version        =   393216
                     CheckBox        =   -1  'True
                     CustomFormat    =   "dd/MM/yyyy,hh:mm:tt"
                     Format          =   66125827
                     CurrentDate     =   41640
                  End
                  Begin MSComCtl2.DTPicker ToEnterDate 
                     Height          =   330
                     Left            =   120
                     TabIndex        =   15
                     Top             =   270
                     Width           =   2295
                     _ExtentX        =   4048
                     _ExtentY        =   582
                     _Version        =   393216
                     CheckBox        =   -1  'True
                     CustomFormat    =   "dd/MM/yyyy,hh:mm:tt"
                     Format          =   66125827
                     CurrentDate     =   41640
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„‰"
                     Height          =   195
                     Index           =   3
                     Left            =   5970
                     RightToLeft     =   -1  'True
                     TabIndex        =   17
                     Top             =   240
                     Width           =   540
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "≈·Ï"
                     Height          =   195
                     Index           =   2
                     Left            =   2550
                     RightToLeft     =   -1  'True
                     TabIndex        =   16
                     Top             =   240
                     Width           =   480
                  End
               End
               Begin VB.TextBox Txtbarcode 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   12
                  Top             =   360
                  Width           =   1785
               End
               Begin VB.TextBox TxtNoImpExp 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   840
                  Width           =   1785
               End
               Begin VB.Frame Frame9 
                  BackColor       =   &H00E2E9E9&
                  Height          =   615
                  Left            =   3120
                  TabIndex        =   6
                  Top             =   120
                  Width           =   4695
                  Begin VB.TextBox TxtToNo 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   8
                     Top             =   240
                     Width           =   1365
                  End
                  Begin VB.TextBox TxtFromNo 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1920
                     RightToLeft     =   -1  'True
                     TabIndex        =   7
                     Top             =   240
                     Width           =   1365
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Ï"
                     Height          =   285
                     Index           =   0
                     Left            =   1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   10
                     Top             =   240
                     Width           =   450
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «·„⁄«„·… „‰"
                     Height          =   285
                     Index           =   4
                     Left            =   3270
                     RightToLeft     =   -1  'True
                     TabIndex        =   9
                     Top             =   240
                     Width           =   1530
                  End
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·Œ÷ «·”‰œ"
                  Height          =   300
                  Index           =   7
                  Left            =   6960
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   1320
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·»«—ﬂÊœ"
                  Height          =   285
                  Index           =   9
                  Left            =   2085
                  TabIndex        =   26
                  Top             =   360
                  Width           =   870
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—ﬁ„"
                  Height          =   300
                  Index           =   1
                  Left            =   1935
                  TabIndex        =   25
                  Top             =   840
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·‰Ê⁄"
                  Height          =   300
                  Index           =   11
                  Left            =   7035
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   840
                  Width           =   675
               End
            End
            Begin VB.CommandButton btnClear 
               BackColor       =   &H00E2E9E9&
               Caption         =   "„”Õ"
               Height          =   375
               Left            =   1680
               TabIndex        =   4
               Top             =   8040
               Width           =   1335
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   495
               Index           =   1
               Left            =   960
               TabIndex        =   90
               Top             =   7320
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   873
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄… "
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
               BackStyle       =   0
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorHoverText  =   16711680
               ColorShadow     =   -2147483637
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               ColorToggledHoverText=   16711680
               LowerToggledContent=   0   'False
               ColorTextShadow =   -2147483637
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   375
               Index           =   2
               Left            =   240
               TabIndex        =   91
               Top             =   8040
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
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
               BackStyle       =   0
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorHoverText  =   16711680
               ColorShadow     =   -2147483637
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               LowerToggledContent=   0   'False
               ColorTextShadow =   -2147483637
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   375
               Index           =   0
               Left            =   13440
               TabIndex        =   92
               Top             =   360
               Width           =   2055
               _Version        =   786432
               _ExtentX        =   3625
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— «·√—‘Ì› "
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   375
               Index           =   1
               Left            =   10800
               TabIndex        =   93
               Top             =   360
               Width           =   2295
               _Version        =   786432
               _ExtentX        =   4048
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   " ﬁ«—Ì— Õ«·… «·”‰œ"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15930
      _cx             =   28099
      _cy             =   1085
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
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   690
         Left            =   0
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   0
         Width           =   16005
         _cx             =   28231
         _cy             =   1217
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
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
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
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   " ﬁ«—Ì—  «·«—‘Ì›           "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   120
            TabIndex        =   97
            Top             =   0
            Width           =   16230
         End
      End
   End
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   360
      Picture         =   "FrmArchReports.frx":0894
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
End
Attribute VB_Name = "FrmArchReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Private Sub ChangeLang()
Frame111.RightToLeft = False
Frame10.RightToLeft = False
Frame2.RightToLeft = False
'Fra(1).RightToLeft = False
XPTab301.Caption = "Reports"
Label5.Caption = "Archiving Reports"
Label11.Caption = "Choose department"
Label23.Caption = "Choose employee"
Label12.Caption = "Choose archive"
Label6.Caption = "Choose room"
Label27.Caption = "Choose box"
Label26.Caption = "Choose Shilf"
Label41.Caption = "Choose procedure type"
lbl(4).Caption = "Procedure No. from"
lbl(0).Caption = "To"
lbl(9).Caption = "Barcode"
lbl(11).Caption = "Type"
lbl(1).Caption = "Document No."
Frame11.Caption = "Entering period"
lbl(3).Caption = "From"
lbl(2).Caption = "To"
Frame12.Caption = "Exiting period"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
Frame10.Caption = "Procedure Date"
lbl(15).Caption = "From"
lbl(13).Caption = "To"
Cmd(1).Caption = "Print"
btnClear.Caption = "Clear"
Cmd(2).Caption = "Exit"
lbl(7).Caption = "Summary"
Rd(0).Caption = "Archiving report"
Rd(1).Caption = "Document status report"
Rd(0).value = True
End Sub



Private Sub FrmRecordDate_Change()
If Not IsNull(FrmRecordDate.value) Then
FrmRecordDateH.value = ToHijriDate(FrmRecordDate.value)
End If
End Sub

Private Sub FrmRecordDateH_LostFocus()
FrmRecordDate.value = ToGregorianDate(FrmRecordDateH.value)
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub
Private Sub btnClear_Click()

clear_all Me
Clearss
End Sub

Private Sub Cmd_Click(Index As Integer)
If Index = 1 Then
print_report
Else
Unload Me
End If
End Sub
Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub

Sub Clearss()
   ListDeptSelect.Clear
   ListUserSelect.Clear
   ListSelectArche.Clear
   ListSelectRoom.Clear
   ListSelectBox.Clear
   ListSelectRaf.Clear
   ListSelectProcess.Clear
   ListAllUsers.Clear
   ListAllArche.Clear
   ListAllRoom.Clear
   ListAllBox.Clear
   ListAllRaf.Clear
   ListAllProcess.Clear
   FrmEnterDate.value = Now
   ToEnterDate.value = Now
   FrmExitDate.value = Now
   ToExitDate.value = Now
   FrmRecordDate.value = Date
   ToRecordDate.value = Date
  FrmEnterDate.value = ""
   ToEnterDate.value = ""
   FrmExitDate.value = ""
   ToExitDate.value = ""
   FrmRecordDate.value = ""
   ToRecordDate.value = ""
End Sub
Private Sub Form_Load()
Clearss
   Rd(0).value = True
If SystemOptions.UserInterface = ArabicInterface Then
    With DcbImportExport
        .Clear
        .AddItem "’«œ—"
        .AddItem "Ê«—œ"
    End With
Else
    With DcbImportExport
        .Clear
        .AddItem "Import"
        .AddItem "Export"
    End With
End If

If SystemOptions.UserInterface = EnglishInterface Then
SetInterface Me
    ChangeLang
End If

FillMylist

Resize_Form Me
End Sub
Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Private Sub Label1_Click()
If Me.ListSelectRoom.ListIndex > -1 Then
Me.ListSelectRoom.RemoveItem (ListSelectRoom.ListIndex)
End If
FillMylist2
End Sub

Private Sub Label10_Click()
 If Me.ListDeptAll.ListIndex > -1 Then
    Me.ListDeptSelect.AddItem ListDeptAll.List(ListDeptAll.ListIndex)
    ListDeptSelect.ItemData(ListDeptSelect.NewIndex) = ListDeptAll.ItemData(ListDeptAll.ListIndex)
End If
FillMylist2
End Sub

Private Sub Label13_Click()
 If Me.ListAllArche.ListIndex > -1 Then
    Me.ListSelectArche.AddItem ListAllArche.List(ListAllArche.ListIndex)
    ListSelectArche.ItemData(ListSelectArche.NewIndex) = ListAllArche.ItemData(ListAllArche.ListIndex)
End If
ClearListArc
FillMylist2
End Sub

Private Sub Label14_Click()
    Dim i As Integer
    Me.ListSelectArche.Clear
    For i = 0 To Me.ListAllArche.ListCount - 1
        Me.ListSelectArche.AddItem ListAllArche.List(i)
        ListSelectArche.ItemData(i) = ListAllArche.ItemData(i)
    Next i
    ClearListArc
   FillMylist2
End Sub

Private Sub Label15_Click()
ListSelectArche.Clear
End Sub

Private Sub Label16_Click()
If Me.ListSelectArche.ListIndex > -1 Then
Me.ListSelectArche.RemoveItem (ListSelectArche.ListIndex)
End If
ClearListArc
FillMylist2
End Sub

Private Sub Label17_Click()
If Me.ListSelectRaf.ListIndex > -1 Then
Me.ListSelectRaf.RemoveItem (ListSelectRaf.ListIndex)
End If
ClearListRaf
FillMylist2
End Sub

Private Sub Label18_Click()
ListSelectRaf.Clear
End Sub

Private Sub Label19_Click()
If Me.ListUserSelect.ListIndex > -1 Then
ListUserSelect.RemoveItem (ListUserSelect.ListIndex)
End If
End Sub

Private Sub Label2_Click()
ListSelectRoom.Clear
End Sub

Private Sub Label20_Click()
Me.ListUserSelect.Clear
End Sub

Private Sub Label21_Click()
    Dim i As Integer
    Me.ListUserSelect.Clear
    For i = 0 To Me.ListAllUsers.ListCount - 1
        Me.ListUserSelect.AddItem ListAllUsers.List(i)
        ListUserSelect.ItemData(i) = ListAllUsers.ItemData(i)
    Next i
End Sub

Private Sub Label22_Click()
 If Me.ListAllUsers.ListIndex > -1 Then
    Me.ListUserSelect.AddItem ListAllUsers.List(ListAllUsers.ListIndex)
    ListUserSelect.ItemData(ListUserSelect.NewIndex) = ListAllUsers.ItemData(ListAllUsers.ListIndex)
End If
End Sub



Private Sub Label24_Click()
    Dim i As Integer
    Me.ListSelectRaf.Clear
    For i = 0 To Me.ListAllRaf.ListCount - 1
        Me.ListSelectRaf.AddItem ListAllRaf.List(i)
        ListSelectRaf.ItemData(i) = ListAllRaf.ItemData(i)
    Next i
    ClearListRaf
   FillMylist2
End Sub

Private Sub Label25_Click()
 If Me.ListAllRaf.ListIndex > -1 Then
    Me.ListSelectRaf.AddItem ListAllRaf.List(ListAllRaf.ListIndex)
    ListSelectRaf.ItemData(ListSelectRaf.NewIndex) = ListAllRaf.ItemData(ListAllRaf.ListIndex)
End If
ClearListRaf
FillMylist2
End Sub

Private Sub Label28_Click()
 If Me.ListAllBox.ListIndex > -1 Then
    Me.ListSelectBox.AddItem ListAllBox.List(ListAllBox.ListIndex)
    ListSelectBox.ItemData(ListSelectBox.NewIndex) = ListAllBox.ItemData(ListAllBox.ListIndex)
End If
ClearListBox
FillMylist2
End Sub

Private Sub Label29_Click()
    Dim i As Integer
    Me.ListSelectBox.Clear
    For i = 0 To Me.ListAllBox.ListCount - 1
        Me.ListSelectBox.AddItem ListAllBox.List(i)
        ListSelectBox.ItemData(i) = ListAllBox.ItemData(i)
    Next i
    ClearListBox
   FillMylist2
End Sub

Private Sub Label3_Click()
    Dim i As Integer
    Me.ListSelectRoom.Clear
    For i = 0 To Me.ListAllRoom.ListCount - 1
        Me.ListSelectRoom.AddItem ListAllRoom.List(i)
        ListSelectRoom.ItemData(i) = ListAllRoom.ItemData(i)
    Next i
    ClearListRoom
   FillMylist2
End Sub

Private Sub Label30_Click()
ListSelectBox.Clear
End Sub

Private Sub Label31_Click()
If Me.ListSelectBox.ListIndex > -1 Then
Me.ListSelectBox.RemoveItem (ListSelectBox.ListIndex)
End If
ClearListBox
FillMylist2
End Sub

Private Sub Label37_Click()
If Me.ListSelectProcess.ListIndex > -1 Then
ListSelectProcess.RemoveItem (ListSelectProcess.ListIndex)
End If
End Sub

Private Sub Label38_Click()
ListSelectProcess.Clear
End Sub

Private Sub Label39_Click()
    Dim i As Integer
    Me.ListSelectProcess.Clear
    For i = 0 To Me.ListAllProcess.ListCount - 1
        Me.ListSelectProcess.AddItem ListAllProcess.List(i)
        ListSelectProcess.ItemData(i) = ListAllProcess.ItemData(i)
    Next i
   FillMylist2
End Sub

Private Sub Label4_Click()
 If Me.ListAllRoom.ListIndex > -1 Then
    Me.ListSelectRoom.AddItem ListAllRoom.List(ListAllRoom.ListIndex)
    ListSelectRoom.ItemData(ListSelectRoom.NewIndex) = ListAllRoom.ItemData(ListAllRoom.ListIndex)
End If
ClearListRoom
FillMylist2
End Sub

Private Sub Label40_Click()
 If Me.ListAllProcess.ListIndex > -1 Then
    Me.ListSelectProcess.AddItem ListAllProcess.List(ListAllProcess.ListIndex)
    ListSelectProcess.ItemData(ListSelectProcess.NewIndex) = ListAllProcess.ItemData(ListAllProcess.ListIndex)
End If
FillMylist2
End Sub

Private Sub Label7_Click()
If Me.ListDeptSelect.ListIndex > -1 Then
Me.ListDeptSelect.RemoveItem (ListDeptSelect.ListIndex)
End If
FillMylist2
End Sub

Private Sub Label8_Click()
Me.ListDeptSelect.Clear
'ClearList
FillMylist2
End Sub

Private Sub Label9_Click()
    Dim i As Integer
    Me.ListDeptSelect.Clear
    For i = 0 To Me.ListDeptAll.ListCount - 1
        Me.ListDeptSelect.AddItem ListDeptAll.List(i)
        ListDeptSelect.ItemData(i) = ListDeptAll.ItemData(i)
    Next i
   FillMylist2
End Sub
Function FillMylist()
    Dim Sql As String
    Dim Rs2 As ADODB.Recordset
    Dim i As Integer
    Set Rs2 = New ADODB.Recordset
    Sql = " SELECT * from  TblEmpDepartments "
    Rs2.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Me.ListDeptAll.Clear
    Me.ListDeptSelect.Clear
    
    If Rs2.RecordCount > 0 Then
        For i = 1 To Rs2.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                ListDeptAll.AddItem IIf(IsNull(Rs2("DepartmentName").value), "", Rs2("DepartmentName").value)
            Else
                ListDeptAll.AddItem IIf(IsNull(Rs2("DepartmentNamee").value), "", Rs2("DepartmentNamee").value)
            End If
            ListDeptAll.ItemData(ListDeptAll.NewIndex) = IIf(IsNull(Rs2("DeparmentID").value), 0, Rs2("DeparmentID").value)
            Rs2.MoveNext
        Next i
    End If
    Rs2.Close
End Function
Sub ClearListRaf()
    ListAllProcess.Clear
    ListSelectProcess.Clear
End Sub
Sub ClearListBox()
    ListAllRaf.Clear
    ListSelectRaf.Clear
    ListAllProcess.Clear
    ListSelectProcess.Clear
End Sub
Sub ClearListRoom()
    ListAllBox.Clear
    ListSelectBox.Clear
    ListAllRaf.Clear
    ListSelectRaf.Clear
    ListAllProcess.Clear
    ListSelectProcess.Clear
End Sub
Sub ClearListArc()
    ListAllRoom.Clear
    ListSelectRoom.Clear
    ListAllBox.Clear
    ListSelectBox.Clear
    ListAllRaf.Clear
    ListSelectRaf.Clear
    ListAllProcess.Clear
    ListSelectProcess.Clear
End Sub
Sub ClearListAll()
    Me.ListAllUsers.Clear
    Me.ListUserSelect.Clear
    ListAllArche.Clear
    ListSelectArche.Clear
    ListAllRoom.Clear
    ListSelectRoom.Clear
    ListAllBox.Clear
    ListSelectBox.Clear
    ListAllRaf.Clear
    ListSelectRaf.Clear
    ListAllProcess.Clear
   ListSelectProcess.Clear
End Sub
Sub RetriveCondition(Optional ByRef ActivID As String, Optional ByRef Arche As String, Optional ByRef Room As String, Optional ByRef Box As String, Optional ByRef Shelf As String)
Dim i As Integer
    ActivID = "0"
    For i = 0 To Me.ListDeptSelect.ListCount - 1
    ActivID = ActivID & "," & Me.ListDeptSelect.ItemData(i)
    Next i
 ''/////////
     Arche = "0"
    For i = 0 To Me.ListSelectArche.ListCount - 1
    Arche = Arche & "," & Me.ListSelectArche.ItemData(i)
    Next i
      ''/////////////
    Room = "0"
    For i = 0 To Me.ListSelectRoom.ListCount - 1
    Room = Room & "," & Me.ListSelectRoom.ItemData(i)
    Next i
      ''/////////////
    Box = "0"
    For i = 0 To Me.ListSelectBox.ListCount - 1
    Box = Box & "," & Me.ListSelectBox.ItemData(i)
    Next i
    ''/////////////
    Shelf = "0"
    For i = 0 To Me.ListSelectRaf.ListCount - 1
    Shelf = Shelf & "," & Me.ListSelectRaf.ItemData(i)
    Next i
    
End Sub
Function FillMylist2()
    Dim Sql As String
    Dim Rs2 As ADODB.Recordset
    Dim i As Integer
    Dim Arche As String
    Dim ActivID As String
    Dim Room As String
    Dim Box As String
    Dim Shelf As String
    RetriveCondition ActivID, Arche, Room, Box, Shelf
   '' ///
    If ActivID = "0" Then Exit Function
    Set Rs2 = New ADODB.Recordset
    Sql = " SELECT     dbo.TblUsers.UserName, dbo.TblUsers.UserID"
    Sql = Sql & "  FROM         dbo.TblUsers LEFT OUTER JOIN"
    Sql = Sql & "                   dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
    Sql = Sql & "   WHERE     dbo.TblEmployee.DepartmentID  in(" & ActivID & ") "
    Rs2.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    ListAllUsers.Clear
    If Rs2.RecordCount > 0 Then
        For i = 1 To Rs2.RecordCount
            Me.ListAllUsers.AddItem IIf(IsNull(Rs2("UserName").value), "", Rs2("UserName").value)
            ListAllUsers.ItemData(ListAllUsers.NewIndex) = IIf(IsNull(Rs2("UserID").value), 0, Rs2("UserID").value)
            Rs2.MoveNext
        Next i

    End If

  '''////////////////
    Set Rs2 = New ADODB.Recordset
    Sql = " SELECT    * from TblXXArch"
    Sql = Sql & "   WHERE   1=1"
    Sql = Sql & " and DepID  in(" & ActivID & ") "
    Rs2.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    ListAllArche.Clear
    If Rs2.RecordCount > 0 Then
        For i = 1 To Rs2.RecordCount
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.ListAllArche.AddItem IIf(IsNull(Rs2("Name").value), "", Rs2("Name").value)
            Else
            Me.ListAllArche.AddItem IIf(IsNull(Rs2("Namee").value), "", Rs2("Namee").value)
            End If
            ListAllArche.ItemData(ListAllArche.NewIndex) = IIf(IsNull(Rs2("ID").value), 0, Rs2("ID").value)
            Rs2.MoveNext
        Next i

    End If
    Rs2.Close
  ''''
      Set Rs2 = New ADODB.Recordset
    Sql = " SELECT    * from TblXXRoom"
    Sql = Sql & "   WHERE   1=1"
    Sql = Sql & " and DepID  in(" & ActivID & ") "
    If Arche <> "0" Then
    Sql = Sql & " and ArchID  in(" & Arche & ") "
    End If
    Rs2.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    ListAllRoom.Clear
    If Rs2.RecordCount > 0 Then
        For i = 1 To Rs2.RecordCount
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.ListAllRoom.AddItem IIf(IsNull(Rs2("Name").value), "", Rs2("Name").value)
            Else
            Me.ListAllRoom.AddItem IIf(IsNull(Rs2("Namee").value), "", Rs2("Namee").value)
            End If
            ListAllRoom.ItemData(ListAllRoom.NewIndex) = IIf(IsNull(Rs2("ID").value), 0, Rs2("ID").value)
            Rs2.MoveNext
        Next i

    End If
    Rs2.Close
      ''///////////
    Set Rs2 = New ADODB.Recordset
    Sql = " SELECT    * from TblXXBox"
    Sql = Sql & "   WHERE   1=1"
    Sql = Sql & " and DepID  in(" & ActivID & ") "
    If Arche <> "0" Then
    Sql = Sql & " and ArchID  in(" & Arche & ") "
    End If
    If Room <> "0" Then
    Sql = Sql & " and RoomID  in(" & Room & ") "
    End If
    Rs2.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    ListAllBox.Clear
    If Rs2.RecordCount > 0 Then
        For i = 1 To Rs2.RecordCount
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.ListAllBox.AddItem IIf(IsNull(Rs2("Name").value), "", Rs2("Name").value)
            Else
            Me.ListAllBox.AddItem IIf(IsNull(Rs2("Namee").value), "", Rs2("Namee").value)
            End If
            ListAllBox.ItemData(ListAllBox.NewIndex) = IIf(IsNull(Rs2("ID").value), 0, Rs2("ID").value)
            Rs2.MoveNext
        Next i

    End If
    Rs2.Close
      ''///////////
    Set Rs2 = New ADODB.Recordset
    Sql = " SELECT    * from TblXXShelf"
    Sql = Sql & "   WHERE   1=1"
    Sql = Sql & " and DepID  in(" & ActivID & ") "
    If Arche <> "0" Then
    Sql = Sql & " and ArchID  in(" & Arche & ") "
    End If
    If Room <> "0" Then
    Sql = Sql & " and RoomID  in(" & Room & ") "
    End If
    If Box <> "0" Then
    Sql = Sql & " and BoxID  in(" & Box & ") "
    End If
    Rs2.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    ListAllRaf.Clear
    If Rs2.RecordCount > 0 Then
        For i = 1 To Rs2.RecordCount
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.ListAllRaf.AddItem IIf(IsNull(Rs2("Name").value), "", Rs2("Name").value)
            Else
            Me.ListAllRaf.AddItem IIf(IsNull(Rs2("Namee").value), "", Rs2("Namee").value)
            End If
            ListAllRaf.ItemData(ListAllRaf.NewIndex) = IIf(IsNull(Rs2("ID").value), 0, Rs2("ID").value)
            Rs2.MoveNext
        Next i

    End If
    Rs2.Close
      ''///////////
    Set Rs2 = New ADODB.Recordset
    Sql = " SELECT    * from TblXXArchDocType"
    Sql = Sql & "   WHERE   1=1"
    Sql = Sql & " and DepID  in(" & ActivID & ") "
    If Arche <> "0" Then
    Sql = Sql & " and ArchID  in(" & Arche & ") "
    End If
    If Room <> "0" Then
    Sql = Sql & " and RoomID  in(" & Room & ") "
    End If
    If Box <> "0" Then
    Sql = Sql & " and BoxID  in(" & Box & ") "
    End If
    If Shelf <> "0" Then
    Sql = Sql & " and ShelfID  in(" & Shelf & ") "
    End If
 
    Rs2.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    ListAllProcess.Clear
    If Rs2.RecordCount > 0 Then
        For i = 1 To Rs2.RecordCount
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.ListAllProcess.AddItem IIf(IsNull(Rs2("Name").value), "", Rs2("Name").value)
            Else
            Me.ListAllProcess.AddItem IIf(IsNull(Rs2("Namee").value), "", Rs2("Namee").value)
            End If
            ListAllProcess.ItemData(ListAllProcess.NewIndex) = IIf(IsNull(Rs2("ID").value), 0, Rs2("ID").value)
            Rs2.MoveNext
        Next i

    End If
    Rs2.Close
End Function

Private Sub ToRecordDate_Change()
If Not IsNull(ToRecordDate.value) Then
ToRecordDateH.value = ToHijriDate(ToRecordDate.value)
End If
End Sub

Private Sub ToRecordDateH_LostFocus()
ToRecordDate.value = ToGregorianDate(ToRecordDateH.value)
End Sub
Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim ActivID As String
    Dim Arche As String
    Dim Room As String
    Dim Box As String
    Dim Shelf As String
    Dim dep As String
    Dim docType As String
    Dim i As Integer
    
    RetriveCondition ActivID, Arche, Room, Box, Shelf
    
 
    docType = "0"
    For i = 0 To Me.ListSelectProcess.ListCount - 1
    docType = docType & "," & Me.ListSelectProcess.ItemData(i)
    Next i
    
MySQL = "SELECT TblTransacRegistr.ID, TblTransacRegistr.BrnchID, TblTransacRegistr.RecordDate, TblTransacRegistr.RecordDateH, TblTransacRegistr.RecordTime, TblTransacRegistr.UserID, TblTransacRegistr.barcode, "
MySQL = MySQL & " TblTransacRegistr.TypTrans, TblTransacRegistr.ImportExport, TblTransacRegistr.NoImpExp, TblTransacRegistr.ImpExpDate, TblTransacRegistr.ImpExpDateH, TblTransacRegistr.Summary, TblTransacRegistr.EnterDate, "
MySQL = MySQL & " TblTransacRegistr.EnterTime, TblTransacRegistr.RequerTime, TblTransacRegistr.ExitTime, TblTransacRegistr.ProcedureReq, TblTransacRegistr.Remarks, TblTransacRegistr.ExitDate, TblTransacRegistr.MHD, "
MySQL = MySQL & " TblTransacRegistr.MHDID, TblTransacRegistr.Posted, TblTransacRegistr.PostedDate, TblTransacRegistr.Approved, TblUsers.UserName AS UserFrom, TblUsers_1.UserName AS UserTo, TblTransacRegistrDet.TransRegID, "
MySQL = MySQL & " TblTransacRegistrDet.FromUser, TblTransacRegistrDet.ID AS TblTransacRegistrDetID, TblTransacRegistrDet.ToUser, TblTransacRegistrDet.FlgTrans, TblTransacRegistrDet.RecDate, "
MySQL = MySQL & " TblTransacRegistrDet.ProcedureReq AS ProcedureReqDet, TblEmpDepartments.DepartmentName, TblEmpDepartments.DepartmentNamee, TblXXArch.Name AS ArchName, TblXXArch.Namee AS ArchNameE, "
MySQL = MySQL & " TblXXRoom.Name AS RoomName, TblXXRoom.Namee AS RoomNameE, TblXXBox.Name AS BoxName, TblXXBox.Namee AS BoxNameE, TblXXShelf.Name AS ShelfName, TblXXShelf.Namee AS ShelfNameE, "
MySQL = MySQL & " TblXXArchDocType.Name AS DocTypeName, TblXXArchDocType.Namee AS DocTypeNameE, TblTransacRegistrDet.Time, TblTransacRegistrDet.TimeUnitID , dbo.TblTransacRegistr.NoteSerial1"
MySQL = MySQL & " FROM TblXXArchDocType RIGHT OUTER JOIN "
MySQL = MySQL & " TblTransacRegistr ON TblXXArchDocType.ID = TblTransacRegistr.TypTrans FULL OUTER JOIN "
MySQL = MySQL & " TblXXShelf RIGHT OUTER JOIN "
MySQL = MySQL & " TblTransacRegistrDet ON TblXXShelf.ID = TblTransacRegistrDet.ShelfID LEFT OUTER JOIN "
MySQL = MySQL & " TblXXRoom ON TblTransacRegistrDet.RoomID = TblXXRoom.ID LEFT OUTER JOIN "
MySQL = MySQL & " TblEmpDepartments ON TblTransacRegistrDet.DepID = TblEmpDepartments.DeparmentID LEFT OUTER JOIN "
MySQL = MySQL & " TblXXBox ON TblTransacRegistrDet.BoxID = TblXXBox.ID LEFT OUTER JOIN "
MySQL = MySQL & " TblXXArch ON TblTransacRegistrDet.ArchID = TblXXArch.ID LEFT OUTER JOIN "
MySQL = MySQL & " TblUsers AS TblUsers_1 ON TblTransacRegistrDet.ToUser = TblUsers_1.UserID LEFT OUTER JOIN "
MySQL = MySQL & " TblUsers ON TblTransacRegistrDet.FromUser = TblUsers.UserID ON TblTransacRegistr.ID = TblTransacRegistrDet.TransRegID where 1 = 1 "
'############################################################################# Where ##########################################################################
If TxtFromNo.Text <> "" Then
    MySQL = MySQL & " and TblTransacRegistr.NoteSerial1 >= " & val(TxtFromNo.Text) & " "
End If

If TxtToNo.Text <> "" Then
    MySQL = MySQL & " and TblTransacRegistr.NoteSerial1 <= " & val(TxtToNo.Text) & " "
End If

If ActivID <> "0" Then
    MySQL = MySQL & " and TblTransacRegistrDet.DepID in (" & ActivID & ") "
End If

If Arche <> "0" Then
    MySQL = MySQL & " and TblTransacRegistrDet.ArchID in (" & Arche & ") "
End If

If Room <> "0" Then
    MySQL = MySQL & " and TblTransacRegistrDet.RoomID in (" & Room & ") "
End If

If Box <> "0" Then
    MySQL = MySQL & " and TblTransacRegistrDet.BoxID in (" & Box & ") "
End If

If Shelf <> "0" Then
    MySQL = MySQL & " and TblTransacRegistrDet.ShelfID  in (" & Shelf & ") "
End If

If Txtbarcode.Text <> "" Then
    MySQL = MySQL & "and TblTransacRegistr.barcode = '" & Txtbarcode.Text & "'"
End If

If TxtNoImpExp.Text <> "" Then
    MySQL = MySQL & "and TblTransacRegistr.NoImpExp = '" & TxtNoImpExp.Text & "'"
End If


If DcbImportExport.Text <> "" And DcbImportExport.ListIndex <> -1 Then
    MySQL = MySQL & " and TblTransacRegistr.ImportExport = (" & DcbImportExport.ListIndex & ") "
End If


If Not IsNull(FrmEnterDate.value) Then
    'MySQL = MySQL & " and 0 < DATEDIFF (n , TblTransacRegistr.EnterDate ," & FrmEnterDate.value & ")"
    MySQL = MySQL & " and TblTransacRegistr.EnterDate >= " & SQLDate(FrmEnterDate.value, True) & " "
End If

If Not IsNull(ToEnterDate.value) Then
    'MySQL = MySQL & " and 0 >  DATEDIFF (n ,TblTransacRegistr.EnterDate ," & ToEnterDate.value & ") "
    MySQL = MySQL & " and TblTransacRegistr.EnterDate <= " & SQLDate(ToEnterDate.value, True) & " "
End If

If Not IsNull(FrmExitDate.value) Then
    'MySQL = MySQL & " and  0 < DATEDIFF (n, TblTransacRegistr.ExitDate ," & FrmExitDate.value & ") "
    MySQL = MySQL & " and  TblTransacRegistr.ExitDate >= " & SQLDate(FrmExitDate.value, True) & ""
End If

If Not IsNull(ToExitDate.value) Then
    'MySQL = MySQL & " and 0 > DATEDIFF (n, TblTransacRegistr.ExitDate ," & ToExitDate.value & ")"
    MySQL = MySQL & " and TblTransacRegistr.ExitDate <=" & SQLDate(ToExitDate.value, True) & " "
End If

If Not IsNull(FrmRecordDate.value) Then
    MySQL = MySQL & " and TblTransacRegistrDet.RecDate >=  " & SQLDate(FrmRecordDate.value, True) & " "
End If

If Not IsNull(ToRecordDate.value) Then
    MySQL = MySQL & " and TblTransacRegistrDet.RecDate <= " & SQLDate(ToRecordDate.value, True) & " "
End If

If Summary.Text <> "" Then
    MySQL = MySQL & " and TblTransacRegistr.Summary like N'% " & Summary.Text & "%' "
End If


    If Rd(0).value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepArch.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepArchE.rpt"
        End If
    ElseIf Rd(1).value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDocCase.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDocCaseE.rpt"
        End If
    Else
        Exit Function
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
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
      Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        StrReportTitle = ""
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
     hide_logo = False
 End Function

