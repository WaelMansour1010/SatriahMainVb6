VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmsandat_ked2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”‰œ ﬁÌœ"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   1080
   ClientWidth     =   13605
   ForeColor       =   &H00404000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   13605
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   765
      Left            =   1440
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   70
      Top             =   3360
      Width           =   10815
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      Height          =   765
      Left            =   1320
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   71
      Top             =   1680
      Width           =   10815
   End
   Begin VB.Frame Frame17 
      Height          =   855
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   600
      Width           =   3735
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   "„·€Ì"
         Height          =   195
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁÌœ œÊ—Ì"
         Height          =   195
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁ«·»"
         Height          =   195
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   " „ «⁄ „«œÂ"
         Height          =   195
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄œÌ„ «· √ÀÌ—"
         Height          =   195
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   720
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label24 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   270
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1920
         TabIndex        =   16
         Top             =   -120
         Width           =   300
      End
      Begin VB.Label Label18 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   0
         Width           =   310
      End
      Begin VB.Label Label17 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   4560
      TabIndex        =   54
      Top             =   9360
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3360
         TabIndex        =   58
         Top             =   240
         Width           =   855
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   57
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   56
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   4560
      TabIndex        =   49
      Top             =   9360
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4800
         TabIndex        =   52
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   51
         Top             =   240
         Width           =   975
      End
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3120
         TabIndex        =   50
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   720
      TabIndex        =   44
      Top             =   8040
      Width           =   12735
      Begin ALLButtonS.ALLButton Command1 
         Height          =   495
         Index           =   1
         Left            =   10680
         TabIndex        =   45
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Õ›Ÿ"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmsandat_ked2.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   495
         Index           =   2
         Left            =   9720
         TabIndex        =   46
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Õ–›"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   255
         BCOLO           =   192
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmsandat_ked2.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   495
         Index           =   0
         Left            =   11640
         TabIndex        =   47
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÃœÌœ"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmsandat_ked2.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton1 
         Height          =   495
         Left            =   1440
         TabIndex        =   65
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "«·„—›ﬁ« "
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   65280
         BCOLO           =   65280
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsandat_ked2.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Print_cmd 
         Height          =   495
         Left            =   2400
         TabIndex        =   66
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÿ»«⁄…"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   65535
         BCOLO           =   65535
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmsandat_ked2.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton2 
         Height          =   495
         Left            =   7800
         TabIndex        =   67
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "»ÕÀ ﬁÌœ"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   65280
         BCOLO           =   65280
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsandat_ked2.frx":008C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   495
         Index           =   3
         Left            =   8760
         TabIndex        =   83
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "«⁄ „«œ"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   255
         BCOLO           =   192
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frmsandat_ked2.frx":00A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         Caption         =   "Label2"
         Height          =   15
         Left            =   -120
         TabIndex        =   48
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.CommandButton Command13 
      Height          =   735
      Left            =   13560
      Picture         =   "frmsandat_ked2.frx":00C4
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "»ÕÀ ⁄‰ ”‰œF3"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   840
      TabIndex        =   31
      Top             =   9600
      Visible         =   0   'False
      Width           =   12495
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "„"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   11760
         TabIndex        =   64
         Top             =   120
         Width           =   135
      End
      Begin VB.Line Line1 
         X1              =   12190
         X2              =   12190
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line9 
         X1              =   11685
         X2              =   11685
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line10 
         X1              =   6390
         X2              =   6390
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line11 
         X1              =   5090
         X2              =   5090
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "œ«∆‰"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   5520
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.Line Line12 
         X1              =   7700
         X2              =   7700
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line13 
         X1              =   9690
         X2              =   9690
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·Õ”«»"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   10080
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·Õ”«»"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   8160
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "„œÌ‰"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   6840
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "«·‘—Õ"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2160
         TabIndex        =   32
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000003&
      Height          =   735
      Left            =   720
      TabIndex        =   37
      Top             =   9480
      Visible         =   0   'False
      Width           =   12735
      Begin VB.Line Line5 
         X1              =   810
         X2              =   810
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "I"
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
         Height          =   375
         Left            =   600
         TabIndex        =   63
         Top             =   120
         Width           =   255
      End
      Begin VB.Line Line14 
         X1              =   7410
         X2              =   7410
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
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
         Height          =   615
         Left            =   6360
         TabIndex        =   42
         Top             =   120
         Width           =   975
      End
      Begin VB.Line Line8 
         X1              =   300
         X2              =   300
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   0
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line6 
         X1              =   2800
         X2              =   2800
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line4 
         X1              =   6110
         X2              =   6110
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "A/C#"
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
         Height          =   375
         Left            =   1440
         TabIndex        =   41
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Name"
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
         Height          =   615
         Left            =   3120
         TabIndex        =   40
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Depit"
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
         Height          =   615
         Left            =   5040
         TabIndex        =   39
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   495
         Left            =   9480
         TabIndex        =   38
         Top             =   120
         Width           =   1815
      End
      Begin VB.Line Line2 
         X1              =   4800
         X2              =   4800
         Y1              =   120
         Y2              =   720
      End
   End
   Begin VB.Frame Frame8 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   8880
      TabIndex        =   28
      Top             =   480
      Width           =   3255
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         DataField       =   "sandat_pc_no"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   -480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo Dckedtype 
         Height          =   315
         Left            =   240
         TabIndex        =   72
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
   End
   Begin VB.Frame Frame10 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   3840
      TabIndex        =   27
      Top             =   600
      Width           =   3735
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   120
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker XPDtbBill 
         Height          =   330
         Left            =   1680
         TabIndex        =   73
         Top             =   600
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   582
         _Version        =   393216
         Format          =   102891521
         CurrentDate     =   38784
      End
   End
   Begin VB.Frame Frame14 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   12120
      TabIndex        =   24
      Top             =   480
      Width           =   1455
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "—ﬁ„ «·ﬁÌœ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "„”·”· «·”‰œ"
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
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   -600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "‰Ê⁄ «·ﬁÌœ"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame15 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   7560
      TabIndex        =   22
      Top             =   480
      Width           =   1935
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   " «—ÌŒ «·ﬁÌœ"
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   62
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "„’œ— «·ﬁÌœ"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   -120
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame16 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2880
      TabIndex        =   21
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   2040
      TabIndex        =   18
      Top             =   9840
      Visible         =   0   'False
      Width           =   7095
      Begin MSAdodcLib.Adodc user_priviliges_adodc 
         Height          =   495
         Left            =   120
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   " "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label screen_name 
         Alignment       =   1  'Right Justify
         Caption         =   "M19"
         Height          =   255
         Left            =   3360
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label adodc4error 
         Alignment       =   1  'Right Justify
         Caption         =   "Label4"
         DataField       =   "employee_id"
         DataSource      =   "user_priviliges_adodc"
         Height          =   495
         Left            =   2160
         TabIndex        =   19
         Top             =   120
         Width           =   495
      End
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   585
      Left            =   2880
      Top             =   8040
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1032
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   " Õ—Ìﬂ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   -5280
      Top             =   6840
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   7
      Top             =   7080
      Width           =   12735
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   600
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DcboUsers 
         Height          =   315
         Left            =   0
         TabIndex        =   84
         Top             =   360
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Õ—— »Ê«”ÿ…"
         Height          =   375
         Left            =   2160
         TabIndex        =   85
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "«Ã„«·Ì «·„œÌ‰"
         Height          =   495
         Left            =   8520
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "«Ã„«·Ì «·œ«∆‰"
         Height          =   495
         Left            =   5640
         TabIndex        =   9
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "«·›—ﬁ"
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDataListLib.DataCombo DcAccount1 
      Height          =   315
      Left            =   11040
      TabIndex        =   86
      Top             =   2880
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcAccount2 
      Height          =   315
      Left            =   7080
      TabIndex        =   87
      Top             =   2880
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmsandat_ked2.frx":1A56
      Height          =   3015
      Left            =   720
      TabIndex        =   88
      Top             =   4200
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483637
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "opr_id"
         Caption         =   "opr_id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "index"
         Caption         =   "„"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "account_no"
         Caption         =   "—ﬁ„ «·Õ”«»"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "account_name"
         Caption         =   "«”„ «·Õ”«»"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "depet_value"
         Caption         =   "„œÌ‰"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "credit_value"
         Caption         =   "œ«∆‰"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "description"
         Caption         =   "«·‘—Õ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "sandat_pc_no"
         Caption         =   "sandat_pc_no"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "sanad_no"
         Caption         =   "sanad_no"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "sanad_type"
         Caption         =   "sanad_type"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "sanad_source"
         Caption         =   "sanad_source"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "box_name"
         Caption         =   "box_name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "bona_3la"
         Caption         =   "bona_3la"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "date"
         Caption         =   "date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "dist"
         Caption         =   "dist"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   5265.071
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1214.929
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1032
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   " Õ—Ìﬂ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   4560
      Picture         =   "frmsandat_ked2.frx":1A6B
      Top             =   2760
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6480
      Picture         =   "frmsandat_ked2.frx":1F1A
      Top             =   2760
      Width           =   360
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      Caption         =   "     ”‰œ ﬁÌœ         "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   -600
      TabIndex        =   68
      Top             =   0
      Width           =   14295
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      Caption         =   "«·‘—Õ «·⁄«„"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11760
      TabIndex        =   60
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Caption         =   "«·‘—Õ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11760
      TabIndex        =   59
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Line Line3 
      X1              =   -600
      X2              =   -600
      Y1              =   480
      Y2              =   1080
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "«· Ê“Ì⁄"
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
      Height          =   495
      Left            =   -1800
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   " ÕœÌÀ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   -1320
      TabIndex        =   12
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   11880
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "„—«ﬂ“ «· ﬂ·›…"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "œ«∆‰"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "„œÌ‰"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·Õ”«»"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·Õ”«»"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11400
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   5040
      TabIndex        =   1
      Top             =   5280
      Width           =   855
   End
End
Attribute VB_Name = "frmsandat_ked2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
    Adodc1.Recordset.AddNew

    Adodc1.Recordset.Fields!NoteID = CStr(new_id("NOTES", "NoteID ", "", True))

    Adodc1.Recordset.Fields!NoteType = 200
    Adodc1.Recordset.Fields!NoteSerial = X
    Adodc1.Recordset.Fields!NoteDate = Date
    Adodc1.Recordset.Fields!Note_Value = 0
    Adodc1.Recordset.Fields!remark = X
    Adodc1.Recordset.Fields!UserID = val(Me.DcboUsers.BoundText)

End Sub

Private Sub DcAccount1_Click(Area As Integer)

    If DcAccount1.text = "" Then Exit Sub
    Dim My_SQL As String

    My_SQL = "select Account_Name from ACCOUNTS where Account_Serial='" & DcAccount1.text & "'"
 
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient

    rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    DcAccount2.text = rec.Fields("Account_Name").value
End Sub

Private Sub DcAccount2_Change()

    If DcAccount2.text = "" Then Exit Sub
    Dim My_SQL As String

    My_SQL = "select Account_Serial from ACCOUNTS where Account_Name ='" & DcAccount2.text & "'"
 
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient

    rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    DcAccount1.text = rec.Fields("Account_Serial").value
    DcAccount2.SetFocus
    SendKeys "{f4}"
End Sub

Private Sub DcAccount2_Click(Area As Integer)

    If DcAccount2.text = "" Then Exit Sub
    Dim My_SQL As String

    My_SQL = "select Account_Serial from ACCOUNTS where Account_Name ='" & DcAccount2.text & "'"
 
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient

    rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    DcAccount1.text = rec.Fields("Account_Serial").value
End Sub

Private Sub Form_Load()
    'On Error GoTo ErrTrap
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Dim i As Integer
    Dim My_SQL As String

    My_SQL = "select * from Notes where NoteType=200"
    'Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient

    RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    'Me.TxtModFlg.text = "R"
    'Resize_Form Me
    'load ACCOUNTS -----------------------------------------------
    My_SQL = "  select Account_Name,Account_Serial from ACCOUNTS  where last_account=1"

    fill_combo DcAccount1, My_SQL
    My_SQL = "  select Account_Serial,Account_Name from ACCOUNTS  where last_account=1"

    fill_combo DcAccount2, My_SQL

    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DcboUsers

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from NOTES  where  NOTE_TYPE=200"
    Adodc1.Refresh

    'load ked type------------------------------------------------------
    'fill_combo DcAccount1, My_SQL
    'My_SQL = "  select ked_no,ked_name from ked_types   "

    'Dckedtype

    'FillGridWithData
    'With Me.Grid
    '    .Cell(flexcpPicture, 0, .ColIndex("name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
    '    .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
    '    For I = 0 To .Cols - 1
    '        .Cell(flexcpPictureAlignment, 0, I) = flexPicAlignRightCenter
    '    Next
    '
    '    .ExtendLastCol = True
    '    .WallPaper = BKGrndPic.Picture
    '    .RowHeight(-1) = 300
    'End With
    'BtnFirst_Click
    'ShowTip
    '
ErrTrap:

End Sub

