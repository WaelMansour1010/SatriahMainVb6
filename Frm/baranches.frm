VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{784C0C13-85E7-4E11-A8FB-F0243A135D03}#2.0#0"; "SuperLablel.ocx"
Begin VB.Form baranches 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  «⁄œ«œ«  «·—»ÿ „⁄  «·Õ”«»« "
   ClientHeight    =   10635
   ClientLeft      =   4800
   ClientTop       =   375
   ClientWidth     =   11745
   Icon            =   "baranches.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10635
   ScaleWidth      =   11745
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame19 
      Caption         =   "Frame19"
      Height          =   4335
      Left            =   13440
      RightToLeft     =   -1  'True
      TabIndex        =   235
      Top             =   600
      Width           =   12135
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":000C
         DataField       =   "a87"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   87
         Left            =   0
         TabIndex        =   236
         Top             =   0
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0021
         DataField       =   "a88"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   88
         Left            =   0
         TabIndex        =   237
         Top             =   360
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0036
         DataField       =   "a89"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   89
         Left            =   0
         TabIndex        =   238
         Top             =   720
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":004B
         DataField       =   "a890"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   90
         Left            =   0
         TabIndex        =   239
         Top             =   1080
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0060
         DataField       =   "a91"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   91
         Left            =   0
         TabIndex        =   240
         Top             =   1440
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0075
         DataField       =   "a16"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   16
         Left            =   0
         TabIndex        =   253
         Top             =   1920
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":008A
         DataField       =   "a53"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   53
         Left            =   0
         TabIndex        =   254
         Top             =   2280
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":009F
         DataField       =   "a54"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   54
         Left            =   0
         TabIndex        =   255
         Top             =   2640
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·«ÃÊ— "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   9240
         TabIndex        =   258
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·Œ’„"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   53
         Left            =   9240
         TabIndex        =   257
         Top             =   2040
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·„ﬂ«›√…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   54
         Left            =   9240
         TabIndex        =   256
         Top             =   2400
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«· √„Ì‰ «·„” —œ «Ì—«œ "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   83
         Left            =   8760
         TabIndex        =   245
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„Ì«Â „ﬁœ„ «Ì—«œ "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   84
         Left            =   8880
         TabIndex        =   244
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«·ﬂÂ—»«¡ „ﬁœ„ «Ì—«œ "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   85
         Left            =   8880
         TabIndex        =   243
         Top             =   1080
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«·”⁄Ì Ê«·—”Ê„ «·«œ«—Ì… «Ì—«œ  "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   86
         Left            =   8760
         TabIndex        =   242
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "—”Ê„ «œ«—Ì…    «Ì—«œ "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   87
         Left            =   8880
         TabIndex        =   241
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Left            =   5640
      TabIndex        =   44
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   " «·„Œ«“‰/«·„»Ì⁄« /«·„‘ —Ì« "
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":00B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "address"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1200
      TabIndex        =   39
      Top             =   -360
      Width           =   4815
   End
   Begin VB.Frame Frame6 
      Caption         =   "œ·«·«  «·«·Ê«‰"
      ClipControls    =   0   'False
      Height          =   615
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   9960
      Width           =   8535
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» —∆Ì”Ì"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
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
         Height          =   255
         Left            =   1440
         TabIndex        =   37
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» ‰Â«∆Ì"
         Height          =   255
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
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
         Height          =   255
         Left            =   4320
         TabIndex        =   35
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.TextBox txtnamee 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "branch_namee"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1200
      TabIndex        =   32
      Top             =   -840
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "tel"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   9480
      TabIndex        =   2
      Top             =   -360
      Width           =   2295
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   8880
      TabIndex        =   22
      Top             =   9960
      Width           =   2655
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   23
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "baranches.frx":00D0
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
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   24
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "baranches.frx":00EC
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
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÿ»«⁄…"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "baranches.frx":0108
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
         Height          =   375
         Index           =   0
         Left            =   3960
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "baranches.frx":0124
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   960
         Top             =   960
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   582
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
         Caption         =   "  Õ—Ìﬂ"
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
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   276
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Œ—ÊÃ"
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "baranches.frx":0140
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         Caption         =   "Label2"
         Height          =   15
         Left            =   -120
         TabIndex        =   27
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   1200
      TabIndex        =   17
      Top             =   12000
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3840
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   1200
      TabIndex        =   12
      Top             =   11640
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   0
      TabIndex        =   11
      ToolTipText     =   " «··€…"
      Top             =   -1320
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "EN"
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
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":015C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   -1  'True
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2175
      Left            =   600
      TabIndex        =   8
      Top             =   10680
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Label Label9 
         Caption         =   "Tel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   0
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Branch#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   10800
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
         Caption         =   "M30"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox txtnameA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "branch_name"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3600
      TabIndex        =   1
      Top             =   -840
      Width           =   2535
   End
   Begin VB.TextBox txtcode 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "branch_id"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   7080
      TabIndex        =   0
      Top             =   -840
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   480
      Top             =   -1320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2175
      Left            =   9960
      TabIndex        =   7
      Top             =   -3360
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   480
      Top             =   11040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   2400
      Top             =   10920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   735
      Left            =   4920
      TabIndex        =   6
      Top             =   -1920
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      Text            =   "»Ì«‰«  «·›—Ê⁄"
      ColorGeneral    =   0
      ColorGeneral    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcemployee 
      DataField       =   "manger_id"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   6960
      TabIndex        =   42
      Top             =   -360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ListField       =   ""
      BoundColumn     =   "Account_Code"
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   4560
      Top             =   10800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
   Begin ALLButtonS.ALLButton ALLButton3 
      Height          =   375
      Left            =   4320
      TabIndex        =   68
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   " ‘∆Ê‰ „ÊŸ›Ì‰"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":0178
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton4 
      Height          =   375
      Left            =   7800
      TabIndex        =   82
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   " «·«’Ê·"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":0194
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton5 
      Height          =   375
      Left            =   8400
      TabIndex        =   83
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ”«»«  «·„‘«—Ì⁄"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":01B0
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
      Height          =   375
      Left            =   240
      TabIndex        =   105
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«·—»ÿ »«·„Ã„Ê⁄« "
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":01CC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton6 
      Height          =   375
      Left            =   8880
      TabIndex        =   106
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   " «·„⁄«„·«  «·„«·Ì…"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":01E8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton7 
      Height          =   375
      Left            =   9960
      TabIndex        =   115
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ”«»«  «·«‰ «Ã"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":0204
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton8 
      Height          =   375
      Left            =   1920
      TabIndex        =   128
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ”«»«  «·«”Â„"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":0220
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton9 
      Height          =   375
      Left            =   6720
      TabIndex        =   129
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ”«»«  «œ«—… «·«„·«ﬂ"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":023C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   375
      Left            =   2640
      Top             =   -1080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
   Begin MSDataListLib.DataCombo DcActivityType 
      DataField       =   "ActivityTypeId"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   9480
      TabIndex        =   140
      Top             =   -840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ListField       =   ""
      BoundColumn     =   "Account_Code"
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ALLButtonS.ALLButton ALLButton10 
      Height          =   375
      Left            =   2880
      TabIndex        =   142
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«·‰ﬁ· «·„œ—”Ì"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":0258
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton11 
      Height          =   375
      Left            =   10320
      TabIndex        =   173
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   " «·Ê”Ìÿ «·«›  «ÕÌ"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":0274
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton12 
      Height          =   375
      Left            =   3600
      TabIndex        =   188
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«·‰ﬁ·Ì«  Ê «·Õ«ÊÌ« "
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":0290
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton13 
      Height          =   375
      Left            =   120
      TabIndex        =   246
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "’Ì«‰… «·„⁄œ« /«·”Ì«—« "
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":02AC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton14 
      Height          =   375
      Left            =   5280
      TabIndex        =   288
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ”«»«  «·„”«Â„« "
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":02C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton15 
      Height          =   375
      Left            =   1440
      TabIndex        =   336
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«·„⁄«Âœ Ê«·„œ«—”"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":02E4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton16 
      Height          =   375
      Left            =   120
      TabIndex        =   344
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«·ÕÃ Ê«·⁄„—…"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranches.frx":0300
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame9 
      Caption         =   "Õ”«»«  «·«‰ «Ã"
      Height          =   8295
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   116
      Top             =   1320
      Width           =   11535
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":031C
         DataField       =   "a37"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   37
         Left            =   120
         TabIndex        =   117
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0331
         DataField       =   "a38"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   38
         Left            =   120
         TabIndex        =   119
         Top             =   840
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0346
         DataField       =   "a39"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   39
         Left            =   120
         TabIndex        =   120
         Top             =   1320
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":035B
         DataField       =   "a68"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   68
         Left            =   120
         TabIndex        =   185
         Top             =   2160
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0370
         DataField       =   "a79"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   79
         Left            =   120
         TabIndex        =   204
         Top             =   1680
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0385
         DataField       =   "a151"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   151
         Left            =   120
         TabIndex        =   378
         Top             =   2640
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "„’«—Ì› ’‰«⁄Ì…"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   144
         Left            =   9360
         TabIndex        =   379
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "„’«—Ì› «·«‰ «Ã   ﬂÂ—»«¡"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   75
         Left            =   9360
         TabIndex        =   205
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "„’—Ê›«   ‘€Ì· «‰ «Ã ‰’› „’‰⁄"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   66
         Left            =   9360
         TabIndex        =   186
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "„’«—Ì› «·«‰ «Ã  «ÃÊ—"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   38
         Left            =   9360
         TabIndex        =   184
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "„’«—Ì› «·«‰ «Ã   ÊﬁÊœ"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   39
         Left            =   9285
         TabIndex        =   121
         Top             =   1365
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "„’«—Ì› «·«‰ «Ã „Ê«œ"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   37
         Left            =   9285
         TabIndex        =   118
         Top             =   405
         Width           =   1695
      End
   End
   Begin VB.Frame Frame20 
      Caption         =   "’Ì«‰… «·„⁄œ« /«·”Ì«—« "
      Height          =   8295
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   247
      Top             =   1440
      Width           =   11535
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":039A
         DataField       =   "a77"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   77
         Left            =   240
         TabIndex        =   248
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":03AF
         DataField       =   "a78"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   78
         Left            =   240
         TabIndex        =   249
         Top             =   600
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ   «Ì—«œ«  «·’Ì«‰…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   73
         Left            =   9480
         TabIndex        =   251
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ   „’—Ê›«  «·’Ì«‰…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   74
         Left            =   9600
         TabIndex        =   250
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Õ”«»«  «·«’Ê·"
      Height          =   8295
      Left            =   -90
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   1440
      Visible         =   0   'False
      Width           =   11535
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":03C4
         DataField       =   "a24"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   24
         Left            =   120
         TabIndex        =   123
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":03D9
         DataField       =   "a25"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   25
         Left            =   120
         TabIndex        =   124
         Top             =   1080
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":03EE
         DataField       =   "a26"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   26
         Left            =   120
         TabIndex        =   125
         Top             =   720
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0403
         DataField       =   "a31"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   31
         Left            =   120
         TabIndex        =   126
         Top             =   1440
         Visible         =   0   'False
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0418
         DataField       =   "a40"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   40
         Left            =   120
         TabIndex        =   127
         Top             =   1800
         Visible         =   0   'False
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":042D
         DataField       =   "a66"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   66
         Left            =   120
         TabIndex        =   182
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0442
         DataField       =   "a67"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   67
         Left            =   120
         TabIndex        =   183
         Top             =   1800
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«»  «·«’·"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   24
         Left            =   9240
         TabIndex        =   78
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   " Œ”«∆— »Ì⁄ √. À«» …"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   40
         Left            =   9240
         TabIndex        =   122
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   " «—»«Õ  »Ì⁄ √. À«» …"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   31
         Left            =   9240
         TabIndex        =   81
         Top             =   1515
         Width           =   1815
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   " „Ã„⁄ «·«Â·«ﬂ"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   26
         Left            =   9240
         TabIndex        =   80
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«»  „’—Ê› «·«Â·«ﬂ"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   25
         Left            =   9240
         TabIndex        =   79
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.Frame Frame17 
      Caption         =   "«·Õ”«»«  «·Ê”Ìÿ… «·«›  «ÕÌ…"
      Height          =   8295
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   156
      Top             =   1440
      Width           =   11535
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0457
         DataField       =   "a19"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   19
         Left            =   120
         TabIndex        =   157
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":046C
         DataField       =   "a41"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   41
         Left            =   120
         TabIndex        =   159
         Top             =   840
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0481
         DataField       =   "a46"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   46
         Left            =   120
         TabIndex        =   161
         Top             =   3000
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0496
         DataField       =   "a57"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   57
         Left            =   120
         TabIndex        =   163
         Top             =   1200
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":04AB
         DataField       =   "a58"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   58
         Left            =   120
         TabIndex        =   165
         Top             =   1560
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":04C0
         DataField       =   "a59"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   59
         Left            =   120
         TabIndex        =   167
         Top             =   1920
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":04D5
         DataField       =   "a60"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   60
         Left            =   120
         TabIndex        =   169
         Top             =   2280
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":04EA
         DataField       =   "a61"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   61
         Left            =   120
         TabIndex        =   171
         Top             =   2640
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":04FF
         DataField       =   "a62"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   62
         Left            =   120
         TabIndex        =   174
         Top             =   3360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0514
         DataField       =   "a73"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   73
         Left            =   120
         TabIndex        =   198
         Top             =   3720
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ Ê”Ìÿ   √›  «ÕÌ «·„‘«—Ì⁄"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   71
         Left            =   9240
         TabIndex        =   199
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ Ê”Ìÿ «·«⁄ „«œ«  «·„” ‰œÌ…"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   62
         Left            =   9240
         TabIndex        =   175
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ Ê”Ìÿ «›  «ÕÌ  „ÊŸ›Ì‰"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   61
         Left            =   9285
         TabIndex        =   172
         Top             =   2685
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ Ê”Ìÿ «›  «ÕÌ  „Ê—œÌ‰"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   60
         Left            =   9285
         TabIndex        =   170
         Top             =   2325
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ Ê”Ìÿ «›  «ÕÌ  ⁄„·«¡"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   59
         Left            =   9285
         TabIndex        =   168
         Top             =   1965
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ Ê”Ìÿ «›  «ÕÌ  »‰Êﬂ"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   58
         Left            =   9285
         TabIndex        =   166
         Top             =   1605
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ Ê”Ìÿ «›  «ÕÌ  Œ“‰ Ê⁄Âœ "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   57
         Left            =   9285
         TabIndex        =   164
         Top             =   1245
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ Ê”Ìÿ «›  «ÕÌ  «·«”Â„"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   46
         Left            =   9360
         TabIndex        =   162
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ Ê”Ìÿ «›  «ÕÌ  √ À«» …"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   41
         Left            =   9285
         TabIndex        =   160
         Top             =   885
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ Ê”Ìÿ «›  «ÕÌ „Œ“Ê‰"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   19
         Left            =   9645
         TabIndex        =   158
         Top             =   525
         Width           =   1695
      End
   End
   Begin VB.Frame Frame21 
      Caption         =   "«·‰ﬁ· «·„œ—”Ì"
      Height          =   8295
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   281
      Top             =   1440
      Visible         =   0   'False
      Width           =   11535
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0529
         DataField       =   "a105"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   105
         Left            =   240
         TabIndex        =   282
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":053E
         DataField       =   "a106"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   106
         Left            =   240
         TabIndex        =   284
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0553
         DataField       =   "a107"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   107
         Left            =   240
         TabIndex        =   286
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "œ›⁄«  «·„ ⁄ÂœÌ‰ «·„” Õﬁ…"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   103
         Left            =   9240
         TabIndex        =   287
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "  ﬂ·›… «·‰ﬁ·"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   102
         Left            =   9240
         TabIndex        =   285
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ì—«œ«  «·‰ﬁ·"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   101
         Left            =   9240
         TabIndex        =   283
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Õ”«»«  «·„‘«—Ì⁄"
      Height          =   8295
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   84
      Top             =   1440
      Width           =   11535
      Begin VB.Frame Frame8 
         Caption         =   "«·Ì… «·„‘«—Ì⁄"
         Height          =   2535
         Left            =   9240
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   240
         Width           =   2175
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Õ”«» Õ”‰ «·«œ«¡"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   145
            Left            =   0
            TabIndex        =   381
            Top             =   2280
            Width           =   2055
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«»   „ﬁ«Ê·Ì «·»«ÿ‰"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   36
            Left            =   480
            TabIndex        =   252
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«» „” Œ·’«  «·„‘«—Ì⁄"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   107
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«» „’—Ê›«  «·„‘«—Ì⁄"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   93
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«» «Ì—«œ«  «·„‘«—Ì⁄"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   15
            Left            =   360
            TabIndex        =   92
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«» „Êœ «·„‘«—Ì⁄"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   27
            Left            =   600
            TabIndex        =   91
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«» «ÃÊ— «·„‘«—Ì⁄"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   28
            Left            =   600
            TabIndex        =   90
            Top             =   1320
            Width           =   1455
         End
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0568
         DataField       =   "a14"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   14
         Left            =   240
         TabIndex        =   85
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":057D
         DataField       =   "a15"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   15
         Left            =   240
         TabIndex        =   86
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0592
         DataField       =   "a27"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   27
         Left            =   240
         TabIndex        =   87
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":05A7
         DataField       =   "a28"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   28
         Left            =   240
         TabIndex        =   88
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":05BC
         DataField       =   "a32"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   32
         Left            =   240
         TabIndex        =   108
         Top             =   1800
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":05D1
         DataField       =   "a36"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   36
         Left            =   240
         TabIndex        =   234
         Top             =   2160
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":05E6
         DataField       =   "a103"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   103
         Left            =   240
         TabIndex        =   277
         Top             =   2880
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":05FB
         DataField       =   "a104"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   104
         Left            =   240
         TabIndex        =   279
         Top             =   3240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0610
         DataField       =   "a142"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   142
         Left            =   240
         TabIndex        =   361
         Top             =   3600
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0625
         DataField       =   "a152"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   152
         Left            =   240
         TabIndex        =   380
         Top             =   2520
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":063A
         DataField       =   "a159"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   159
         Left            =   240
         TabIndex        =   394
         Top             =   3960
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ì—«œ «· —ﬂÌ»« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   152
         Left            =   9000
         TabIndex        =   395
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "„‘—Ê⁄«   Õ  «· ‰›Ì–"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   136
         Left            =   9120
         TabIndex        =   360
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "œ›⁄«  „ﬁœ„… ⁄„·«¡"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   100
         Left            =   9120
         TabIndex        =   280
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ”„Ì«  „” Œ·’« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   99
         Left            =   9120
         TabIndex        =   278
         Top             =   2880
         Width           =   2055
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Õ”«»«  «·„ÊŸ›Ì‰"
      Height          =   8295
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   1320
      Visible         =   0   'False
      Width           =   11535
      Begin VB.Frame Frame5 
         Caption         =   "Õ”«»«  «·„ÊŸ›Ì‰"
         Height          =   5535
         Left            =   9240
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   240
         Width           =   2175
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Œ’„ «·«Ã«“… «·„—÷Ì…"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   155
            Left            =   330
            TabIndex        =   400
            Top             =   4860
            Width           =   1425
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "«Ã«“«  »œÊ‰ —« »"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   135
            Left            =   360
            TabIndex        =   359
            Top             =   4440
            Width           =   1215
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Œ’Ê„«  ‰Â«Ì… «·Œœ„…"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   134
            Left            =   120
            TabIndex        =   358
            Top             =   4080
            Width           =   1455
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "«÷«›«  ‰Â«Ì… «·Œœ„…"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   133
            Left            =   120
            TabIndex        =   357
            Top             =   3720
            Width           =   1455
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«»  „’—Ê› «· –«ﬂ—"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   90
            Left            =   120
            TabIndex        =   232
            Top             =   3000
            Width           =   1935
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "„Œ’’ «· –«ﬂ—"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   89
            Left            =   120
            TabIndex        =   227
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "„Œ’’ ‰Â«Ì… «·Œœ„…"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   72
            Left            =   120
            TabIndex        =   201
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„œ›Ê⁄«  «·„ﬁœ„Â"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   65
            Left            =   120
            TabIndex        =   180
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "«·»œ·«  «·„ﬁœ„…"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   64
            Left            =   360
            TabIndex        =   178
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«»  „’—Ê›  —ﬂ «·Œœ„…"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   56
            Left            =   120
            TabIndex        =   153
            Top             =   2640
            Width           =   1935
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«» „’—Ê› «·«Ã«“…"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   55
            Left            =   240
            TabIndex        =   152
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "«·«ÃÊ— «·„” Õﬁ… "
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   76
            Top             =   550
            Width           =   1455
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "–„„  «·„ÊŸ›Ì‰"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   75
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "„Œ’’ «·«Ã«“« "
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   74
            Top             =   840
            Width           =   1455
         End
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":064F
         DataField       =   "a7"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   7
         Left            =   240
         TabIndex        =   70
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0664
         DataField       =   "a29"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   29
         Left            =   240
         TabIndex        =   71
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0679
         DataField       =   "a30"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   30
         Left            =   240
         TabIndex        =   72
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":068E
         DataField       =   "a55"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   55
         Left            =   240
         TabIndex        =   154
         Top             =   2520
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":06A3
         DataField       =   "a56"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   56
         Left            =   240
         TabIndex        =   155
         Top             =   2880
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":06B8
         DataField       =   "a64"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   64
         Left            =   240
         TabIndex        =   179
         Top             =   3600
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":06CD
         DataField       =   "a65"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   65
         Left            =   240
         TabIndex        =   181
         Top             =   2160
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":06E2
         DataField       =   "a74"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   74
         Left            =   240
         TabIndex        =   200
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":06F7
         DataField       =   "a93"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   93
         Left            =   240
         TabIndex        =   226
         Top             =   1800
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":070C
         DataField       =   "a94"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   94
         Left            =   240
         TabIndex        =   233
         Top             =   3240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0721
         DataField       =   "a139"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   139
         Left            =   240
         TabIndex        =   354
         Top             =   3960
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0736
         DataField       =   "a140"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   140
         Left            =   240
         TabIndex        =   355
         Top             =   4320
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":074B
         DataField       =   "a141"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   141
         Left            =   240
         TabIndex        =   356
         Top             =   4680
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0760
         DataField       =   "a204"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   204
         Left            =   240
         TabIndex        =   401
         Top             =   5130
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "«œ«—… «·«„·«ﬂ"
      Height          =   8535
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   206
      Top             =   1440
      Visible         =   0   'False
      Width           =   11535
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0775
         DataField       =   "a47"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   47
         Left            =   120
         TabIndex        =   207
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":078A
         DataField       =   "a48"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   48
         Left            =   120
         TabIndex        =   208
         Top             =   600
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":079F
         DataField       =   "a80"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   80
         Left            =   120
         TabIndex        =   217
         Top             =   1320
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":07B4
         DataField       =   "a81"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   81
         Left            =   120
         TabIndex        =   218
         Top             =   3120
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":07C9
         DataField       =   "a82"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   82
         Left            =   120
         TabIndex        =   219
         Top             =   960
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":07DE
         DataField       =   "a83"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   83
         Left            =   120
         TabIndex        =   220
         Top             =   3480
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":07F3
         DataField       =   "a84"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   84
         Left            =   120
         TabIndex        =   221
         Top             =   3840
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0808
         DataField       =   "a85"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   85
         Left            =   120
         TabIndex        =   222
         Top             =   4200
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":081D
         DataField       =   "a86"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   86
         Left            =   120
         TabIndex        =   223
         Top             =   4560
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0832
         DataField       =   "a95"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   95
         Left            =   120
         TabIndex        =   259
         Top             =   5280
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0847
         DataField       =   "a92"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   92
         Left            =   120
         TabIndex        =   263
         Top             =   4920
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":085C
         DataField       =   "a123"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   123
         Left            =   120
         TabIndex        =   320
         Top             =   5640
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0871
         DataField       =   "a124"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   124
         Left            =   120
         TabIndex        =   322
         Top             =   6000
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0886
         DataField       =   "a125"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   125
         Left            =   120
         TabIndex        =   324
         Top             =   6360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":089B
         DataField       =   "a143"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   143
         Left            =   120
         TabIndex        =   362
         Top             =   6720
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":08B0
         DataField       =   "a153"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   153
         Left            =   120
         TabIndex        =   382
         Top             =   1680
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":08C5
         DataField       =   "a154"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   154
         Left            =   120
         TabIndex        =   384
         Top             =   2040
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":08DA
         DataField       =   "a155"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   155
         Left            =   120
         TabIndex        =   386
         Top             =   2400
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":08EF
         DataField       =   "a156"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   156
         Left            =   120
         TabIndex        =   388
         Top             =   2760
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0904
         DataField       =   "a161"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   161
         Left            =   120
         TabIndex        =   398
         Top             =   7080
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0919
         DataField       =   "a207"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   162
         Left            =   120
         TabIndex        =   402
         Top             =   7440
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":092E
         DataField       =   "a163"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   163
         Left            =   120
         TabIndex        =   404
         Top             =   7800
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0943
         DataField       =   "a166"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   166
         Left            =   120
         TabIndex        =   410
         Top             =   8160
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·‹ √„Ì‰ «·„” —œ ··„” √Ã—Ì‰"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   160
         Left            =   8760
         TabIndex        =   411
         Top             =   8160
         Width           =   2535
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» „’—Ê›«  «·«„·«ﬂ"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   157
         Left            =   9000
         TabIndex        =   405
         Top             =   7830
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·⁄„Ê·« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   156
         Left            =   9000
         TabIndex        =   403
         Top             =   7470
         Width           =   2055
      End
      Begin VB.Shape Shape2 
         Height          =   1815
         Left            =   8880
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» ⁄„Ê·«  «·„‰«œÌ»"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   154
         Left            =   9000
         TabIndex        =   399
         Top             =   7080
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ì—«œ „” Õﬁ Œœ„« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   149
         Left            =   9000
         TabIndex        =   389
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ì—«œ „” Õﬁ ﬂÂ—»«¡"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   148
         Left            =   9000
         TabIndex        =   387
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ì—«œ „” Õﬁ „Ì«Â"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   147
         Left            =   9000
         TabIndex        =   385
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "”⁄Ì „” Õﬁ ··‘—ﬂ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   146
         Left            =   9000
         TabIndex        =   383
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "„’—Ê›«  Ê›Ê« Ì— «·ﬂÂ—»«¡"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   137
         Left            =   9000
         TabIndex        =   363
         Top             =   6720
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   975
         Left            =   8640
         Top             =   5640
         Width           =   2775
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄„Ê·«  „” Õﬁ… „‰ «„·«ﬂ «·€Ì—"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   121
         Left            =   8760
         TabIndex        =   325
         Top             =   6360
         Width           =   2415
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Ê”Ìÿ  √ «·«ÌÃ«—«  «·„” Õﬁ… ··€Ì—"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   120
         Left            =   8880
         TabIndex        =   323
         Top             =   6000
         Width           =   2415
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«ÌÃ«—«  «·„” Õﬁ… ··€Ì—"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   119
         Left            =   9000
         TabIndex        =   321
         Top             =   5640
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "œ›⁄«  ÕÃ“"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   91
         Left            =   9000
         TabIndex        =   260
         Top             =   5280
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ì—«œ Œœ„« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   81
         Left            =   9000
         TabIndex        =   209
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ê”Ìÿ √  √„Ì‰«  ··€Ì— ·ﬂ· «·„” √Ã—Ì‰"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   88
         Left            =   8520
         TabIndex        =   225
         Top             =   4920
         Width           =   2655
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ì—«œ«  -«·«ÌÃ«—« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   82
         Left            =   9000
         TabIndex        =   224
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·„” √Ã—Ì‰ Ê «·„‘ —Ì‰"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   48
         Left            =   8880
         TabIndex        =   216
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·„·«ﬂ"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   47
         Left            =   9480
         TabIndex        =   215
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«ÌÃ«—«  «·„” Õﬁ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   76
         Left            =   9000
         TabIndex        =   214
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«· √„Ì‰ «·„” —œ"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   77
         Left            =   9000
         TabIndex        =   213
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ì—«œ «·„Ì«Â "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   78
         Left            =   9000
         TabIndex        =   212
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ì—«œ «·ﬂÂ—»«¡ "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   79
         Left            =   9000
         TabIndex        =   211
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ì—«œ «·”⁄Ì Ê«·⁄„Ê·« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   80
         Left            =   9000
         TabIndex        =   210
         Top             =   3120
         Width           =   2055
      End
   End
   Begin VB.Frame Frame22 
      Caption         =   "«·„”«Â„« "
      Height          =   8295
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   289
      Top             =   1410
      Visible         =   0   'False
      Width           =   11535
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0958
         DataField       =   "a108"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   108
         Left            =   120
         TabIndex        =   290
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":096D
         DataField       =   "a110"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   110
         Left            =   120
         TabIndex        =   292
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0982
         DataField       =   "a111"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   111
         Left            =   120
         TabIndex        =   294
         Top             =   1200
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0997
         DataField       =   "a112"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   112
         Left            =   120
         TabIndex        =   296
         Top             =   1560
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":09AC
         DataField       =   "a121"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   121
         Left            =   120
         TabIndex        =   308
         Top             =   4920
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":09C1
         DataField       =   "a122"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   122
         Left            =   120
         TabIndex        =   310
         Top             =   5280
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":09D6
         DataField       =   "a113"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   113
         Left            =   120
         TabIndex        =   312
         Top             =   1920
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":09EB
         DataField       =   "a114"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   114
         Left            =   120
         TabIndex        =   313
         Top             =   2280
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0A00
         DataField       =   "a117"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   117
         Left            =   120
         TabIndex        =   314
         Top             =   2640
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0A15
         DataField       =   "a118"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   118
         Left            =   120
         TabIndex        =   315
         Top             =   3000
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0A2A
         DataField       =   "a119"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   119
         Left            =   120
         TabIndex        =   316
         Top             =   3360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0A3F
         DataField       =   "a120"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   120
         Left            =   120
         TabIndex        =   317
         Top             =   3720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0A54
         DataField       =   "a109"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   109
         Left            =   120
         TabIndex        =   318
         Top             =   5640
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0A69
         DataField       =   "a127"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   127
         Left            =   120
         TabIndex        =   328
         Top             =   4080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0A7E
         DataField       =   "a128"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   128
         Left            =   120
         TabIndex        =   329
         Top             =   4440
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0A93
         DataField       =   "a129"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   129
         Left            =   120
         TabIndex        =   330
         Top             =   6000
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0AA8
         DataField       =   "a130"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   130
         Left            =   120
         TabIndex        =   332
         Top             =   6360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0ABD
         DataField       =   "a131"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   131
         Left            =   120
         TabIndex        =   334
         Top             =   6840
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0AD2
         DataField       =   "a116"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   116
         Left            =   -7080
         TabIndex        =   302
         Top             =   4440
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0AE7
         DataField       =   "a115"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   115
         Left            =   -7080
         TabIndex        =   300
         Top             =   4080
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄„Ê·«  »Ì⁄ "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   125
         Left            =   9120
         TabIndex        =   335
         Top             =   6840
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Ê”Ìÿ „’«—Ì›  ÿÊÌ—"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   124
         Left            =   9120
         TabIndex        =   333
         Top             =   6360
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ì—«œ«   Œ’Ì’ «—«÷Ì"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   123
         Left            =   9120
         TabIndex        =   331
         Top             =   6000
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Ê”Ìÿ √ «—«÷Ì ."
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   105
         Left            =   9120
         TabIndex        =   319
         Top             =   5640
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Ê”Ìÿ „”«Â„«  ⁄ﬁ«—« "
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   118
         Left            =   9120
         TabIndex        =   311
         Top             =   5280
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Ê”Ìÿ „”«Â„«  «—«÷Ì"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   117
         Left            =   9120
         TabIndex        =   309
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» Œ”«∆— ⁄ﬁ«—« "
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   116
         Left            =   9120
         TabIndex        =   307
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» Œ”«∆— «—«÷Ì "
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   115
         Left            =   9120
         TabIndex        =   306
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «—»«Õ ⁄ﬁ«—« "
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   114
         Left            =   9120
         TabIndex        =   305
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «—»«Õ «—«÷Ì"
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   113
         Left            =   9120
         TabIndex        =   304
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» „‘ —Ì«  ⁄ﬁ«—« "
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   112
         Left            =   9120
         TabIndex        =   303
         Top             =   4560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» „‘ —Ì«  «—«÷Ì"
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   111
         Left            =   9120
         TabIndex        =   301
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» „»Ì⁄«  ⁄ﬁ«—« "
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   110
         Left            =   9120
         TabIndex        =   299
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» „»Ì⁄«  «—«÷Ì"
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   109
         Left            =   9120
         TabIndex        =   298
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·„”«Â„«  ⁄ﬁ«—« "
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   108
         Left            =   9120
         TabIndex        =   297
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·„”«Â„«  «—«÷Ì"
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   107
         Left            =   9120
         TabIndex        =   295
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·„”«ÂÌ‰ ⁄ﬁ«—« "
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   106
         Left            =   9120
         TabIndex        =   293
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·„”«ÂÌ‰ «—«÷Ì"
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   104
         Left            =   9120
         TabIndex        =   291
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "Õ”«»«  «·„⁄«„·«  «·„«·Ì…"
      Height          =   8295
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   94
      Top             =   1380
      Visible         =   0   'False
      Width           =   11535
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0AFC
         DataField       =   "a18"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   18
         Left            =   240
         TabIndex        =   95
         Top             =   2760
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0B11
         DataField       =   "a6"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   96
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0B26
         DataField       =   "a20"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   20
         Left            =   240
         TabIndex        =   97
         Top             =   600
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0B3B
         DataField       =   "a21"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   22
         Left            =   240
         TabIndex        =   98
         Top             =   3120
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0B50
         DataField       =   "a22"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   21
         Left            =   240
         TabIndex        =   99
         Top             =   3480
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0B65
         DataField       =   "a33"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   33
         Left            =   240
         TabIndex        =   111
         Top             =   3840
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0B7A
         DataField       =   "a34"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   34
         Left            =   240
         TabIndex        =   112
         Top             =   4200
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0B8F
         DataField       =   "a35"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   35
         Left            =   240
         TabIndex        =   113
         Top             =   4560
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0BA4
         DataField       =   "a50"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   50
         Left            =   240
         TabIndex        =   144
         Top             =   960
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0BB9
         DataField       =   "a51"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   51
         Left            =   240
         TabIndex        =   146
         Top             =   1320
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0BCE
         DataField       =   "a52"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   52
         Left            =   240
         TabIndex        =   149
         Top             =   1680
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0BE3
         DataField       =   "a63"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   63
         Left            =   240
         TabIndex        =   176
         Top             =   2400
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0BF8
         DataField       =   "a49"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   49
         Left            =   240
         TabIndex        =   187
         Top             =   4920
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0C0D
         DataField       =   "a72"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   72
         Left            =   240
         TabIndex        =   196
         Top             =   5280
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0C22
         DataField       =   "a126"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   126
         Left            =   240
         TabIndex        =   326
         Top             =   2040
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0C37
         DataField       =   "a145"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   145
         Left            =   240
         TabIndex        =   366
         Top             =   5640
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0C4C
         DataField       =   "a146"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   146
         Left            =   240
         TabIndex        =   368
         Top             =   6000
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0C61
         DataField       =   "a147"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   147
         Left            =   240
         TabIndex        =   370
         Top             =   6360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0C76
         DataField       =   "a148"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   148
         Left            =   240
         TabIndex        =   372
         Top             =   6720
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0C8B
         DataField       =   "a149"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   149
         Left            =   240
         TabIndex        =   374
         Top             =   7080
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0CA0
         DataField       =   "a150"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   150
         Left            =   240
         TabIndex        =   376
         Top             =   7440
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0CB5
         DataField       =   "a157"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   157
         Left            =   240
         TabIndex        =   390
         Top             =   7800
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ‹ ⁄„Ê·… „»Ì⁄«  Œ«—ÃÌ…"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   150
         Left            =   9600
         TabIndex        =   391
         Top             =   7800
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ‹ ⁄„Ê·… „»Ì⁄«   Õ  «· ’—Ì›"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   143
         Left            =   9600
         TabIndex        =   377
         Top             =   7440
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ‹ „»Ì⁄«   Õ  «· ’—Ì›"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   142
         Left            =   9360
         TabIndex        =   375
         Top             =   7080
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«·ﬁÌ„… «·„÷«›… ··Ã„«—ﬂ"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   141
         Left            =   9720
         TabIndex        =   373
         Top             =   6720
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «‘⁄«— œ«∆‰"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   140
         Left            =   9840
         TabIndex        =   371
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «‘⁄«— „œÌ‰"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   139
         Left            =   9840
         TabIndex        =   369
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "ÂÌ∆… «·“ﬂ«…- «·ﬁÌ„… «·„÷«›Â"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   138
         Left            =   9840
         TabIndex        =   367
         Top             =   5640
         Width           =   1335
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "‘Ìﬂ«  „— œ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   122
         Left            =   9480
         TabIndex        =   327
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã«—Ì «·›—Ê⁄"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   70
         Left            =   9960
         TabIndex        =   197
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ   «Ì—«œ«  «· ﬁ”Ìÿ"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   63
         Left            =   9525
         TabIndex        =   177
         Top             =   2565
         Width           =   1815
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ  „’—Ê›«  »‰ﬂÌ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   52
         Left            =   9525
         TabIndex        =   148
         Top             =   1725
         Width           =   1815
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ «·«⁄ „«œ«  «·„” ‰œÌÂ"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   51
         Left            =   9525
         TabIndex        =   147
         Top             =   1365
         Width           =   1815
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» ⁄„Ê·Â «·»‰Êﬂ"
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   50
         Left            =   9525
         TabIndex        =   145
         Top             =   1005
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "√ Œ «·⁄«„"
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   49
         Left            =   9960
         TabIndex        =   139
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·⁄Âœ"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   35
         Left            =   9570
         TabIndex        =   114
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«Ì—«œ« "
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   34
         Left            =   9600
         TabIndex        =   110
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„’—Ê›« "
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   33
         Left            =   9600
         TabIndex        =   109
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·»‰Êﬂ"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   20
         Left            =   9525
         TabIndex        =   104
         Top             =   645
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "›—Êﬁ«  ⁄„·…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   18
         Left            =   9555
         TabIndex        =   103
         Top             =   2940
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄Ã“ ›Ì «·‰ﬁœÌ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   21
         Left            =   9480
         TabIndex        =   101
         Top             =   3300
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·’‰œÊﬁ"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   6
         Left            =   9570
         TabIndex        =   102
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "“Ì«œ… ›Ì «·‰ﬁœÌ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   22
         Left            =   9600
         TabIndex        =   100
         Top             =   3600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame24 
      Caption         =   "Õ”«»«   «·ÕÃ Ê«·⁄„—…"
      Height          =   8295
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   345
      Top             =   1380
      Width           =   11535
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0CCA
         DataField       =   "a135"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   135
         Left            =   240
         TabIndex        =   346
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0CDF
         DataField       =   "a136"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   136
         Left            =   240
         TabIndex        =   347
         Top             =   600
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0CF4
         DataField       =   "a137"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   137
         Left            =   240
         TabIndex        =   348
         Top             =   3840
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0D09
         DataField       =   "a138"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   138
         Left            =   240
         TabIndex        =   352
         Top             =   960
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0D1E
         DataField       =   "a144"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   144
         Left            =   240
         TabIndex        =   364
         Top             =   1320
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”„Ì«  «·ÕÃ"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   68
         Left            =   9240
         TabIndex        =   365
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ì—«œ«  «·ÕÃ"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   132
         Left            =   9240
         TabIndex        =   353
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "‰ﬁ«»… «·ÕÃ"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   131
         Left            =   9240
         TabIndex        =   351
         Top             =   3840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ì—«œ«  ⁄„—…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   130
         Left            =   9120
         TabIndex        =   350
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ’„ „”«—"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   129
         Left            =   9165
         TabIndex        =   349
         Top             =   525
         Width           =   1935
      End
   End
   Begin VB.Frame Frame23 
      Caption         =   "Õ”«»«   «·„⁄«Âœ"
      Height          =   8295
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   337
      Top             =   1440
      Width           =   11535
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0D33
         DataField       =   "A132"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   132
         Left            =   240
         TabIndex        =   338
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0D48
         DataField       =   "A133"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   133
         Left            =   240
         TabIndex        =   339
         Top             =   600
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0D5D
         DataField       =   "a134"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   134
         Left            =   240
         TabIndex        =   340
         Top             =   960
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·⁄„Ê·« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   128
         Left            =   9165
         TabIndex        =   343
         Top             =   645
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "„»Ì⁄«  «·„⁄«Âœ"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   127
         Left            =   9120
         TabIndex        =   342
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» „’«—Ì› «· ”ÊÌﬁ"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   126
         Left            =   9120
         TabIndex        =   341
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Õ”«»«  «·„Œ«“‰/«·„»Ì⁄« /«·„‘ —Ì« "
      Height          =   8295
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   1350
      Width           =   11535
      Begin VB.Frame Frame4 
         Caption         =   "«·Ì… «·„Œ«“‰"
         Height          =   1575
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   240
         Width           =   1935
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "«· ”ÊÌ«  «·Ã—œÌ…"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   67
            ToolTipText     =   "Ì »⁄ «·«’Ê·"
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Œ”«∆— ›ﬁœ Ê ·›"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   66
            ToolTipText     =   "Ì »⁄ «·„’—Ê›« "
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«» «·„Œ“Ê‰"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   65
            ToolTipText     =   "Ì »⁄ «·«’Ê·"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Labelx 
            Caption         =   "Õ”«»  Âœ«Ì« Ê⁄Ì‰« "
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   17
            Left            =   360
            TabIndex        =   64
            Top             =   1200
            Width           =   1455
         End
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0D72
         DataField       =   "a0"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   45
         ToolTipText     =   "Ì »⁄ «·«’Ê·"
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0D87
         DataField       =   "a10"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   10
         Left            =   300
         TabIndex        =   46
         ToolTipText     =   "Ì »⁄ «·„’—Ê›« "
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0D9C
         DataField       =   "a11"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   11
         Left            =   240
         TabIndex        =   47
         ToolTipText     =   "Ì »⁄ «·«’Ê·"
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0DB1
         DataField       =   "a1"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   48
         ToolTipText     =   "Ì »⁄ «·„’—Ê›« "
         Top             =   1800
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0DC6
         DataField       =   "a2"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   49
         ToolTipText     =   "Ì »⁄ «·«Ì—«œ« "
         Top             =   2160
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0DDB
         DataField       =   "a3"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   50
         Top             =   2520
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0DF0
         DataField       =   "a4"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   51
         Top             =   2880
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0E05
         DataField       =   "a5"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   5
         Left            =   240
         TabIndex        =   52
         Top             =   3240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0E1A
         DataField       =   "a12"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   12
         Left            =   240
         TabIndex        =   53
         Top             =   3600
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0E2F
         DataField       =   "a13"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   13
         Left            =   240
         TabIndex        =   54
         Top             =   3960
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0E44
         DataField       =   "a17"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   17
         Left            =   240
         TabIndex        =   55
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0E59
         DataField       =   "a23"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   23
         Left            =   240
         TabIndex        =   150
         Top             =   4320
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0E6E
         DataField       =   "a75"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   75
         Left            =   240
         TabIndex        =   202
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0E83
         DataField       =   "a76"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   76
         Left            =   240
         TabIndex        =   203
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0E98
         DataField       =   "a8"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   8
         Left            =   240
         TabIndex        =   228
         Top             =   4680
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0EAD
         DataField       =   "a9"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   9
         Left            =   240
         TabIndex        =   229
         Top             =   5040
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0EC2
         DataField       =   "a96"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   96
         Left            =   240
         TabIndex        =   261
         Top             =   5400
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0ED7
         DataField       =   "a97"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   97
         Left            =   240
         TabIndex        =   264
         Top             =   5760
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0EEC
         DataField       =   "a98"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   98
         Left            =   240
         TabIndex        =   266
         Top             =   6120
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0F01
         DataField       =   "a99"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   99
         Left            =   240
         TabIndex        =   268
         Top             =   6480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0F16
         DataField       =   "a100"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   100
         Left            =   240
         TabIndex        =   270
         Top             =   6840
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0F2B
         DataField       =   "a101"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   101
         Left            =   240
         TabIndex        =   272
         Top             =   7200
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0F40
         DataField       =   "a102"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   102
         Left            =   240
         TabIndex        =   274
         Top             =   7560
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0F55
         DataField       =   "a158"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   158
         Left            =   240
         TabIndex        =   392
         Top             =   7920
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ‹ ⁄„·«¡ œ›⁄«  „ﬁœ„…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   151
         Left            =   9240
         TabIndex        =   393
         Top             =   7920
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ê«„— «·‘—«¡ œ«∆‰"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   98
         Left            =   9240
         TabIndex        =   275
         Top             =   7560
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ê«„— «·‘—«¡ „œÌ‰"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   97
         Left            =   9240
         TabIndex        =   273
         Top             =   7200
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   " ÕÊÌ·«  œ«∆‰… ‘Õ‰ "
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   96
         Left            =   9240
         TabIndex        =   271
         Top             =   6840
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   " ÕÊÌ·«  „œÌ‰… ‘Õ‰"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   95
         Left            =   9240
         TabIndex        =   269
         Top             =   6480
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁÿ⁄ „” »œ·… ·”‰œ «·Â«·ﬂ"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   94
         Left            =   9240
         TabIndex        =   267
         Top             =   6120
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "œÌÊ‰ „⁄œÊ„…"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   93
         Left            =   9240
         TabIndex        =   265
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» ⁄„Ê·«  «·„‘ —Ì« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   92
         Left            =   9240
         TabIndex        =   262
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«»   «·„Ê—œÌ‰"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   9
         Left            =   9480
         TabIndex        =   231
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«»   «·⁄„·«¡"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   8
         Left            =   9360
         TabIndex        =   230
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «Ì—«œ«  «·Œœ„« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   23
         Left            =   9240
         TabIndex        =   151
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ’„ „ﬂ ”»"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   13
         Left            =   9240
         TabIndex        =   62
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ’„ „”„ÊÕ »…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   9240
         TabIndex        =   61
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«»  ﬂ·›… «·„»Ì⁄« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   9240
         TabIndex        =   60
         ToolTipText     =   "Ì »⁄ «·„’—Ê›« "
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·„»Ì⁄« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   9240
         TabIndex        =   59
         ToolTipText     =   "Ì »⁄ «·«Ì—«œ« "
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» „—œÊœ«  «·„»Ì⁄« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   9240
         TabIndex        =   58
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·„‘ —Ì« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   9240
         TabIndex        =   57
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» „—œÊœ«  «·„‘ —Ì« "
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   9240
         TabIndex        =   56
         Top             =   3240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame18 
      Caption         =   "Õ”«»«  ﬁÿ«⁄ «·‰ﬁ·"
      Height          =   8295
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   189
      Top             =   1320
      Width           =   11535
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0F6A
         DataField       =   "a69"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   69
         Left            =   240
         TabIndex        =   190
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0F7F
         DataField       =   "a70"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   70
         Left            =   240
         TabIndex        =   191
         Top             =   600
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0F94
         DataField       =   "a71"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   71
         Left            =   240
         TabIndex        =   194
         Top             =   960
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0FA9
         DataField       =   "a160"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   160
         Left            =   240
         TabIndex        =   396
         Top             =   1320
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0FBE
         DataField       =   "a209"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   164
         Left            =   270
         TabIndex        =   406
         Top             =   1740
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0FD3
         DataField       =   "a210"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   165
         Left            =   270
         TabIndex        =   408
         Top             =   2100
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «· ›—Ì€« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   159
         Left            =   9180
         TabIndex        =   409
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ì—«œ«  «·Õ«ÊÌ« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   158
         Left            =   9240
         TabIndex        =   407
         Top             =   1740
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ’Ê„« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   153
         Left            =   9120
         TabIndex        =   397
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ì—«œ«  ⁄„Ê·« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   69
         Left            =   9120
         TabIndex        =   195
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label l 
         Alignment       =   1  'Right Justify
         Caption         =   "„’—Ê› «·œÌ“·"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   68
         Left            =   9120
         TabIndex        =   193
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "„’—Ê› „ﬂ«›√… «·”«∆ﬁÌ‰"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   67
         Left            =   9165
         TabIndex        =   192
         Top             =   645
         Width           =   2055
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   "Õ”«»«  «·«”Â„"
      Height          =   8295
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   130
      Top             =   1320
      Width           =   11535
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0FE8
         DataField       =   "a42"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   42
         Left            =   120
         TabIndex        =   131
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":0FFD
         DataField       =   "a43"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   43
         Left            =   120
         TabIndex        =   132
         Top             =   840
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":1012
         DataField       =   "a44"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   44
         Left            =   120
         TabIndex        =   133
         Top             =   1320
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranches.frx":1027
         DataField       =   "a45"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   45
         Left            =   120
         TabIndex        =   134
         Top             =   1680
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "“Ì«œ… Ê‰ﬁ’ ›Ì ﬁÌ„Â «·«”Â„ - √ Œ"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   45
         Left            =   8685
         TabIndex        =   138
         Top             =   1725
         Width           =   2415
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "„‘ —Ì«  «·«”Â„-„Ì“«‰Ì…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   43
         Left            =   9285
         TabIndex        =   136
         Top             =   885
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "“Ì«œ… Ê‰ﬁ’ ›Ì ﬁÌ„Â «·«”Â„ - „Ì“«‰Ì… "
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   42
         Left            =   8565
         TabIndex        =   135
         Top             =   1365
         Width           =   2535
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "„»Ì⁄«  «·«”Â„"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   44
         Left            =   9285
         TabIndex        =   137
         Top             =   405
         Width           =   1695
      End
   End
   Begin VB.Label LblHeader 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "  «⁄œ«œ«  «·—»ÿ „⁄  «·Õ”«»« "
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
      TabIndex        =   143
      Top             =   0
      Width           =   11955
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "«·‰‘«ÿ"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   11880
      TabIndex        =   141
      Top             =   -840
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "„œÌ— «·›—⁄"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8520
      TabIndex        =   41
      Top             =   -360
      Width           =   855
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "⁄‰Ê«‰ «·›—⁄"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6000
      TabIndex        =   40
      Top             =   -360
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Eng Name"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   -690
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·›—⁄"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6120
      TabIndex        =   30
      Top             =   -840
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   " ·Ì›Ê‰ «·›—⁄"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   11400
      TabIndex        =   31
      Top             =   -360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·›—⁄"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8280
      TabIndex        =   29
      Top             =   -840
      Width           =   1095
   End
End
Attribute VB_Name = "baranches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag_mode As String
Dim MyFormLoaded  As Boolean
Dim checksave As Boolean


Private Sub ALLButton1_Click()
    Dim rsOut As New ADODB.Recordset
    Dim Current_case As Integer
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!opt_group = True And rsOut!opt_inv_and_branch_create_account = 1 Then
   
        Else

            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "to done this process change option in system manger", vbCritical: Exit Sub
            Else
                MsgBox "·« Ì„ﬂ‰ « „«„ Â–… «·⁄„·Ì… ·«‰ﬂ «Œ —  —»ÿ «·„Œ«“‰ »«·„Ã„Ê⁄«  ›ﬁÿ ›Ì „œÌ— «·‰Ÿ«„", vbCritical: Exit Sub
            End If
        End If
    End If

    If TxtCode.Text = "" Then Exit Sub

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    sql = "Select * from groups_account_in_inventory where branch_id='" & TxtCode & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount > 0 Then
        If SystemOptions.UserInterface = EnglishInterface Then
            MsgBox "This Branch Already linked With groups", vbCritical: Exit Sub
        Else
            MsgBox " „ —»ÿ Â–« «·›—⁄ »«·„Ã„Ê⁄«  „‰ ﬁ»·", vbCritical: Exit Sub
        End If
    End If

    Rs3.Close
    sql = "Select * from Groups "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub

    For i = 1 To Rs3.RecordCount

        If create_Branch_group(TxtCode.Text, Rs3("GroupID").value, Rs3("GroupName").value) = True Then
        End If

        Rs3.MoveNext
    Next i

    Rs3.Close

    If SystemOptions.UserInterface = EnglishInterface Then
        MsgBox "Link was Done", vbInformation
    Else
        MsgBox " „ «·—»ÿ", vbInformation
    End If

End Sub

Function hide_all_frame()
    Frame24.Visible = False
    Frame10.Visible = False
    Frame11.Visible = False
    Frame12.Visible = False
    Frame13.Visible = False
    Frame14.Visible = False
    Frame9.Visible = False
    Frame15.Visible = False
    Frame16.Visible = False
    Frame17.Visible = False
    Frame18.Visible = False
        Frame20.Visible = False
        Frame21.Visible = False
        Frame22.Visible = False
        Frame23.Visible = False
        
End Function

Private Sub ALLButton10_Click()
    hide_all_frame
   Frame21.Visible = True
End Sub

Private Sub ALLButton11_Click()
    hide_all_frame
    Frame17.Visible = True
End Sub

Private Sub ALLButton12_Click()
    hide_all_frame
    Frame18.Visible = True
End Sub

Private Sub ALLButton13_Click()
   hide_all_frame
    Frame20.Visible = True
End Sub

Private Sub ALLButton14_Click()
    hide_all_frame
    Frame22.Visible = True

End Sub

Private Sub ALLButton15_Click()
   hide_all_frame
   Frame23.Visible = True
End Sub

Private Sub ALLButton16_Click()
   hide_all_frame
   Frame24.Visible = True
End Sub

Private Sub ALLButton2_Click()
    hide_all_frame
    Frame10.Visible = True

End Sub

Private Sub ALLButton3_Click()
    hide_all_frame
    Frame11.Visible = True
End Sub

Private Sub ALLButton4_Click()
    hide_all_frame
    Frame12.Visible = True
End Sub

Private Sub ALLButton5_Click()
    hide_all_frame
    Frame13.Visible = True
End Sub

Private Sub ALLButton6_Click()
    hide_all_frame
    Frame14.Visible = True
End Sub

Private Sub ALLButton7_Click()
    hide_all_frame
    Frame9.Visible = True
End Sub

Private Sub ALLButton8_Click()
    hide_all_frame
    Frame15.Visible = True
End Sub

Private Sub ALLButton9_Click()
    hide_all_frame
    Frame16.Visible = True

End Sub

Private Sub CMD_language_Click()
    On Error Resume Next

    If CMD_language.Caption = "EN" Then
        my_language = "E"
 
        'Call Reload(Me)
 
    Else
        my_language = "A"
 
        'Call Reload(Me)
    End If

End Sub

Function create_accounts(Optional ActivityTypeId As Integer) As Boolean
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset
    Dim i As Integer
    Dim StrNewAccountCode As String
    Dim namea As String
    Dim NameE As String
    Dim currency_code As String
    Dim mowazna As Boolean
    Dim cost_center As Boolean
    Set rs = New ADODB.Recordset
    Set Rs1 = New ADODB.Recordset

    rs.Open "Select * from ACCOUNTS where Sum_account=1  AND " & " ActivityTypeId =" & ActivityTypeId, Cn, adOpenStatic, adLockOptimistic, adCmdText
      If rs.RecordCount = 0 Then
              If SystemOptions.UserInterface = EnglishInterface Then
            
                  MsgBox "·«»œ „‰  ⁄—Ì› Õ”«»«   Ã„Ì⁄Ì… «Ê·« ›Ì œ·»· «·Õ”«»« ", vbCritical, "": create_accounts = False: Exit Function
              Else
                 
                  MsgBox "Must define Summary Accounts first ", vbCritical, "": create_accounts = False: Exit Function
              End If

Exit Function
End If
    rs.MoveFirst
 
    Rs1.Open "ACCOUNTS", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    For i = 1 To rs.RecordCount
        namea = rs("Account_Name").value & "  ›—⁄   " & txtnameA.Text
        NameE = rs("Account_NameEng").value & " " & TxtNameE.Text & "  Branch"
        currency_code = IIf(IsNull(rs("currenct_code").value), 1, rs("currenct_code").value)
        mowazna = IIf(IsNull(rs("mowazna").value), 0, rs("mowazna").value)
        cost_center = IIf(IsNull(rs("cost_center").value), 0, rs("cost_center").value)

        StrNewAccountCode = ModAccounts.AddNewAccount(rs("Account_Code").value, namea, 0, False, NameE, currency_code, mowazna, cost_center, False, val(TxtCode.Text))
        rs.MoveNext
    Next i

    If SystemOptions.UserInterface = EnglishInterface Then
        MsgBox "Branch Created With Accounts", vbInformation, ""
    Else
        MsgBox " „ «‰‘«¡ «·›—⁄ ÊÕ”«»« …", vbInformation, ""
    End If


    create_accounts = True
End Function

Private Sub Command1_Click(Index As Integer)
    On Error Resume Next
    
    
    
checksave = True


If Index = 4 Then Unload Me
    If Index = 0 Then
        Adodc1.Recordset.AddNew
        TxtCode.Text = CStr(new_id("branches", "branch_id", "", True))
        flag_mode = "N"

    Else

        If Index = 1 Then
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
        
            LogTextA = " Õ›Ÿ ‘«‘… " & " «⁄œ«œ«  «·—»ÿ „⁄ «·Õ”«»«  "
            LogTexte = " Save" & "   Settings link with accounts "

            AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""
     
            '     If txtnamee.text = "" Then MsgBox "write  branch name first", vbCritical: Exit Sub
      
            '     If txtnameA.text = "" Then MsgBox "«ﬂ »  «”„ «·›—⁄ «Ê·«  ", vbCritical: Exit Sub
      
            '     If DcActivityType.BoundText = "" Then
            '      MsgBox "Õœœ      ‰‘«ÿ «·›—⁄ «Ê·«  ", vbCritical
            '      DcActivityType.SetFocus
            '      SendKeys ("{F4}")
            '      Exit Sub
            '
            '     End If
    
            '   Adodc1.Recordset.Fields!inventory = DataCombo2.text
            Adodc1.Recordset.update
            Adodc1.Recordset.MoveLast
   
            If flag_mode = "N" Then
   
                If create_accounts(val(Me.DcActivityType.BoundText)) = False Then
                    Exit Sub
                End If

                flag_mode = "E"
     
            End If
 
            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "Saved", vbInformation, ""
            Else
                MsgBox " „ «·Õ›Ÿ", vbInformation, ""
            End If
  
        Else

            If Index = 2 Then
 
                Dim X As Integer

                If my_language = "E" Then
                    X = MsgBox("Confirm delete", vbCritical + vbYesNo)
                Else
                    X = MsgBox("Â· «‰  „ √ﬂœ „‰ «·Õ–›", vbCritical + vbYesNo)
              
                End If

                If X = vbNo Then
                    Exit Sub
                End If

                If Adodc1.Recordset.RecordCount > 0 Then
                    Adodc1.Recordset.delete
                    Adodc1.Refresh
                Else

                    If my_language = "E" Then
                        MsgBox "No Departement to delete", vbCritical
                    Else
                        MsgBox "·« ÌÊÃœ „« Ì„ﬂ‰ Õ–›…", vbCritical
                    End If
                
                End If

                Exit Sub

            End If
        End If
    End If

End Sub
 
Private Sub DataCombo1_Change(Index As Integer)
    On Error Resume Next

    If MyFormLoaded = True Then
        LogTextA = "    €ÌÌ— «·Õ”«» «·Œ«’ » " & Labelx(Index).Caption & " «·Ï " & DataCombo1(Index).Text & "  „‰ «⁄œ«œ«  «·—»ÿ „⁄ «·Õ”«»«  "
        LogTexte = " Change Account " & Labelx(Index).Caption & " To " & DataCombo1(Index).Text & "  From  Settings link with accounts "

        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "E", "", ""
    End If

End Sub

Private Sub DataCombo1_KeyUp(Index As Integer, _
                             KeyCode As Integer, _
                             Shift As Integer)
 
    'On Error Resume Next

    If KeyCode = vbKeyF3 Then
        Account_search.show
' Index = 113 Or Index = 114 Or Index = 115 Or Index = 116 Or Index = 117 Or Index = 118 Or Index = 119 Or Index = 120 Or Index = 121 Or Index = 122 Or
        If Index = 109 Or Index = 123 Or Index = 125 Or Index = 143 Or Index = 153 Or Index = 154 Or Index = 124 Or Index = 107 Or Index = 105 Or Index = 106 Or Index = 104 Or Index = 99 Or Index = 100 Or Index = 103 Or Index = 101 Or Index = 102 Or Index = 94 Or Index = 97 Or Index = 92 Or Index = 96 Or Index = 95 Or Index = 86 Or Index = 79 Or Index = 37 Or Index = 38 Or Index = 39 Or Index = 151 Or Index = 19 Or Index = 18 Or Index = 22 Or Index = 21 Or Index = 23 Or Index = 41 Or Index = 16 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 4 Or Index = 5 Or Index = 12 Or Index = 13 Or Index = 52 Or Index = 53 Or Index = 54 Or Index = 55 Or Index = 56 Or Index = 57 Or Index = 58 Or Index = 59 Or Index = 60 Or Index = 61 Or Index = 62 Or Index = 63 Or Index = 64 Or Index = 66 Or Index = 67 Or Index = 68 Or Index = 69 Or Index = 70 Or Index = 71 Or Index = 73 Or Index = 75 Or Index = 76 Or Index = 77 Or Index = 78 Or Index = 80 Or Index = 81 Or Index = 82 Or Index = 83 Or Index = 84 Or Index = 85 _
       Or Index = 135 Or Index = 136 Or Index = 137 Or Index = 138 Or Index = 139 Or Index = 140 Or Index = 141 Or Index = 144 Or Index = 145 Or Index = 146 Or Index = 147 Or Index = 148 Or Index = 149 Or Index = 150 Or Index = 155 Or Index = 156 Or Index = 157 Or Index = 158 Or Index = 159 Or Index = 160 Or Index = 161 Or Index = 204 Or Index = 163 Then                                            'Õ”«»«  ‰Â«∆Ì… ›ﬁÿ

            Account_search.case_id = 1700 'last Account
         ElseIf Index = 134 Or Index = 133 Or Index = 132 Or Index = 131 Or Index = 130 Or Index = 129 Or Index = 126 Or Index = 115 Or Index = 116 Or Index = 121 Or Index = 122 Or Index = 136 Or Index = 135 Or Index = 137 Or Index = 138 Or Index = 139 Or Index = 140 Or Index = 141 Or Index = 144 Or Index = 145 Or Index = 146 Or Index = 147 Or Index = 148 Or Index = 149 Or Index = 150 Or Index = 157 Or Index = 158 Then
             Account_search.case_id = 1700 'last Accoun
        Else
            Account_search.case_id = 700
        End If

        Account_search.case_index = Index
    End If

    If KeyCode = vbKeyF6 Then
'        account_index.show
    End If

    If KeyCode = vbKeyF5 Then
        Adodc2.Refresh
        DataCombo1(Index).ReFill
    End If

End Sub

'Private Sub DataCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF6 Then
'frmcities.Show
'End If
'End Sub

Private Sub Form_Activate()

End Sub

Private Sub Form_Load()
    Dim My_SQL As String
    MyFormLoaded = False
    checksave = False
    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " «⁄œ«œ«  «·—»ÿ „⁄ «·Õ”«»«  "
    LogTexte = " Open Window " & "   Settings link with accounts "

    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    My_SQL = "  select Emp_ID,Emp_Name from TblEmployee   "
    fill_combo DCEmployee, My_SQL 'On Error Resume Next

    My_SQL = "  SELECT id ,name FROM tblActivitesType order by name"
    fill_combo Me.DcActivityType, My_SQL

    hide_all_frame
   Frame17.Visible = True
If SystemOptions.StoreAccountHaveSettelment = True Then
        DataCombo1(11).Visible = False
Else

        DataCombo1(11).Visible = True
 End If


If SystemOptions.AssetAccount = True Then
        DataCombo1(26).Visible = False
Else

        DataCombo1(26).Visible = True
 End If
 
 
    If my_language = "E" Then
        CMD_language.ToolTipText = "change Language"
 
        Me.dept_lbl = departement_name
        Me.emp_name_lbl = current_user_name
        InfoE.Visible = True
        infoA.Visible = False
    Else

        emp_a.Caption = current_user_name
        dep_a.Caption = departement_name
   
        infoA.Visible = True
        InfoE.Visible = False
    End If

    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    '  On Error Resume Next
    If SystemOptions.UserInterface = EnglishInterface Then
        chag_lung
    End If

    connection_string = Cn.ConnectionString

    'Adodc5.ConnectionString = connection_string
    ' Adodc5.CommandType = adCmdText
    'Adodc5.RecordSource = "select * from cities where not(city_name is null) "
    'Adodc5.Refresh
    '

    'where  NOT (branch_name='')

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select *  from ACCOUNTS WHERE last_account=1" '
    Adodc2.Refresh

    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText

    Dim rsOut As New ADODB.Recordset
    Dim Msg As String
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!opt_group = True Then
     
            If rsOut!Opt_Inventory_create_account = 1 Then
                Adodc4.RecordSource = "select *  from ACCOUNTS WHERE last_account=1" '
            ElseIf rsOut!opt_inv_and_branch_create_account = 1 Then
                Adodc4.RecordSource = "select *  from ACCOUNTS WHERE last_account=0" '
                recolor
     
            End If

        Else
            Adodc4.RecordSource = "select *  from ACCOUNTS WHERE last_account=1" '
        End If
    End If

    Adodc4.Refresh

    Adodc6.ConnectionString = connection_string
    Adodc6.CommandType = adCmdText

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!Arrows_group = True Then
            Adodc6.RecordSource = "select *  from ACCOUNTS WHERE last_account=0" '
            recolor 1
        Else
            Adodc6.RecordSource = "select *  from ACCOUNTS WHERE last_account=1" '
        End If
    End If

    Adodc6.Refresh

    Adodc5.ConnectionString = connection_string
    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "select *  from ACCOUNTS WHERE last_account=0" '
    Adodc5.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "select * from  TblStore" '  where branch_no=" & branch_no
    Adodc3.Refresh

    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from   branches   " ' where departement_no=0"
    Adodc1.Refresh

    If SystemOptions.AssetAccount1 = True Then 'Õ”«»«  «·«—»« Õ Ê «·Œ”«∆— —∆Ì”Ì…
        DataCombo1(31).Visible = True
        DataCombo1(40).Visible = True

        DataCombo1(66).Visible = False
        DataCombo1(67).Visible = False

        Labelx(31).ForeColor = vbRed
        Labelx(40).ForeColor = vbRed

    Else

        DataCombo1(31).Visible = False
        DataCombo1(40).Visible = False
        DataCombo1(66).Visible = True
        DataCombo1(67).Visible = True

        Labelx(31).ForeColor = vbBlack
        Labelx(40).ForeColor = vbBlack

    End If
    
    
    
    
        If SystemOptions.eachStoreHaveLossAccount = True Then 'Õ”«»«     ›ﬁœ Ê ·› Õ”«» —∆Ì”Ì
        DataCombo1(10).Visible = True
    DataCombo1(75).Visible = False
     Labelx(10).ForeColor = vbRed
 
    Else
    DataCombo1(10).Visible = False
    DataCombo1(75).Visible = True
  Labelx(10).ForeColor = vbBlack
 

    End If
    
    
    
    
            If SystemOptions.eachStoreHaveGiftAccount = True Then 'Õ”«»«     Âœ«Ì« Ê⁄Ì‰«    Õ”«» —∆Ì”Ì
        DataCombo1(17).Visible = True
   DataCombo1(76).Visible = False
    Labelx(17).ForeColor = vbRed
 
    Else
    DataCombo1(17).Visible = False
    DataCombo1(76).Visible = True
  Labelx(17).ForeColor = vbBlack
 
    End If
    
    
    

    MyFormLoaded = True
    
    If SystemOptions.UserInterface = EnglishInterface Then
    DataCombo1(72).ListField = "Account_NameEng"
    
    End If
End Sub

Function recolor(Optional Index As Integer = 0)

    Select Case Index

        Case 0
            Labelx(1).ForeColor = &HFF&
            Labelx(2).ForeColor = &HFF&
            Labelx(3).ForeColor = &HFF&
            Labelx(4).ForeColor = &HFF&
            Labelx(5).ForeColor = &HFF&
            Labelx(17).ForeColor = &HFF&
            Labelx(12).ForeColor = &HFF&
            Labelx(13).ForeColor = &HFF&

        Case 1
            Labelx(42).ForeColor = &HFF&
            Labelx(43).ForeColor = &HFF&
            Labelx(44).ForeColor = &HFF&
            Labelx(45).ForeColor = &HFF&

    End Select

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If checksave = False Then

    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
 
 
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & CHR(13)
                    StrMSG = StrMSG & " the new data  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
        
                End If
  
 

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)

        Select Case IntResult

            Case vbYes
                Cancel = True
       
                Command1_Click (1)

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
 


 

End Sub

Private Sub Form_Unload(Cancel As Integer)
    LogTextA = "     «·Œ—ÊÃ „‰  ‘«‘… " & " «⁄œ«œ«  «·—»ÿ „⁄ «·Õ”«»«  "
    LogTexte = " Exit Window " & "   Settings link with accounts "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

End Sub

Private Sub txtnameA_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtNameE_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Function chag_lung()

    Labelx(71).Caption = "Projects Opening Balance  "
        Me.Caption = "Accounts Link"
        LblHeader.Caption = Me.Caption
        Labelx(35).Caption = "Covenant Acc."
        Labelx(49).Caption = "Profits and losses"
        Labelx(31).Caption = "Sale Profit"
        Labelx(40).Caption = "Sale Loss"
      Labelx(139).Caption = "Notice of the Debt"
      Labelx(140).Caption = "Notice of the Credit"
      Labelx(141).Caption = "Customs Account"
      Labelx(142).Caption = "Sales Secretariat"
      Labelx(143).Caption = "Commission Secretariat "
      
      Command1(4).Caption = "Exit"
        Label1.Caption = "Branch NO"
        Label2.Caption = "Branch Name"
        Label4.Caption = "Branch Tel"
        '    Label3.Caption = "Basic Store"
        Label14.Caption = "Address"
        Label15.Caption = "Manger"
    
        Labelx(0).Caption = "Store Account"
        Labelx(10).Caption = "Damage Account"
        Labelx(11).Caption = "Inventory adjustment"
     
        Labelx(1).Caption = "Sale cost Account"
     
        Labelx(2).Caption = "Sale Account"
      
        Labelx(3).Caption = "sale return Account"
        Labelx(4).Caption = "Purchase Account"
        Labelx(5).Caption = "purchase return Account"
       
        Labelx(6).Caption = "Box Account"
        Labelx(20).Caption = "Banks Account "
        Labelx(19).Caption = "Opening Balance "
        Labelx(7).Caption = "staff Accounts "
        Labelx(29).Caption = "Due salaries Acc."
        Labelx(16).Caption = "salaries"
            
        Labelx(30).Caption = "Apportionment"
        
        Frame10.Caption = "Store Accounts"
        Frame11.Caption = "receivables Accounts "
        Frame12.Caption = "Assets Accounts "
        Frame13.Caption = "Projects Accounts "
        Frame14.Caption = "Another Accounts "
        ALLButton2.Caption = "Store Accounts"
        ALLButton3.Caption = "receivables Accounts "
        ALLButton4.Caption = "Assets Accounts"
        ALLButton5.Caption = "Projects Accounts"
        ALLButton6.Caption = "Another Accounts"
         
        Labelx(8).Caption = "Customer Account"
        Labelx(9).Caption = "Vendor Account"
           
        Labelx(12).Caption = "Allowed discount"
        Labelx(13).Caption = "Unearned discount"
        Labelx(153).Caption = "Discount"
        Labelx(21).Caption = "Increase in cash "
        Labelx(22).Caption = "Shortfall in cash "
        Labelx(24).Caption = "Assets Account "
        Labelx(25).Caption = "Depreciation expense "
        Labelx(26).Caption = "Accumu. depreciation"
        
        Labelx(14).Caption = "Project Expanses"
        Labelx(15).Caption = "Projects Revenu"
        Labelx(27).Caption = "Project Materials"
        Labelx(28).Caption = "Projects salaries"
    
        Labelx(23).Caption = "Service revenue "
        Labelx(17).Caption = "Gifts and  Samples "
        Labelx(18).Caption = "Currency differences "
      
        ' TabControl1.Item(0).Caption = "Inventory"
         
        ALLButton1.Caption = "Link With Group"
        SetInterface Me
        'Labelx(31).Caption = "Fixed Asset"

        Labelx(32).Caption = "Legal Accounts"
        ALLButton7.Caption = "Production Acc"
        Frame9.Caption = ALLButton7.Caption
        Labelx(37).Caption = "Materials EXP"
        Labelx(38).Caption = "Salaries EXP"
        Labelx(39).Caption = "Operating  EXP"
        Labelx(144).Caption = "Depreciation"
        Labelx(145).Caption = "Good Performance"
        ALLButton9.Caption = "RealState Acc"
        Frame16.Caption = ALLButton9.Caption
        Labelx(47).Caption = "Ownere Acc"
        Labelx(48).Caption = "Buyer Acc"

        ALLButton11.Caption = "Opening Balance Acc"
        Frame17.Caption = ALLButton11.Caption
        Labelx(19).Caption = "Stock O B"
        Labelx(41).Caption = "Fixed Assets O B"
        Labelx(57).Caption = "Boxes O B"
        Labelx(58).Caption = "Banks    O B"
        Labelx(59).Caption = " Customers O B"
        Labelx(60).Caption = " Suppliers O B"
        Labelx(61).Caption = " Employees O B"

        Labelx(46).Caption = " Arrows O B"
        Labelx(62).Caption = " LC O B"
        ALLButton8.Caption = "Arrows Accounts"
        Frame15.Caption = ALLButton8.Caption
        Labelx(44).Caption = "Arrow sale"
        Labelx(43).Caption = "Arrow Purchase"
        Labelx(42).Caption = "Inc/dec Balance Sheet"
        Labelx(45).Caption = "Inc/dec Income Stat"
        Labelx(50).Caption = "Bank Comm."
        Labelx(51).Caption = "LC Acc"
        Labelx(52).Caption = "Bank Expenses"

        Labelx(53).Caption = "Discount Acc"
        Labelx(54).Caption = "bonus Acc"
        Labelx(55).Caption = "Leave entitlements"
        Labelx(56).Caption = "End of service benefits"

        Me.Left = (mdifrmmain.Width - Me.Width) / 2
        Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500

        CMD_language.Caption = "⁄—»Ì"
        ALLButton12.Caption = "Transportation Acc"
 
        ' Text1.Alignment = 0
        '  Text2.Alignment = 0
        ' DataCombo1.RightToLeft = False
  
        Frame2.Visible = False
        Frame3.Visible = True
        SuperLabel1.Text = "Branches Data"
        Me.Caption = SuperLabel1.Text
        Command1(0).Caption = "new"
        Command1(1).Caption = "save"
        Command1(2).Caption = "delete"
        Adodc1.Caption = "move"
        Frame8.Caption = "Projects"
        Frame4.Caption = "Stores"
        Frame5.Caption = "Employees"
        Frame6.Caption = "Colors"
        Label8.Caption = "Last Account"
        Label13.Caption = "Master account"
        Labelx(33).Caption = "Expenses"
        Labelx(34).Caption = "Revenues"
        Labelx(36).Caption = "Sub-contractor"
        Labelx(37).Caption = "Under Implementation"
        ALLButton13.Caption = "Cars Maintenance"
        
        ALLButton14.Caption = "Contributions Accounts"
        ALLButton10.Caption = "School Transportation"
        ALLButton15.Caption = "Schools and Institutes"
        ALLButton16.Caption = "Hijj And Omraa"
        
        
        
        Labelx(92).Caption = "Purchasing Commission Accounts"
        Labelx(93).Caption = "Bad debts"
        Labelx(94).Caption = "Replacement Parts For Deceased Voucher"
        Labelx(95).Caption = "Debt Account For Transportation Transaction"
        Labelx(96).Caption = "Credit Account For Transportation Transaction"
        Labelx(97).Caption = "Purchase Order Debt Account"
        Labelx(98).Caption = "Purchase Order Credit Account"
        
        Frame18.Caption = "Transportation Sector Accounts"
        l(68).Caption = "Diesel Expense"
        Labelx(67).Caption = "Driver Reward Expense"
        Labelx(69).Caption = "Commission income Account"
        
        Labelx(75).Caption = "Production Expenses (Electricity)"
        Labelx(66).Caption = " Production Expenses for Half the Factory"
        
        
        Labelx(73).Caption = "Maintenance Revenues Account"
        Labelx(73).Caption = "Maintenance Expenses Account"
        
        Labelx(99).Caption = "Paid Progress Bill Account"
        Labelx(100).Caption = "Advances to Customers"
        Labelx(136).Caption = "Projects Under Implementation"
        
        Labelx(72).Caption = "End of Service Benefits Allotment Account"
        Labelx(89).Caption = "Travel Tickets Allotment Account "
        Labelx(65).Caption = "Prepaid"
        Labelx(64).Caption = "Prepaid Allowances"
        Labelx(133).Caption = "End of Service Benefits Bonuses"
        Labelx(134).Caption = "End of Service Benefits Penalties"
        Labelx(135).Caption = "Vacations Without Salary"
        
        Labelx(76).Caption = "Owed Rents"
        Labelx(80).Caption = "Commissions"
        Labelx(77).Caption = "Insurance"
        Labelx(78).Caption = "Water Bill Advance"
        Labelx(79).Caption = "Electricity Bill Advance"
        Labelx(81).Caption = "Services"
        Labelx(82).Caption = "Revenues"
        Labelx(88).Caption = "Insurance Middle Account for Third Parties for All Tenants"
        Labelx(91).Caption = "Reservation Payments"
        Labelx(119).Caption = "Rents Owed to Others"
        Labelx(120).Caption = "Middle Account for Owed Rents to Others"
        Labelx(121).Caption = "Commissions for Property of Others"
        Labelx(137).Caption = "Water and Electricity Expenses"
        Labelx(138).Caption = "VAT"
        Frame22.Caption = "Contributions"
        Labelx(104).Caption = "Account of Lands Shareholders"
        Labelx(106).Caption = "Real Estate Shareholders Account"
        Labelx(107).Caption = "Account of Land Contributions"
        Labelx(108).Caption = "Real Estate Contributions Account"
        Labelx(109).Caption = "Lands Sales Account"
        Labelx(110).Caption = "Real Estate Sales Account"
        Labelx(113).Caption = "Lands Profit Account"
        Labelx(114).Caption = "Real Estate Profit Account"
        Labelx(115).Caption = "Lands Losses Account"
        Labelx(116).Caption = "Real Estate Losses Account"
        Labelx(111).Caption = "Lands Purchases Account"
        Labelx(112).Caption = "Real Estate Purchases Account"
        Labelx(117).Caption = "Middle Account for Lands Contribution"
        Labelx(118).Caption = "Middle Account for Real Estate Contribution"
        Labelx(105).Caption = "Opening Account for Lands"
        Labelx(123).Caption = "Land Allocated Revenue"
        Labelx(124).Caption = "Middle Account for Development Expenses"
        Labelx(125).Caption = "Sales Commissions"
        
        Labelx(122).Caption = "Returned Checks"
        Labelx(63).Caption = "Installment Income Account"
        Labelx(70).Caption = "Current Accounts for Brunches "
        
        Frame24.Caption = "Hijj and Omra Accounts"
        Labelx(130).Caption = "Omra Revenue"
        Labelx(129).Caption = "Route Discount"
        Labelx(131).Caption = "Hijj Association"
        Labelx(132).Caption = "Hijj Revenue"
        
        Frame23.Caption = "Institutes Accounts"
        Labelx(127).Caption = "Institutes Sales"
        Labelx(128).Caption = "Commissions Account"
        Labelx(126).Caption = "Marketing Expenses Account"
        
        Frame21.Caption = "School Transportation"
        Labelx(101).Caption = "Transportation Revenue"
        Labelx(102).Caption = "Transportation Expenses"
        Labelx(103).Caption = "Payments Owed to Contractors"
        
        Labelx(90).Caption = "Travel Ticket Entitlements Account"
        
        Frame20.Caption = "Car Maintenance Accounts"
        Labelx(74).Caption = "Maintenance Expenses Account "
End Function
