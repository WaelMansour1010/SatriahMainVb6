VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form mofradat2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„›—œ«  «·—« »"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5760
   Icon            =   "frmmofradnew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   5760
   Visible         =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.CheckBox ChkChanged 
      Alignment       =   1  'Right Justify
      Caption         =   "„ €Ì—"
      Height          =   255
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text2e 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   1320
      Width           =   3855
   End
   Begin VB.OptionButton Option5 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬁÌ„… ·ﬂ· „ÊŸ›"
      Height          =   195
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   3000
      Value           =   -1  'True
      Width           =   4035
   End
   Begin VB.Frame Frame1 
      Height          =   370
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   2385
      Visible         =   0   'False
      Width           =   2415
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         Caption         =   "«÷«›…"
         Height          =   195
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ’„"
         Height          =   195
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CheckBox Monthly 
      Alignment       =   1  'Right Justify
      Caption         =   "‘Â—Ì"
      Height          =   255
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   600
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      Caption         =   "⁄·«ﬁ… „⁄ „ﬂÊ‰ «Œ—"
      Height          =   255
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   3480
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬁÌ„… À«» … ·ﬂ· «·„ÊŸ›Ì‰"
      Height          =   195
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2220
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Frame Frame5 
      Caption         =   "ÿ—Ìﬁ… «·Õ”«»"
      Enabled         =   0   'False
      Height          =   2775
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3840
      Width           =   4935
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmd_clear 
         Caption         =   "„”Õ"
         Height          =   315
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TXT_VALUE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   720
         Width           =   495
      End
      Begin MSDataListLib.DataCombo Dcmofrdat 
         Height          =   315
         Left            =   2760
         TabIndex        =   41
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Frame Frame6 
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1935
         Begin VB.CommandButton Command1 
            Caption         =   "/"
            Height          =   315
            Index           =   3
            Left            =   480
            TabIndex        =   14
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            Caption         =   "*"
            Height          =   315
            Index           =   2
            Left            =   840
            TabIndex        =   13
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            Caption         =   "-"
            Height          =   315
            Index           =   1
            Left            =   1200
            TabIndex        =   12
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            Caption         =   "+"
            Height          =   315
            Index           =   0
            Left            =   1560
            TabIndex        =   11
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H0000FF00&
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   120
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "«Õ„«·Ì «·„ﬁ«„"
            ForeColor       =   &H000000FF&
            Height          =   15
            Left            =   0
            TabIndex        =   15
            Top             =   2520
            Width           =   1935
         End
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Õœ «ﬁ’Ï"
         Height          =   255
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Õœ «œ‰Ï"
         Height          =   255
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁÌ„Â"
         Height          =   255
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   480
         Width           =   495
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   285
         Index           =   44
         Left            =   1680
         TabIndex        =   19
         Top             =   3000
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·‰ ÌÃ…"
         Height          =   285
         Index           =   43
         Left            =   3480
         TabIndex        =   18
         Top             =   2880
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ÿ—Ìﬁ… «·Õ”«»"
         Height          =   285
         Index           =   42
         Left            =   2880
         TabIndex        =   17
         Top             =   480
         Width           =   1875
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "ÌŒ÷⁄ ·· √„Ì‰« "
      Height          =   375
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   0
      Left            =   1185
      TabIndex        =   23
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "frmmofradnew.frx":000C
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "frmmofradnew.frx":03A6
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   1
      Left            =   1710
      TabIndex        =   25
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "frmmofradnew.frx":0740
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   3
      Left            =   645
      TabIndex        =   26
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "frmmofradnew.frx":0ADA
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   27
      Top             =   7515
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÃœÌœ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   3825
      TabIndex        =   28
      Top             =   7515
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " ⁄œÌ·"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   2
      Left            =   3090
      TabIndex        =   29
      Top             =   7515
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ›Ÿ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   3
      Left            =   2355
      TabIndex        =   30
      Top             =   7515
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " —«Ã⁄"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   4
      Left            =   1620
      TabIndex        =   31
      Top             =   7515
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ–›"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   6
      Left            =   0
      TabIndex        =   32
      Top             =   7515
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton CmdHelp 
      Height          =   345
      Left            =   735
      TabIndex        =   33
      Top             =   7530
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   609
      ButtonStyle     =   1
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo Dcmofrdat1 
      Height          =   315
      Left            =   360
      TabIndex        =   45
      Top             =   1860
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCAccounts 
      Height          =   315
      Left            =   0
      TabIndex        =   57
      Top             =   6600
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«·«”„ «‰Ã·Ì“Ì"
      Height          =   255
      Index           =   3
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Õ”«» «·—»ÿ"
      Height          =   255
      Index           =   2
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ì »⁄"
      Height          =   255
      Index           =   0
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "%"
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ «·„›—œ"
      Height          =   255
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   2820
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   7080
      Width           =   825
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   870
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   7080
      Width           =   705
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   4
      Left            =   1620
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   7080
      Width           =   1035
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   2
      Left            =   3780
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   7080
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "„›—œ«  «·—« »"
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
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   0
      Width           =   5715
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "‰”»… «· √„Ì‰« "
      Height          =   255
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "ÿ—Ìﬁ… Õ”«» «·„›—œ"
      Height          =   255
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«·«”„ ⁄—»Ì"
      Height          =   255
      Index           =   1
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "mofradat2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim Dcombos As ClsDataCombos

Private Sub Check2_Click()

    If Check2.value = vbChecked Then
        Text3.Enabled = True
    Else
        Text3.Enabled = False
        Text3.Text = ""
    End If

End Sub

Private Sub cmd_clear_Click()
    Text14.Text = ""
    Text15.Text = ""
End Sub

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim Msg As String

    Select Case Index

        Case 0
            ' If DoPremis(Do_New, Me.name, True) = False Then
            '     Exit Sub
            ' End If
            TxtModFlg.Text = "N"
            clear_all Me
            Frame5.Enabled = True
            Me.TxtSerial.Text = CStr(new_id("mofrdat", "mofrad_code", "", True))
            ' XPTxtStoreName.SetFocus
            Me.Monthly.value = vbChecked
            Me.ChkChanged.value = vbUnchecked
         
Option5_Click
   Option5.value = True
        Case 1
            ' If DoPremis(Do_Edit, Me.name, True) = False Then
            '     Exit Sub
            ' End If
            '   If TxtSerial.text = 1 Then
            '    Msg = "·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰«  Â–« «·”Ã·"
            '       MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '       Exit Sub
            '   End If
            TxtModFlg.Text = "E"
            Frame5.Enabled = True

        Case 2
            SaveData

        Case 3
            Undo

        Case 4
            ' If DoPremis(Do_Delete, Me.name, True) = False Then
            '     Exit Sub
            ' End If
            ' If TxtSerial.text = 1 Then
            '     Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·”Ã·"
            '     MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '     Exit Sub
            ' End If
            Del_Company

        Case 5

        Case 6
            Unload Me
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub Command1_Click(Index As Integer)

    If IsNumeric(TXT_VALUE.Text) Then

        If Command1(Index).Caption <> "=" Then

            If Text14.Text <> "" Then
                Text14.Text = Text14.Text & TXT_VALUE.Text & Command1(Index).Caption
                Text15.Text = Text15.Text & TXT_VALUE.Text & Command1(Index).Caption
            Else
                Text14.Text = TXT_VALUE.Text & Command1(Index).Caption
                Text15.Text = TXT_VALUE.Text & Command1(Index).Caption
            End If

        Else

            If Text14.Text <> "" Then
                Text14.Text = Text14.Text & TXT_VALUE.Text & Command1(Index).Caption
                Text15.Text = Text15.Text & TXT_VALUE.Text
            Else
                Text14.Text = TXT_VALUE.Text & Command1(Index).Caption
                Text15.Text = TXT_VALUE.Text
            End If

            '  If Text14.text <> "" Then
            '  Text14.text = Text14.text & TXT_VALUE.text
            '  Text15.text = Text15.text & TXT_VALUE.text
            '  Else
            '  Text14.text = ""
            '  Text15.text = ""
            '  End If
        End If

        TXT_VALUE.Text = ""
        Exit Sub
    End If

    If Command1(Index).Caption <> "=" Then

        If Text14.Text <> "" Then
            Text14.Text = Text14.Text & "A" & Dcmofrdat.BoundText & Command1(Index).Caption
            Text15.Text = Text15.Text & Dcmofrdat.Text & Command1(Index).Caption
        Else
            Text14.Text = "A" & Dcmofrdat.BoundText & Command1(Index).Caption
            Text15.Text = Dcmofrdat.Text & Command1(Index).Caption
        End If

    Else

        If Text14.Text <> "" Then
            Text14.Text = Text14.Text & "A" & Dcmofrdat.BoundText
            Text15.Text = Text14.Text & Dcmofrdat.Text
        Else
            Text14.Text = ""
            Text15.Text = ""
        End If
    End If

End Sub

Private Sub Form_Activate()
    'XPTxtStoreID.SetFocus
End Sub

Private Sub ChangeLang()
    Me.Caption = "SAL Components"
    Label2(0).Caption = "Type"
    Frame5.Caption = "Calculation Method"
    Label4.Caption = Me.Caption
    cmd_clear.Caption = "Clear"
    Label2(2).Caption = "Account"
    Option5.Caption = "Fixed Value"
    Label1.Caption = "code"
    Label2(1).Caption = "Name AR"
    Label2(3).Caption = "Name Eng"

    Check2.Caption = "Insurance"
    Me.Monthly.Caption = "Monthly"

    ChkChanged.Caption = "Changed"

    Label6.Caption = "Percentage"
    Label7.Caption = "Min"
    Label8.Caption = "Max"
    Label3.Caption = "Method Of Cal."
    Option1.Caption = "Fixed Value For All"
    Option2.Caption = "Equation"
    XPLbl(42).Caption = "select"
    Label5.Caption = "Value"

    XPLbl(43).Caption = "Result"
    lbl(2).Caption = "curr. rec."
    lbl(4).Caption = "rec. Count."
    Option4.Caption = "Addition"
    Option3.Caption = "Deduction"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    'Cmd(5).Caption = "Search"

    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

End Sub

Private Sub Form_Load()
    'On Error GoTo ErrTrap
    Dim My_SQL As String
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture

    Resize_Form Me
    Set rs = New ADODB.Recordset

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  id,name  from mofrad  where  FixedOrChanged<>1 "
    Else
        My_SQL = "  select  id,namee  from mofrad where  FixedOrChanged<>1 "
    End If
    My_SQL = My_SQL & " and  (ViewComp=1 or AllowIntrod=1)"

    fill_combo Dcmofrdat1, My_SQL

    rs.Open "[mofrdat]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"
    AddTip

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetAccountingCodes Me.DCAccounts

    My_SQL = "  select  mofrad_code,mofrad_name  from mofrdat  "

    fill_combo Dcmofrdat, My_SQL

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
   
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub Option1_Click()

    If Me.Option1.value = True Then
        Frame5.Visible = False
        Text6.Enabled = True
        cmd_clear_Click
        Text4.Text = ""
        Text5.Text = ""
        Text4.Enabled = False
        Text5.Enabled = False
    Else
        Text6.Enabled = False
        Text6.Text = ""
        Frame5.Visible = True
   
        Text4.Enabled = True
        Text5.Enabled = True
    
    End If

End Sub

Private Sub Option2_Click()

    If Me.Option2.value = True Then
        Frame5.Visible = True
        Text6.Text = ""
        Text6.Enabled = False
        Text4.Enabled = True
        Text5.Enabled = True
    Else
  
        Frame5.Visible = False
    End If

End Sub

Private Sub Option5_Click()
 
    Frame5.Visible = False
    Text6.Enabled = True
    cmd_clear_Click
    Text4.Text = ""
    Text5.Text = ""
    Text4.Enabled = False
    Text5.Enabled = False
 
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
       
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«‰Ê«⁄ «·„ﬁ—œ« "
            Else
                Me.Caption = "Salary component types"
            End If
        
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            '   Me.XPTxtStoreID.locked = True
            '   Me.XPTxtStoreName.locked = True
            '   Me.XPTxtStoreAddress.locked = True
            '   Me.XPTxtStorePhone.locked = True
            '   Me.XPMTxtRemark.locked = True
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

        Case "N"
        
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«‰Ê«⁄ «·„ﬁ—œ«  (ÃœÌœ)"
            Else
                Me.Caption = "Salary component types (New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            '     Me.XPBtnMove(0).Enabled = False
            '     Me.XPBtnMove(1).Enabled = False
            '     Me.XPBtnMove(2).Enabled = False
            '     Me.XPBtnMove(3).Enabled = False
        
            'Me.XPTxtStoreID.locked = True
            'Me.XPTxtStoreName.locked = False
            'Me.XPTxtStoreAddress.locked = False
            'Me.XPTxtStorePhone.locked = False
            'Me.XPMTxtRemark.locked = False
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«‰Ê«⁄ «·„ﬁ—œ«  ( ⁄œÌ·)"
            Else
                Me.Caption = "Salary component types (Edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            'Me.XPTxtStoreID.locked = True
            'Me.XPTxtStoreName.locked = False
            'Me.XPTxtStoreAddress.locked = False
            'Me.XPTxtStorePhone.locked = False
            'Me.XPMTxtRemark.locked = False
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    'If Lngid <> 0 Then
    '    Rs.find "StoreID=" & Lngid, , adSearchForward, adBookmarkFirst
    '    If Rs.BOF Or Rs.EOF Then
    '        Exit Sub
    '    End If
    'End If
    Me.TxtSerial.Text = IIf(IsNull(rs.Fields("mofrad_code").value), "", rs.Fields("mofrad_code").value)
    Text2.Text = IIf(IsNull(rs.Fields("mofrad_name").value), "", rs.Fields("mofrad_name").value)
    Text2e.Text = IIf(IsNull(rs.Fields("mofrad_namee").value), "", rs.Fields("mofrad_namee").value)

    Text3.Text = IIf(IsNull(rs.Fields("percentage").value), "", rs.Fields("percentage").value)
    Text4.Text = IIf(IsNull(rs.Fields("min_val").value), "", rs.Fields("min_val").value)
    Text5.Text = IIf(IsNull(rs.Fields("max_val").value), "", rs.Fields("max_val").value)

    Text14.Text = IIf(IsNull(rs.Fields("eq_text").value), "", rs.Fields("eq_text").value)
    Text15.Text = IIf(IsNull(rs.Fields("eq_sys").value), "", rs.Fields("eq_sys").value)
    Me.Check2.value = IIf(rs("assurance").value = True, vbChecked, vbUnchecked)
    Me.Monthly.value = IIf(rs("Monthly").value = True, vbChecked, vbUnchecked)
    Me.ChkChanged.value = IIf(rs("Changed").value = True, vbChecked, vbUnchecked)

    Text3.Text = IIf(IsNull(rs.Fields("percentage").value), "", rs.Fields("percentage").value)
    'Me.Check1.value = IIf(Rs("for_all").value = True, vbChecked, vbUnchecked)
    Text6.Text = IIf(IsNull(rs.Fields("specific_value").value), "", rs.Fields("specific_value").value)
 
    Dcmofrdat1.BoundText = IIf(IsNull(rs("mofrad_type").value), "", rs("mofrad_type").value)
    Me.DCAccounts.BoundText = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)

    If rs("is_fixed").value = 1 Then
        Option1.value = True
    ElseIf rs("is_fixed").value = 0 Then
          
        Option2.value = True
    ElseIf rs("is_fixed").value = 2 Then
        Option5.value = True
    End If
        
    If Me.Option1.value = True Then
        Frame5.Visible = False
    Else
  
        Frame5.Visible = True
    End If
        
    Text6.Text = IIf(IsNull(rs.Fields("specific_value").value), "", rs.Fields("specific_value").value)

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case Index

        Case 0

            If Not (rs.EOF Or rs.BOF) Then
                rs.MovePrevious

                If rs.BOF Then rs.MoveFirst
            End If

        Case 1

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
            End If

        Case 2

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If

        Case 3

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveNext

                If rs.EOF Then rs.MoveLast
            End If

    End Select

    Retrive
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean

    If Me.TxtModFlg.Text <> "R" Then
        If Text2.Text = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·„›—œ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Text2.SetFocus
            Exit Sub
        End If
 
        If val(Dcmofrdat1.BoundText) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "„‰ ›÷·ﬂ Õœœ «·„›—œ  «»⁄ «·Ì ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Plz Specify Component Type ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            
            Dcmofrdat1.SetFocus
            SendKeys ("{F4}")
            Exit Sub
            
        End If
 
    End If

    Select Case Me.TxtModFlg.Text

        Case "N"
            StrSQL = "select * From  mofrdat where mofrad_name='" & Trim(Text2.Text) & "'"
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If RsTemp.RecordCount > 0 Then
                Msg = "ÌÊÃœ „›—œ „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„›—œ"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Text2.SetFocus
                Exit Sub
            End If

        Case "E"
            StrSQL = "select * From  mofrdat where mofrad_name='" & Trim(Text2.Text) & "'"
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If RsTemp.RecordCount > 0 Then
                If RsTemp("mofrad_code").value <> val(Me.TxtSerial.Text) Then
                    Msg = "ÌÊÃœ „›—œ „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„›—œ"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Text2.SetFocus
                    Exit Sub
                End If
            End If

    End Select
    
    Select Case Me.TxtModFlg.Text

        Case "N"
            rs.AddNew
            rs.Fields("mofrad_code").value = Me.TxtSerial.Text
 
    End Select

    Cn.BeginTrans
    BeginTrans = True
    rs.Fields("mofrad_name").value = Me.Text2.Text
    rs.Fields("mofrad_namee").value = Me.Text2e.Text
            
    If Check2.value = vbChecked Then
        rs.Fields("assurance").value = 1
        rs.Fields("percentage").value = IIf(IsNumeric(Text3.Text), val(Text3.Text), 0)
            
    Else
            
        rs.Fields("assurance").value = 0
        rs.Fields("percentage").value = Null
    End If
            
    If Monthly.value = vbChecked Then
        rs.Fields("Monthly").value = 1
    Else
        rs.Fields("Monthly").value = 0
    End If
            
    If ChkChanged.value = vbChecked Then
        rs.Fields("Changed").value = 1
    Else
        rs.Fields("Changed").value = 0
    End If
            
    rs.Fields("min_val").value = IIf(Text4.Text <> "", Trim(Text4.Text), Null)
            
    rs.Fields("max_val").value = IIf(Text5.Text <> "", Trim(Text5.Text), Null)
            
    If Me.Option1.value = True Then
        rs.Fields("is_fixed").value = 1
    ElseIf Me.Option2.value = True Then
        rs.Fields("is_fixed").value = 0
    ElseIf Me.Option5.value = True Then
        rs.Fields("is_fixed").value = 2
    End If
            
    rs.Fields("specific_value").value = IIf(Text6.Text <> "", Trim(Text6.Text), Null)
    rs.Fields("eq_text").value = IIf(Text14.Text <> "", Trim(Text14.Text), Null)
            
    rs.Fields("eq_sys").value = IIf(Text15.Text <> "", Trim(Text15.Text), Null)
        
    If val(Dcmofrdat1.BoundText) = 0 Then
        rs.Fields("mofrad_type").value = Null
    Else
        rs.Fields("mofrad_type").value = val(Dcmofrdat1.BoundText)
    End If
       
    rs.Fields("Account_code").value = IIf(Me.DCAccounts.BoundText = "", Null, Me.DCAccounts.BoundText)

    rs.Update
    Cn.CommitTrans
    BeginTrans = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
                
    Select Case Me.TxtModFlg.Text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then

                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„›—œ" & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " This Project Data Was Saved" & Chr(13)
                Msg = Msg + "Do you want To enter Another Component"
            End If

            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If
            
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Amendments have been saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If

    End Select

    TxtModFlg.Text = "R"
    Dim My_SQL As String
    My_SQL = "  select  mofrad_code,mofrad_name  from mofrdat  "

    fill_combo Dcmofrdat, My_SQL

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If SystemOptions.UserInterface = EnglishInterface Then
        Msg = "Error During saving  """
    Else
        Msg = "ÕœÀ Œÿ√ «À‰«¡  Õ›Ÿ «·»Ì«‰« """
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Function create_accounts(inv_id As Integer) As Boolean
    Dim rsOut As New ADODB.Recordset
    Dim Current_case As Integer
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!opt_group = False Then
            create_accounts = False: Exit Function
        End If

    Else
        create_accounts = False: Exit Function
    End If

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    sql = "Select * from Groups where not(ParentID is null)"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function

    For i = 1 To Rs3.RecordCount

        If create_inventory_group(inv_id, Rs3("GroupID").value, Rs3("GroupName").value) = True Then
        End If

        Rs3.MoveNext
    Next i

    Rs3.Close
    create_accounts = True
End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "mofrad_code='" & val(Me.TxtSerial.Text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If Me.TxtSerial.Text <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  «·„›—œ —ﬁ„ " & Chr(13)
        Msg = Msg + (Me.TxtSerial.Text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    On Error GoTo ErrTrap
    Dim Wrap As String
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„›—œ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  „›—œ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„›—œ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·„›—œ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„›—œ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·„›—œ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„›—œ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„›—œ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·„›—œ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„›—œ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ „›—œ" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„›—œ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„›—œ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„›—œ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„›—œ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„›—œ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„›—œ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.Text = "R" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
            XPBtnMove_Click (2)
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
            XPBtnMove_Click (1)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
            XPBtnMove_Click (3)
        ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
            XPBtnMove_Click (0)
        End If
    End If

    If KeyCode = vbKeyF12 Then
        If Cmd(0).Enabled = False Then Exit Sub
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If Cmd(1).Enabled = False Then Exit Sub
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If Cmd(2).Enabled = False Then Exit Sub
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If Cmd(3).Enabled = False Then Exit Sub
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then

        Select Case Me.TxtModFlg.Text

            Case "N"
    
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                
                End If

        End Select

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)

        Select Case IntResult

            Case vbYes
                Cancel = True
       
                'SaveData
                '        btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

