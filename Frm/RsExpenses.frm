VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form RsExpenses 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·„’—Ê›« "
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12810
   HelpContextID   =   280
   Icon            =   "RsExpenses.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   12810
   Begin VB.Frame FramePay 
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·„»·€ «·„œ›Ê⁄"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   138
      Top             =   600
      Visible         =   0   'False
      Width           =   12855
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFFF&
         Height          =   5055
         Left            =   120
         TabIndex        =   157
         Top             =   480
         Width           =   5535
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   0
            Left            =   4320
            TabIndex        =   158
            Top             =   3970
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
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
            ButtonImage     =   "RsExpenses.frx":038A
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   1
            Left            =   2160
            TabIndex        =   159
            Top             =   3000
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsExpenses.frx":0B4A
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   2
            Left            =   3240
            TabIndex        =   160
            Top             =   3000
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsExpenses.frx":114C
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   3
            Left            =   4320
            TabIndex        =   161
            Top             =   3000
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsExpenses.frx":1933
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   4
            Left            =   2160
            TabIndex        =   162
            Top             =   2040
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsExpenses.frx":2148
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   5
            Left            =   3240
            TabIndex        =   163
            Top             =   2040
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsExpenses.frx":28D3
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   6
            Left            =   4320
            TabIndex        =   164
            Top             =   2040
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsExpenses.frx":3092
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   7
            Left            =   2160
            TabIndex        =   165
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsExpenses.frx":382C
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   8
            Left            =   3240
            TabIndex        =   166
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsExpenses.frx":3F2F
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   9
            Left            =   4320
            TabIndex        =   167
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsExpenses.frx":474A
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   10
            Left            =   3240
            TabIndex        =   168
            Top             =   3970
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsExpenses.frx":4ED9
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   11
            Left            =   2160
            TabIndex        =   169
            Top             =   3970
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsExpenses.frx":5A20
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   12
            Left            =   120
            TabIndex        =   170
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsExpenses.frx":5F12
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   975
            Index           =   13
            Left            =   1200
            TabIndex        =   171
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsExpenses.frx":6779
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton CmdNos 
            Height          =   2895
            Index           =   14
            Left            =   120
            TabIndex        =   172
            Top             =   2040
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   5106
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsExpenses.frx":6E8A
            ButtonImageDisabled=   "RsExpenses.frx":8238
            ColorButton     =   16777215
         End
         Begin VB.Label LBLPayVal 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   1200
            TabIndex        =   173
            Top             =   360
            Width           =   3375
         End
         Begin VB.Image Image13 
            Height          =   1035
            Left            =   120
            Picture         =   "RsExpenses.frx":85D3
            Stretch         =   -1  'True
            Top             =   120
            Width           =   5295
         End
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "1500"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   3000
         TabIndex        =   156
         Top             =   7320
         Width           =   1215
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "2000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   4200
         TabIndex        =   155
         Top             =   7320
         Width           =   1335
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   1935
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   147
         Top             =   4440
         Width           =   7080
         Begin VB.CommandButton Command1 
            Caption         =   "⁄—÷ «·ﬂ·"
            Height          =   375
            Left            =   5280
            TabIndex        =   151
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox TxtNetValue2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   600
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   150
            Top             =   240
            Width           =   2460
         End
         Begin VB.TextBox TxtPayedValue2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   555
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   149
            Top             =   840
            Width           =   2445
         End
         Begin VB.TextBox TxtRemainValue2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   555
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   148
            Top             =   1320
            Width           =   2445
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·«Ã„«·Ì"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   101
            Left            =   3600
            TabIndex        =   154
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„œ›Ê⁄"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   100
            Left            =   3600
            TabIndex        =   153
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ »ﬁÌ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   99
            Left            =   3600
            TabIndex        =   152
            Top             =   1440
            Width           =   855
         End
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   1560
         TabIndex        =   146
         Top             =   7320
         Width           =   1455
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "500"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   240
         TabIndex        =   145
         Top             =   7320
         Width           =   1335
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   4200
         TabIndex        =   144
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   3000
         TabIndex        =   143
         Top             =   6720
         Width           =   1215
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   1560
         TabIndex        =   142
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   141
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   8640
         TabIndex        =   140
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton CmdValue 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   5760
         TabIndex        =   139
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin ImpulseButton.ISButton CMDPAy 
         Height          =   1215
         Left            =   240
         TabIndex        =   174
         Top             =   5450
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2143
         Caption         =   "”œ«œ"
         ForeColor       =   16777215
         FontSize        =   24
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "RsExpenses.frx":8989
         ColorHoverText  =   16777215
         ColorToggledText=   16777215
         ColorToggledHoverText=   16777215
         AlignmentIgnoreImage=   -1  'True
      End
      Begin VSFlex8UCtl.VSFlexGrid Grid22 
         Height          =   3885
         Left            =   5760
         TabIndex        =   175
         Top             =   600
         Width           =   6885
         _cx             =   12144
         _cy             =   6853
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483640
         ForeColor       =   65280
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483641
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483640
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   650
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"RsExpenses.frx":8F03
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblexit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   90
         Left            =   9120
         TabIndex        =   177
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Height          =   495
         Left            =   10440
         TabIndex        =   176
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame20 
      Height          =   2055
      Left            =   0
      TabIndex        =   126
      Top             =   4890
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox TxtCurrentBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   131
         Top             =   480
         Width           =   2115
      End
      Begin VB.TextBox TxtPaymentValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   130
         Top             =   840
         Width           =   2115
      End
      Begin VB.TextBox TxtPercentage 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   129
         Top             =   1200
         Width           =   1995
      End
      Begin VB.TextBox TxtPercentageValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   128
         Top             =   1560
         Width           =   2115
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000FF&
         Caption         =   "X"
         Height          =   255
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·—’Ìœ «·Õ«·Ì"
         Height          =   285
         Index           =   59
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   137
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„»·€ «·”œ«œ"
         Height          =   285
         Index           =   60
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   136
         Top             =   840
         Width           =   1635
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «·Œ’„"
         Height          =   285
         Index           =   61
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   135
         Top             =   1200
         Width           =   1635
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·Œ’„"
         Height          =   285
         Index           =   62
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   134
         Top             =   1560
         Width           =   1635
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   24
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   133
         Top             =   1200
         Width           =   195
      End
      Begin VB.Label Label64 
         Alignment       =   2  'Center
         Caption         =   "”Ì«”…  ⁄ÃÌ· «·œ›⁄"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   390
         TabIndex        =   132
         Top             =   150
         Width           =   3255
      End
   End
   Begin VB.CommandButton CMDSENDSMS 
      Caption         =   "«—”«· —”«·Â"
      Height          =   375
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   117
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   4215
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   720
      Width           =   12855
      Begin VB.TextBox TxtRemarks 
         Alignment       =   1  'Right Justify
         Height          =   555
         Left            =   1800
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   124
         Top             =   3600
         Width           =   7095
      End
      Begin VB.TextBox TxtValue 
         Alignment       =   1  'Right Justify
         Height          =   555
         Left            =   9960
         RightToLeft     =   -1  'True
         TabIndex        =   122
         Top             =   3600
         Width           =   1935
      End
      Begin VB.TextBox TxtSearch2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   118
         Top             =   600
         Width           =   825
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·› —Â"
         Height          =   615
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   103
         Top             =   3240
         Visible         =   0   'False
         Width           =   6015
         Begin Dynamic_Byte.NourHijriCal FrmPriodDateH 
            Height          =   315
            Left            =   3120
            TabIndex        =   104
            Top             =   240
            Width           =   1215
            _extentx        =   2143
            _extenty        =   556
         End
         Begin MSComCtl2.DTPicker FrmPriodDate 
            Height          =   315
            Left            =   4350
            TabIndex        =   105
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   99155969
            CurrentDate     =   41640
         End
         Begin Dynamic_Byte.NourHijriCal ToPriodDateH 
            Height          =   315
            Left            =   120
            TabIndex        =   106
            Top             =   240
            Width           =   1215
            _extentx        =   2143
            _extenty        =   556
         End
         Begin MSComCtl2.DTPicker ToPriodDate 
            Height          =   315
            Left            =   1350
            TabIndex        =   107
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   99155969
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   285
            Index           =   63
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   240
            Width           =   285
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   285
            Index           =   64
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   240
            Width           =   285
         End
      End
      Begin VB.TextBox TxtSearch 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   5160
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   94
         Top             =   3000
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox XPMTxtRemarks 
         Alignment       =   1  'Right Justify
         Height          =   645
         Left            =   240
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   80
         Top             =   1350
         Width           =   5835
      End
      Begin VB.TextBox XPTxtID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   150
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   150
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Text            =   "Text1"
         Top             =   990
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtSerial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   120
         Width           =   1695
      End
      Begin VB.ComboBox CboPaymentType 
         Height          =   315
         ItemData        =   "RsExpenses.frx":90E9
         Left            =   7200
         List            =   "RsExpenses.frx":90EB
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   600
         Width           =   4335
      End
      Begin VB.Frame FraNote 
         BackColor       =   &H00E2E9E9&
         Height          =   2205
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   990
         Width           =   5835
         Begin VB.CheckBox chkvat 
            Alignment       =   1  'Right Justify
            Caption         =   "«·Õ”«» Œ«÷⁄ ··÷—Ì»Â"
            Height          =   255
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   178
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox TxtAccount 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3570
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   1560
            Width           =   705
         End
         Begin VB.TextBox TxtChequeNumber 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   840
            Width           =   4035
         End
         Begin MSComCtl2.DTPicker DtpChequeDueDate 
            Height          =   315
            Left            =   240
            TabIndex        =   67
            Top             =   1140
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   556
            _Version        =   393216
            Format          =   99155969
            CurrentDate     =   39614
         End
         Begin MSDataListLib.DataCombo DcboBankName 
            Height          =   315
            Left            =   240
            TabIndex        =   68
            Top             =   480
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   240
            TabIndex        =   69
            Top             =   120
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbAccount 
            Height          =   315
            Left            =   240
            TabIndex        =   116
            Top             =   1560
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Õ”«»"
            Height          =   285
            Index           =   91
            Left            =   4350
            TabIndex        =   115
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·Œ“‰…"
            Height          =   285
            Index           =   16
            Left            =   4350
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·»‰ﬂ"
            Height          =   285
            Index           =   17
            Left            =   4350
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   510
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·‘Ìﬂ"
            Height          =   285
            Index           =   18
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
            Height          =   285
            Index           =   19
            Left            =   4380
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   1140
            Width           =   1215
         End
      End
      Begin VB.TextBox txtto 
         Alignment       =   1  'Right Justify
         Height          =   765
         Left            =   240
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
         Top             =   2040
         Width           =   5835
      End
      Begin VB.TextBox TxtSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txt_general_des 
         Alignment       =   1  'Right Justify
         Height          =   645
         Left            =   240
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         Top             =   2910
         Width           =   5835
      End
      Begin VB.TextBox txt_ORDER_NO 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   13200
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   2790
         Width           =   2775
      End
      Begin VB.TextBox TXT_A_NoteID 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Text            =   "Text2"
         Top             =   3390
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   5100
         TabIndex        =   81
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   99155969
         CurrentDate     =   38784
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   7
         Left            =   30
         TabIndex        =   82
         Top             =   30
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "«·⁄—÷ «·ÃœÊ·Ï"
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
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin MSDataListLib.DataCombo dcproject 
         Height          =   315
         Left            =   240
         TabIndex        =   83
         Top             =   1110
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcCostCenter 
         Bindings        =   "RsExpenses.frx":90ED
         Height          =   315
         Left            =   360
         TabIndex        =   84
         Top             =   630
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
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
      Begin MSDataListLib.DataCombo DcbIqara 
         Height          =   315
         Left            =   240
         TabIndex        =   95
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
         Top             =   3000
         Visible         =   0   'False
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbUnitNo 
         Height          =   315
         Left            =   2160
         TabIndex        =   96
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbUnitType 
         Height          =   315
         Left            =   4680
         TabIndex        =   97
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcsupplier 
         Height          =   315
         Left            =   240
         TabIndex        =   98
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   960
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Height          =   315
         Left            =   1800
         TabIndex        =   111
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   120
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbIqara2 
         Height          =   315
         Left            =   240
         TabIndex        =   119
         Top             =   600
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ALLButtonS.ALLButton ALLButton2 
         Height          =   315
         Left            =   240
         TabIndex        =   121
         Tag             =   "Delete Row"
         Top             =   3720
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "≈÷«›… "
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
         BCOL            =   16776960
         BCOLO           =   16776960
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "RsExpenses.frx":9102
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·‘—Õ"
         Height          =   285
         Index           =   23
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   125
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁÌ„…"
         Height          =   285
         Index           =   22
         Left            =   11280
         RightToLeft     =   -1  'True
         TabIndex        =   123
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄ﬁ«—"
         Height          =   285
         Index           =   48
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   120
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   14
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   110
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·ÊÕœ…"
         Height          =   195
         Index           =   2
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   102
         Top             =   3360
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄ﬁ«—"
         Height          =   195
         Index           =   4
         Left            =   5865
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   3000
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·ÊÕœ…"
         Height          =   195
         Index           =   15
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   100
         Top             =   3360
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «·„«·ﬂ"
         Height          =   165
         Index           =   1
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   99
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·«„—"
         Height          =   285
         Index           =   5
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   1590
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·”‰œ"
         Height          =   285
         Index           =   4
         Left            =   11400
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„’—Ê›« "
         Height          =   285
         Index           =   3
         Left            =   14520
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   1110
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   1
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   135
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÌœ"
         Height          =   255
         Index           =   0
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   150
         Width           =   975
      End
      Begin VB.Image ImgNote 
         Height          =   240
         Left            =   0
         Picture         =   "RsExpenses.frx":911E
         Top             =   750
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
         Height          =   195
         Index           =   15
         Left            =   11460
         RightToLeft     =   -1  'True
         TabIndex        =   88
         Top             =   630
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»‰«¡ ⁄·Ï"
         Height          =   285
         Index           =   0
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   2310
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·‘—Õ «·⁄«„"
         Height          =   285
         Index           =   20
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   3030
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÿ·»Ì…"
         Height          =   285
         Index           =   21
         Left            =   15000
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   2790
         Width           =   1155
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   2340
      Left            =   12840
      TabIndex        =   43
      Top             =   4920
      Visible         =   0   'False
      Width           =   12795
      _cx             =   22569
      _cy             =   4128
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   19
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"RsExpenses.frx":96A8
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.Frame Frame3 
         Caption         =   "Õœœ —ﬁ„ «·ﬁÌœ «·„—«œ ‰”Œ…"
         Height          =   1215
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   3720
         Visible         =   0   'False
         Width           =   4215
         Begin VB.CommandButton Command5 
            Caption         =   "‰”Œ"
            Height          =   255
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   255
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3915
         Left            =   2550
         RightToLeft     =   -1  'True
         ScaleHeight     =   3915
         ScaleWidth      =   9405
         TabIndex        =   44
         Top             =   810
         Visible         =   0   'False
         Width           =   9405
         Begin VB.TextBox TxtDese 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   1485
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   48
            Top             =   2040
            Width           =   8955
         End
         Begin VB.TextBox txtcodesub 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   3600
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Add des"
            Height          =   255
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   3600
            Width           =   1350
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Call des"
            Height          =   255
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   3600
            Width           =   1095
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   3900
            Left            =   120
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   0
            Width           =   10905
            _cx             =   19235
            _cy             =   6879
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   20.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   16777215
            ForeColor       =   4210688
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
            CaptionPos      =   7
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   7
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
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               Height          =   1605
               Left            =   0
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   50
               Top             =   480
               Visible         =   0   'False
               Width           =   8955
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000C&
               Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
               ForeColor       =   &H0000C8FF&
               Height          =   315
               Left            =   6840
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   0
               Width           =   2445
            End
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            Height          =   495
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   3480
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Height          =   495
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            Height          =   255
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   1320
            Width           =   735
         End
      End
      Begin VDSCOMBOLibCtl.SmartCombo CboDes 
         Height          =   315
         Left            =   0
         TabIndex        =   112
         ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
         Top             =   0
         Visible         =   0   'False
         Width           =   2955
         _cx             =   1973752924
         _cy             =   1973748268
         Alignment       =   0
         Appearance      =   3
         AutoSearch      =   0   'False
         BackColor       =   -2147483624
         BackgroundColor =   -2147483633
         BorderColor     =   0
         BorderVisible   =   -1  'True
         Caption         =   "SmartCombo1"
         CaptionAlignment=   4
         CaptionBackColor=   -2147483633
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionForeColor=   -2147483630
         CaptionHeight   =   15
         CaptionOnTop    =   0   'False
         CaptionMultiLine=   0
         Checkbox3D      =   0   'False
         CheckboxAlignment=   5
         CheckboxBackColor=   16777215
         CheckboxSize    =   13
         CheckboxValue   =   0
         BrowsePictureAlignment=   5
         BrowsePictureStretchH=   0
         BrowsePictureStretchV=   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Gap             =   0
         HideSelection   =   -1  'True
         Locked          =   0   'False
         MaxLength       =   0
         MultiLine       =   0
         OnFocus         =   3
         PasswordChar    =   ""
         Picture         =   "RsExpenses.frx":99BF
         PictureAlignment=   5
         PictureBackColor=   -2147483624
         PictureStretchH =   0
         PictureStretchV =   0
         Redraw          =   -1  'True
         ScrollBar       =   0
         Style           =   0
         Text            =   ""
         UnderLine       =   0   'False
         Enabled0        =   -1  'True
         Position0       =   0
         Tip0            =   "Caption"
         Visible0        =   0   'False
         Width0          =   90
         Enabled1        =   -1  'True
         Position1       =   1
         Tip1            =   ""
         Visible1        =   -1  'True
         Width1          =   32
         Enabled2        =   -1  'True
         Position2       =   2
         Tip2            =   "Check Box (Space, Ctrl + Space)"
         Visible2        =   0   'False
         Width2          =   16
         Enabled3        =   -1  'True
         Position3       =   3
         Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
         Visible3        =   -1  'True
         Width3          =   145
         Enabled4        =   -1  'True
         Position4       =   4
         Tip4            =   "Left Spinner (Alt + Left)"
         Visible4        =   0   'False
         Width4          =   16
         Enabled5        =   -1  'True
         Position5       =   5
         Tip5            =   "Right Spinner (Alt + Right)"
         Visible5        =   0   'False
         Width5          =   16
         Enabled6        =   -1  'True
         Position6       =   6
         Tip6            =   "Up Spinner (Ctrl + Up)"
         Visible6        =   0   'False
         Width6          =   16
         Enabled7        =   -1  'True
         Position7       =   7
         Tip7            =   "Down Spinner (Ctrl + Down)"
         Visible7        =   0   'False
         Width7          =   16
         Enabled8        =   -1  'True
         Position8       =   8
         Tip8            =   "Browse (Alt + Enter)"
         Visible8        =   0   'False
         Width8          =   16
         Enabled9        =   -1  'True
         Position9       =   9
         Tip9            =   " (Alt + Down, F4)"
         Visible9        =   -1  'True
         Width9          =   16
         Enabled10       =   -1  'True
         Position10      =   10
         Tip10           =   "Right Arrow (Alt + >)"
         Visible10       =   0   'False
         Width10         =   16
      End
   End
   Begin VB.TextBox Txt_Numorder 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌœ «·„Õ«”»Ì"
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
      Height          =   1035
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   9420
      Width           =   6465
      Begin MSDataListLib.DataCombo DcboDebitSide 
         Height          =   315
         Left            =   90
         TabIndex        =   27
         Top             =   270
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboCreditSide 
         Height          =   315
         Left            =   90
         TabIndex        =   29
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Index           =   12
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   570
         Width           =   1485
      End
      Begin VB.Label LblDevID 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·› —… :"
         Height          =   315
         Index           =   13
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÌœ:"
         Height          =   315
         Index           =   11
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› œ«∆‰"
         Height          =   285
         Index           =   10
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› „œÌ‰"
         Height          =   285
         Index           =   9
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.TextBox XPTxtVal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   7440
      Width           =   1905
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   765
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   12855
      _cx             =   22675
      _cy             =   1349
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
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "RsExpenses.frx":9F59
      Caption         =   "«·„’—Ê›«  "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
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
      PicturePos      =   0
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
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1695
         TabIndex        =   4
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "RsExpenses.frx":AC33
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
         Height          =   375
         Index           =   2
         Left            =   630
         TabIndex        =   5
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "RsExpenses.frx":AFCD
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
         Height          =   375
         Index           =   1
         Left            =   2220
         TabIndex        =   6
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "RsExpenses.frx":B367
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
         Height          =   375
         Index           =   3
         Left            =   1155
         TabIndex        =   7
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "RsExpenses.frx":B701
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin MSAdodcLib.Adodc numbering 
         Height          =   585
         Left            =   4680
         Top             =   0
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
      Begin MSAdodcLib.Adodc detect_no 
         Height          =   585
         Left            =   2640
         Top             =   0
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
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   7320
         Picture         =   "RsExpenses.frx":BA9B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label LblShortcutKeys 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ÃœÌœ F12 Or Enter ,  ⁄œÌ· F11 , Õ›Ÿ F10 ,  —«Ã⁄ F9 ,Õ–› F8 ,»ÕÀ F7 "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   510
         Width           =   5445
      End
   End
   Begin MSDataListLib.DataCombo XPCboExpensesType 
      Height          =   315
      Left            =   12960
      TabIndex        =   0
      Top             =   2760
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   8640
      TabIndex        =   9
      Top             =   8490
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   10020
      TabIndex        =   16
      Top             =   7800
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   1
      Left            =   9120
      TabIndex        =   17
      Top             =   7800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   2
      Left            =   8310
      TabIndex        =   18
      Top             =   7800
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   3
      Left            =   7155
      TabIndex        =   19
      Top             =   7830
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   4
      Left            =   6240
      TabIndex        =   20
      Top             =   7800
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   6
      Left            =   2160
      TabIndex        =   21
      Top             =   7830
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdHelp 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   3120
      TabIndex        =   22
      Top             =   7830
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   5
      Left            =   5190
      TabIndex        =   23
      Top             =   7800
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
      Height          =   2340
      Left            =   0
      TabIndex        =   34
      Top             =   4920
      Width           =   12840
      _cx             =   22648
      _cy             =   4128
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   32
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"RsExpenses.frx":F703
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.PictureBox PicDes 
         BorderStyle     =   0  'None
         Height          =   1635
         Left            =   240
         RightToLeft     =   -1  'True
         ScaleHeight     =   1635
         ScaleWidth      =   2925
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   2925
         Begin VB.TextBox TxtDes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   1125
            Left            =   30
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   37
            Top             =   360
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.Label LblDes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000C&
            Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
            ForeColor       =   &H0000C8FF&
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   0
            Width           =   2445
         End
      End
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   255
      Left            =   11040
      TabIndex        =   35
      Top             =   8160
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "„—«ﬂ“ «· ﬂ·›…"
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
      MICON           =   "RsExpenses.frx":FB97
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   8
      Left            =   4200
      TabIndex        =   39
      Top             =   7920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â"
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   9
      Left            =   5640
      TabIndex        =   40
      Top             =   8400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·‘Ìﬂ"
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ALLButtonS.ALLButton CmdRemove 
      Height          =   375
      Left            =   9600
      TabIndex        =   41
      Tag             =   "Delete Row"
      Top             =   7320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ–› ”ÿ—"
      ENAB            =   0   'False
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
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   255
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "RsExpenses.frx":FBB3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   10
      Left            =   3840
      TabIndex        =   42
      Top             =   8400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   11
      Left            =   120
      TabIndex        =   113
      Top             =   7920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«·„—›ﬁ« "
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
   Begin VB.Label LblValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   405
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   7380
      Width           =   5175
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   435
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   8370
      Width           =   555
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   435
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   8370
      Width           =   525
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "/"
      Height          =   435
      Index           =   6
      Left            =   690
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   8370
      Width           =   165
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
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
      Height          =   435
      Index           =   7
      Left            =   1500
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   8370
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   390
      Index           =   8
      Left            =   11745
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   8505
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«Ã„«·Ì"
      Height          =   285
      Index           =   2
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   7440
      Width           =   1515
   End
End
Attribute VB_Name = "RsExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim numbering_type As Integer
Dim departement_name  As String
Dim branch_no  As String
Dim RsNotes As ADODB.Recordset
Public LongRow As Long
Public LngRow As Double
Public LngCol As Double
Dim OtherInformation As New ClsGLOther
Dim Line1 As Double




Private Sub Command1_Click()
FillGridWithData222
End Sub

Private Sub Fg_Journal_Click()
   If Fg_Journal.Row <> 0 Then
        DcbIqara2.BoundText = val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("iqarid")))
                       TxtValue.Text = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("value"))
                TxtRemarks.Text = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("des"))

    End If
End Sub

Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    If Index = 18 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(18).ToolTipText = "ﬁÌ„… „»·€ «·„ﬁ»Ê÷« :" & lbl(18).Caption
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(18).ToolTipText = "Notes Recivable Value:" & lbl(18).Caption
        End If
    End If

End Sub

Private Sub CmdNos_Click(Index As Integer)
  If Index <= 9 Then
LBLPayVal.Caption = LBLPayVal.Caption & Index

ElseIf Index = 10 Then
LBLPayVal.Caption = LBLPayVal.Caption & "00"

ElseIf Index = 11 Then
LBLPayVal.Caption = LBLPayVal.Caption & "."

ElseIf Index = 12 Then 'ar
If Len(LBLPayVal.Caption) > 1 Then
LBLPayVal.Caption = mId(LBLPayVal.Caption, 1, Len(LBLPayVal.Caption) - 1)
Else
LBLPayVal.Caption = ""
End If
ElseIf Index = 13 Then 'ar
 LBLPayVal.Caption = ""

TxtPayedValue2.Text = ""
cleargrid

ElseIf Index = 14 Then
TxtPayedValue2.Text = val(LBLPayVal)

 
        With Grid22
          .TextMatrix(.Row, .ColIndex("Value")) = TxtPayedValue2.Text
          End With
    ReLineGrid2
     
 TxtRemainValue2.Text = val(Me.TxtPayedValue2.Text) - val(Me.TxtNetValue2.Text)
 

End If

 ReLineGrid2
 
End Sub
Private Sub Grid22_AfterEdit(ByVal Row As Long, ByVal Col As Long)
ReLineGrid2
End Sub
Private Sub CmdValue_Click(Index As Integer)
LBLPayVal.Caption = 0
'TxtPayedValue.text = CmdValue(Index).Caption
LBLPayVal.Caption = CmdValue(Index).Caption
        With Grid22
          .TextMatrix(.Row, .ColIndex("Value")) = LBLPayVal.Caption
          End With
     ReLineGrid2
End Sub
Private Sub lblexit_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
If Index = 90 Then
If val(CboPayMentType.ListIndex) = 4 Then
If val(TxtRemainValue2.Text) <> 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œŒ·… €Ì— ’ÕÌÕ…"
Else
MsgBox "The  value is incorrect"
End If
Exit Sub
End If
FramePay.Visible = False
End If
End If
Else
FramePay.Visible = False
End If
End Sub
Private Sub Grid22_Click()
If TxtPayedValue2.Text = "" Or val(TxtPayedValue2.Text) = 0 Then
With Me.Grid22
.TextMatrix(.Row, .ColIndex("Value")) = LBLPayVal.Caption
ReLineGrid2
End With
End If
End Sub
Sub DeleteBillBuy()
Dim i As Integer
Dim StrSQL As String
With VSFlexGrid1
 For i = .FixedRows To .Rows - 1
 If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
      StrSQL = "Update Transactions Set  TotalPayed=0 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
     End If
     Next i
 End With
End Sub
Function saveBillBuy()
    Dim StrSQL As String
   ' Dim StrSQL  As String
    Dim i As Integer
    Dim Diff As Double
    Dim Note_Value1 As Double
    Dim TxtValueTemp As Double
    
    Diff = 0
Dim RsDetails As ADODB.Recordset
      If Me.TxtModFlg.Text = "E" Then
    StrSQL = "Delete From TblNotesBillBuyPayment2 Where NoteID1=" & val(Me.XPTxtID.Text) & " and TransType is null"
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtID.Text) & " and TransType is null"
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblNotesBillBuyPayment2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid1
    TxtValueTemp = val(XPTxtVal.Text)
    For i = .FixedRows To .Rows - 1
        If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID1").value = val(XPTxtID.Text)
            RsDetails("NoteID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
            RsDetails("branch_no").value = val(.TextMatrix(i, .ColIndex("branch_no")))
            RsDetails("NoteSerial1").value = val(.TextMatrix(i, .ColIndex("NoteSerial1")))
            RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
            Note_Value1 = val(.TextMatrix(i, .ColIndex("RemainingValue")))
            Diff = 0
            If val(TxtValueTemp) > 0 Then
          If val(TxtValueTemp) <= Note_Value1 Then
          Diff = val(TxtValueTemp)
          TxtValueTemp = val(TxtValueTemp) - Note_Value1
          Else
          Diff = Note_Value1
          TxtValueTemp = val(TxtValueTemp) - Note_Value1
          End If
            End If
          ' .TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("RemainingValue")))
            .TextMatrix(i, .ColIndex("TransPayedValue")) = Diff
            
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("PayedValue")))
            
            RsDetails("too").value = (.TextMatrix(i, .ColIndex("too")))
            RsDetails("NoteDate").value = IIf((.TextMatrix(i, .ColIndex("NoteDate"))) = "", Null, (.TextMatrix(i, .ColIndex("NoteDate"))))
            If .TextMatrix(i, .ColIndex("DueDate")) <> "" And .TextMatrix(i, .ColIndex("DueDate")) <> " " Then
            RsDetails("DueDate").value = IIf((.TextMatrix(i, .ColIndex("DueDate"))) = "", Null, (.TextMatrix(i, .ColIndex("DueDate"))))
            Else
            RsDetails("DueDate").value = Null
            End If
            RsDetails("TransPayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
           .TextMatrix(i, .ColIndex("NetValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
            RsDetails("RemainingValue").value = val(.TextMatrix(i, .ColIndex("RemainingValue")))
            RsDetails.update
                
            If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 Then
            StrSQL = "Update Transactions Set  TotalPayed=1 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
             Else
                 StrSQL = "Update Transactions Set  TotalPayed=0 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
      End If
    Next i
End With
    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblBillBuyPayment2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid1
    For i = .FixedRows To .Rows - 1
        If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID").value = val(XPTxtID.Text)
            RsDetails("RecDate").value = XPDtbTrans.value
            RsDetails("Serial").value = TxtSerial1.Text
            RsDetails("Transaction_ID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
            RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails.update
        End If
    Next i
End With

End Function
Public Sub RetriveBillBuyData(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String


   ' On Error GoTo ErrTrap
    Set RsDetails = New ADODB.Recordset
  StrSQL = "   SELECT     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblNotesBillBuyPayment2.*"
  StrSQL = StrSQL & "  FROM         dbo.TblNotesBillBuyPayment2 LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblNotesBillBuyPayment2.branch_no = dbo.TblBranchesData.branch_id"
  StrSQL = StrSQL & "  Where (dbo.TblNotesBillBuyPayment2.NoteID1 = " & val(XPTxtID.Text) & ")"
    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid1
    .Clear flexClearScrollable, flexClearEverything
    .Rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
      '  Fra(2).Visible = True
      '               lbl(47).Visible = True
      '  TxtAdvance.Visible = True
        RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To RsDetails.RecordCount
        .TextMatrix(i, .ColIndex("Ser")) = i

            .TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(RsDetails("branch_no").value), 0, RsDetails("branch_no").value)
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_name").value), "", RsDetails("branch_name").value)
            Else
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_namee").value), 0, RsDetails("branch_namee").value)
            End If
            .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsDetails("NoteID").value), 0, RsDetails("NoteID").value)
            .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsDetails("NoteSerial1").value), 0, RsDetails("NoteSerial1").value)
            .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsDetails("Note_Value").value), 0, RsDetails("Note_Value").value)
            .TextMatrix(i, .ColIndex("PayedValue")) = IIf(IsNull(RsDetails("PayedValue").value), 0, RsDetails("PayedValue").value)
            .TextMatrix(i, .ColIndex("TransPayedValue")) = IIf(IsNull(RsDetails("TransPayedValue").value), 0, RsDetails("TransPayedValue").value)
            .TextMatrix(i, .ColIndex("too")) = IIf(IsNull(RsDetails("too").value), "", RsDetails("too").value)
            .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(RsDetails("NetValue").value), 0, RsDetails("NetValue").value)
            .TextMatrix(i, .ColIndex("RemainingValue")) = IIf(IsNull(RsDetails("RemainingValue").value), 0, RsDetails("RemainingValue").value)
            
           ' .TextMatrix(i, .ColIndex("PartValue")) = Round(RsDetails("PartValue").value, 2)
             .TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(((RsDetails("DueDate").value))), " ", ((RsDetails("DueDate").value)))
           ' .TextMatrix(i, .ColIndex("NoteDate")) = DisplayDate(CDate(RsDetails("NoteDate").value))
            .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(((RsDetails("NoteDate").value))), "", ((RsDetails("NoteDate").value)))
            .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
            RsDetails.MoveNext
        Next i
        

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
End Sub
Sub RetriveBillBuy(Optional CuID As Double = 0)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Dim i As Integer
Set Rs8 = New ADODB.Recordset
With VSFlexGrid1
.Clear flexClearScrollable, flexClearEverything
.Rows = 1
End With
sql = " SELECT     TOP 100 PERCENT dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, "
sql = sql & "                      dbo.Transactions.ManualNO, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.CusID,"
sql = sql & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Transactions.TotalPayed, dbo.Transactions.OldContID,"
sql = sql & "                      dbo.transactions.OldValue , dbo.transactions.dueDate, dbo.transactions.Vat, dbo.transactions.Transaction_NetValue"
sql = sql & " FROM         dbo.Transactions LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
sql = sql & "  WHERE     (dbo.Transactions.PaymentType = 1) AND (dbo.Transactions.Transaction_Type = 21 OR"
sql = sql & "                       dbo.Transactions.Transaction_Type = 2 or dbo.Transactions.Transaction_Type = 71) AND (dbo.Transactions.TotalPayed IS NULL OR"
sql = sql & "                       dbo.Transactions.TotalPayed = 0) AND (dbo.Transactions.CusID = " & CuID & ")"
sql = sql & "  ORDER BY dbo.Transactions.DueDate ,dbo.Transactions.NoteSerial1"

Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
VSFlexGrid1.Enabled = True


        VSFlexGrid1.Enabled = True
With VSFlexGrid1
.Clear flexClearScrollable, flexClearEverything
.Rows = 1
    .Rows = .Rows + Rs8.RecordCount
.Rows = .FixedRows + Rs8.RecordCount
Rs8.MoveFirst
For i = .FixedRows To Rs8.RecordCount
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(Rs8("BranchId").value), 0, Rs8("BranchId").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), 0, Rs8("branch_name").value)
Else
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), 0, Rs8("branch_namee").value)
End If

.TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(Rs8("DueDate").value), "", Rs8("DueDate").value)
.TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(Rs8("Transaction_ID").value), 0, Rs8("Transaction_ID").value)
.TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs8("Transaction_Date").value), "", Rs8("Transaction_Date").value)
.TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs8("NoteSerial1").value), "", Rs8("NoteSerial1").value)
.TextMatrix(i, .ColIndex("too")) = IIf(IsNull(Rs8("ManualNO").value), "", Rs8("ManualNO").value)
.TextMatrix(i, .ColIndex("Note_Value")) = val(IIf(IsNull(Rs8("Transaction_NetValue").value), IIf(IsNull(Rs8("OldValue").value), 0, Rs8("OldValue").value), Rs8("Transaction_NetValue").value))
If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
.TextMatrix(i, .ColIndex("PayedValue")) = GeteBillBuy(val(.TextMatrix(i, .ColIndex("NoteID"))))
Else
.TextMatrix(i, .ColIndex("PayedValue")) = 0
End If
.TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("Note_Value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
Rs8.MoveNext
Next i
End With
End If
End Sub
Function GeteBillBuy(Optional Transaction_ID As Double = 0) As Double
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT   SUM(PayedValue) AS Smatiobn"
sql = sql & " From dbo.TblBillBuyPayment2"
sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"
sql = sql & " GROUP BY Transaction_ID"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GeteBillBuy = IIf(IsNull(Rs8("Smatiobn").value), 0, Rs8("Smatiobn").value)
Else
GeteBillBuy = 0
End If
End Function
Private Sub CMDPAy_Click()
If val(CboPayMentType.ListIndex) = 4 Then
If val(TxtRemainValue2.Text) <> 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œŒ·… €Ì— ’ÕÌÕ…"
Else
MsgBox "The  value is incorrect"
End If
Exit Sub
End If
FramePay.Visible = False
End If
End Sub
Private Sub TxtNetValue2_Change()
    TxtRemainValue2.Text = val(Me.TxtPayedValue2.Text) - val(Me.TxtNetValue2.Text)
End Sub

Private Sub ALLButton2_Click()
If DcbIqara2.Text = "" Then Exit Sub
FillData
DcbIqara2.Text = ""
dcsupplier.Text = ""

End Sub

Private Sub CMDSENDSMS_Click()
'0 manual
'1 save
'2 Print

SendMessage (0)
End Sub
Function SendMessage(currentOpt As Integer)
            Dim subject As String
            Dim Msg As String
            Dim msgstatus As Boolean
           Dim CompanyName As String
           Dim cOptions As ClsCompanyInfo
           Set cOptions = New ClsCompanyInfo
           Dim companyphone As String
           Dim Opt As Integer
            Dim CurrentMessage As String
            Dim t As String
    CurrentMessage = ComposMessage(Me.Name, 0, "", "", Opt)
 Dim i As Integer
  If Opt = currentOpt Then
  
  
  
      CompanyName = cOptions.ArabCompanyName '& CHR(13) & CurrentBranchName
     companyphone = cOptions.Company_Mobile
  '
       For i = Me.Fg_Journal.FixedRows To Fg_Journal.Rows - 1
       
                 If (Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("iqarid"))) <> "" Then
 '
  '«·„«·ﬂ
 Msg = "     „ ’—› „»·€  " & (Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("value"))) & "  „ﬁ«»· " & (Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des"))) & "  ··ÊÕœÂ " & CHR(13) & (Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("Unitss"))) & CHR(13) & "  »«·⁄ﬁ«—   " & (Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("iqar")))
t = sendMessageM("user", "password", Msg, "", GetCustomerNumber(getownerId(val(Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("iqarid"))))))

 DoEvents
       'GetCustomerNumber(dcsupplier.BoundText))
       
                End If
  
        Next i
  
MsgBox " „ «·«—”«·"
     
     
     End If
 
End Function



Sub SaveUnitNo(Optional ID As Long, Optional i As Integer)
   Dim RsDetails As ADODB.Recordset
   Dim astrSplit2tems2() As String
   Dim astrSplitItems() As String
   Dim sql As String
   Dim j As Integer
    Dim st As String
    Dim nElements As Integer
    
      If Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("StrUnit")) <> "" Then
          st = Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("StrUnit"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         sql = "Select * from TblExpUnitNo where 1=-1"
         Set RsDetails = New ADODB.Recordset
         RsDetails.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
         For j = 0 To nElements - 1
          RsDetails.AddNew
         astrSplit2tems2 = Split(astrSplitItems(j), "#")
         RsDetails("ExpID").value = val(XPTxtID.Text)
         RsDetails("ExpDetails").value = ID
         RsDetails("UnitID").value = val(astrSplit2tems2(1))
         RsDetails("Valu").value = val(astrSplit2tems2(2))
         RsDetails.update
         Next j
          End If
End Sub
Public Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteType, dbo.notes_all.NoteSerial, dbo.notes_all.Note_Value, dbo.notes_all.BankID, "
MySQL = MySQL & "                       dbo.notes_all.ChqueNum, dbo.notes_all.DueDate, dbo.notes_all.UserID, dbo.notes_all.Remark, dbo.notes_all.ExpensesID, dbo.notes_all.BoxID, dbo.notes_all.too,"
MySQL = MySQL & "                       dbo.notes_all.note_value_by_characters, dbo.notes_all.general_des, dbo.notes_all.NoteSerial1, dbo.notes_all.ToPriodDateH, dbo.notes_all.FrmPriodDateH,"
MySQL = MySQL & "                       dbo.notes_all.ToPriodDate, dbo.notes_all.FrmPriodDate, dbo.notes_all.Iqar, dbo.notes_all.UnitType, dbo.notes_all.NoteDateH, dbo.notes_all.CashingType,"
MySQL = MySQL & "                       dbo.notes_all.NoteCashingType, dbo.notes_all.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblExpensesDet.Unitss,"
MySQL = MySQL & "                       dbo.TblExpensesDet.StrUnit, dbo.TblExpensesDet.AccountCode, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng,"
MySQL = MySQL & "                       dbo.TblExpensesDet.des, dbo.TblExpensesDet.order_no, dbo.TblExpensesDet.opr_fullcode, dbo.TblExpensesDet.[value], dbo.TblExpensesDet.iqarid,"
MySQL = MySQL & "                       dbo.TblAqar.aqarname, dbo.TblExpensesDet.uintid, dbo.TblAqarDetai.unitno, dbo.TblExpensesDet.type, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee,"
MySQL = MySQL & "                       dbo.notes_all.NoteHijriDate, dbo.BanksData.BankName, dbo.BanksData.BankNamee, dbo.TblExpensesDet.ExpID, dbo.TblExpensesDet.ID, dbo.TblExpUnitNo.Valu,"
MySQL = MySQL & "                       dbo.TblExpUnitNo.UnitID, TblAqarDetai_1.unitno AS UnitnoName"
MySQL = MySQL & "  FROM         dbo.TblAqar RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblExpUnitNo LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAqarDetai TblAqarDetai_1 ON dbo.TblExpUnitNo.UnitID = TblAqarDetai_1.Id RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblExpensesDet ON dbo.TblExpUnitNo.ExpDetails = dbo.TblExpensesDet.ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAkarUnit ON dbo.TblExpensesDet.type = dbo.TblAkarUnit.id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAqarDetai ON dbo.TblExpensesDet.uintid = dbo.TblAqarDetai.Id ON dbo.TblAqar.Aqarid = dbo.TblExpensesDet.iqarid LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.ACCOUNTS ON dbo.TblExpensesDet.AccountCode = dbo.ACCOUNTS.Account_Code RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.BanksData RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.notes_all ON dbo.BanksData.BankID = dbo.notes_all.BankID ON dbo.TblExpensesDet.ExpID = dbo.notes_all.NoteID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.notes_all.NoteID = " & val(XPTxtID.Text) & ")"

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Expenses_order13.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Expenses_order13.rpt"
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
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(XPTxtVal.Text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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
Private Sub ALLButton1_Click()
    On Error GoTo ErrTrap

    If DcCostCenter.BoundText <> "" Then

        MsgBox "·«Ì„ﬂ‰ «· Ê“Ì⁄ ⁄·Ï „—«ﬂ“ «· ﬂ·›… ·«‰ﬂ «Œ —   Ê“Ì⁄ ⁄«„ ⁄·Ï „—ﬂ“  ﬂ·›… „Õœœ", vbCritical
        Exit Sub
    End If

    Dim opr_id As Double

    If Not IsNumeric(Text1.Text) Then Exit Sub
    'If Me.TxtModFlg.text = "N" Then
    opr_id = val(Me.Text1.Text)
    'Else
    'opr_id = TxtDEV_NO.text
    'End If

    If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) = "" Then
        If Not val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE"))) = 0 Then
            marakes_taklefa_tawze3.show
            
            marakes_taklefa_tawze3.value.Caption = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE")) ' Text4.Text
            marakes_taklefa_tawze3.depit_or_credit.Caption = "„œÌ‰"
            marakes_taklefa_tawze3.kedno = opr_id
            marakes_taklefa_tawze3.account_no = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode"))
            marakes_taklefa_tawze3.account_name = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountName"))
            marakes_taklefa_tawze3.lineno = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„… «Ê·« ", vbCritical
            Else
                MsgBox "Enter Value First ", vbCritical
            End If

            Exit Sub
        End If

        marakes_taklefa_tawze3.opr_type = "”‰œ ’—›"
        marakes_taklefa_tawze3.opr_id = opr_id 'TxtDEV_NO.text 'Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo"))  'Text5.Text
        marakes_taklefa_tawze3.Adodc3.ConnectionString = connection_string
        marakes_taklefa_tawze3.Adodc3.CommandType = adCmdText
        marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        marakes_taklefa_tawze3.Adodc3.Refresh
        '    Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("distributed")) = "1"

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CboDes_AfterAutoCloseUp()
    PutData
    CboDes.Visible = False
End Sub

Private Sub CboPayMentType_Change()
chkvat.Enabled = False
    If Me.CboPayMentType.ListIndex = 0 Then
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        DcboBankName.Text = ""
        TxtChequeNumber.Text = ""
        TxtAccount.Text = ""
        TxtAccount.Enabled = False
        DcbAccount.BoundText = ""
        DcbAccount.Enabled = False
    ElseIf Me.CboPayMentType.ListIndex = 1 Then
        Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.Text = ""
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
                Me.lbl(18).Caption = "—ﬁ„ «·‘Ìﬂ"
        TxtAccount.Text = ""
        TxtAccount.Enabled = False
        DcbAccount.BoundText = ""
        DcbAccount.Enabled = False
     ElseIf Me.CboPayMentType.ListIndex = 2 Then
        TxtAccount.Enabled = True
        DcbAccount.Enabled = True
            chkvat.Enabled = True
            
        Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.Text = ""
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
    ElseIf Me.CboPayMentType.ListIndex = 3 Then
         Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.Text = ""
        Me.lbl(19).Enabled = True
        Me.lbl(18).Caption = "—ﬁ„ «·ÕÊ«·…"
        Me.lbl(18).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        TxtAccount.Text = ""
        TxtAccount.Enabled = False
        DcbAccount.BoundText = ""
        DcbAccount.Enabled = False
    ElseIf Me.CboPayMentType.ListIndex = 4 Then

      TxtAccount.Enabled = False
       Me.DcboBox.Enabled = True
        DcbAccount.Enabled = False
        TxtAccount.Text = ""
        DcbAccount.BoundText = ""
     FramePay.Visible = True
  '   If Me.TxtModFlg.Text <> "R" And Me.TxtModFlg.Text <> "" Then
     If Me.TxtModFlg.Text = "N" Then
     If val(XPTxtVal.Text) > 0 Then
     FramePay.Visible = True
     FillGridWithData222
     LBLPayVal.Caption = 0
LBLPayVal.Caption = val(XPTxtVal.Text)
TxtNetValue2.Text = val(LBLPayVal.Caption)
    With Grid22
          .TextMatrix(.Row, .ColIndex("Value")) = 0
    End With
     ReLineGrid2
     End If
     Else
      FramePay.Visible = True
      
    If FillGridWithDataPayment() = True Then
     LBLPayVal.Caption = val(XPTxtVal.Text)
     TxtNetValue2.Text = val(LBLPayVal.Caption)
     ReLineGrid2
     Else
     '''
    If val(XPTxtVal.Text) > 0 Then
     FramePay.Visible = True
     FillGridWithData222
     LBLPayVal.Caption = 0
LBLPayVal.Caption = val(XPTxtVal.Text)
TxtNetValue2.Text = val(LBLPayVal.Caption)
    With Grid22
          .TextMatrix(.Row, .ColIndex("Value")) = 0
    End With
     ReLineGrid2
     End If
     End If
     
     End If
          Me.lbl(9).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Frame3.Enabled = False
    Else
    
    
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        TxtAccount.Enabled = False
        DcbAccount.Enabled = False
         TxtAccount.Text = ""
        TxtAccount.Enabled = False
        DcbAccount.BoundText = ""
        DcbAccount.Enabled = False
    End If

End Sub

Private Sub ReLineGrid2()
If Me.TxtModFlg = "R" Then Exit Sub
    On Error Resume Next
    Dim i As Integer
    Dim IntCounter As Integer
    Dim totalPayed As Double
    Dim visapayed As Double
 totalPayed = 0
 visapayed = 0
  With Grid22

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("value")) <> "" Then
               ' IntCounter = IntCounter + 1
                totalPayed = totalPayed + .TextMatrix(i, .ColIndex("value"))
                If totalPayed > val(Me.TxtNetValue2.Text) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·ﬁÌ„… «ﬂ»— „‰ «·«Ã„«·Ì"
                Else
                 MsgBox "ERROR Incorrect Value" & CHR(13)
                End If
                .TextMatrix(i, .ColIndex("value")) = 0
                Exit Sub
                End If
            End If

        Next i

    End With
  TxtPayedValue2.Text = totalPayed
    TxtRemainValue2.Text = val(Me.TxtPayedValue2.Text) - val(Me.TxtNetValue2.Text)
End Sub
Public Sub FillGridWithData222()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "SELECT     dbo.TblPaymentType.PaymentID, dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId, dbo.TblPaymentType.Accountsus, "
My_SQL = My_SQL & "  dbo.TblPaymentType.Accountcom, dbo.TblPaymentType.commision, dbo.TblPaymentType.PaymentNamee, dbo.BanksData.Account_Code AS bankAccount_Code ,dbo.TblPaymentType.MaxValue"
My_SQL = My_SQL & " FROM         dbo.TblPaymentType LEFT OUTER JOIN"
My_SQL = My_SQL & " dbo.BanksData ON dbo.TblPaymentType.BankId = dbo.BanksData.BankID"
My_SQL = My_SQL & " where (dbo.TblPaymentType.TypTran=2 or dbo.TblPaymentType.TypTran is null)  "
If SystemOptions.LinkUsersWithPayment = True Then
My_SQL = My_SQL & " and dbo.TblPaymentType.PaymentID in (SELECT     PaynetID"
My_SQL = My_SQL & " From dbo.TblPaymentUser"
My_SQL = My_SQL & " Where (UserID = " & user_id & "))"
End If
My_SQL = My_SQL & " order by PaymentID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid22
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 2
            rs.MoveFirst
      If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(1, .ColIndex("PaymentName")) = " ‰ﬁœÌ"
               Else
               .TextMatrix(1, .ColIndex("PaymentName")) = " Cash"
               End If
               
                .TextMatrix(1, .ColIndex("PaymentID")) = 0
           
           
            For i = 2 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentName").value), "", rs.Fields("PaymentName").value)
               Else
               .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentNamee").value), "", rs.Fields("PaymentNamee").value)
               End If
               
                .TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(rs.Fields("PaymentID").value), "", rs.Fields("PaymentID").value)
           
                .TextMatrix(i, .ColIndex("BankId")) = IIf(IsNull(rs.Fields("BankId").value), "", rs.Fields("BankId").value)
            
            .TextMatrix(i, .ColIndex("Accountsus")) = IIf(IsNull(rs.Fields("Accountsus").value), "", rs.Fields("Accountsus").value)
            .TextMatrix(i, .ColIndex("Accountcom")) = IIf(IsNull(rs.Fields("Accountcom").value), "", rs.Fields("Accountcom").value)
            .TextMatrix(i, .ColIndex("MaxValue")) = IIf(IsNull(rs.Fields("MaxValue").value), 0, rs.Fields("MaxValue").value)
            .TextMatrix(i, .ColIndex("commision")) = IIf(IsNull(rs.Fields("commision").value), "", rs.Fields("commision").value)
           .TextMatrix(i, .ColIndex("bankAccount_Code")) = IIf(IsNull(rs.Fields("bankAccount_Code").value), "", rs.Fields("bankAccount_Code").value)
            
                rs.MoveNext
            Next

            rs.Close
        End If

  '      .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub
Private Sub cleargrid()
    On Error Resume Next
    Dim i As Integer
 
  With Grid22

       ' For I = .FixedRows To .Rows - 1

         .TextMatrix(.Row, .ColIndex("value")) = 0
          
       ' Next I

    End With
     TxtPayedValue2 = 0
    
End Sub

Sub SaveMultyPayment(Optional NoteID As Double)
Dim Rs3 As ADODB.Recordset
Dim i As Integer
Set Rs3 = New ADODB.Recordset
Dim sql As String
sql = "select * from TblMultuPayment where 1=-1 "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With Grid22
For i = 1 To .Rows - 1
If (.TextMatrix(i, .ColIndex("PaymentName"))) <> "" Then
Rs3.AddNew
Rs3("NoteID").value = NoteID
Rs3("PaymentID").value = val(.TextMatrix(i, .ColIndex("PaymentID")))
Rs3("Value").value = val(.TextMatrix(i, .ColIndex("Value")))
Rs3("CardNo").value = (.TextMatrix(i, .ColIndex("CardNo")))
Rs3("maxvalue").value = val((.TextMatrix(i, .ColIndex("MaxValue"))))

Rs3.update
End If
Next i
End With
End Sub


Private Sub TxtPaymentValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPaymentValue.Text, 0)
End Sub

 Function FillGridWithDataPayment() As Boolean
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = " SELECT     TOP 100 PERCENT dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId, dbo.TblPaymentType.Accountsus, dbo.TblPaymentType.Accountcom, "
    My_SQL = My_SQL & "                   dbo.TblPaymentType.commision, dbo.TblPaymentType.PaymentNamee, dbo.BanksData.Account_Code AS bankAccount_Code, dbo.TblMultuPayment.[Value],"
    My_SQL = My_SQL & "                   dbo.TblMultuPayment.CardNo, dbo.TblMultuPayment.PaymentID , dbo.TblMultuPayment.NoteID,dbo.TblMultuPayment.MaxValue"
    My_SQL = My_SQL & "      FROM         dbo.TblPaymentType RIGHT OUTER JOIN"
    My_SQL = My_SQL & "                   dbo.TblMultuPayment ON dbo.TblPaymentType.PaymentID = dbo.TblMultuPayment.PaymentID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                   dbo.BanksData ON dbo.TblPaymentType.BankId = dbo.BanksData.BankID"
    My_SQL = My_SQL & "     Where (dbo.TblMultuPayment.NoteID = " & val(XPTxtID.Text) & ")"
    My_SQL = My_SQL & "   ORDER BY dbo.TblPaymentType.PaymentID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid22
        .Rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 2
            rs.MoveFirst
FillGridWithDataPayment = True
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentName").value), "‰ﬁœÌ", rs.Fields("PaymentName").value)
               Else
               .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentNamee").value), "Cash", rs.Fields("PaymentNamee").value)
               End If
               .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(rs.Fields("Value").value), "", rs.Fields("Value").value)
               .TextMatrix(i, .ColIndex("CardNo")) = IIf(IsNull(rs.Fields("CardNo").value), "", rs.Fields("CardNo").value)
               .TextMatrix(i, .ColIndex("MaxValue")) = IIf(IsNull(rs.Fields("MaxValue").value), 0, rs.Fields("MaxValue").value)
                .TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(rs.Fields("PaymentID").value), "", rs.Fields("PaymentID").value)
           
                .TextMatrix(i, .ColIndex("BankId")) = IIf(IsNull(rs.Fields("BankId").value), "", rs.Fields("BankId").value)
            
            .TextMatrix(i, .ColIndex("Accountsus")) = IIf(IsNull(rs.Fields("Accountsus").value), "", rs.Fields("Accountsus").value)
            .TextMatrix(i, .ColIndex("Accountcom")) = IIf(IsNull(rs.Fields("Accountcom").value), "", rs.Fields("Accountcom").value)
            .TextMatrix(i, .ColIndex("commision")) = IIf(IsNull(rs.Fields("commision").value), "", rs.Fields("commision").value)
           .TextMatrix(i, .ColIndex("bankAccount_Code")) = IIf(IsNull(rs.Fields("bankAccount_Code").value), "", rs.Fields("bankAccount_Code").value)
            
                rs.MoveNext
            Next

            rs.Close
            Else
            FillGridWithDataPayment = False
        End If

  '      .RowHeight(-1) = 300
    End With

ErrTrap:
End Function

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Function setfoxy()
    Text1.Text = CStr(new_id("foxy", "id", "", True))

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id").value = Text1.Text
 
    rs.update
    
End Function

Private Sub Cmd_Click(Index As Integer)
Dim Msg As String
    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            DcCostCenter.Text = ""
            ' XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
            ' Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=3"))
        
            Me.DCboUserName.BoundText = user_id
            '        XPDtbTrans.SetFocus
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.Rows = 2
          
            Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            Me.VSFlexGrid1.Rows = 2
          
            Fg_Journal.Enabled = True
            DtpChequeDueDate.value = Date
            
            setfoxy
            DcbBranch.BoundText = Current_branch
          
        Case 1
                   If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id
            Fg_Journal.Rows = Fg_Journal.Rows + 1
            Fg_Journal.Enabled = True
         
            VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
            VSFlexGrid1.Enabled = True

        Case 2
            If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
               End If
              Exit Sub
             End If
        
        If Me.CboPayMentType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
            Else
                Msg = "Select Payment method ...!!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CboPayMentType.SetFocus
            Exit Sub
        End If

        If Me.CboPayMentType.ListIndex = 0 Or Me.CboPayMentType.ListIndex = 4 Then
            If Trim(Me.DcboBox.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
                Else
                    Msg = "Select Box..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBox.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

        ElseIf Me.CboPayMentType.ListIndex = 1 Then

            If Me.DcboBankName.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                Else
                    Msg = "Select Bank...!!"
        
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBankName.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.Text) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                Else
                    Msg = "Enter Cheque No:...!!"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If

            If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
                Else
                    Msg = "Cheque Due Date Not Valid...!!"
        
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DtpChequeDueDate.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If
         ElseIf Me.CboPayMentType.ListIndex = 2 Then

            If Me.DcbAccount.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·Õ”«»...!!"
                Else
                    Msg = "Select Account...!!"
        
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcbAccount.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If
        End If
        my_branch = val(Me.DcbBranch.BoundText)
            If TxtSerial.Text = "" Then
                If Notes_coding(val(val(DcbBranch.BoundText)), XPDtbTrans.value) = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                    Else
                        MsgBox " Cant't Create Journal Entry to this Process no You exceed the maximum number ": Exit Sub
                    End If

                Else
         
                    If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                        Else
                            MsgBox "You must Define JE Coding ": Exit Sub
                        End If

                    Else
                        TxtSerial.Text = Notes_coding(val(my_branch), XPDtbTrans.value)
                    End If
                End If
            End If

            ' TxtSerial.text = Notes_coding(Val(my_branch), XPDtbTrans.value) 'kk
            If TxtSerial1.Text = "" Then
                If Voucher_coding(val(DcbBranch.BoundText), XPDtbTrans.value, 1, 3) = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—› ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    Else
                        MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
                    End If

                Else
         
                    If Voucher_coding(val(DcbBranch.BoundText), XPDtbTrans.value, 1, 3) = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                        Else
                            MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                        End If

                    Else
                        TxtSerial1.Text = Voucher_coding(val(DcbBranch.BoundText), XPDtbTrans.value, 1, 3)
                    End If
                End If
            End If
    
            SaveData
           SendMessage (1)
        Case 3
            Undo

        Case 4
                   If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            Load FrmNotesSearch
            FrmNotesSearch.SearchType = 333
            FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7
            ViewDataList

        Case 8
            print_report
SendMessage (2)
        Case 9
            print_Cheque TxtChequeNumber.Text, get_Cheque_report_no(val(DcboBankName.BoundText)), TxtSerial.Text

        Case 10
            ShowGL_cc TxtSerial.Text, , 3, , , TxtSerial1.Text
          Case 11
                      On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
ShowAttachments val(XPTxtID.Text) & TxtSerial1.Text, "16102017001"
    
    End Select

    Exit Sub
ErrTrap:
End Sub

Function print_Cheque(Optional ChqueNum As String = "", Optional report_no As String = "", Optional serial As String)
    hide_logo = True
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From notes  where ChqueNum='" & ChqueNum & "' and noteserial='" & TxtSerial & "'"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
    Else
        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
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
    'MsgBox ToHijriDate(Date)

    xReport.ParameterFields(5).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 1, 2)
    xReport.ParameterFields(6).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 4, 2)
    xReport.ParameterFields(7).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 9, 2)

    xReport.ParameterFields(8).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 1, 2)
    xReport.ParameterFields(9).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 4, 2)
    xReport.ParameterFields(10).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 9, 2)
    xReport.ParameterFields(11).AddCurrentValue CStr(txtto.Text)
    xReport.ParameterFields(12).AddCurrentValue CStr(XPTxtVal.Text)
    xReport.ParameterFields(13).AddCurrentValue CStr(Me.XPMTxtRemarks.Text)
    xReport.ParameterFields(14).AddCurrentValue CStr(LblValue.Caption)
 
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

'Function print_report(Optional NoteSerial As String)
'
'    Dim MySQL As String
'    Dim RsData As New ADODB.Recordset
'    Dim xApp As New CRAXDRT.Application
'    Dim xReport As CRAXDRT.Report
'    Dim CViewer As ClsReportViewer
'    Dim StrReportTitle As String
'    Dim StrFileName As String
'    Dim Msg As String
'
'    MySQL = "Select * From Expanses_Order  where noteserial='" & NoteSerial & "'"
'
'    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    'End If

'    If SystemOptions.UserInterface = ArabicInterface Then
'        StrFileName = App.path & "\Reports\" & "Expenses_order.rpt"
'    Else
'        StrFileName = App.path & "\Reports\" & "Expenses_order.rpt"
'    End If
'
'    If Dir(StrFileName) = "" Then
'        'GetMsgs 139, vbExclamation
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'
'    Set RsData = New ADODB.Recordset
'    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If RsData.BOF Or RsData.EOF Then
'        'GetMsgs 138, vbExclamation
'        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        RsData.Close
'        Set RsData = Nothing
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'
'    Screen.MousePointer = vbArrowHourglass
'    Set xReport = xApp.OpenReport(StrFileName)
'    xReport.Database.SetDataSource RsData

'    Dim cCompanyInfo As New ClsCompanyInfo

'    If SystemOptions.UserInterface = ArabicInterface Then
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
'        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
'    Else
 
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
'        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
'        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
'    End If
'
'    xReport.ParameterFields(3).AddCurrentValue user_name
'    xReport.reporttitle = StrReportTitle
'    xReport.EnableParameterPrompting = False
'    xReport.ApplicationName = App.title
'    xReport.ReportAuthor = App.title
'    Set CViewer = New ClsReportViewer
'    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
'
'    RsData.Close
'    Set RsData = Nothing
'    Screen.MousePointer = vbDefault
'
'End Function

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    Dim sql As String

    sql = "Delete  marakes_taklefa_temp where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
    Cn.Execute sql, , adExecuteNoRecords
    
    If Fg_Journal.Rows > 1 Then
        If Fg_Journal.Rows = 2 Then
            Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Fg_Journal.Rows > 1 Then
                If Me.Fg_Journal.Row <> Me.Fg_Journal.FixedRows - 1 Then
                    Me.Fg_Journal.RemoveItem (Me.Fg_Journal.Row)
                End If
            End If
        End If
    End If
            
    With Fg_Journal
        Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    End With

End Sub

Private Sub DcbAccount_Change()
DcbAccount_Click (0)
End Sub

Private Sub DcbAccount_Click(Area As Integer)
TxtAccount.Text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount.BoundText)
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        DcboCreditSide.BoundText = DcbAccount.BoundText
    End If
End Sub

Private Sub DcbAccount_KeyUp(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyF3 Then
                      Account_search.show
                     Account_search.case_id = 350054

                   End If
End Sub

Private Sub DcbBranch_Change()
DcbBranch_Click (0)
End Sub

Private Sub DcbBranch_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
TxtSerial1.Text = ""
TxtSerial.Text = ""
End If
End Sub

Private Sub DcbIqara_KeyUp(KeyCode As Integer, Shift As Integer)

                If KeyCode = vbKeyF3 Then
                    FrmAqarSearch.show
                    FrmAqarSearch.m_RetrunType = 2
                        
                End If
End Sub

Private Sub DcbIqara2_Change()
DcbIqara2_Click (0)
End Sub

Private Sub DcbIqara2_Click(Area As Integer)
      If val(DcbIqara2.BoundText) = 0 Then: Exit Sub
Dim EmpCode  As String
Dim ownerid As Double

GetIqarCode , , DcbIqara2.BoundText, EmpCode, ownerid
    Me.TxtSearch2.Text = EmpCode
    dcsupplier.BoundText = ownerid
End Sub

Private Sub DcboBankName_Change()
    On Error Resume Next

    If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        '    Me.DcboCreditSide.BoundText = "a2a3a2"
    
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value

    End If

End Sub

Private Sub DcboBox_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    End If

End Sub

Private Sub DcboBox_Click(Area As Integer)
    DcboBox_Change
End Sub

Private Sub DcbUnitNo_Click(Area As Integer)
If val(Me.DcbIqara.BoundText) = 0 Then
MsgBox "ÌÃ» «Œ Ì«— «·⁄ﬁ«—«Ê·«"
DcbIqara.SetFocus
Exit Sub
End If

If val(Me.DcbUnitType.BoundText) = 0 Then
MsgBox "ÌÃ» «Œ Ì«— ‰Ê⁄ «·ÊÕœÂ «Ê·«"
DcbUnitType.SetFocus
Exit Sub
End If
If val(Me.DcbUnitNo.BoundText) <> 0 Then
With Fg_Journal
.TextMatrix(.Rows - 1, .ColIndex("uintid")) = Me.DcbUnitNo.BoundText
.TextMatrix(.Rows - 1, .ColIndex("unitno")) = Me.DcbUnitNo.Text
.TextMatrix(.Rows - 1, .ColIndex("type")) = Me.DcbUnitType.BoundText
.TextMatrix(.Rows - 1, .ColIndex("unittype")) = Me.DcbUnitType.Text
.TextMatrix(.Rows - 1, .ColIndex("iqarid")) = Me.DcbIqara.BoundText
.TextMatrix(.Rows - 1, .ColIndex("iqar")) = Me.DcbIqara.Text
.Rows = .Rows + 1
End With
End If
End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 4
    End If

End Sub

Private Sub dcproject_Change()

    If DCproject.Text = "" Then
        VSFlexGrid1.Visible = False
        Me.Fg_Journal.Visible = True
    End If
 
End Sub

Private Sub dcproject_Click(Area As Integer)

    If SystemOptions.gldetails_or_gl_general = 0 Then 'Õ”«»«  «·„‘—Ê⁄
        VSFlexGrid1.Visible = True
        Me.Fg_Journal.Visible = False
    Else
        VSFlexGrid1.Visible = False
        Me.Fg_Journal.Visible = True
    End If

End Sub

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim Rs3 As ADODB.Recordset
    With Fg_Journal

        Select Case .ColKey(Col)
        Case "Vatyo"
        If val(.TextMatrix(Row, .ColIndex("Vatyo"))) = 0 Then
        .TextMatrix(Row, .ColIndex("Vat")) = 0
        If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) <> 0 Then
        .TextMatrix(Row, .ColIndex("value")) = val(.TextMatrix(Row, .ColIndex("PriceTotal")))
        End If
        If .Rows > Row Then
        If val(.TextMatrix(Row + 1, .ColIndex("FlgVat"))) = 1 Then
        .RemoveItem Row + 1
        End If
        End If
        End If
         Case "PriceTotal"
                AddVATExp Row
            Case "branch_name"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("BrnchID")) = StrAccountCode
                AddVATExp Row
          Case "project"
         
                If val(.TextMatrix(Row, .ColIndex("projectid"))) <> 0 Then
               StrSQL = "Select Fullcode from  Projects where ID =" & val(.TextMatrix(Row, .ColIndex("projectid"))) & ""
               Set Rs3 = New ADODB.Recordset
               Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
               If Rs3.RecordCount > 0 Then
               .TextMatrix(Row, .ColIndex("ProjectCode")) = IIf(IsNull(Rs3("Fullcode").value), "", Rs3("Fullcode").value)
               Else
               .TextMatrix(Row, .ColIndex("ProjectCode")) = ""
               End If
               End If
               AddVATExp Row
         Case "ProjectCode"
               If .TextMatrix(Row, .ColIndex("ProjectCode")) <> "" Then
               StrSQL = "select * from  Projects where Fullcode ='" & .TextMatrix(Row, .ColIndex("ProjectCode")) & "'"
                Set Rs3 = New ADODB.Recordset
               Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
               If Rs3.RecordCount > 0 Then
               .TextMatrix(Row, .ColIndex("projectid")) = IIf(IsNull(Rs3("ID").value), "", Rs3("ID").value)
               If SystemOptions.UserInterface = ArabicInterface Then
               .TextMatrix(Row, .ColIndex("project")) = IIf(IsNull(Rs3("Project_name").value), "", Rs3("Project_name").value)
               Else
               .TextMatrix(Row, .ColIndex("project")) = IIf(IsNull(Rs3("Project_nameE").value), "", Rs3("Project_nameE").value)
               End If
               Else
               .TextMatrix(Row, .ColIndex("projectid")) = 0
               .TextMatrix(Row, .ColIndex("project")) = ""
               End If
               End If
               AddVATExp Row
                  Case "pand"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("pandid")) = StrAccountCode
                AddVATExp Row
                  Case "oper"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("operid")) = StrAccountCode
                AddVATExp Row
         Case "iqar"
        ' If val(.TextMatrix(Row, .ColIndex("iqarid"))) <> 0 Then
'                  If val(.TextMatrix(Row, .ColIndex("iqarid"))) <> 0 Then
                
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("iqarid")) = StrAccountCode
                 DcbIqara2.BoundText = val(Fg_Journal.TextMatrix(Row, Fg_Journal.ColIndex("iqarid")))
                 TxtValue.Text = Fg_Journal.TextMatrix(Row, Fg_Journal.ColIndex("value"))
                TxtRemarks.Text = Fg_Journal.TextMatrix(Row, .ColIndex("des"))
        '        End If
               If SystemOptions.NoCreatJLInRentContract Then
                    FillData Row
                End If
              '  StrAccountCode = .ComboData
              '   LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("iqarid"), False, True)
               ' .TextMatrix(Row, .ColIndex("iqarid")) = StrAccountCode
                AddVATExp Row
         Case "unittype"
                    
                StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("type"), False, True)
                .TextMatrix(Row, .ColIndex("type")) = StrAccountCode
                AddVATExp Row
            Case "unitno"
                    
                StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("uintid"), False, True)
                .TextMatrix(Row, .ColIndex("uintid")) = StrAccountCode
 AddVATExp Row
            Case "ExpensesID"
              
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
                AddVATExp Row
           Case "Account_Serial"

                StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(.TextMatrix(Row, Col)) & "'"
                StrSQL = StrSQL & GetAccountByBarnchUser
                StrSQL = StrSQL & GetAccountCodeHiding
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                   

                    .TextMatrix(Row, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    Else
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                    
                    End If
                  End If
.TextMatrix(Row, .ColIndex("uintid")) = Me.DcbUnitNo.BoundText
.TextMatrix(Row, .ColIndex("unitno")) = Me.DcbUnitNo.Text
.TextMatrix(Row, .ColIndex("type")) = Me.DcbUnitType.BoundText
.TextMatrix(Row, .ColIndex("unittype")) = Me.DcbUnitType.Text
.TextMatrix(Row, .ColIndex("iqarid")) = Me.DcbIqara.BoundText
.TextMatrix(Row, .ColIndex("iqar")) = Me.DcbIqara.Text
                  
    AddVATExp Row
    ''///////////
             Case "AccountCode"

                StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Serial, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Code='" & Trim(.TextMatrix(Row, Col)) & "'"
                StrSQL = StrSQL & GetAccountByBarnchUser
                StrSQL = StrSQL & GetAccountCodeHiding
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                   

                    .TextMatrix(Row, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    Else
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                    
                    End If
                  End If
                  
                  

   ' AddVATExp Row
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
     
                StrAccountCode = .ComboData
            
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
                .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)

               ' If SystemOptions.UserInterface = ArabicInterface Then
               '     StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
               ' Else
               '     StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
               ' End If
            '
            '    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
            '    If rs.RecordCount > 0 Then
            '        .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
            '    Else
            '        .TextMatrix(Row, .ColIndex("des")) = ""
            '    End If
''//
.TextMatrix(Row, .ColIndex("uintid")) = Me.DcbUnitNo.BoundText
.TextMatrix(Row, .ColIndex("unitno")) = Me.DcbUnitNo.Text
.TextMatrix(Row, .ColIndex("type")) = Me.DcbUnitType.BoundText
.TextMatrix(Row, .ColIndex("unittype")) = Me.DcbUnitType.Text
.TextMatrix(Row, .ColIndex("iqarid")) = Me.DcbIqara.BoundText
.TextMatrix(Row, .ColIndex("iqar")) = Me.DcbIqara.Text
AddVATExp Row
'//
            Case "value", "opr_fullcode"
                Dim sgl As String
                Dim project_id As Integer
                project_id = get_project_id(DCproject.BoundText, "expanses_account")
                
                If checkitems(project_id, .TextMatrix(Row, .ColIndex("opr_fullcode")), val(.TextMatrix(Row, .ColIndex("Value")))) = False Then
                    .TextMatrix(Row, .ColIndex("Value")) = 0
                End If
               
                Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
                sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                Cn.Execute sgl, , adExecuteNoRecords
        AddVATExp Row
                '   Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        ' AddVATExp Row

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .Rows - 1 And .ColKey(Col) <> "AccountCode" Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

   

                
        Select Case .ColKey(Col)
        Case "iqar"
        If SystemOptions.cantCahngeAkarinExpenses = True Then
            Cancel = True
        End If
        
        
       Case "AccountName"
        If SystemOptions.cantCahngeAkarinExpenses = True Then
            Cancel = True
        End If
        
        
              Case "Account_Serial"
        If SystemOptions.cantCahngeAkarinExpenses = True Then
            Cancel = True
        End If
        
        
        Case "Vat"
                 Cancel = True
        Case "Vatyo"
              If val(.TextMatrix(Row, .ColIndex("ForcedFlg"))) = 1 Then
                 Cancel = True
              Else
              .ComboList = ""
              End If
              Case "PriceTotal"
                .ComboList = ""
         Case "ProjectCode"
                .ComboList = ""

            Case "value", "Account_Serial"
                .ComboList = ""

            Case "Unitss"
                .ComboList = ""
            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
        
                '  Cancel = True
Case "iqar"
            If SystemOptions.cantCahngeAkarinExpenses = True Then
            Cancel = True
            End If

        End Select

    End With

End Sub

Private Sub Fg_Journal_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
  With Fg_Journal
   Select Case .ColKey(Col)
      Case "unitno"
    
       If val(.TextMatrix(Row, .ColIndex("type"))) <> 0 And val(.TextMatrix(Row, .ColIndex("iqarid"))) <> 0 Then
           LngRow = Row
           LngCol = Col
           FrmIqarUnitNo.TypIndex = 1
           Load FrmIqarUnitNo
           FrmIqarUnitNo.TypIndex = 1
           FrmIqarUnitNo.show vbModal
           
        Else
       If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄ﬁ«— Ê«·‰Ê⁄"
       Else
        MsgBox "Please Select Real Estate"
       End If
       Exit Sub
        End If
  End Select
End With
End Sub

Private Sub Fg_Journal_DblClick()
    Exit Sub
  
    Static lNoteRow&, lNoteCol&, r&, c&

    With Fg_Journal
        ' clicking? no work
        'If Button <> 0 Then Exit Sub
        ' get mouse coordinates
        r = Fg_Journal.Row
        c = Fg_Journal.Col

        If Fg_Journal.ColKey(c) <> "Des" Then
         CboDes.Visible = False
            Exit Sub
        End If

        If Fg_Journal.TextMatrix(r, c) = "" Then
            'Exit Sub
        End If

        If .TextMatrix(r, .ColIndex("AccountCode")) = "" Then
            Exit Sub
        End If

        ' same cell or neighbour? no work
        '    If r = lNoteRow And C = lNoteCol Then Exit Sub
        '    If r = lNoteRow And C = lNoteCol + 1 Then Exit Sub

        ' other cell, hide current note, if any
        If lNoteRow >= 0 And lNoteCol >= 0 Then
            Fg_Journal.SetFocus
            lNoteRow = -1
            lNoteCol = -1
        End If

        ' no note to show? then bail out
        If r <= 0 Or c <= 0 Then Exit Sub
        If typename(Fg_Journal.Cell(flexcpData, r, c)) <> "String" Then
            TxtDes.Text = ""
        Else
            '
            TxtDes.Text = Fg_Journal.Cell(flexcpData, r, c)
        End If

        ' show new note
        CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
      CboDes.Visible = True
          CboDes.ZOrder 0
         CboDes.SetFocus
        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c
    End With

End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    With Fg_Journal

        Select Case .ColKey(.Col)

            Case "Order_No"
                           
                If KeyCode = vbKeyF3 Then

                    Order_no_search.show
                       Order_no_search.RetrunType = 4
                End If

            Case "AccountName", "Account_Serial"
                   If KeyCode = vbKeyF3 Then
                        If SystemOptions.cantCahngeAkarinExpenses = True Then
                            Exit Sub
                            Else
                      Account_search.show
                     Account_search.case_id = 350053
                          End If

                   End If
                   
         '       If KeyCode = vbKeyF3 Then
         '           FrmExpensesSearch.show
         '           FrmExpensesSearch.RetrunType = 1915
                        
         '       End If
             Case "iqar"
LongRow = .Row
                If KeyCode = vbKeyF3 Then
                      If SystemOptions.cantCahngeAkarinExpenses = True Then
                            Exit Sub
                      Else
                                          FrmAqarSearch.show
                    FrmAqarSearch.m_RetrunType = 3
               
               
                          End If
                          
         
                End If
 
        End Select

    End With

End Sub
Sub DeleteGridCurrRowExp(Optional CurrRow As Long)
Dim i As Integer
With Fg_Journal
i = .Rows
Do
i = i - 1
If val(.TextMatrix(i, .ColIndex("CurrRow"))) = CurrRow Then
.RemoveItem i
End If
Loop While i > 1
End With
End Sub
Sub AddVATExp(Optional ByRef Row As Long)
If True = True Then
Dim ForcedFlg As Integer
Dim valuee As Double
Dim AccountVATDept As String
Dim i As Integer
Dim k As Integer
Dim ClsAcc  As New ClsAccounts
With Fg_Journal

.TextMatrix(Row, .ColIndex("Vatyo")) = PercentgValueAddedAccount(XPDtbTrans.value, .TextMatrix(Row, .ColIndex("AccountCode")), val(DcbBranch.BoundText), ForcedFlg)
.TextMatrix(Row, .ColIndex("Rate")) = val(.TextMatrix(Row, .ColIndex("Vatyo"))) / 100 + 1
If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) > 0 And val(.TextMatrix(Row, .ColIndex("Rate"))) > 0 Then
.TextMatrix(Row, .ColIndex("value")) = Round(val(.TextMatrix(Row, .ColIndex("PriceTotal"))) / val(.TextMatrix(Row, .ColIndex("Rate"))), 2)
End If
valuee = val(.TextMatrix(Row, .ColIndex("Value")))
.TextMatrix(Row, .ColIndex("ForcedFlg")) = ForcedFlg
.TextMatrix(Row, .ColIndex("Vat")) = Round((val(.TextMatrix(Row, .ColIndex("Vatyo"))) * valuee) / 100, 2)
GetValueAddedAccount XPDtbTrans.value, AccountVATDept
If AccountVATDept = "" And val(.TextMatrix(Row, .ColIndex("Vat"))) > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· «·Õ”«» «·„œÌ‰ ›Ì ‘«‘… «⁄œ«œ  «·›« "
Else
MsgBox "Please Enter Account In VAT Settings"
End If
.TextMatrix(Row, .ColIndex("Vat")) = 0
.TextMatrix(Row, .ColIndex("Vatyo")) = 0
Exit Sub
End If
If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) = 0 Then
.TextMatrix(Row, .ColIndex("PriceTotal")) = val(.TextMatrix(Row, .ColIndex("Vat"))) + val(.TextMatrix(Row, .ColIndex("value")))
End If
''/////////////
If val(.TextMatrix(Row, .ColIndex("Vat"))) > 0 Then
   If Not .TextMatrix(.Row, .ColIndex("AccountCode")) = "" Then
   DeleteGridCurrRowExp Row
   For i = 1 To 1
         .AddItem " ", Row + 1
  k = Row + i
.TextMatrix(k, .ColIndex("CurrRow")) = Row
 
If i = 1 Then
.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(AccountVATDept)
.TextMatrix(k, .ColIndex("AccountName")) = Get_Account_name(, AccountVATDept)
.TextMatrix(k, .ColIndex("AccountCode")) = AccountVATDept
.TextMatrix(k, .ColIndex("Value")) = .TextMatrix(Row, .ColIndex("Vat"))
Else
.TextMatrix(k, .ColIndex("AccountCode")) = DcboCreditSide.BoundText
.TextMatrix(k, .ColIndex("AccountName")) = Get_Account_name(, DcboCreditSide.BoundText)
.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(DcboCreditSide.BoundText)
.TextMatrix(k, .ColIndex("Value")) = .TextMatrix(Row, .ColIndex("Vat"))
End If
.TextMatrix(k, .ColIndex("LineNo1")) = setfoxy_Line
.TextMatrix(k, .ColIndex("Des")) = .TextMatrix(Row, .ColIndex("Des")) & " " & " ﬁÌ„… „÷«›…"

.TextMatrix(k, .ColIndex("FlgVat")) = 1
Next i
End If
End If
End With
End If
End Sub
Public Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                ByVal Col As Long, _
                                Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String
    Dim Rs3 As ADODB.Recordset
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)
                    Case "project"

               
                StrSQL = " SELECT  LTRIM(RTRIM( Project_name )) as Project_name ,Project_nameE, id From dbo.Projects  "
                    If SystemOptions.UserInterface = ArabicInterface Then
    
        StrSQL = StrSQL & " where  not (Project_name is null)and Project_name<>N'""'"
    Else
        
        StrSQL = StrSQL & " where  not (Project_nameE is null)and Project_nameE<>N'""'"
    End If
    StrSQL = StrSQL & " and (Not (Fullcode Is Null))"
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL & " order by  Project_name"
    Else
    StrSQL = StrSQL & " order by  Project_nameE"
    End If
    
                Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "Project_name", "id")
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
             Case "pand"
             If .TextMatrix(Row, .ColIndex("projectid")) = "" Then
             MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
             Exit Sub
             End If

                StrSQL = " SELECT     des, oprid From projects_des "
                 StrSQL = StrSQL & "    Where (project_id =" & val(.TextMatrix(Row, .ColIndex("projectid"))) & ")"
                Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "des", "oprid")
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                  Case "oper"
                   
If .TextMatrix(Row, .ColIndex("projectid")) = "" Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
.TextMatrix(Row, .ColIndex("oper")) = ""
Exit Sub
End If
If .TextMatrix(Row, .ColIndex("pandid")) = "" Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·»‰œ «Ê·«"
.TextMatrix(Row, .ColIndex("oper")) = ""
Exit Sub
End If
           
                If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = "SELECT     dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "    FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "      dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
               Else
               StrSQL = "SELECT     dbo.TblProcessDEF.ProcessNameE, dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "    FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "      dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEF"
                End If
               StrSQL = StrSQL & "    Where (ProjectDes_ID = " & val(.TextMatrix(Row, .ColIndex("pandid"))) & ") And (project_id = " & val(.TextMatrix(Row, .ColIndex("projectid"))) & ")"
               Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "ProcessName", "TblProcessDEFID")
                    Else
                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "ProcessNameE", "TblProcessDEFID")
                    End If
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
                
             Case "unitno"
                 .ColComboList(.ColIndex("unitno")) = "..."
        Case "iqar"
                StrSQL = "SELECT  Aqarid,aqarname from TblAqar "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
                    StrComboList = Fg_Journal.BuildComboList(rs, "aqarname", "Aqarid")
              
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
      
            Case "unittype"
     StrSQL = "SELECT  * from TblAkarUnit "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
                Case "unitno"
                Dim unittype As Integer
               Dim Aqarid As Integer
               If val(.TextMatrix(Row, .ColIndex("iqarid"))) <> 0 Then
                Aqarid = .TextMatrix(Row, .ColIndex("iqarid"))
                Else
                MsgBox "ÌÃ» ≈Œ Ì«—  «·⁄ﬁ«— «Ê·«"
                Exit Sub
                End If
                
               If val(.TextMatrix(Row, .ColIndex("type"))) <> 0 Then
                unittype = .TextMatrix(Row, .ColIndex("type"))
                Else
                MsgBox "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·ÊÕœÂ «Ê·«"
                Exit Sub
                End If
     StrSQL = "SELECT  * from TblAqarDetai where ( Aqarid =" & Aqarid & ")and(unittype=" & unittype & ") "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList = Fg_Journal.BuildComboList(rs, "unitno", "id")
             
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
    
                

         '   Case "AccountName"
              '  StrSQL = "select * from Expenses_accounts"
              '  rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
              '  StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
'
'                If StrComboList <> "" Then
'                    StrComboList = "|" & StrComboList
'                End If
'
'                .ComboList = StrComboList
            Case "AccountName"
               ' Exit Sub
                'Full Path Display
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '   If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                    
                    '   End If
        
                
                Else
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                 
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                 
              StrComboList = Fg_Journal.BuildComboList(rs, "RootName,ParentName,*FirstName", "Account_Code")
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
            Case "opr_fullcode"
                Dim project_id As Integer
                project_id = get_project_id(DCproject.BoundText, "expanses_account")

                If SystemOptions.Items_or_operation = 1 Then
                    StrSQL = "  select fullcode,name from terms_operations where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode,name", "fullcode")
                ElseIf SystemOptions.Items_or_operation = 0 Then
                    StrSQL = "  select fullcode,des from projects_des where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode,des", "fullcode")
         
                End If

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
         
        End Select

    End With

End Sub
Private Sub DcbIqara_Click(Area As Integer)
      If val(DcbIqara.BoundText) = 0 Then dcsupplier.BoundText = 0: Exit Sub

    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , DcbIqara.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.Text = EmpCode
    dcsupplier.BoundText = ownerid
    'DcbUnitType_Change
End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String

   ' On Error GoTo ErrTrap

    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
    fill_combo Me.DcCostCenter, StrSQL

       If SystemOptions.SpecialVersion = True Then
Cmd(10).Visible = False
'Fra(1).Visible = False
   End If
   
   
    Set TTD = New clstooltipdemand
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("FillData").Picture
    Resize_Form Me
    AddTip
    SetDtpickerDate XPDtbTrans
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetExpensesType XPCboExpensesType
    Dcombos.GetAccountingCodes Me.DcbAccount, True, False
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.XPCboExpensesType
    Dcombos.GetCustomersSuppliers 257, Me.dcsupplier
    Dcombos.GetIqar DcbIqara
    Dcombos.getAkarUnit Me.DcbUnitType
    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetBranches DcbBranch
    Dcombos.GetIqar DcbIqara2
    With Me.CboPayMentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "‘Ìﬂ"
        .AddItem "Õ”«»"
        .AddItem "ÕÊ«·…"
        .AddItem "„ ⁄œœ"
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    StrSQL = " select expanses_account,Project_name from projects  where not(expanses_account is null)"
    fill_combo DCproject, StrSQL

    Set rs = New ADODB.Recordset
    StrSQL = "select * From notes_all where notetype=3"
    StrSQL = StrSQL & " and Branch_NO in(" & Current_branchSql & ") and not (ToPriodDateH is null)"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
    'MsgBox ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    hide_logo = False

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
        Set rs = Nothing
    End If

    Set TTP = Nothing
    'Set EmpReport = Nothing
    TTD.Destroy
    Exit Sub
ErrTrap:
End Sub

Private Sub CboDes_ButtonClick(ByVal ButtonID As VDSCOMBOLibCtl.vdsButtonID, _
                               ByVal SpinningEnded As Boolean)

    If ButtonID = vdsDownArrow Then
        If CboDes.IsDropped = False Then
            If PicHeight > 0 Then
                PicDes.Height = PicHeight
                PicDes.Width = PicWidth
            Else
                PicDes.Width = CboDes.Width - 10
                PicDes.Height = CboDes.Height * 8
            End If

            Debug.Print PicHeight
            Debug.Print PicWidth
            TxtDes.Visible = True
            TxtDes.Text = Fg_Journal.Cell(flexcpData, Fg_Journal.Row, Fg_Journal.ColIndex("Des"))
            CboDes.DropDown PicDes.hwnd, vdsRightToLeft, vdsBottomToDown, vdsDownArrow, True, vdsSoftResize
            Debug.Print PicDes.Height & "Pic H " & "-----" & PicDes.Width & "Pic W"
        Else
            CboDes.CloseUp
        End If
    End If

End Sub

Private Sub CboDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys "{F4}"
    End If

End Sub

 Public Sub FillData(Optional ByVal mRow As Long = 0)
 Dim IarType As Integer
 Dim Account_Code_dynamic As String
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
IarType = AqarCommisionType(val(DcbIqara2.BoundText))
If IarType <> 1 And SystemOptions.OpenAccountAqar = True Then
           Account_Code_dynamic = get_account_code_branch(163, my_branch)
        If Account_Code_dynamic = "NO branch" Then
          MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            
            Else
           If Account_Code_dynamic = "NO account" Then
              MsgBox "·„ Ì „  ÕœÌœ Õ”«»     «Ì—«œ«  «·”⁄Ì Ê«·⁄„Ê·« ", vbCritical
             Exit Sub
    
           End If
        End If
 End If
With Me.Fg_Journal

If mRow = 0 And .Rows >= 2 And .TextMatrix(.Rows - 1, .ColIndex("AccountCode")) <> "" Then .Rows = .Rows + 1
Dim mRow2 As Long
If mRow = 0 Then mRow2 = .Rows - 1 Else mRow2 = mRow

.TextMatrix(mRow2, .ColIndex("iqarid")) = val(Me.DcbIqara2.BoundText)
.TextMatrix(mRow2, .ColIndex("iqar")) = Me.DcbIqara2.Text
.TextMatrix(mRow2, .ColIndex("value")) = val(TxtValue.Text)
.TextMatrix(mRow2, .ColIndex("des")) = TxtRemarks.Text

         If SystemOptions.OpenAccountAqar = False Then
          .TextMatrix(mRow2, .ColIndex("AccountCode")) = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(dcsupplier.BoundText))
         Else
           
            
            If IarType <> 0 Then
              .TextMatrix(mRow2, .ColIndex("AccountCode")) = GetAqarAcountCode(val(DcbIqara2.BoundText))
              Else
              .TextMatrix(mRow2, .ColIndex("AccountCode")) = get_account_code_branch(163, my_branch)
              End If
         End If
         
'.TextMatrix(.Rows - 1, .ColIndex("AccountCode")) = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(dcsupplier.BoundText))
Fg_Journal_AfterEdit mRow2, .ColIndex("AccountCode")
If mRow = 0 Then
   ' Fg_Journal_AfterEdit .Rows - 1, .ColIndex("iqar")
End If
End With
End If
End Sub

Private Sub FrmPriodDate_Change()
   If Me.TxtModFlg.Text <> "R" Then
     
    FrmPriodDateH.value = ToHijriDate(FrmPriodDate.value)
    
End If
End Sub

Private Sub FrmPriodDateH_LostFocus()
 If Me.TxtModFlg.Text <> "R" Then
             
             FrmPriodDate.value = ToGregorianDate(FrmPriodDateH.value)

               
        End If
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub PicDes_Resize()

    With PicDes
        LblDes.Move .ScaleLeft, .ScaleTop, .ScaleWidth, LblDes.Height
        TxtDes.Move .ScaleLeft, .ScaleTop + LblDes.Height, .ScaleWidth, .ScaleHeight - LblDes.Height
        '    PicHeight = PicDes.Height
        '    PicWidth = PicDes.Width
    End With

End Sub

Private Sub ToPriodDate_Change()
   If Me.TxtModFlg.Text <> "R" Then
     
    ToPriodDateH.value = ToHijriDate(ToPriodDate.value)
    
End If
End Sub

Private Sub ToPriodDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
             
             ToPriodDate.value = ToGregorianDate(ToPriodDateH.value)

               
        End If
End Sub

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Order_no_search.show
        Order_no_search.RetrunType = 0
    End If

End Sub

Private Sub TxtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
DcbAccount.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount.Text)
End If
End Sub

Private Sub TxtDes_LostFocus()
    PicHeight = PicDes.Height
    PicWidth = PicDes.Width
    CboDes.CloseUp
    CboDes.Visible = False
End Sub

Private Sub TxtDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyEscape Then
        PutData
        CboDes.CloseUp
    End If

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«·„’—Ê›« "
            Else
                Me.Caption = "Expenses"
            End If
        
            Me.VSFlexGrid1.Enabled = False
          '  Me.Fg_Journal.Enabled = False
            Frame1.Enabled = False
        
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            CmdRemove.Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            XPTxtVal.locked = True
            '        XPCboProfLevel.Locked = True
            '        XPTxtProfMail.Locked = True
            '        XPTxtPhone.Locked = True
            '        XPTxtMobile.Locked = True
            XPMTxtRemarks.locked = True
            XPCboExpensesType.locked = True
            Me.DcboBox.locked = True
            XPDtbTrans.Enabled = False

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
                Me.Caption = "«·„’—Ê›« (ÃœÌœ)"
            Else
                Me.Caption = "Expenses(New Record)"
            End If
        
            Me.VSFlexGrid1.Enabled = True
           ' Me.Fg_Journal.Enabled = True
            Frame1.Enabled = True
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            CmdRemove.Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            'Me.XPBtnMove(0).Enabled = False
            'Me.XPBtnMove(1).Enabled = False
            'Me.XPBtnMove(2).Enabled = False
            'Me.XPBtnMove(3).Enabled = False
        
            ' XPTxtVal.locked = False
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            XPCboExpensesType.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«·„’—Ê›« (  ⁄œÌ· )"
            Else
                Me.Caption = "Expenses(Edit Current Record)"
            End If
        
            Me.VSFlexGrid1.Enabled = True
          '  Me.Fg_Journal.Enabled = True
            Frame1.Enabled = True
        
            CmdRemove.Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            XPTxtVal.locked = False
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            XPCboExpensesType.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtSearch2_KeyPress(KeyAscii As Integer)
Dim EmpID As Double
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch2.Text, EmpID
        DcbIqara2.BoundText = EmpID
        DcbIqara2_Click (0)
    End If
End Sub

Public Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                 ByVal Col As Long)
    'check_cost_center
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
    Dim project_id As Integer

    With VSFlexGrid1

        Select Case .ColKey(Col)
    
            Case "Value", "opr_fullcode"
    
                project_id = get_project_id(DCproject.BoundText, "expanses_account")
    
                If checkitems(project_id, .TextMatrix(Row, .ColIndex("opr_fullcode")), val(.TextMatrix(Row, .ColIndex("Value")))) = False Then
                    .TextMatrix(Row, .ColIndex("Value")) = 0
                End If

                Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
   
            Case "DebitValue", "CreditValue"

                'remove destribution
     
                ' sgl = "update  marakes_taklefa_temp  set value=0 where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                ' Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))
            
                If .ColKey(Col) = "DebitValue" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0
                    ' Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                 
                    '    Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValue" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0
                    ' Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '     Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                End If

                .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
            
            Case "DebitValueE", "CreditValueE"
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))

                If .ColKey(Col) = "DebitValueE" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE"))
                    End If

                    '
                    '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '      Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '        Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValueE" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE"))
                    End If
                 
                    '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '      Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '        Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                End If
            
            Case "Account_Serial"
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
       
                    .TextMatrix(Row, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    Else
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                    
                    End If
                    
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), 0, rs("cost_center").value)
                    
                    Dim rs2 As ADODB.Recordset
                    Dim My_SQL As String

                    If IsNull(rs("currenct_code").value) Then

                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                    
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo xx
                    End If

                    My_SQL = "  select * from currency WHERE id=" & val(rs("currenct_code").value)

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), 1, rs2.Fields("rate").value)
xx:
                Else
                    GetMsgs 130, vbExclamation
                    .TextMatrix(Row, Col) = ""
                    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    Exit Sub
                End If

                rs.Close
                Set rs = Nothing

            Case "AccountName"
        
                'sgl = "Delete  marakes_taklefa_temp  where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                'Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)

                If LngRow <> -1 Then
                    'Msg = "Â–« «·Õ”«» „ÊÃÊœ „”»ﬁ«  ›Ï «·”ÿ— " & .TextMatrix(LngRow, .ColIndex("LineNo"))
                    'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    '.TextMatrix(Row, Col) = ""
                    '.TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    'Exit Sub
                End If

                Set ClsAcc = New ClsAccounts

                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                'End If
           
                Set ClsAcc = Nothing
            
                StrSQL = "SELECT ACCOUNTS.cost_center ,ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Name='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), vbFalse, rs("cost_center").value)
            
                    'Dim rs2 As ADODB.Recordset
                    'Dim My_SQL As String
                    If IsNull(rs("currenct_code").value) Then
                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo ll
                    End If

                    My_SQL = "  select * from currency WHERE id=" & rs("currenct_code").value

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), "", rs2.Fields("rate").value)
ll:
                End If

        End Select

        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ReLineGrid

    End With

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid1

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

        Select Case .ColKey(Col)

            Case "Value"
                .ComboList = ""

            Case "Account_Serial"
                .ComboList = ""
        
            Case "Des"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub VSFlexGrid1_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 80

    End If

End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String
    Dim Msg As String
    Dim project_id As Integer
    Dim whrstring As String

    With VSFlexGrid1

        Select Case .ColKey(Col)

            Case "opr_fullcode"
                    
                project_id = get_project_id(DCproject.BoundText, "expanses_account")

                If SystemOptions.Items_or_operation = 1 Then
                    StrSQL = "  select fullcode,name from terms_operations where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = .BuildComboList(rs, "fullcode,name", "fullcode")
                ElseIf SystemOptions.Items_or_operation = 0 Then
                    StrSQL = "  select fullcode,des from projects_des where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = .BuildComboList(rs, "fullcode,des", "fullcode")
         
                End If

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
            
            Case "AccountName"
         
                project_id = get_project_id(DCproject.BoundText, "expanses_account")
                whrstring = getProjectAccountwhereString(project_id)
                
                'Full Path Display
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '   If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If

                    '   End If
                    StrSQL = StrSQL + "and (" + whrstring + ")"
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                
                Else
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '     If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If

                    '     End If
                    StrSQL = StrSQL + "and (" + whrstring + ")"
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                
                End If
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "RootName,ParentName,*FirstName", "Account_Code")
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap
'
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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

 '   On Error GoTo ErrTrap
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 3
                 
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
          
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        'Lngid
          If Lngid <> 0 Then
              rs.find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst
             If rs.EOF Or rs.BOF Then
                  Exit Sub
              End If
         End If
    End If

    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    Else
        Me.DcCostCenter.BoundText = ""
    End If
     If Not IsNull(rs("Branch_NO").value) Then
        Me.DcbBranch.BoundText = IIf(rs("Branch_NO").value = "", "", rs("Branch_NO").value)
    Else
        Me.DcbBranch.BoundText = ""
    End If
    
''//
  FrmPriodDate.value = IIf(IsNull(rs("FrmPriodDate").value), Date, rs("FrmPriodDate").value)
    FrmPriodDateH.value = IIf(IsNull(rs("FrmPriodDateH").value), ToHijriDate(FrmPriodDate.value), rs("FrmPriodDateH").value)
    XPDtbTrans.value = IIf(IsNull(rs("ToPriodDate").value), Date, rs("ToPriodDate").value)
    ToPriodDateH.value = IIf(IsNull(rs("ToPriodDateH").value), ToHijriDate(ToPriodDate.value), rs("ToPriodDateH").value)
''//
If IsNull(rs("chkvat").value) Then
Me.chkvat.value = vbUnchecked
Else
If rs("chkvat").value = 0 Then
Me.chkvat.value = vbUnchecked
Else
Me.chkvat.value = vbChecked
End If


End If

    Me.Text1.Text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    Me.TXT_order_no.Text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
    TXT_A_NoteID.Text = IIf(IsNull(rs("A_NoteID").value), "", val(rs("A_NoteID").value))
    XPTxtID.Text = IIf(IsNull(rs("NoteID").value), "", val(rs("NoteID").value))
    Me.TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    XPTxtVal.Text = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
    XPMTxtRemarks.Text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
    txtto.Text = IIf(IsNull(rs("too").value), "", rs("too").value)
    txt_general_des.Text = IIf(IsNull(rs("general_des").value), "", rs("general_des").value)
''//
    Me.DcbIqara.BoundText = IIf(IsNull(rs("Iqar").value), "", rs("Iqar").value)
    Me.DcbUnitType.BoundText = IIf(IsNull(rs("UnitType").value), "", rs("UnitType").value)
    Me.dcsupplier.BoundText = IIf(IsNull(rs("Owner").value), "", rs("Owner").value)
    Me.DcbUnitNo.BoundText = IIf(IsNull(rs("UnitNo").value), "", rs("UnitNo").value)
    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    XPCboExpensesType.BoundText = IIf(IsNull(rs("ExpensesID").value), "", rs("ExpensesID").value)
Me.DcbIqara2.BoundText = IIf(IsNull(rs("IqarID2").value), "", rs("IqarID2").value)
    If IsNull(rs("NoteCashingType").value) Then
        Me.CboPayMentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
        DcbAccount.BoundText = ""
        TxtAccount.Text = ""
    ElseIf rs("NoteCashingType").value = 0 Then
        Me.CboPayMentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
        DcbAccount.BoundText = ""
        TxtAccount.Text = ""
    ElseIf rs("NoteCashingType").value = 1 Then
        Me.CboPayMentType.ListIndex = 1
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        DcbAccount.BoundText = ""
        TxtAccount.Text = ""
  ElseIf rs("NoteCashingType").value = 3 Then
        Me.CboPayMentType.ListIndex = 3
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        DcbAccount.BoundText = ""
        TxtAccount.Text = ""
  ElseIf rs("NoteCashingType").value = 2 Then
        Me.CboPayMentType.ListIndex = 2
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
        DcbAccount.BoundText = IIf(IsNull(rs("AccountCode2").value), "", rs("AccountCode2").value)
  ElseIf rs("NoteCashingType").value = 5 Then
        Me.CboPayMentType.ListIndex = 4
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
      
    End If

    CboPayMentType_Change

   ' Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
    DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))

    If rs("NoteCashingType").value = 0 Or rs("NoteCashingType").value = 5 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    ElseIf rs("NoteCashingType").value = 1 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.DcboBankName.BoundText))
    ElseIf rs("NoteCashingType").value = 3 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.DcboBankName.BoundText))
    ElseIf rs("NoteCashingType").value = 2 Then
        DcboCreditSide.BoundText = DcbAccount.BoundText
    End If

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.Txt_Numorder.Text = IIf(IsNull(rs("NumOrderInpot").value), "", rs("NumOrderInpot").value)
    Me.TxtSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    Me.TxtSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)

    Me.DCproject.BoundText = IIf(IsNull(rs("project_Expensen_account").value), "", rs("project_Expensen_account").value)


    Me.Fg_Journal.Visible = True

    '«·„’—Ê›« 
    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then

    '-----------------------------------------------------------------------------
     StrSQL = " SELECT     dbo.TblExpensesDet.ExpID, dbo.TblExpensesDet.ID, dbo.TblExpensesDet.Unitss, dbo.TblExpensesDet.StrUnit, dbo.TblExpensesDet.AccountCode, "
     StrSQL = StrSQL & "                  dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.TblExpensesDet.[value], dbo.TblExpensesDet.opr_fullcode,"
     StrSQL = StrSQL & "                  dbo.TblExpensesDet.order_no, dbo.TblExpensesDet.des, dbo.TblExpensesDet.TypeTrans, dbo.TblExpensesDet.iqarid, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname,"
     StrSQL = StrSQL & "                  dbo.TblExpensesDet.uintid, dbo.TblAqarDetai.unitno, dbo.TblExpensesDet.type, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblExpensesDet.Vatyo,"
     StrSQL = StrSQL & "                  dbo.TblExpensesDet.Vat, dbo.TblExpensesDet.PriceTotal, dbo.TblExpensesDet.Rate, dbo.TblExpensesDet.projectid, dbo.projects.Project_name,"
     StrSQL = StrSQL & "                  dbo.TblExpensesDet.pandid, dbo.projects_des.des AS PandDes, dbo.TblExpensesDet.operid, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE,"
     StrSQL = StrSQL & "                  dbo.TblExpensesDet.FlgVat , dbo.TblExpensesDet.ForcedFlg, dbo.TblExpensesDet.CurrRow , dbo.projects.Fullcode"
     StrSQL = StrSQL & "       FROM         dbo.TblExpensesDet LEFT OUTER JOIN"
     StrSQL = StrSQL & "                  dbo.TblProcessDEF ON dbo.TblExpensesDet.operid = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                  dbo.projects_des ON dbo.TblExpensesDet.pandid = dbo.projects_des.oprid LEFT OUTER JOIN"
     StrSQL = StrSQL & "                  dbo.projects ON dbo.TblExpensesDet.projectid = dbo.projects.id LEFT OUTER JOIN"
     StrSQL = StrSQL & "                  dbo.TblAkarUnit ON dbo.TblExpensesDet.type = dbo.TblAkarUnit.id LEFT OUTER JOIN"
     StrSQL = StrSQL & "                  dbo.TblAqarDetai ON dbo.TblExpensesDet.uintid = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
     StrSQL = StrSQL & "                  dbo.TblAqar ON dbo.TblExpensesDet.iqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
     StrSQL = StrSQL & "                  dbo.ACCOUNTS ON dbo.TblExpensesDet.AccountCode = dbo.ACCOUNTS.Account_Code"
     StrSQL = StrSQL & "       Where (dbo.TblExpensesDet.ExpID = " & XPTxtID.Text & ")"
  Set RsDev = New ADODB.Recordset
      RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            With Me.Fg_Journal

                If Me.DCproject.BoundText = "" Then
                    .Rows = .FixedRows + RsDev.RecordCount
                Else
                    .Rows = .FixedRows + RsDev.RecordCount - 1
                End If

                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, .ColIndex("ProjectCode")) = IIf(IsNull(RsDev("Fullcode").value), "", RsDev("Fullcode").value)
                    .TextMatrix(i, .ColIndex("PriceTotal")) = IIf(IsNull(RsDev("PriceTotal").value), 0, RsDev("PriceTotal").value)
                    .TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(RsDev("Vat").value), 0, RsDev("Vat").value)
                    .TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(RsDev("Vatyo").value), 0, RsDev("Vatyo").value)
                    .TextMatrix(i, .ColIndex("FlgVat")) = IIf(IsNull(RsDev("FlgVat").value), 0, RsDev("FlgVat").value)
                    .TextMatrix(i, .ColIndex("ForcedFlg")) = IIf(IsNull(RsDev("ForcedFlg").value), 0, RsDev("ForcedFlg").value)
                    .TextMatrix(i, .ColIndex("CurrRow")) = IIf(IsNull(RsDev("CurrRow").value), 0, RsDev("CurrRow").value)
                    .TextMatrix(i, .ColIndex("Rate")) = IIf(IsNull(RsDev("Rate").value), 0, RsDev("Rate").value)
                    .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(RsDev("Project_name").value), "", RsDev("Project_name").value)
                    .TextMatrix(i, .ColIndex("pand")) = IIf(IsNull(RsDev("PandDes").value), "", RsDev("PandDes").value)
                    .TextMatrix(i, .ColIndex("oper")) = IIf(IsNull(RsDev("ProcessName").value), "", RsDev("ProcessName").value)
                    .TextMatrix(i, .ColIndex("operid")) = IIf(IsNull(RsDev("operid").value), 0, RsDev("operid").value)
                    .TextMatrix(i, .ColIndex("projectid")) = IIf(IsNull(RsDev("projectid").value), 0, RsDev("projectid").value)
                    .TextMatrix(i, .ColIndex("pandid")) = IIf(IsNull(RsDev("pandid").value), 0, RsDev("pandid").value)
                    .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(RsDev("Account_Serial").value), "", RsDev("Account_Serial").value)
                    .TextMatrix(i, .ColIndex("StrUnit")) = IIf(IsNull(RsDev("StrUnit").value), "", RsDev("StrUnit").value)
                    .TextMatrix(i, .ColIndex("Unitss")) = IIf(IsNull(RsDev("Unitss").value), "", RsDev("Unitss").value)
                    .TextMatrix(i, .ColIndex("opr_fullcode")) = IIf(IsNull(RsDev("opr_fullcode").value), "", RsDev("opr_fullcode").value)
                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("AccountCode").value), "", RsDev("AccountCode").value)
                    .TextMatrix(i, .ColIndex("unitno")) = IIf(IsNull(RsDev("unitno").value), "", RsDev("unitno").value)
                    .TextMatrix(i, .ColIndex("type")) = IIf(IsNull(RsDev("type").value), 0, RsDev("type").value)
                    .TextMatrix(i, .ColIndex("uintid")) = IIf(IsNull(RsDev("uintid").value), "", RsDev("uintid").value)
                    .TextMatrix(i, .ColIndex("iqarid")) = IIf(IsNull(RsDev("iqarid").value), "", RsDev("iqarid").value)
                    .TextMatrix(i, .ColIndex("iqar")) = IIf(IsNull(RsDev("aqarname").value), "", RsDev("aqarname").value)
                    .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("des").value), "", RsDev("des").value)
                    .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
                    .TextMatrix(i, .ColIndex("Order_No")) = IIf(IsNull(RsDev("Order_No").value), "", RsDev("Order_No").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                     .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                    .TextMatrix(i, .ColIndex("unittype")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
                    Else
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
                    .TextMatrix(i, .ColIndex("unittype")) = IIf(IsNull(RsDev("namee").value), "", RsDev("namee").value)
                    End If
                    RsDev.MoveNext
                Next i
                If .Rows > 1 Then
                   Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
                End If
            End With

        End If

 '    RetriveBillBuyData
 FillGridWithDataPayment
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long
    Dim rs2 As ADODB.Recordset
    Dim astrSplit2tems2() As String
    Dim astrSplitItems() As String
    Dim sql As String
    Dim j As Integer
    Dim st As String
    Dim des As String
    Dim nElements As Integer
    Dim OtherInformation As New ClsGLOther
    'On Error GoTo ErrTrap
     Dim Posted As Integer
      
  
            If (CboPayMentType.ListIndex = 1 Or CboPayMentType.ListIndex = 3) And Me.DcboBankName.BoundText = "" Then
                            Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'                DcboBox.SetFocus
                 SendKeys "{F4}"
                Exit Sub

            End If
            
            If Me.DcboBox.BoundText = "" And (CboPayMentType.ListIndex = 4 Or CboPayMentType.ListIndex = 0) Then
                Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'                DcboBox.SetFocus
                 SendKeys "{F4}"
                Exit Sub
            End If
         If CboPayMentType.ListIndex = 4 And CheckMult_Cash() = True Then  '›Ì Õ«·Â «·„ ⁄œœ «· √ﬂœ „‰ ÿ—Ìﬁ… «·œ›⁄
           If val(TxtPayedValue2) <> val(XPTxtVal) Then
             Msg = " Õ·Ì· «·ﬁÌ„Â «·„œ›Ê⁄Â «·„ ⁄œœÂ €Ì— „ÿ«»ﬁ… ·ﬁÌ„Â «·”‰œ "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
             If Me.DcboBox.BoundText = "" And (CboPayMentType.ListIndex = 4 Or CboPayMentType.ListIndex = 0) Then
                Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBox.SetFocus
                 SendKeys "{F4}"
                Exit Sub
            End If
           End If
        End If
            If CheckAprroveScreen(Me.Name) = True Then
            Posted = 1
            Else
            Posted = 0
            End If
    
        If Me.TxtModFlg.Text = "N" Then
            If Me.CboPayMentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.Text), XPDtbTrans.value) = False Then
                        Exit Sub
                    End If
                End If
            End If

        ElseIf Me.TxtModFlg.Text = "E" Then

            If Me.CboPayMentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.Text), XPDtbTrans.value, , , val(Me.XPTxtID.Text)) = False Then
                        Exit Sub
                    End If
                End If
            End If
        End If

        Dim xrow As Integer

        With Fg_Journal

            For xrow = .Rows - 1 To 2 Step -1

                If .TextMatrix(xrow, .ColIndex("AccountCode")) = "" Then

                    .Rows = .Rows - 1
                End If

            Next xrow

        End With
    
        With Me.VSFlexGrid1

            For xrow = .Rows - 1 To 2 Step -1

                If .TextMatrix(xrow, .ColIndex("AccountCode")) = "" Then

                    .Rows = .Rows - 1
                End If

            Next xrow

        End With

        If SystemOptions.gldetails_or_gl_general = 0 And Me.DCproject.BoundText <> "" Then
            GoTo xx
        End If

        Dim i As Integer

        With Fg_Journal

            For i = .FixedRows To .Rows - 1

                If .TextMatrix(i, .ColIndex("AccountCode")) = "" Then
                    '////////////////////////////////////////notes
               
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·«Ì ÌÊÃœ „’—Ê› ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                    Else
                        MsgBox "Select Expenses in line no" & i, vbCritical
                    End If

                    Exit Sub
              
                End If
        
            Next i

        End With

        With Fg_Journal

            For i = .FixedRows To .Rows - 1

                If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Then
                    '////////////////////////////////////////notes
               
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·«Ì ÌÊÃœ ﬁÌ„… ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                    Else
                        MsgBox "Enter Value in line no" & i, vbCritical
                    End If
               
                    Exit Sub
                End If
        
            Next i

        End With
          Dim ISVAT As Boolean
    ISVAT = False
         With Fg_Journal
    For i = .FixedRows To .Rows - 1
      If val(.TextMatrix(i, .ColIndex("Vat"))) > 0 Then
      ISVAT = True
      End If
     Next i
 End With
 
Dim AccountVATDept As String
If ISVAT = True And True = True Then
If GetValueAddedAccount(XPDtbTrans.value, AccountVATDept) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›…"
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If

xx:
        calcnets     '-------------------------------------------------------------------------------------------
  
        '-------------------------------------------------------------------------------------------
        Cn.BeginTrans
        BeginTrans = True
        Dim A_NoteID As Long

        '///////////////NOTESALL
        If TxtModFlg.Text = "N" Then
            XPTxtID.Text = CStr(new_id("notes_all", "NoteID", "", True))
            Me.TxtNoteSerial.Text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=3"))
            rs.AddNew
            rs("NoteID").value = val(XPTxtID.Text)
        ElseIf Me.TxtModFlg.Text = "E" Then
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where notes_all=" & val(XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            Cn.Execute " Delete from TblExpUnitNo where  ExpID =" & val(XPTxtID.Text)
            Cn.Execute " Delete from TblExpensesDet where  ExpID =" & val(XPTxtID.Text)
            StrSQL = "Delete From notes Where notes_all=" & val(XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
         StrSQL = "Delete From TblMultuPayment Where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
            If DcCostCenter.BoundText <> "" Then
                StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
        
            StrSQL = "Delete From ExpensesDetails Where NoteSerial1='" & Me.TxtSerial1.Text & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
        
        End If
    ''//
     rs("FrmPriodDate").value = Me.FrmPriodDate.value
        rs("FrmPriodDateH").value = Me.FrmPriodDateH.value
        rs("ToPriodDate").value = Me.ToPriodDate.value
        rs("ToPriodDateH").value = Me.ToPriodDateH.value
    '''/
    
    If Me.chkvat.value = vbUnchecked Then
    rs("chkvat").value = 0
    Else
    rs("chkvat").value = 1
    End If
    
    
        '  Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("Branch_NO").value = IIf(Me.DcbBranch.BoundText = "", 0, val(Me.DcbBranch.BoundText))
        rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
        rs("foxy_no").value = val(Text1.Text)
        rs("order_no").value = TXT_order_no.Text
         rs("IqarID2").value = val(Me.DcbIqara2.BoundText)
        rs("Note_Value").value = IIf(XPTxtVal.Text = "", Null, XPTxtVal.Text)
        rs("Remark").value = IIf(XPMTxtRemarks.Text = "", "", Trim(XPMTxtRemarks.Text))
        rs("too").value = IIf(txtto.Text = "", "", Trim(txtto.Text))
        rs("general_des").value = IIf(txt_general_des.Text = "", "", Trim(txt_general_des.Text))
    ''////
    rs("Iqar").value = IIf(DcbIqara.Text = "", Null, DcbIqara.BoundText)
 rs("UnitNo").value = IIf(Me.DcbUnitNo.Text = "", Null, DcbUnitNo.BoundText)
 rs("UnitType").value = IIf(Me.DcbUnitType.Text = "", Null, DcbUnitType.BoundText)
    
  rs("Owner").value = IIf(Me.dcsupplier.Text = "", Null, dcsupplier.BoundText)
   
  
    '''////
    
        rs("CusID").value = Null
        rs("NoteType").value = 3
        rs("NoteDate").value = XPDtbTrans.value
        rs("NoteHijriDate").value = ToHijriDate(XPDtbTrans.value)
        rs("UserID").value = user_id
        rs("ExpensesID").value = IIf(XPCboExpensesType.Text = "", Null, XPCboExpensesType.BoundText)
   Dim lineno As Integer
 lineno = 1
            '«·ÿ—› «·„œÌ‰
       Dim newdes As String
       newdes = ""
               Line1 = setfoxy_Line
'      If val(DCboCashType.ListIndex) = 4 Then
'      If SystemOptions.UserInterface = ArabicInterface Then
'         newdes = newdes & " " & " ·„‘—Ê⁄ "
'         'newdes = newdes & DBCboClientName.Text
'         'newdes = newdes & " " & " ﬂÊœ «·„‘—Ê⁄ "
'         'newdes = newdes & TxtCustCode.Text
'       Else
'         ' newdes = newdes & " " & " project "
'         'newdes = newdes & DBCboClientName.Text
'         'newdes = newdes & " " & " Code "
'         'newdes = newdes & TxtCustCode.Text
'      End If
        If Me.CboPayMentType.ListIndex = 0 Then
            rs("BoxID").value = val(DcboBox.BoundText)
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
            rs("NoteCashingType").value = 0
            rs("AccountCode2").value = Null
        ElseIf Me.CboPayMentType.ListIndex = 1 Then
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("NoteCashingType").value = 1
            rs("AccountCode2").value = Null
    ElseIf Me.CboPayMentType.ListIndex = 3 Then
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("NoteCashingType").value = 3
            rs("AccountCode2").value = Null
        ElseIf Me.CboPayMentType.ListIndex = 2 Then
            rs("BoxID").value = Null
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
            rs("NoteCashingType").value = 2
            rs("AccountCode2").value = IIf(Me.DcbAccount.BoundText = "", "", Me.DcbAccount.BoundText)
         ElseIf val(CboPayMentType.ListIndex) = 4 Then
            rs("NoteCashingType").value = 5
            rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, DcboBox.BoundText)
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
        
        End If
    
        rs("project_Expensen_account").value = IIf(Me.DCproject.BoundText = "", "", Me.DCproject.BoundText)
        rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.Text) = "", Null, Trim$(Me.Txt_Numorder.Text))
        rs("Buy").value = "0"
        rs("Remark").value = XPMTxtRemarks.Text
        rs("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
        rs("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
        rs("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
        rs("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
         rs("Branch_NO").value = IIf(Me.DcbBranch.BoundText = "", 0, val(Me.DcbBranch.BoundText))
        rs("sanad_year").value = year(XPDtbTrans.value)
        rs("sanad_month").value = Month(XPDtbTrans.value)
        rs("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
    
        If Me.TxtModFlg.Text = "N" Then
            A_NoteID = CStr(new_id("Notes", "NoteID", "", True))
            TXT_A_NoteID.Text = A_NoteID
        Else
            A_NoteID = val(TXT_A_NoteID.Text)
        End If
    
        rs("A_NoteID").value = val(A_NoteID)
     
        rs.update
        Dim project_id As Integer
        project_id = get_project_id(DCproject.BoundText, "expanses_account")
        '/////////////////////Accounts Õ”«Ì« 
        Dim line_no  As Integer
        '„’—Ê›« 
        '//////////////////////////////////////Notes////////////////////////////////////
        Set RsNotes = New ADODB.Recordset
        sql = " select * from Notes where 1=-1"
        RsNotes.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
       
            Set RsDev = New ADODB.Recordset
            sql = "select * from DOUBLE_ENTREY_VOUCHERS where 1=-1"
            RsDev.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
             sql = "select * from TblExpensesDet where 1=-1"
             Set rs2 = New ADODB.Recordset
            rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
            '«·ÿ—› «·„œÌ‰
  
            Dim ExpensesID As Double
 
            Dim NoteID As String

            With Fg_Journal
                
                line_no = lineno + 1

                For i = .FixedRows To .Rows - 1
 
                    If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                        '////////////////////////////////////////notes
                
                        If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·« Ì„ﬂ‰ « „«„ ⁄„·Ì… «·Õ›Ÿ ·⁄œ„ «œŒ«· ﬁÌ„… ›Ì «·”ÿ— —ﬁ„  " & i - 1, vbCritical: GoTo ErrTrap
                            Else
                                MsgBox "Cant save no value in line no:  " & i - 1, vbCritical: GoTo ErrTrap
                            End If
               
                        End If

                        RsNotes.AddNew
                        NoteID = CStr(new_id("Notes", "NoteID", "", True))
                        RsNotes("NoteID").value = CStr(NoteID)
                
                        RsNotes("Note_Value").value = .TextMatrix(i, .ColIndex("value"))
                        RsNotes("Remark").value = txt_general_des.Text
                        RsNotes("foxy_no").value = val(Text1.Text)

                        If TXT_order_no.Text <> "" Then
                            RsNotes("order_no").value = TXT_order_no.Text
                        Else
                            RsNotes("order_no").value = IIf(.TextMatrix(i, .ColIndex("Order_No")) = "", Null, .TextMatrix(i, .ColIndex("Order_No")))
                        End If

                        RsNotes("CusID").value = Null
                        RsNotes("NoteType").value = 3
                        RsNotes("NoteDate").value = XPDtbTrans.value
                        RsNotes("UserID").value = user_id
                        RsNotes("ExpensesID").value = val(.TextMatrix(i, .ColIndex("ExpensesID")))
                        RsNotes("notes_all").value = Me.XPTxtID.Text
                        RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
                        RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
                        RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
                        RsNotes("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
                        RsNotes("sanad_year").value = year(XPDtbTrans.value)
                        RsNotes("sanad_month").value = Month(XPDtbTrans.value)
                        RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(.TextMatrix(I, .ColIndex("value")), "0.00"), 0, True, ".")
                        RsNotes("remark").value = txt_general_des.Text
                        RsNotes("Branch_NO").value = IIf(Me.DcbBranch.BoundText = "", 0, val(Me.DcbBranch.BoundText))
                        RsNotes.update
              
                        '////////////////////////////////////////notes
   
                 
                       '''////////////////////////////////////
                      rs2.AddNew
                      rs2("ExpID").value = val(Me.XPTxtID.Text)
                      rs2("Unitss").value = IIf(.TextMatrix(i, .ColIndex("Unitss")) = "", "", .TextMatrix(i, .ColIndex("Unitss")))
                      rs2("StrUnit").value = IIf(.TextMatrix(i, .ColIndex("StrUnit")) = "", "", .TextMatrix(i, .ColIndex("StrUnit")))
                      rs2("AccountCode").value = IIf(.TextMatrix(i, .ColIndex("AccountCode")) = "", "", .TextMatrix(i, .ColIndex("AccountCode")))
                      rs2("uintid").value = IIf(.TextMatrix(i, .ColIndex("uintid")) = "", 0, val(.TextMatrix(i, .ColIndex("uintid"))))
                      rs2("type").value = IIf(.TextMatrix(i, .ColIndex("type")) = "", 0, val(.TextMatrix(i, .ColIndex("type"))))
                      rs2("iqarid").value = IIf(.TextMatrix(i, .ColIndex("iqarid")) = "", 0, val(.TextMatrix(i, .ColIndex("iqarid"))))
                      rs2("value").value = IIf(.TextMatrix(i, .ColIndex("value")) = "", 0, val(.TextMatrix(i, .ColIndex("value"))))
                      rs2("opr_fullcode").value = IIf(.TextMatrix(i, .ColIndex("opr_fullcode")) = "", "", .TextMatrix(i, .ColIndex("opr_fullcode")))
                      rs2("order_no").value = IIf(.TextMatrix(i, .ColIndex("order_no")) = "", "", .TextMatrix(i, .ColIndex("order_no")))
                      rs2("des").value = IIf(.TextMatrix(i, .ColIndex("des")) = "", "", .TextMatrix(i, .ColIndex("des")))
                      rs2("FlgVat").value = IIf(.TextMatrix(i, .ColIndex("FlgVat")) = "", 0, val(.TextMatrix(i, .ColIndex("FlgVat"))))
                      rs2("ForcedFlg").value = IIf(.TextMatrix(i, .ColIndex("ForcedFlg")) = "", 0, val(.TextMatrix(i, .ColIndex("ForcedFlg"))))
                      rs2("CurrRow").value = IIf(.TextMatrix(i, .ColIndex("CurrRow")) = "", 0, val(.TextMatrix(i, .ColIndex("CurrRow"))))
                      rs2("projectid").value = IIf(.TextMatrix(i, .ColIndex("projectid")) = "", 0, val(.TextMatrix(i, .ColIndex("projectid"))))
                      rs2("pandid").value = IIf(.TextMatrix(i, .ColIndex("pandid")) = "", 0, val(.TextMatrix(i, .ColIndex("pandid"))))
                      rs2("operid").value = IIf(.TextMatrix(i, .ColIndex("operid")) = "", 0, val(.TextMatrix(i, .ColIndex("operid"))))
                      rs2("Rate").value = IIf(.TextMatrix(i, .ColIndex("Rate")) = "", 0, val(.TextMatrix(i, .ColIndex("Rate"))))
                      rs2("Vatyo").value = IIf(.TextMatrix(i, .ColIndex("Vatyo")) = "", 0, val(.TextMatrix(i, .ColIndex("Vatyo"))))
                      rs2("Vat").value = IIf(.TextMatrix(i, .ColIndex("Vat")) = "", 0, val(.TextMatrix(i, .ColIndex("Vat"))))
                      rs2("PriceTotal").value = IIf(.TextMatrix(i, .ColIndex("PriceTotal")) = "", 0, val(.TextMatrix(i, .ColIndex("PriceTotal"))))
                      
                      rs2.update
                      
SaveUnitNo rs2("id").value, i
       project_id = get_project_id(DCproject.BoundText, "expanses_account")
                       
                    OtherInformation.UnitString = .TextMatrix(i, .ColIndex("StrUnit"))
                    OtherInformation.Unitss = .TextMatrix(i, .ColIndex("Unitss"))
                    OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
                    OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
                    OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
                    OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
                    OtherInformation.PriceTotal = val(.TextMatrix(i, .ColIndex("PriceTotal")))
                    OtherInformation.Rate = val(.TextMatrix(i, .ColIndex("Rate")))
                    
      If .TextMatrix(i, .ColIndex("StrUnit")) <> "" Then
          st = .TextMatrix(i, .ColIndex("StrUnit"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         For j = 0 To nElements - 1
         astrSplit2tems2 = Split(astrSplitItems(j), "#")
         des = ""
         des = .TextMatrix(i, .ColIndex("des"))
         des = des & "  " & " "
         des = des & .TextMatrix(i, .ColIndex("iqar")) & "\ "
         des = des & " "
         des = des & .TextMatrix(i, .ColIndex("unittype")) & "\ "
         des = des & " "
         des = des & astrSplit2tems2(0) & " "
         
         
         If val(astrSplit2tems2(2)) <> 0 Then
                         LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                        If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), val(astrSplit2tems2(2)), 0, des + "  " + txtto.Text, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , val(.TextMatrix(i, Fg_Journal.ColIndex("LineNo1"))), val(Me.XPTxtID.Text), val(.TextMatrix(i, .ColIndex("projectid"))), .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , , , , val(Me.DcbBranch.BoundText), , , , , , , , , , val(.TextMatrix(i, Fg_Journal.ColIndex("iqarid"))), val(.TextMatrix(i, Fg_Journal.ColIndex("type"))), val(astrSplit2tems2(1)), , , , , , , , , , , , , , , OtherInformation) = False Then
                            GoTo ErrTrap
                    
                        End If

                        line_no = line_no + 1
         End If
         Next j
        Else
            If val(.TextMatrix(i, .ColIndex("value"))) <> 0 Then
                         LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                        If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), val(.TextMatrix(i, .ColIndex("value"))), 0, .TextMatrix(i, .ColIndex("des")) + "  " + txtto.Text, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , val(.TextMatrix(i, Fg_Journal.ColIndex("LineNo1"))), val(Me.XPTxtID.Text), val(.TextMatrix(i, .ColIndex("projectid"))), .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , , , , val(Me.DcbBranch.BoundText), , , , , , , , , , val(.TextMatrix(i, Fg_Journal.ColIndex("iqarid"))), val(.TextMatrix(i, Fg_Journal.ColIndex("type"))), , , , , , , , , , , , , , , , OtherInformation) = False Then
                            GoTo ErrTrap
                    
                        End If

                        line_no = line_no + 1
            End If
        End If
        
                    End If

                Next i

            End With
    
            '«·ÿ—› «·œ«∆‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
            RsNotes.AddNew
            NoteID = CStr(new_id("Notes", "NoteID", "", True))
            RsNotes("NoteID").value = CStr(NoteID)
            RsNotes("Branch_NO").value = IIf(Me.DcbBranch.BoundText = "", 0, val(Me.DcbBranch.BoundText))
            RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0)
            RsNotes("Remark").value = txt_general_des.Text 'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
            RsNotes("foxy_no").value = val(Text1.Text)

            If Me.CboPayMentType.ListIndex = 0 Then
                RsNotes("BoxID").value = val(DcboBox.BoundText)
                RsNotes("BankID").value = Null
                RsNotes("ChqueNum").value = Null
                RsNotes("DueDate").value = Null
                RsNotes("NoteCashingType").value = 0
            ElseIf Me.CboPayMentType.ListIndex = 1 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 1
            ElseIf Me.CboPayMentType.ListIndex = 2 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = Null
                RsNotes("ChqueNum").value = Null
                RsNotes("DueDate").value = Null
                RsNotes("NoteCashingType").value = 2
            ElseIf Me.CboPayMentType.ListIndex = 3 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 3
            ElseIf Me.CboPayMentType.ListIndex = 4 Then
            
             
              
               
                
                RsNotes("NoteCashingType").value = 5
                RsNotes("BoxID").value = IIf(DcboBox.BoundText = "", Null, DcboBox.BoundText)
                RsNotes("BankID").value = Null
                RsNotes("ChqueNum").value = Null
                RsNotes("DueDate").value = Null
            End If
                        
            '                       If txt_ORDER_NO.text <> "" Then
            '           RsNotes("order_no").value = txt_ORDER_NO.text
            '       Else
            '        RsNotes("order_no").value = IIf(Me.Fg_Journal.TextMatrix(i, .ColIndex("Order_No")) = "", Null, .TextMatrix(i, .ColIndex("Order_No")))
            '       End If
            
            RsNotes("CusID").value = Null
            RsNotes("NoteType").value = 3
            RsNotes("NoteDate").value = XPDtbTrans.value
            RsNotes("UserID").value = user_id
            ' rsnotes("ExpensesID").value = .TextMatrix(I, .ColIndex("ExpensesID"))
            RsNotes("notes_all").value = Me.XPTxtID.Text
            RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
            RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
            RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
            RsNotes("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
            RsNotes("sanad_year").value = year(XPDtbTrans.value)
            RsNotes("sanad_month").value = Month(XPDtbTrans.value)
            RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
            RsNotes("Remark").value = txt_general_des.Text 'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
            RsNotes.update
    
            '«·ÿ—› «·œ«∆‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
      
            If Me.CboPayMentType.ListIndex <> 4 Then
            Dim TotalValue As Double
            Dim VATValue As Double
            Dim X As Integer
            Dim VatsalesAccount As String
            
            If Me.chkvat = vbUnchecked Then
            TotalValue = IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0)
            Else
            TotalValue = IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0)
            TotalValue = TotalValue / 1.05
            VATValue = TotalValue * 0.05
            X = GetValueAddedAccount(XPDtbTrans.value, , VatsalesAccount, 1, 21)
            End If
            
            RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = DcboCreditSide.BoundText
            RsDev("branch_id").value = val(Me.DcbBranch.BoundText)
      
            RsDev("Value").value = TotalValue ' IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = txtto ' .TextMatrix(I, .ColIndex("des"))
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
            RsNotes("Remark").value = txt_general_des.Text 'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.Text
            RsDev.update
     
     
     If Me.chkvat = vbChecked Then
    'ﬁÌœ ÷—Ì»Â «·«Ì—«œ
    line_no = line_no + 1
            RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = VatsalesAccount
            RsDev("branch_id").value = val(Me.DcbBranch.BoundText)
      
            RsDev("Value").value = VATValue ' IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = txtto ' .TextMatrix(I, .ColIndex("des"))
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
            RsNotes("Remark").value = txt_general_des.Text 'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.Text
            RsDev.update
     End If
     
     
     Else '„ ⁄œœ
               PGMultyPayment val(NoteID), line_no, Line1, XPMTxtRemarks.Text & CHR(13) & newdes, Posted
          
            
     End If
            'GoTo ll
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            If Me.DCproject.BoundText <> "" Then
                '«·ÿ—› «·„œÌ‰   „’—Ê›«  «·„‘—Ê⁄
                RsNotes.AddNew
                NoteID = CStr(new_id("Notes", "NoteID", "", True))
                RsNotes("NoteID").value = CStr(NoteID)
          
                RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0)
                RsNotes("Remark").value = txt_general_des.Text

                If Me.CboPayMentType.ListIndex = 0 Then
                    RsNotes("BoxID").value = val(DcboBox.BoundText)
                    RsNotes("BankID").value = Null
                    RsNotes("ChqueNum").value = Null
                    RsNotes("DueDate").value = Null
                    RsNotes("NoteCashingType").value = 0
                ElseIf Me.CboPayMentType.ListIndex = 1 Then
                    RsNotes("BoxID").value = Null
                    RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                    RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                    RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                    RsNotes("NoteCashingType").value = 1
                End If
               
                ' If txt_ORDER_NO.text <> "" Then
                '       RsNotes("order_no").value = txt_ORDER_NO.text
                '   Else
                '   RsNotes("order_no").value = IIf(.TextMatrix(i, .ColIndex("Order_No")) = "", Null, .TextMatrix(i, .ColIndex("Order_No")))
                '  End If
            
                RsNotes("CusID").value = Null
                RsNotes("NoteType").value = 3
                RsNotes("NoteDate").value = XPDtbTrans.value
                RsNotes("UserID").value = user_id
                ' rsnotes("ExpensesID").value = .TextMatrix(I, .ColIndex("ExpensesID"))
                RsNotes("notes_all").value = Me.XPTxtID.Text
                RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
                RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
                RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
                RsNotes("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
                RsNotes("sanad_year").value = year(XPDtbTrans.value)
                RsNotes("sanad_month").value = Month(XPDtbTrans.value)
                
                RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
                
                RsNotes.update
          
                RsDev.AddNew
                RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
                RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
                RsDev("DEV_ID_Line_No1").value = setfoxy_Line
                RsDev("Account_Code").value = DCproject.BoundText
                RsDev("Value").value = IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
                RsDev("Credit_Or_Debit").value = 0
                RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.Text  ' .TextMatrix(I, .ColIndex("des"))
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
                       
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev("notes_all").value = Me.XPTxtID.Text
                '                      RsDev("project_id").value = project_id
                        
                RsDev.update
                    
                line_no = line_no + 1

                With Fg_Journal

                    For i = .FixedRows To .Rows - 1
                        ' line_no = 2
        
                        If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                            '////////////////////////////////////////notes
                
                            If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "·« Ì„ﬂ‰ « „«„ ⁄„·Ì… «·Õ›Ÿ ·⁄œ„ «œŒ«· ﬁÌ„… ›Ì «·”ÿ— —ﬁ„  " & i - 1, vbCritical: GoTo ErrTrap
                                Else
                                    MsgBox "Cant save enter value in line :  " & i - 1, vbCritical: GoTo ErrTrap
                                End If
               
                            End If
 
                            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                            If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("value")), 1, txt_general_des.Text, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , setfoxy_Line, val(Me.XPTxtID.Text)) = False Then
                                GoTo ErrTrap
                    
                            End If

                            line_no = line_no + 1
        
                        End If

                    Next i
    
                End With

               ' Dim sql As String
                sql = "Update notes    set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.Text) * 2, "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.Text) & " and notetype=3" & "and NoteSerial1=" & TxtSerial1
                Cn.Execute sql
                sql = "Update   notes_all  set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.Text) * 2, "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.Text) & " and notetype=3" & "and NoteSerial1=" & TxtSerial1
                Cn.Execute sql
 
            End If

            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            LblDevID.Caption = LngDevID
            lbl(12).Caption = SystemOptions.SysCurrentAccountIntervalID
        End If
'saveBillBuy
  Dim PayDes As String
 Dim RowNum As Double
      PayDes = ""
    For RowNum = 1 To Grid22.Rows - 1
            
                       If val(Grid22.TextMatrix(RowNum, Grid22.ColIndex("Value"))) <> 0 Then
                        
                                    'Check Repeat Serial
                                     
If PayDes <> "" Then
          PayDes = PayDes & CHR(13) & Grid22.TextMatrix(RowNum, Grid22.ColIndex("PaymentName")) & ":" & Grid22.TextMatrix(RowNum, Grid22.ColIndex("value"))
 Else
           PayDes = Grid22.TextMatrix(RowNum, Grid22.ColIndex("PaymentName")) & ":" & Grid22.TextMatrix(RowNum, Grid22.ColIndex("value"))
 End If
 End If
 Next RowNum
         SaveMultyPayment val(XPTxtID.Text)
 Cn.Execute "update Notes set PayDes ='" & PayDes & "'   where NoteID=" & val(XPTxtID.Text)
 '''' **************save Paydes***********************

ll:
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Saved... " & CHR(13)
                    Msg = Msg + "Do you want to enter another operation?"
        
                End If

             '   Fg_Journal.Enabled = False

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
                End If

              '  Fg_Journal.Enabled = False
        End Select

        '«· Ê“Ì⁄ ⁄·Ï „—ﬂ“ «· ﬂ·›… «·⁄«„
        If Me.DcCostCenter.BoundText <> "" Then
            save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.Text, "”‰œ ’—›", Me.XPDtbTrans.value
        End If
        
        'Õ›Ÿ «·„’«—Ì› › ÃœÊ· «·„’«—Ì›
     
        'If saveExpensesDetails(0, TxtSerial.text, TxtSerial1.text, txt_ORDER_NO.text, XPDtbTrans.value) = True Then
        'End If
    
        TxtModFlg.Text = "R"
    

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "cant save " & CHR(13)
            Msg = Msg + "Invalid entry value " & CHR(13)
            Msg = Msg + "Check data and try again"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorr.... Error during saving " & CHR(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Sub PGMultyPayment(Optional general_noteid As Long, Optional ByRef lineno As Integer, Optional ByRef Line1 As Double, Optional StrTempDes As String, Optional Posted As Integer)
Dim StrMSG As String
Dim Commisionvalue As Double
Dim StrTempAccountCode As String
Dim i As Integer
Dim ValuGird As Double
Dim maxvalue As Double
Dim commision As Double
   Dim LngDevID As Long
    'Dim LngDevNO  As Integer
  '  Dim StrTempDes As String
    Dim SngTemp As Variant
    Dim TotalValue As Double
    On Error GoTo ErrTrap
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
  With Grid22
      For i = 1 To .Rows - 1
      StrMSG = ""
      If val(.TextMatrix(i, .ColIndex("Value"))) <> 0 Then
      ValuGird = val(.TextMatrix(i, .ColIndex("Value"))) '* val(txt_Currency_rate.Text)
      StrMSG = " " & (.TextMatrix(i, .ColIndex("PaymentName")))
      If val(.TextMatrix(i, .ColIndex("PaymentID"))) = 0 Then
      StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
      OtherInformation.NextAccount_Code = DcboCreditSide.BoundText
      lineno = lineno + 1
      
            If ModAccounts.AddNewDev(LngDevID, lineno, StrTempAccountCode, ValuGird, 1, StrTempDes & StrMSG, general_noteid, , , , Me.XPDtbTrans.value, DCboUserName.BoundText, val(XPTxtID), , , , , , , , , , , , , , , , , val(Me.DcbBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                End If
            
      ElseIf val(.TextMatrix(i, .ColIndex("PaymentID"))) > 0 Then
      commision = val(.TextMatrix(i, .ColIndex("commision")))
      maxvalue = val(get_TblPaymentTypet(val(.TextMatrix(i, .ColIndex("PaymentID"))), "MaxValue"))
          StrTempAccountCode = .TextMatrix(i, .ColIndex("bankAccount_Code"))
      If SystemOptions.AllowCommtionJEFromValueVisa = True Then

            
      If commision > 0 And .TextMatrix(i, .ColIndex("Accountcom")) <> "" Then
                Commisionvalue = (ValuGird * commision) / 100
                If maxvalue <> 0 And maxvalue < Commisionvalue Then
                Commisionvalue = maxvalue
                End If
        lineno = lineno + 1
            If ModAccounts.AddNewDev(LngDevID, lineno, .TextMatrix(i, .ColIndex("Accountcom")), Commisionvalue, 1, StrTempDes & "   " & .TextMatrix(i, .ColIndex("PaymentName")) & "⁄„Ê·… ", general_noteid, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , , , , , , , , , val(Me.DcbBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                End If
        '    lineno = lineno + 1
      End If
      ValuGird = ValuGird - Commisionvalue
       lineno = lineno + 1
                If ModAccounts.AddNewDev(LngDevID, lineno, StrTempAccountCode, ValuGird, 1, StrTempDes & StrMSG, general_noteid, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, val(Me.XPTxtID.Text), , , , , , , , , , , , , , , , , val(Me.DcbBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                End If
         
      Else
      lineno = lineno + 1
                   If ModAccounts.AddNewDev(LngDevID, lineno, StrTempAccountCode, ValuGird, 1, StrTempDes & StrMSG, general_noteid, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, val(Me.XPTxtID.Text), , , , , , , , , , , , , , , , , val(Me.DcbBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                End If
             
            
      If commision > 0 And .TextMatrix(i, .ColIndex("Accountcom")) <> "" Then
                Commisionvalue = (ValuGird * commision) / 100
                If maxvalue <> 0 And maxvalue < Commisionvalue Then
                Commisionvalue = maxvalue
                End If
        lineno = lineno + 1
            If ModAccounts.AddNewDev(LngDevID, lineno, .TextMatrix(i, .ColIndex("Accountcom")), Commisionvalue, 1, StrTempDes & "   " & .TextMatrix(i, .ColIndex("PaymentName")) & "⁄„Ê·… ", general_noteid, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, val(Me.XPTxtID.Text), , , , , , , , , , , , , , , , , val(Me.DcbBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                End If
            lineno = lineno + 1
            OtherInformation.NextAccount_Code = DcboDebitSide.BoundColumn
                 If ModAccounts.AddNewDev(LngDevID, lineno, StrTempAccountCode, Commisionvalue, 1, StrTempDes & "   " & .TextMatrix(i, .ColIndex("PaymentName")) & "⁄„Ê·… ", general_noteid, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, val(Me.XPTxtID.Text), , , , , , , , , , , , , , , , , val(Me.DcbBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                End If
             
      End If
      End If
   
      End If
          
          End If
     Next i
      End With
ErrTrap:
End Sub
Public Function save_General_cost_center(cost_center_id As String, _
                                         cost_center, _
                                         opr_type As String, _
                                         record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
 
    StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Fg_Journal
 
        .Rows = .Rows + 1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
        
                rs.AddNew
                rs("cost_center_id").value = cost_center_id
                rs("cost_center").value = cost_center
                rs("value").value = .TextMatrix(i, .ColIndex("value"))
                rs("depit_or_credit").value = "„œÌ‰"
                rs("opr_id").value = Me.Text1.Text
                rs("kedno").value = Me.Text1.Text
                rs("opr_type").value = opr_type
                rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
                rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
                rs("line_no").value = .TextMatrix(i, .ColIndex("LineNo1"))
                rs("record_date").value = record_date
                rs.update
        
            End If

        Next i

    End With

    rs.Close
End Function

Function calcnets()

    With Fg_Journal
        Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    End With

    If SystemOptions.gldetails_or_gl_general = 0 And Me.DCproject.BoundText <> "" Then

        With Me.VSFlexGrid1
            Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        End With

    End If

End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "NoteID='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If XPTxtID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From notes Where NoteID=" & val(TXT_A_NoteID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            Cn.Execute " Delete from TblExpensesDet where  ExpID =" & val(XPTxtID.Text)
            StrSQL = "Delete From ExpensesDetails Where NoteSerial1='" & val(TxtSerial1.Text) & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
            Cn.Execute " Delete from TblExpUnitNo where  ExpID =" & val(XPTxtID.Text)
            Cn.Execute " Delete from TblExpensesDet where  ExpID =" & val(XPTxtID.Text)
            
              StrSQL = "Delete From TblMultuPayment Where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
   'DeleteBillBuy
              StrSQL = "Delete From TblNotesBillBuyPayment2 Where NoteID1=" & val(Me.XPTxtID.Text) & " and TransType is null"
              Cn.Execute StrSQL, , adExecuteNoRecords
              StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtID.Text) & " and TransType is null"
              Cn.Execute StrSQL, , adExecuteNoRecords
              
            If Not rs.RecordCount < 1 Then
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    Fg_Journal.Clear flexClearScrollable, flexClearEverything
                    Fg_Journal.Rows = 3
                 '   Fg_Journal.Enabled = False
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Function FillGridWithData()

End Function

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

    IntCounter = 0

    With Me.VSFlexGrid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

End Sub

Private Sub PutData()

    'MsgBox Fg_Journal.Row & "---" & Fg_Journal.ColKey(Fg_Journal.Col)
    With Fg_Journal

        If Len(TxtDes.Text) > 0 Then
            .Cell(flexcpData, .Row, .ColIndex("Des")) = TxtDes.Text
            .Cell(flexcpPicture, .Row, .ColIndex("Des")) = ImgNote.Picture
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        Else
            .Cell(flexcpData, .Row, .ColIndex("Des")) = ""
            .Cell(flexcpPicture, .Row, .ColIndex("Des")) = Empty
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        End If

    End With

End Sub

Function sand_numbering() As String
    On Error Resume Next
    Dim start_at As Integer
    Dim end_at As Integer
    Dim auto_sanad_no As String
    Dim NO As String
    auto_sanad_no = ""
    departement_name = 1
    branch_no = 1
    connection_string = Cn.ConnectionString
    numbering.ConnectionString = connection_string
    numbering.CommandType = adCmdText
    numbering.RecordSource = "select * from sanad_numbering where branch_no=" & my_branch & " and departement='" & departement_name & "' and  sanad_no=1"
    numbering.Refresh

    If numbering.Recordset.RecordCount = 0 Then
        numbering_type = 0
    Else
        numbering_type = numbering.Recordset.Fields!numbering_id
        start_at = numbering.Recordset.Fields!start_at
        end_at = numbering.Recordset.Fields!end_at

    End If

    If numbering_type = 1 Then
        detect_no.ConnectionString = connection_string
        detect_no.CommandType = adCmdText
        detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=3 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
        detect_no.Refresh

        If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
 
            If end_at = 0 Then end_at = detect_no.Recordset.Fields!last_sand_no + 1
 
            If detect_no.Recordset.Fields!last_sand_no >= end_at Then
                sand_numbering = "error"
                Exit Function
            End If
        End If

    Else

        If numbering_type = 2 Then
 
            detect_no.ConnectionString = connection_string
            detect_no.CommandType = adCmdText
            detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=3 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & mId(Format$(Now, "dd/mm/yyyy"), 4, 2)
            'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
            detect_no.Refresh

            If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)

                If end_at = 0 Then end_at = NO + 1
                If NO >= end_at Then
                    sand_numbering = "error"
                    Exit Function
                End If
            End If

        Else

            If numbering_type = 3 Then
 
                detect_no.ConnectionString = connection_string
                detect_no.CommandType = adCmdText
                detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=3 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(Now, "dd/mm/yyyy"), 7, 4)
                'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "'  and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
                detect_no.Refresh

                If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)

                    If end_at = 0 Then end_at = NO + 1
                    If NO >= end_at Then
                        sand_numbering = "error"
                        Exit Function
                    End If
                End If
 
            End If
 
        End If
    End If

    If detect_no.Recordset.RecordCount = 0 Or IsNull(detect_no.Recordset.Fields!last_sand_no) Then

        If numbering_type = 0 Then
            ' auto_sanad_no = 1
        Else

            If numbering_type = 1 Then
                auto_sanad_no = start_at
            Else
                
                If numbering_type = 2 Then
                    auto_sanad_no = mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & mId(Format$(Now, "dd/mm/yyyy"), 4, 2) & start_at

                Else

                    If numbering_type = 3 Then
                        auto_sanad_no = mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & start_at

                    End If
                End If
            End If
        End If

    Else

        If numbering_type = 0 Then
            'auto_sanad_no = x + 1
        Else

            If numbering_type = 1 Then
                auto_sanad_no = detect_no.Recordset.Fields!last_sand_no + 1
            Else
                
                If numbering_type = 2 Then
                    '  If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) Then
                    ' no = 1
                    '  auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) & "1"
                    '  Else
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
                    auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (NO + 1)
                    '  End If
                      
                Else

                    If numbering_type = 3 Then
                        '    If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) Then
                        'no = 1
                        '    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "1"
                        '    Else
                        NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
                        auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (NO + 1)

                        '    End If

                    End If
                End If
            End If
        End If

    End If

    sand_numbering = auto_sanad_no

    'MsgBox auto_sanad_no

End Function

Function setfoxy_Line() As Double
    
    Dim X As Double
    X = CStr(new_id("foxy", "id1", "", True))
    setfoxy_Line = X
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id1").value = X ' last_line_id
 
    rs.update
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    'Exit Sub
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
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub AddTip()
    Dim Wrap As String
    Dim BolRtl As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    If BolRtl = True Then

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    Else

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "Add New Record..." & Wrap & "Shortcut Key F12 OR Enter" & Wrap & "OR Alt+N", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit the Current Record..." & Wrap & "Shortcut Key F11 " & Wrap & "OR Alt+E", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save the New Record OR Save the Editing in the Current Record..." & Wrap & "Shortcut Key F10 " & Wrap & "OR Alt+S", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Cancel the New Record OR Cancel Editing in the Current Record..." & Wrap & "Shortcut Key F9 " & Wrap & "OR Alt+U", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete the Current Record..." & Wrap & "Shortcut Key F8 " & Wrap & "OR Alt+D", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Close this Screen" & Wrap & "OR Alt+X", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help..." & Wrap & "Display Help for this Screen" & Wrap & "Shortcut Key F1" & Wrap, BolRtl
        End With

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
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & CHR(13)
                    StrMSG = StrMSG & " the Modifications  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
                
                End If

        End Select

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)

        Select Case IntResult

            Case vbYes
                Cancel = True
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPCboExpensesType_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("ExpensesType", "ID", val(Me.XPCboExpensesType.BoundText))
    End If

End Sub

Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtSerial1.Text = ""
TxtSerial.Text = ""
End If
End Sub

Private Sub XPTxtVal_Change()

    'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0)
    'Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".")
    If SystemOptions.UserInterface = ArabicInterface Then
        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.Text, "0.00"), 0, True, ".", , 0)
    Else

        Me.LblValue.Caption = WriteNo(XPTxtVal.Text, 0, , , , 1)
    End If
    
End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)
End Sub

Private Sub XPTxtVal_Validate(Cancel As Boolean)
    'If Val(XPTxtVal.Text) = 0 Then
    '    Set TTD = New clstooltipdemand
    '    TTD.Style = TTBalloon
    '    TTD.Icon = TTIconWarning
    '    TTD.Centered = True
    '    TTD.RightToLeft = True
    '    TTD.VisibleTime = 600
    '    TTD.BackColor = 0
    '    TTD.Title = "ﬁÌ„… «·„’—Ê›« "
    '    TTD.TipText = "»—Ã«¡ ﬂ «»… ﬁÌ„… «·„’—Ê›« "
    '    TTD.PopupOnDemand = True
    '    TTD.CreateToolTip XPTxtVal.hwnd
    '    TTD.Show 0, XPTxtVal.Height / Screen.TwipsPerPixelX - 1    '//In Pixel only
    '    Cancel = True
    'Else
    '    TTD.Destroy
    'End If
End Sub
Private Sub DcbIqara_Change()
DcbUnitType_Change
End Sub
Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Double
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        DcbIqara.BoundText = EmpID
        DcbIqara_Click (0)
    End If
End Sub
Private Sub DcbUnitType_Click(Area As Integer)
DcbUnitType_Change
End Sub
 

Private Sub DcbUnitType_Change()
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
   Set Dcombos = New ClsDataCombos

If val(DcbIqara.BoundText) > 0 Then
idd = val(DcbIqara.BoundText)

idd1 = val(DcbUnitType.BoundText)
If Me.TxtModFlg = "R" Then
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
Else
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
End If
End If
End Sub
Private Sub ViewDataList()
  Dim FrmView As FrmViewList
    Dim FG As VSFlex8UCtl.VSFlexGrid
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    'Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set FG = FrmView.vsfGroup1.VSFlexGrid

    With FG
        .Cols = 18
        .RowHeightMin = 320
        .ExplorerBar = flexExSortShowAndMove
        .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
        .ColKey(0) = "NoteID"
        .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
        .ColKey(1) = "NoteSerial"
        .TextMatrix(0, 2) = "«· «—ÌŒ"
        .ColKey(2) = "NoteDate"
        .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
        .ColKey(3) = "Name"
        .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
        .ColKey(4) = "Note_Value"
        .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
        .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
        .ColKey(5) = "BoxName"
        .TextMatrix(0, 6) = "„·«ÕŸ« "
        .ColKey(6) = "Remark"
        .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
        .ColKey(7) = "UserName"
    
        StrSQL = "SELECT NoteID, NoteSerial, NoteDate, Name, Note_Value, BoxName," & "Remark, UserName From ExpensesReport"
        StrSQL = StrSQL + " Order By NoteID"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
        'Â‰« Ìﬂ » ﬂÊœ ·⁄„· „⁄œ·  Õ„Ì· «·»Ì«‰« 
        '------------------------------------
        '
        '
        '
        '
    
        '------------------------------------
        Set .DataSource = rs
        .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
        .ColKey(0) = "NoteID"
        .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
        .ColKey(1) = "NoteSerial"
        .TextMatrix(0, 2) = "«· «—ÌŒ"
        .ColKey(2) = "NoteDate"
        .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
        .ColKey(3) = "Name"
        .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
        .ColKey(4) = "Note_Value"
        .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
        .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
        .ColKey(5) = "BoxName"
        .TextMatrix(0, 6) = "„·«ÕŸ« "
        .ColKey(6) = "Remark"
        .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
        .ColKey(7) = "UserName"
    
        'Rs.Close
        'Set Rs = Nothing
        .AutoSize 0, .Cols - 1, False
    End With

    Set GrdBack = New ClsBackGroundPic
    FrmView.vsfGroup1.VSFlexGrid.WallPaper = GrdBack.Picture
    FrmView.vsfGroup1.SetRTL = True
    FrmView.vsfGroup1.TotalOnColKey = "Note_Value"
    FrmView.vsfGroup1.sql = StrSQL
    FrmView.vsfGroup1.ShowTreeGroups = True
    FrmView.vsfGroup1.update
    FrmView.SetDblClickRetrun Me, "NoteID"
    FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·»Ì«‰«  «·„’—Ê›« "
    FrmView.show
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    LblValue.Visible = False
    lbl(14).Caption = "Project#"
    'Label1.Caption = "Voucher #"
    Me.ALLButton1.Caption = "Cost Center"
    lbl(15).Caption = "Payment Method"
    lbl(16).Caption = "Box Name"
    lbl(20).Caption = "General Des"
    lbl(21).Caption = "Order No:"
    lbl(91).Caption = "Account"
   ' Label8.Caption = "General C. C."

    With Me.CboPayMentType
        .Clear
        .AddItem "Cash"
        .AddItem "Cheque"
        .AddItem "Account"
    End With

    CmdRemove.Caption = "Delete Row"
    Me.Caption = "Expenses"
    Me.Ele.Caption = "Expenses"
    Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
    Me.lbl(4).Caption = "Operation ID"
    Me.lbl(1).Caption = "Operation Date"
    Me.lbl(3).Caption = "Expenses Type"
    Me.lbl(2).Caption = "Expenses Value"
    Me.lbl(0).Caption = "Based On"
    Me.lbl(5).Caption = "TO"
    Me.lbl(8).Caption = "Issued By."
    Me.lbl(7).Caption = "Current Record."
    Fra.Caption = "GL"
    lbl(11).Caption = "GL#"
    lbl(13).Caption = "interval"
    lbl(9).Caption = "Depit"
    lbl(10).Caption = "Credit"
    lbl(17).Caption = "Bank"
    lbl(18).Caption = "Cheque#"
    lbl(19).Caption = "Due Date"

    Me.Cmd(0).Caption = "&New"
    Me.Cmd(1).Caption = "&Edit"
    Me.Cmd(2).Caption = "&Save"
    Me.Cmd(3).Caption = "&Undo"
    Me.Cmd(4).Caption = "&Delete"
    Me.Cmd(5).Caption = "Sear&ch"
    Me.Cmd(6).Caption = "E&xit"
    Me.Cmd(7).Caption = "&Table View"
    Cmd(8).Caption = "Print"
    Cmd(9).Caption = "Cheque Print"
    Cmd(10).Caption = "GL Print "
    Label1(4).Caption = "Real Estate"
    Label1(15).Caption = "Type"
    Label1(1).Caption = "Owner"
    Label1(2).Caption = "Unit No"
    Me.CmdHelp.Caption = "&Help"
    lbl(63).Caption = "From"
    lbl(64).Caption = "To"
    Frame8.Caption = "Period"
    With Me.Fg_Journal
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("AccountName")) = " Account Name "
        .TextMatrix(0, .ColIndex("Account_Serial")) = " Account  Code "
        .TextMatrix(0, .ColIndex("value")) = "value"
        .TextMatrix(0, .ColIndex("des")) = "description"
        .TextMatrix(0, .ColIndex("opr_fullcode")) = "Operation"
        .TextMatrix(0, .ColIndex("order_no")) = "Order no"
        .TextMatrix(0, .ColIndex("iqar")) = " Real Estate "
        .TextMatrix(0, .ColIndex("unittype")) = " Type  "
        .TextMatrix(0, .ColIndex("unitno")) = "Unit No "

    End With

End Sub
Function CheckMult_Cash() As Boolean
Dim i As Integer
  With Grid22
      For i = 1 To .Rows - 1
      If val(.TextMatrix(i, .ColIndex("Value"))) <> 0 And val(.TextMatrix(i, .ColIndex("PaymentID"))) = 0 Then
      CheckMult_Cash = True
      Exit Function
      End If
      Next i
      CheckMult_Cash = False
   End With
End Function
