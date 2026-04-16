VERSION 5.00
Begin VB.Form mofradat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·„ﬁ—œ« "
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   5400
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame4 
      Caption         =   "„›—œ«  «·—« »"
      Height          =   6255
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox txtsakn 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   31
         Top             =   390
         Width           =   1335
      End
      Begin VB.TextBox txtbus 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   30
         Top             =   750
         Width           =   1335
      End
      Begin VB.TextBox txtfood 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   29
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtanother 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   28
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtsaknm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtbusm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   26
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtfoodm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   25
         Top             =   1170
         Width           =   1335
      End
      Begin VB.TextBox txtanotherm 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   24
         Top             =   1650
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   480
         Width           =   615
      End
      Begin VB.Frame Frame5 
         Caption         =   "ÿ—Ìﬁ… «·Õ”«»"
         Height          =   3255
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   2280
         Visible         =   0   'False
         Width           =   4935
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "Frmmofrad.frx":0000
            Left            =   2160
            List            =   "Frmmofrad.frx":0007
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            Caption         =   "‘Â—Ì"
            Height          =   195
            Index           =   0
            Left            =   2400
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "”‰ÊÌ"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   1200
            Width           =   4695
         End
         Begin VB.Frame Frame6 
            Height          =   495
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Width           =   1935
            Begin VB.CommandButton Command1 
               Caption         =   "="
               Height          =   315
               Index           =   4
               Left            =   120
               TabIndex        =   13
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "+"
               Height          =   315
               Index           =   3
               Left            =   1560
               TabIndex        =   12
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "-"
               Height          =   315
               Index           =   2
               Left            =   1200
               TabIndex        =   11
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "*"
               Height          =   315
               Index           =   1
               Left            =   840
               TabIndex        =   10
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "/"
               Height          =   315
               Index           =   5
               Left            =   480
               TabIndex        =   9
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               Caption         =   "«Õ„«·Ì «·„ﬁ«„"
               ForeColor       =   &H000000FF&
               Height          =   15
               Left            =   0
               TabIndex        =   14
               Top             =   2520
               Width           =   1935
            End
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   1560
            Width           =   4695
         End
         Begin VB.CommandButton Command6 
            Caption         =   "„Ê«›ﬁ"
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   285
            Index           =   41
            Left            =   3720
            TabIndex        =   22
            Top             =   240
            Width           =   915
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·Õ”«»"
            Height          =   285
            Index           =   42
            Left            =   3480
            TabIndex        =   21
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰ ÌÃ…"
            Height          =   285
            Index           =   43
            Left            =   3600
            TabIndex        =   20
            Top             =   2040
            Width           =   1155
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   285
            Index           =   44
            Left            =   1680
            TabIndex        =   19
            Top             =   2040
            Width           =   1155
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Õ”«»"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Œ—ÊÃ"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»œ· «·”ﬂ‰"
         Height          =   285
         Index           =   34
         Left            =   3840
         TabIndex        =   37
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»œ· „Ê«’·« "
         Height          =   285
         Index           =   35
         Left            =   4200
         TabIndex        =   36
         Top             =   870
         Width           =   915
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»œ· ÿ⁄«„"
         Height          =   285
         Index           =   37
         Left            =   4200
         TabIndex        =   35
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»œ·«  «Œ—Ï"
         Height          =   285
         Index           =   38
         Left            =   4200
         TabIndex        =   34
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁÌ„… «·”‰ÊÌ…"
         Height          =   285
         Index           =   39
         Left            =   2520
         TabIndex        =   33
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁÌ„… «·‘Â—Ì…"
         Height          =   285
         Index           =   40
         Left            =   480
         TabIndex        =   32
         Top             =   120
         Width           =   1275
      End
   End
End
Attribute VB_Name = "mofradat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
