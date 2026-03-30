VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form imaged 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄—÷  «·„—›ﬁ« "
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18810
   Icon            =   "imaged.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10650
   ScaleWidth      =   18810
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      Caption         =   "Õœœ «”„ «·„—›ﬁ"
      Height          =   2055
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   1920
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton Command3 
         Caption         =   "Õ›Ÿ"
         Height          =   375
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtAttach 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Text            =   "„—›ﬁ1"
         Top             =   360
         Width           =   5775
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   15720
      Top             =   9360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Height          =   2775
      Left            =   15480
      TabIndex        =   45
      Top             =   3480
      Width           =   3375
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "imaged.frx":000C
         Height          =   1695
         Left            =   120
         TabIndex        =   46
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2990
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "operation_no"
            Caption         =   "operation_no"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "subject_no"
            Caption         =   "subject_no"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "image_NAME"
            Caption         =   "image_NAME"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "image_date"
            Caption         =   "image_date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "DEPARTEMENT"
            Caption         =   "DEPARTEMENT"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "image_no"
            Caption         =   "image_no"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "emp"
            Caption         =   "emp"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "contract_no"
            Caption         =   "contract_no"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "image_Title"
            Caption         =   "«”„ «·„·›"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   -1  'True
               ColumnWidth     =   2505.26
            EndProperty
         EndProperty
      End
      Begin ALLButtonS.ALLButton CmdNewFile 
         Height          =   375
         Left            =   2400
         TabIndex        =   47
         Top             =   2400
         Width           =   975
         _ExtentX        =   1720
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "imaged.frx":0021
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
         Left            =   1320
         TabIndex        =   48
         Top             =   2400
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "⁄—÷"
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "imaged.frx":003D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton CmdDeleteDoc 
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   2400
         Width           =   975
         _ExtentX        =   1720
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   255
         BCOLO           =   8421631
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "imaged.frx":0059
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„·›«  «·„—ﬁﬁ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   50
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Height          =   855
      Left            =   15720
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   43
      Top             =   9600
      Width           =   2895
   End
   Begin VB.TextBox txtopeation_type 
      Height          =   615
      Left            =   22200
      TabIndex        =   42
      Text            =   "Text8"
      Top             =   7920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      Height          =   3015
      Left            =   15360
      TabIndex        =   27
      Top             =   6240
      Width           =   3615
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "imaged.frx":0075
         Height          =   2175
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483648
         HeadLines       =   2
         RowHeight       =   24
         FormatLocked    =   -1  'True
         RightToLeft     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "id"
            Caption         =   "id"
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
            DataField       =   "subject_no"
            Caption         =   "—ﬁ„ «·„Ê÷Ê⁄"
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
            DataField       =   "template_id"
            Caption         =   "—ﬁ„ «·‰„Ê–Ã"
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
            DataField       =   "template_name"
            Caption         =   "«”„ «·‰„Ê–Ã"
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
            DataField       =   "date_added"
            Caption         =   " «—ÌŒ «·«÷«›…"
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
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin ALLButtonS.ALLButton Command2 
         Height          =   375
         Left            =   2400
         TabIndex        =   30
         Top             =   2400
         Width           =   975
         _ExtentX        =   1720
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
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "imaged.frx":008A
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
         Left            =   1320
         TabIndex        =   31
         Top             =   2400
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "⁄—÷"
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
         BCOL            =   65535
         BCOLO           =   65535
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "imaged.frx":00A6
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
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   2400
         Width           =   975
         _ExtentX        =   1720
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
         BCOL            =   255
         BCOLO           =   8421631
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "imaged.frx":00C2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label8 
         Caption         =   "«·‰„«–Ã   "
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
         Height          =   375
         Left            =   2160
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   15480
      TabIndex        =   20
      Top             =   0
      Width           =   3615
      Begin VB.Frame Frame4 
         Height          =   2295
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   1335
         Begin ALLButtonS.ALLButton Command1 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   " €ÌÌ— ’Ê—…"
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
            BCOL            =   255
            BCOLO           =   255
            FCOL            =   16777215
            FCOLO           =   16777215
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "imaged.frx":00DE
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
            Index           =   4
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   " ﬂ»Ì—"
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
            BCOL            =   8438015
            BCOLO           =   8438015
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "imaged.frx":00FA
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
            Index           =   5
            Left            =   120
            TabIndex        =   36
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   " ’€Ì—"
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
            BCOL            =   8438015
            BCOLO           =   8438015
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "imaged.frx":0116
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
            Index           =   6
            Left            =   120
            TabIndex        =   37
            Top             =   1680
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "œÊ—«‰"
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
            BCOL            =   8438015
            BCOLO           =   8438015
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "imaged.frx":0132
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2295
         Left            =   1800
         TabIndex        =   21
         Top             =   1080
         Width           =   1335
         Begin ALLButtonS.ALLButton Command1 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   240
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
            BCOL            =   16711680
            BCOLO           =   16711680
            FCOL            =   16777215
            FCOLO           =   16777215
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "imaged.frx":014E
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
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   720
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
            BCOL            =   16711680
            BCOLO           =   16711680
            FCOL            =   16777215
            FCOLO           =   16777215
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "imaged.frx":016A
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
            Left            =   120
            TabIndex        =   24
            Top             =   1200
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
            BCOL            =   255
            BCOLO           =   8421631
            FCOL            =   16777215
            FCOLO           =   16777215
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "imaged.frx":0186
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
            Index           =   7
            Left            =   120
            TabIndex        =   40
            Top             =   1680
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   8454016
            BCOLO           =   8454016
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "imaged.frx":01A2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   240
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
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
         Caption         =   "  Õ—Ìﬂ «·’Ê—"
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
      Begin VB.Label Label5 
         Caption         =   "«·„” ‰œ«    "
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
         Height          =   375
         Left            =   2040
         TabIndex        =   25
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.TextBox Text7 
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   23520
      TabIndex        =   19
      Text            =   "Text7"
      Top             =   6960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "⁄—÷  «·‰„Ê–Ã"
      Height          =   615
      Left            =   21240
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "«÷«›… ‰„Ê–Ã ··„” ‰œ"
      Height          =   615
      Left            =   20640
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   " ⁄œÌ· «·’Ê—…"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "⁄—÷ «·’Ê—"
      Height          =   375
      Left            =   21480
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      DataField       =   "image_no"
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   9000
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      DataField       =   "subject_no"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   600
      TabIndex        =   12
      Text            =   "Text6"
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "DEPARTEMENT"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "image_NAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   -120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      DataField       =   "subject_no"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   -240
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "image_date"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin DBPIXLib.DBPix20 DBPix201 
      Height          =   8895
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   14895
      _Version        =   131072
      _ExtentX        =   26273
      _ExtentY        =   15690
      _StockProps     =   1
      _Image          =   "imaged.frx":01BE
      ImageResampleWidth=   100
      ImageResampleHeight=   100
      ImageResampleMode=   1
      ImageSaveFormat =   0
      JPEGQuality     =   75
      JPEGEncoding    =   0
      JPEGColorMode   =   0
      JPEGNoRecompress=   -1  'True
      JPEGRotateWarning=   0
      PNGColorDepth   =   0
      PNGCompression  =   0
      PNGFilter       =   0
      PNGInterlace    =   1
      ImageDitherMethod=   3
      ImagePaletteMethod=   4
      ImagePreviewMode=   0   'False
      ImageKeepMetaData=   -1  'True
      UseAmbientBackcolor=   -1  'True
      ViewAsyncDecoding=   -1  'True
      ViewEnableMouseZoom=   -1  'True
      ViewInitialZoom =   0
      ViewHAlign      =   1
      ViewVAlign      =   1
      ViewMenuMode    =   0
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   21360
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   21360
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      Caption         =   "  Õ—Ìﬂ «·’Ê—"
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
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "„·«ÕŸ« "
      Height          =   255
      Left            =   17400
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   9240
      Width           =   855
   End
   Begin VB.Label lbl_emp 
      Caption         =   "0"
      Height          =   255
      Left            =   8760
      TabIndex        =   41
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "’Ê—… —ﬁ„"
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
      Height          =   615
      Left            =   10440
      TabIndex        =   39
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "‘«‘… ⁄—÷ „—›ﬁ«  «·„ÊŸ› —ﬁ„"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   7560
      TabIndex        =   38
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label7 
      Caption         =   "«·’Ê—"
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
      Height          =   375
      Left            =   12360
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label screen_name 
      Alignment       =   1  'Right Justify
      Caption         =   "Label4"
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„—›ﬁ"
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label DEPARTEMENT 
      Caption         =   "DEPARTEMENT"
      Height          =   255
      Left            =   9840
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬁ”„"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ «·„—›ﬁ"
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label SUBJECT_NO 
      Alignment       =   2  'Center
      Caption         =   "SUBJECT_NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·„—›ﬁ"
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·„” ‰œ"
      Height          =   375
      Left            =   11160
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "imaged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NEW_IMAGE As Boolean
Dim loadform As Boolean

Private Sub ALLButton1_Click()

    If Adodc3.Recordset.RecordCount <> 0 Then
        loading_temolates.show
        loading_temolates.Command11.Enabled = False
        loading_temolates.Command16.Enabled = False

        loading_temolates.Label6.Caption = Adodc3.Recordset.Fields!template_id
        'loading_temolates.Label7.Caption = Adodc1.Recordset.Fields!IMAGE_NAME
        loading_temolates.Label9.Caption = Adodc3.Recordset.Fields!no_of_images
        'Me.Hide
    End If

End Sub

Private Sub ALLButton2_Click()
    X = MsgBox("Â· «‰  „ √ﬂœ „‰ «·Õ–›", vbCritical + vbYesNo)

    If X = vbNo Then
        Exit Sub
    End If

    If Adodc3.Recordset.RecordCount > 0 Then
        Adodc3.Recordset.delete
        Adodc3.Refresh
    End If

End Sub

Private Sub ALLButton4_Click()
If (Adodc4.Recordset.RecordCount > 0) Then

ShellExecute 0&, vbNullString, App.path & "\Doc\" & Adodc4.Recordset.Fields!IMAGE_NAME, vbNullString, vbNullString, vbNormalFocus
 


'Shell Adodc3.Recordset.Fields!help_list_path, vbNormalFocus
Else
MsgBox "·« ÌÊÃœ —«Ìÿ", vbCritical
End If
End Sub

Private Sub CmdDeleteDoc_Click()
If Adodc4.Recordset.RecordCount > 0 Then
Adodc4.Recordset.delete
Adodc4.Refresh
End If

End Sub

Private Sub CmdNewFile_Click()
Frame6.Visible = True
End Sub

Private Sub Command1_Click(Index As Integer)

    If Index = 0 Then
        NEW_IMAGE = True
        DBPix201.ImageClear
        DBPix201.Visible = True
        Dim LASTIMAGENO As Integer
        Dim X As Integer

        Adodc2.CommandType = adCmdText
 
        Adodc2.RecordSource = "SELECT MAX(image_no)  AS LASTIMAGENO FROM subjects_images WHERE operation_type ='" & txtopeation_type.Text & " ' and subject_no= '" & SUBJECT_NO.Caption & "'"
        Adodc2.Refresh

        If Adodc2.Recordset.RecordCount = 0 Or IsNull(Adodc2.Recordset.Fields!LASTIMAGENO) Then
            LASTIMAGENO = 1
        Else
            LASTIMAGENO = (Adodc2.Recordset.Fields!LASTIMAGENO) + 1
        End If

        Adodc1.Recordset.AddNew
 
         Text1.Text = SUBJECT_NO.Caption
        Text2.Text = txtopeation_type & "#" & SUBJECT_NO & "-" & LASTIMAGENO & "#" & day(Date) & "-" & Month(Date) & "-" & year(Date)
        Text3.Text = Now
        Text4.Text = LASTIMAGENO
        Text5.Text = Departement.Caption

        X = MsgBox("Â·  —Ìœ ’Ê—… „‰ „·›", vbExclamation + vbYesNoCancel)

        If X = vbYes Then
            DBPix201.ImageLoad

            DoEvents
            MsgBox " „  Õ„Ì· «·’Ê—…"
        Else

            If X = vbNo Then
                DBPix201.TWAINAcquire
                MsgBox " „ „”Õ ÷Ê∆Ì  ··’Ê—…"

                DoEvents
            Else

                Exit Sub
            End If
        End If

        DBPix201.ImageSaveFile (system_path & "\" & SystemOptions.ImagesPath & "\" & Text2.Text & ".JPG")
        NEW_IMAGE = False

        Adodc1.Recordset.Fields!operation_type = txtopeation_type.Text
        Adodc1.Recordset.update

        Adodc1.Recordset.MoveLast
        Adodc1.Refresh
        '      log_files_form.Adodc1.Recordset.AddNew: log_files_form.Adodc1.Recordset.Fields!log_date = Date
        '      log_files_form.Adodc1.Recordset.Fields!log_time = Time
        '      log_files_form.Adodc1.Recordset.Fields!User_Name = login.Adodc1.Recordset.Fields!name
        '
        '      log_files_form.Adodc1.Recordset.Fields!process_name = "    ‘«‘…   " & Me.Caption
        '       log_files_form.Adodc1.Recordset.Fields!process_text = "  „  «÷«›… ’Ê—… —ﬁ„ " & LASTIMAGENO & "  ··„” ‰œ —ﬁ„  " & SUBJECT_NO
        '
        '        log_files_form.Adodc1.Recordset.update: DoEvents
        Exit Sub
    End If

    If Index = 1 Then
        'DBPix201.ImageSaveFile (system_path & "\images\" & Text2.text & ".JPG")
        NEW_IMAGE = False
        Exit Sub
    End If

    If Index = 2 Then

        X = MsgBox("Â· «‰  „ √ﬂœ „‰ «·Õ–›", vbCritical + vbYesNo)

        If X = vbNo Then
            Exit Sub
        End If

        If Adodc1.Recordset.RecordCount > 0 Then
            Adodc1.Recordset.delete
            Adodc1.Refresh
            'DBPix201.Visible = False
        End If

    End If

    If Index = 3 Then
        'Dim x As Integer
        X = MsgBox("Â·  —Ìœ ’Ê—… „‰ „·›", vbExclamation + vbYesNoCancel)

        If X = vbYes Then
            DBPix201.ImageLoad

        Else

            If X = vbNo Then
                DBPix201.TWAINAcquire
            Else

                Exit Sub
            End If
        End If

        DBPix201.ImageSaveFile (system_path & "\" & SystemOptions.ImagesPath & "\" & Text2.Text & ".JPG")
    End If

    If Index = 4 Then
        DBPix201.ViewZoomIn
    End If

    If Index = 5 Then
        DBPix201.ViewZoomOut
    End If

    If Index = 6 Then
        DBPix201.ImageRotate ImageRotate90
    End If

    If Index = 7 Then
        On Error Resume Next
        loading_temolates.show
        loading_temolates.Frame2.Visible = False
        loading_temolates.Frame3.Visible = False
        loading_temolates.Frame4.Visible = False
        loading_temolates.Frame5.Visible = False
        loading_temolates.Frame6.Top = 4800
        loading_temolates.Image1.Picture = LoadPicture(system_path & "\" & SystemOptions.ImagesPath & "\" & Text2.Text & ".JPG")
        loading_temolates.Image1.Enabled = False
    End If

End Sub

Private Sub Command2_Click()
    frm_templates.show
    frm_templates.case_id = 0
 
End Sub

Private Sub Command3_Click()
     If txtAttach.Text = "" Then
                      If SystemOptions.UserInterface = ArabicInterface Then
                             MsgBox "Õœœ «”„ «·„—›ﬁ", vbCritical
                             Exit Sub
                                     
                Else
                         
                        MsgBox "Enter Attachment No ", vbCritical
                             Exit Sub
                End If
       End If
       
    Dim sourcefilename As String
        Dim desFilename As String
On Error Resume Next
        Dim pos As Integer
        Dim filename As String
        
    dlgFile.filter = "All files(*.*) | *.*"
    dlgFile.ShowOpen
    If dlgFile.filename <> vbNullString And dlgFile.filename <> "" Then
      
        filename = dlgFile.filename
        sourcefilename = (filename)
        
         
        
        
        pos = InStrRev(filename, "\")
        If pos > 0 Then
            filename = mId(filename, pos + 1)
        End If
        
         filename = TimeStamp(Date) & filename
   
         
         desFilename = App.path & "\Doc\" & filename
         FileCopy sourcefilename, desFilename
         
 
       

       
            Adodc4.Recordset.AddNew
            imaged.Adodc4.Recordset.Fields!SUBJECT_NO = SUBJECT_NO
            imaged.Adodc4.Recordset.Fields!IMAGE_NAME = filename
           imaged.Adodc4.Recordset.Fields!image_date = DateValue(Now)
           
            imaged.Adodc4.Recordset.Fields!Departement = Departement.Caption
             imaged.Adodc4.Recordset.Fields!image_Title = txtAttach.Text
             imaged.Adodc4.Recordset.Fields!operation_type = txtopeation_type.Text
             
             
            
            
            imaged.Adodc4.Recordset.update
            imaged.Adodc4.Refresh


Adodc4.Recordset.update
DataGrid2.Refresh
    
    End If
Frame6.Visible = False

End Sub

Private Sub Command4_Click()
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE subject_no=" & SUBJECT_NO.Caption
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then

        DBPix201.Visible = True
    Else
        DBPix201.Visible = False
    End If

End Sub

Private Sub Command6_Click()

    'loading_temolates.Show
    'loading_temolates.SUBJECT_NO.Caption = Me.SUBJECT_NO.Caption
End Sub

Private Sub Form_Activate()
    If loadform = False Then
    Exit Sub
    Adodc4.ConnectionString = connection_string
Adodc4.CommandType = adCmdText
'Adodc4.RecordSource = "select * from  Subject_doc WHERE   operation_type =" & val(txtopeation_type.Text) & " and subject_no='" & SUBJECT_NO.Caption & "'"   '
Position = InStr(1, SUBJECT_NO, "-")

If Position > 0 Then
Dim str1 As String
Dim str2 As String
Dim ConnectionStr As String
str1 = mId(SUBJECT_NO, 1, Position - 1)
str2 = mId(SUBJECT_NO, Position + 1, Len(TxtSerial1))
ConnectionStr = "SELECT * FROM subjects_images WHERE operation_type = '" & txtopeation_type
ConnectionStr = ConnectionStr & "' and ( subject_no='" & (str1) & "'" & " or  subject_no='" & (SUBJECT_NO) & "') "
  imaged.Adodc4.RecordSource = ConnectionStr
  
Else
imaged.Adodc4.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (SUBJECT_NO) & "'"
End If

'

Adodc4.Refresh
    loadform = True
    End If
    
End Sub

Private Sub Form_Load()
 '   On Error Resume Next
    'LoadSettings
     loadform = False
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500

    system_path = App.path ' "D:\my works\accountant\28  01 2011\SourceCode\SourceCode"
    connection_string = Cn.ConnectionString
    Adodc4.ConnectionString = connection_string
Adodc4.CommandType = adCmdText
Adodc4.RecordSource = "select * from  Subject_doc WHERE subject_no='" & SUBJECT_NO.Caption & "'"

'position = InStr(1, SUBJECT_NO, "-")

'If position > 0 Then
Dim str1 As String
Dim str2 As String
Dim ConnectionStr As String
'Str1 = mId(SUBJECT_NO, 1, position - 1)
'str2 = mId(SUBJECT_NO, position + 1, Len(TxtSerial1))
'ConnectionStr = "SELECT * FROM Subject_doc WHERE operation_type = '" & txtopeation_type
'ConnectionStr = ConnectionStr & "' and ( subject_no='" & (Str1) & "'" & " or  subject_no='" & (SUBJECT_NO) & "') "
'  imaged.Adodc4.RecordSource = ConnectionStr
'
'Else
'imaged.Adodc4.RecordSource = "SELECT * FROM Subject_doc WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (SUBJECT_NO) & "'"
'End If


Adodc4.Refresh

    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE subject_no='0'"
    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from  subjects_images"
    Adodc2.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "select * from  subject_templates"
    Adodc3.Refresh

Adodc4.ConnectionString = connection_string
Adodc4.CommandType = adCmdText
Adodc4.RecordSource = "select * from  Subject_doc WHERE subject_no='0'"
Adodc4.Refresh

    NEW_IMAGE = False
 
    '   log_files_form.Adodc1.Recordset.AddNew: log_files_form.Adodc1.Recordset.Fields!log_date = Date
    '   log_files_form.Adodc1.Recordset.Fields!log_time = Time
    '   log_files_form.Adodc1.Recordset.Fields!User_Name = login.Adodc1.Recordset.Fields!name
    '
    '      log_files_form.Adodc1.Recordset.Fields!process_name = " œŒÊ· «·Ï  ‘«‘…  " & Me.Caption
    '       log_files_form.Adodc1.Recordset.Fields!process_text = ""
    '
    '        log_files_form.Adodc1.Recordset.update: DoEvents
    ' Command4_Click
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
       
End Sub

Private Sub ChangeLang()
    Command1(3).Caption = "Change Image"
    Command1(0).Caption = "New Image"
    Command1(1).Caption = "save Image"
    Command1(2).Caption = "Delete Image"
    Command1(7).Caption = "Print Image"
    Command1(4).Caption = "Zoom IN"
    Command1(5).Caption = "Zoom Out"
    Command1(6).Caption = "Rotate"
    Label10.Caption = "Image #"
    Label5.Caption = "Documents"
    Label8.Caption = "Forms"
    DataGrid1.Columns(2).Caption = "Form ID"
    DataGrid1.Columns(3).Caption = "Form Name"
    DataGrid1.RightToLeft = False
    Command2.Caption = "New Form"
    ALLButton1.Caption = "View Form"
    ALLButton2.Caption = "Delete Form"
    Adodc1.Caption = "Move"
Label11.Caption = "Remarks"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    '       log_files_form.Adodc1.Recordset.AddNew: log_files_form.Adodc1.Recordset.Fields!log_date = Date
    '       log_files_form.Adodc1.Recordset.Fields!log_time = Time
    '       log_files_form.Adodc1.Recordset.Fields!User_Name = login.Adodc1.Recordset.Fields!name
    '
    '      log_files_form.Adodc1.Recordset.Fields!process_name = "  Œ—ÊÃ „‰  ‘«‘…" & Me.Caption
    '       log_files_form.Adodc1.Recordset.Fields!process_text = ""
    '
    '        log_files_form.Adodc1.Recordset.update: DoEvents
End Sub

Private Sub Text2_Change()

    If Text2.Text = "" Or NEW_IMAGE = True Then Exit Sub

    DBPix201.ImageClear

    If Dir(system_path & "\" & SystemOptions.ImagesPath & "\" & Text2.Text & ".JPG") <> "" Then
        DBPix201.ImageLoadFile (system_path & "\" & SystemOptions.ImagesPath & "\" & Text2.Text & ".JPG")
    End If

    'If DBPix201.ImageLoadFile(system_path & "\images\" & Text2.text & ".JPG") = True Then

    'End If
End Sub

