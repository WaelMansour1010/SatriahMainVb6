VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form imaged2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄—÷  «·„—›ﬁ« "
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20595
   Icon            =   "imaged2.frx":0000
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10650
   ScaleWidth      =   20595
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   10440
      Visible         =   0   'False
      Width           =   3375
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "imaged2.frx":000C
         Height          =   1695
         Left            =   120
         TabIndex        =   7
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
         TabIndex        =   8
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
         MICON           =   "imaged2.frx":0021
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
         TabIndex        =   9
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
         MICON           =   "imaged2.frx":003D
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
         TabIndex        =   10
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
         MICON           =   "imaged2.frx":0059
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
         Left            =   -1560
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Height          =   975
      Left            =   8520
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   9600
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.TextBox txtopeation_type 
      Height          =   615
      Left            =   22200
      TabIndex        =   4
      Text            =   "Text8"
      Top             =   7920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text7 
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   23520
      TabIndex        =   3
      Text            =   "Text7"
      Top             =   6960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "⁄—÷  «·‰„Ê–Ã"
      Height          =   615
      Left            =   21240
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "«÷«›… ‰„Ê–Ã ··„” ‰œ"
      Height          =   615
      Left            =   20640
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "⁄—÷ «·’Ê—"
      Height          =   375
      Left            =   21480
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
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
   Begin C1SizerLibCtl.C1Elastic ELe 
      Height          =   10650
      Index           =   15
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   20595
      _cx             =   36327
      _cy             =   18785
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
      BackColor       =   14871017
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
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
      Frame           =   7
      FrameStyle      =   4
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.CommandButton Command9 
         Caption         =   "»ÕÀ"
         Height          =   615
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   3330
         Width           =   1080
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   600
         Left            =   120
         TabIndex        =   60
         Top             =   4065
         Width           =   19830
         _Version        =   786432
         _ExtentX        =   34978
         _ExtentY        =   1058
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Option1"
            Height          =   375
            Index           =   11
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   -120
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Option1"
            Height          =   375
            Index           =   10
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   -120
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Option1"
            Height          =   375
            Index           =   9
            Left            =   -120
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Option1"
            Height          =   375
            Index           =   8
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Option1"
            Height          =   375
            Index           =   7
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   120
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Option1"
            Height          =   375
            Index           =   6
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Option1"
            Height          =   375
            Index           =   5
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Option1"
            Height          =   375
            Index           =   4
            Left            =   7920
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Option1"
            Height          =   375
            Index           =   3
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   120
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Option1"
            Height          =   375
            Index           =   2
            Left            =   11640
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Option1"
            Height          =   375
            Index           =   1
            Left            =   13200
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Option1"
            Height          =   375
            Index           =   0
            Left            =   15480
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   120
            Value           =   -1  'True
            Width           =   1935
         End
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
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox Text6 
         DataField       =   "subject_no"
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   915
         TabIndex        =   58
         Text            =   "Text6"
         Top             =   600
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         DataField       =   "image_date"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4875
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         DataField       =   "image_NAME"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4470
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         DataField       =   "DEPARTEMENT"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1455
         TabIndex        =   37
         Top             =   720
         Visible         =   0   'False
         Width           =   3165
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
         Left            =   9870
         TabIndex        =   36
         Top             =   720
         Width           =   1470
      End
      Begin VB.CommandButton Command5 
         Caption         =   " ⁄œÌ· «·’Ê—…"
         Height          =   255
         Left            =   270
         TabIndex        =   35
         Top             =   1785
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Left            =   14865
         TabIndex        =   23
         Top             =   1425
         Width           =   5685
         Begin VB.Frame Frame4 
            Height          =   2295
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   1335
            Begin ALLButtonS.ALLButton Command1 
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   30
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
               MICON           =   "imaged2.frx":0075
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
               TabIndex        =   31
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
               MICON           =   "imaged2.frx":0091
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
               TabIndex        =   32
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
               MICON           =   "imaged2.frx":00AD
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
               TabIndex        =   33
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
               MICON           =   "imaged2.frx":00C9
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
            Left            =   3720
            TabIndex        =   24
            Top             =   120
            Width           =   1335
            Begin ALLButtonS.ALLButton Command1 
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   25
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
               MICON           =   "imaged2.frx":00E5
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
               TabIndex        =   26
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
               MICON           =   "imaged2.frx":0101
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
               TabIndex        =   27
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
               MICON           =   "imaged2.frx":011D
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
               TabIndex        =   28
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
               MICON           =   "imaged2.frx":0139
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
            Left            =   1440
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
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
            TabIndex        =   34
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2655
         Left            =   9480
         TabIndex        =   17
         Top             =   1425
         Width           =   5280
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "imaged2.frx":0155
            Height          =   1935
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   3413
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
            TabIndex        =   19
            Top             =   2160
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
            MICON           =   "imaged2.frx":016A
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
            TabIndex        =   20
            Top             =   2160
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
            MICON           =   "imaged2.frx":0186
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
            TabIndex        =   21
            Top             =   2160
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
            MICON           =   "imaged2.frx":01A2
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
            TabIndex        =   22
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Õœœ «”„ «·„—›ﬁ"
         Height          =   2055
         Left            =   1185
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1065
         Visible         =   0   'False
         Width           =   6855
         Begin VB.CommandButton Command8 
            Caption         =   "Õ›Ÿ"
            Height          =   375
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Õ›Ÿ"
            Height          =   375
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   15
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
            TabIndex        =   14
            Text            =   "„—›ﬁ1"
            Top             =   360
            Width           =   5775
         End
      End
      Begin MSComDlg.CommonDialog dlgFile 
         Left            =   15720
         Top             =   9600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin DBPIXLib.DBPix20 DBPix201 
         Height          =   2145
         Left            =   270
         TabIndex        =   40
         Top             =   1065
         Visible         =   0   'False
         Width           =   8115
         _Version        =   131072
         _ExtentX        =   14314
         _ExtentY        =   3784
         _StockProps     =   1
         _Image          =   "imaged2.frx":01BE
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   375
         Left            =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   3300
         _ExtentX        =   5821
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
      Begin VSFlex8Ctl.VSFlexGrid FG 
         Height          =   5700
         Left            =   105
         TabIndex        =   41
         Top             =   4770
         Width           =   19905
         _cx             =   35110
         _cy             =   10054
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
         BackColorFixed  =   14871017
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   100
         Cols            =   20
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"imaged2.frx":01D6
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
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   8
         Left            =   2730
         TabIndex        =   42
         Top             =   3450
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "«÷«›… ”ÿ—"
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
         MICON           =   "imaged2.frx":04F9
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
         Index           =   9
         Left            =   1425
         TabIndex        =   43
         Top             =   3450
         Width           =   1200
         _ExtentX        =   2117
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
         MICON           =   "imaged2.frx":0515
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton3 
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   3450
         Width           =   1065
         _ExtentX        =   1879
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
         MICON           =   "imaged2.frx":0531
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   315
         Left            =   6360
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   3690
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   1612120067
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   315
         Left            =   6360
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   3330
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   1612120067
         CurrentDate     =   37140
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ï  «—ÌŒ"
         Height          =   285
         Index           =   0
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   3690
         Width           =   810
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰  «—ÌŒ"
         Height          =   285
         Index           =   63
         Left            =   8055
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   3330
         Width           =   1170
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·„” ‰œ"
         Height          =   375
         Left            =   12240
         TabIndex        =   57
         Top             =   360
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·„—›ﬁ"
         Height          =   375
         Left            =   8955
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   930
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
         Left            =   1980
         TabIndex        =   55
         Top             =   240
         Width           =   5925
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·„—›ﬁ"
         Height          =   375
         Left            =   5790
         TabIndex        =   54
         Top             =   360
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
         Height          =   375
         Left            =   5925
         TabIndex        =   53
         Top             =   720
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label DEPARTEMENT 
         Caption         =   "DEPARTEMENT"
         Height          =   255
         Left            =   10785
         TabIndex        =   52
         Top             =   1200
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„—›ﬁ"
         Height          =   375
         Left            =   8955
         TabIndex        =   51
         Top             =   600
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label screen_name 
         Alignment       =   1  'Right Justify
         Caption         =   "Label4"
         Height          =   255
         Left            =   3285
         TabIndex        =   50
         Top             =   1200
         Visible         =   0   'False
         Width           =   1860
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
         Left            =   13560
         TabIndex        =   49
         Top             =   240
         Visible         =   0   'False
         Width           =   1200
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
         Left            =   8295
         TabIndex        =   48
         Top             =   240
         Width           =   4350
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
         Left            =   11445
         TabIndex        =   47
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label lbl_emp 
         Caption         =   "0"
         Height          =   255
         Left            =   9600
         TabIndex        =   46
         Top             =   1440
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "„·«ÕŸ« "
         Height          =   255
         Left            =   19350
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   9555
         Visible         =   0   'False
         Width           =   930
      End
   End
End
Attribute VB_Name = "imaged2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NEW_IMAGE As Boolean
Dim loadform As Boolean
Public mIDD As String
Public LngRow2 As Long
Dim mTypeID As Long
Public IsFromDate As Integer
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

Private Sub ALLButton3_Click()
If FG.Row > 0 Then
    FG.RemoveItem FG.Row
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

Private Sub cmdEditImageType_Click()
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
        Text5.Text = DEPARTEMENT.Caption

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
    If Index = 8 Then
        FG.Rows = FG.Rows + 2
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

If Index = 9 Then
    'saveimage

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
Private Sub saveimage(Optional ByVal mRow As Long = 0)
    Dim s As String
    
    
    
   
    
    's = "Delete Subject_doc WHERE subject_no='" & mIDD & "' and  Subject_doc.Type1ID =" & mTypeID & " and operation_no = "
    'Cn.Execute s
    
'    s = "select    * from  Subject_doc WHERE subject_no='" & mIDD & "' and operation_no = " & val(FG.TextMatrix(mRow, FG.ColIndex("operation_no")))
s = "select    * from  Subject_doc WHERE operation_no = " & val(FG.TextMatrix(mRow, FG.ColIndex("operation_no")))
    
    Dim RsNew As New ADODB.Recordset
    RsNew.Open s, Cn, adOpenKeyset, adLockOptimistic
    If Not RsNew.EOF Then
        
    Else
        RsNew.AddNew
    End If
        RsNew!SUBJECT_NO = Trim(FG.TextMatrix(mRow, FG.ColIndex("SUBJECT_NO")))
        RsNew!IMAGE_NAME = Trim(FG.TextMatrix(mRow, FG.ColIndex("IMAGE_NAME")))
        RsNew!operation_type = Trim(FG.TextMatrix(mRow, FG.ColIndex("operation_type")))
        If Trim(FG.TextMatrix(mRow, FG.ColIndex("Datee1"))) = "" Then
            RsNew!Datee1 = Date
        Else
            RsNew!Datee1 = Trim(FG.TextMatrix(mRow, FG.ColIndex("Datee1")))
        End If
        If Trim(FG.TextMatrix(mRow, FG.ColIndex("Datee2"))) = "" Then
            RsNew!Datee2 = Date
        Else
            RsNew!Datee2 = Trim(FG.TextMatrix(mRow, FG.ColIndex("Datee2")))
        End If
        RsNew!Type1ID = val(FG.TextMatrix(mRow, FG.ColIndex("Type1ID")))
        RsNew!Type2ID = val(FG.TextMatrix(mRow, FG.ColIndex("Type2ID")))
        RsNew!NameFile = Trim(FG.TextMatrix(mRow, FG.ColIndex("NameFile")))
        
        RsNew.update
        FG.TextMatrix(mRow, FG.ColIndex("operation_no")) = val(RsNew!operation_no & "")
    
    'saveGrid s, FG, "NameFile", "ContNo", "subject_no", mIDD
    mTypeID = GetTypeId
    
    
End Sub
Private Sub Command2_Click()
    frm_templates.show
    frm_templates.case_id = 0
 
End Sub
Private Function GetTypeId() As Long
Dim i As Long
GetTypeId = 0
For i = 0 To Option1.count - 1
    If Option1(i).value = True Then
        GetTypeId = i + 1
        Exit Function
    End If
Next
If GetTypeId = 0 Then GetTypeId = 1: Option1(0).value = True
End Function
Private Sub Command3_Click()
     If TxtAttach.Text = "" Then
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
           
            imaged.Adodc4.Recordset.Fields!DEPARTEMENT = DEPARTEMENT.Caption
             imaged.Adodc4.Recordset.Fields!image_Title = TxtAttach.Text
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

Private Sub Command8_Click()
     If TxtAttach.Text = "" Then
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
           
            imaged.Adodc4.Recordset.Fields!DEPARTEMENT = DEPARTEMENT.Caption
             imaged.Adodc4.Recordset.Fields!image_Title = TxtAttach.Text
             imaged.Adodc4.Recordset.Fields!operation_type = txtopeation_type.Text
             
             
            
            
            imaged.Adodc4.Recordset.update
            imaged.Adodc4.Refresh


Adodc4.Recordset.update
DataGrid2.Refresh
    
    End If
Frame6.Visible = False


End Sub

Private Sub Command9_Click()
Dim s As String

s = "select TblTypeImage.Name as Type1Name,TblTypeImage2.Name as Type2Name, Subject_doc.* from  Subject_doc Left outer join TblTypeImage On Subject_doc.Type1ID =  TblTypeImage.Id "
s = s & " Left outer join TblTypeImage2 On Subject_doc.Type2ID =  TblTypeImage2.Id "
s = s & " WHERE subject_no='" & mIDD & "' and IsNull(NameFile,'') <> ''"

s = s & " and IsNull(IsDeleted,0) = 0 "
If Not IsNull(FromDate.value) Then
        s = s & "  and Subject_doc.Datee2 >= " & SQLDate(FromDate.value, True)
End If

If Not IsNull(ToDate.value) Then
        s = s & "  and Subject_doc.Datee2 <= " & SQLDate(ToDate.value, True)
End If

s = s & " "

loadgrid s, FG, True, True
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
  With FG
  mTypeID = GetTypeId
  .TextMatrix(Row, .ColIndex("Type1ID")) = mTypeID
  Select Case .ColKey(Col)
    Case "Type1Name"
           'If val(.TextMatrix(Row, .ColIndex("PrMainDesID"))) = 0 Then
            If 0 = 0 Then
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Type1ID"), False, True)
                .TextMatrix(Row, .ColIndex("Type1ID")) = StrAccountCode
            Else
                    
                    
                    
            End If
Case "Type2Name"
           'If val(.TextMatrix(Row, .ColIndex("PrMainDesID"))) = 0 Then
            If 0 = 0 Then
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Type2ID"), False, True)
                .TextMatrix(Row, .ColIndex("Type2ID")) = StrAccountCode
            Else
                    
                    
                    
            End If
End Select
End With
saveimage Row
End Sub

Private Sub FG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

Dim Frm As New FrmDateOpProject, mDate As Date, mTime As Date

  mTypeID = GetTypeId
  FG.TextMatrix(Row, FG.ColIndex("Type1ID")) = mTypeID
  
Select Case FG.ColKey(Col)
    Case "CmdSave"
    saveimage Row
    Case "CmdDelete"
        s = "Update    Subject_doc Set IsDeleted = 1 where operation_no = " & val(FG.TextMatrix(Row, FG.ColIndex("operation_no")))
        Cn.Execute s
        s = "select TblTypeImage.Name as Type1Name,TblTypeImage2.Name as Type2Name, Subject_doc.* from  Subject_doc Left outer join TblTypeImage On Subject_doc.Type1ID =  TblTypeImage.Id "
        s = s & " Left outer join TblTypeImage2 On Subject_doc.Type2ID =  TblTypeImage2.Id "
        s = s & " WHERE subject_no='" & mIDD & "' and IsNull(NameFile,'') <> ''"
        s = s & " and Subject_doc.Type1ID = " & GetTypeId
        s = s & " and IsNull(IsDeleted,0) = 0 "

        loadgrid s, FG, True, True

        
        
    Case "NameFile"

'     If FG.TextMatrix(Row, FG.ColIndex("Type1Name")) = "" Then
'                      If SystemOptions.UserInterface = ArabicInterface Then
'                             MsgBox "Õœœ «”„ «·„—›ﬁ", vbCritical
'                             Exit Sub
'
'                Else
'
'                        MsgBox "Enter Attachment No ", vbCritical
'                             Exit Sub
'                End If
'       End If
       
    Dim sourcefilename As String
        Dim desFilename As String
On Error Resume Next
        Dim pos As Integer
        Dim filename As String
        
            NEW_IMAGE = True
        DBPix201.ImageClear
        DBPix201.Visible = True
        Dim LASTIMAGENO As Integer
        Dim X As Integer
'
'             x = MsgBox("Â·  —Ìœ ’Ê—… „‰ „·›", vbExclamation + vbYesNoCancel)
'
'        If x = vbYes Then
'            DBPix201.ImageLoad
'sourcefilename = (filename)
'            DoEvents
'            MsgBox " „  Õ„Ì· «·’Ê—…"
'        Else
'
'            If x = vbNo Then
'                DBPix201.TWAINAcquire
'                MsgBox " „ „”Õ ÷Ê∆Ì  ··’Ê—…"
'
'                DoEvents
'            Else
'
'                Exit Sub
'            End If
'        End If

       
'
'
    dlgFile.filter = "All files(*.*) | *.*"
    dlgFile.ShowOpen
    If dlgFile.filename <> vbNullString And dlgFile.filename <> "" Then

        filename = dlgFile.filename
        sourcefilename = (filename)


        sourcefilename = (filename)
        
        pos = InStrRev(filename, "\")
        If pos > 0 Then
            filename = mId(filename, pos + 1)
        End If
        
         filename = mIDD & "_" & Trim(FG.TextMatrix(Row, FG.ColIndex("Type1Name"))) & "_" & FG.TextMatrix(Row, FG.ColIndex("Type2Name")) & "_" & filename
         'TimeStamp(Date) & filename
   
 End If
         desFilename = App.path & "\Doc\" & filename
         FileCopy sourcefilename, desFilename
      '   DBPix201.ImageSaveFile (App.path & "\Doc\" & filename & ".JPG")
        NEW_IMAGE = False
        FG.TextMatrix(Row, FG.ColIndex("NameFile")) = filename
       

  

           FG.TextMatrix(Row, FG.ColIndex("SUBJECT_NO")) = SUBJECT_NO
                      FG.TextMatrix(Row, FG.ColIndex("IMAGE_NAME")) = filename
           FG.TextMatrix(Row, FG.ColIndex("DEPARTEMENT")) = DEPARTEMENT.Caption
           FG.TextMatrix(Row, FG.ColIndex("image_Title")) = TxtAttach.Text
           FG.TextMatrix(Row, FG.ColIndex("operation_type ")) = txtopeation_type.Text
           FG.TextMatrix(Row, FG.ColIndex("image_date")) = DateValue(Now)
           
             
             saveimage Row
            
            
         

Case "ShowAt"
If Trim(FG.TextMatrix(Row, FG.ColIndex("NameFile"))) <> "" Then
    ShellExecute 0&, vbNullString, App.path & "\Doc\" & Trim(FG.TextMatrix(FG.Row, FG.ColIndex("NameFile"))), vbNullString, vbNullString, vbNormalFocus
End If

    
'    End If
'Frame6.Visible = False
Case "Datee1"
IsFromDate = 10
        Set Frm = New FrmDateOpProject
        Frm.Index = 897
        LngRowP = Row
        imaged2.LngRow2 = Row
        Frm.show 1
        FG.TextMatrix(Row, FG.ColIndex("Datee1")) = mDateP
        imaged2.LngRow2 = Row
        saveimage Row
   Case "Datee2"
IsFromDate = 10
        Set Frm = New FrmDateOpProject
       LngRowP = Row
        Frm.Index = 898
        imaged2.LngRow2 = Row
        DoEvents
        Frm.show 1
        DoEvents
        DoEvents
       FG.TextMatrix(Row, FG.ColIndex("Datee2")) = mDateP
       DoEvents
       saveimage Row
End Select

     
 
End Sub

Private Sub fg_Click()
 NEW_IMAGE = True
        DBPix201.ImageClear
        DBPix201.Visible = True
        Dim LASTIMAGENO As Integer
        Dim X As Integer
        Dim mFileName As String
        mFileName = App.path & "\Doc\" & Trim(FG.TextMatrix(FG.Row, FG.ColIndex("NameFile"))) & ".JPG"
       ' DBPix201.ImageSaveFile (system_path & "\" & SystemOptions.ImagesPath & "\" & mFileName & "")
     '    DBPix201.Image = mFileName
         DBPix201.ImageClear
         DBPix201.ImageViewFile mFileName
         
   '      DBPix201.ImageLoad

            DoEvents
            
       
End Sub

Private Sub Form_Activate()

If IsFromDate = 10 Then Exit Sub
Dim s As String
If mIDD <> "" Then
s = "select TblTypeImage.Name as Type1Name,TblTypeImage2.Name as Type2Name, Subject_doc.* from  Subject_doc Left outer join TblTypeImage On Subject_doc.Type1ID =  TblTypeImage.Id "
s = s & " Left outer join TblTypeImage2 On Subject_doc.Type2ID =  TblTypeImage2.Id "
s = s & " WHERE subject_no='" & mIDD & "' and IsNull(NameFile,'') <> ''"
s = s & " and Subject_doc.Type1ID = " & GetTypeId
s = s & " and IsNull(IsDeleted,0) = 0 "
Else
s = "select    * from  Subject_doc WHERE subject_no='" & SUBJECT_NO & "'"
s = s & " and IsNull(IsDeleted,0) = 0 "
End If
loadgrid s, FG, True, True
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
IsFromDate = 0
    system_path = App.path ' "D:\my works\accountant\28  01 2011\SourceCode\SourceCode"
    connection_string = Cn.ConnectionString
    Adodc4.ConnectionString = connection_string
Adodc4.CommandType = adCmdText
Adodc4.RecordSource = "select * from  Subject_doc WHERE subject_no='" & SUBJECT_NO.Caption & "'"


Dim rs As New ADODB.Recordset
Dim StrSQL  As String
StrSQL = "select ID,Name,Namee from TblTypeImage "
rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
Dim i As Long
i = 0
Do While Not rs.EOF
    Option1(i).Caption = rs!Name & ""
    Option1(i).Tag = rs!ID & ""
    i = i + 1
    rs.MoveNext
    
Loop
Option1(0).value = True
mTypeID = GetTypeId
Dim j As Long
'j = i + 1
FromDate.value = Date
ToDate.value = Date
For j = i To Option1.count - 1
    Option1(j).Visible = False
Next


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

   
'   x = MsgBox("Â·  —Ìœ «·Õ›Ÿ ﬁ»· «€·«ﬁ «·‘«‘…", vbCritical + vbYesNo)
'
'    If x = vbNo Then
'        Exit Sub
'    Else
'        Command1_Click 9
'    End If

End Sub
Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FG

   Select Case .ColKey(Col)
        Case "Date1", "Date2", "NameFile", "Discount", "Total", "Vat", "VatValue"
            .ComboList = ""
        Case "NoteNo"
            .ComboList = ""
        Case "DayMeter"
            .ComboList = ""
        Case "CustName", "Total"
            Cancel = True
        End Select
        
    End With
End Sub


Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  
   Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With FG

        Select Case .ColKey(Col)
 
            Case "Type1Name"
             .TextMatrix(Row, .ColIndex("Type2Name")) = ""
             .TextMatrix(Row, .ColIndex("Type2ID")) = ""
                StrSQL = "select ID,Name,Namee from TblTypeImage "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = FG.BuildComboList(rs, "Name", "ID")
                Else
                    StrComboList = FG.BuildComboList(rs, "Namee", "ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
            Case "Type2Name"
            ' .TextMatrix(Row, .ColIndex("Type2Name")) = ""
                mTypeID = GetTypeId
                
                FG.TextMatrix(Row, FG.ColIndex("Type1ID")) = mTypeID
                
                StrSQL = "select ID,Name,Namee from TblTypeImage2 Where MasterId =  " & val(FG.ValueMatrix(Row, FG.ColIndex("Type1ID"))) & ""
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = FG.BuildComboList(rs, "Name", "ID")
                Else
                    StrComboList = FG.BuildComboList(rs, "Namee", "ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
                FG.TextMatrix(Row, FG.ColIndex("Datee1")) = Date
                FG.TextMatrix(Row, FG.ColIndex("Datee2")) = Date
          
            
            End Select
        End With
End Sub


Private Sub Option1_Click(Index As Integer)
    If Option1(Index).value = True Then
     '   saveimage
        mTypeID = GetTypeId
        
        
        Dim s As String

        s = "select TblTypeImage.Name as Type1Name,TblTypeImage2.Name as Type2Name, Subject_doc.* from  Subject_doc Left outer join TblTypeImage On Subject_doc.Type1ID =  TblTypeImage.Id "
        s = s & " Left outer join TblTypeImage2 On Subject_doc.Type2ID =  TblTypeImage2.Id "
        s = s & " WHERE subject_no='" & mIDD & "' and IsNull(NameFile,'') <> ''"
        s = s & " and Subject_doc.Type1ID = " & mTypeID
        s = s & " and IsNull(IsDeleted,0) = 0 "
'        If Not IsNull(FromDate.value) Then
'                s = s & "  and Subject_doc.Datee2 >= " & SQLDate(FromDate.value, True)
'        End If
'
'        If Not IsNull(ToDate.value) Then
'                s = s & "  and Subject_doc.Datee2 <= " & SQLDate(ToDate.value, True)
'        End If
        
        s = s & " "
        
        loadgrid s, FG, True, True

    End If
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

Private Sub VSFlexGrid3_Click()

End Sub
