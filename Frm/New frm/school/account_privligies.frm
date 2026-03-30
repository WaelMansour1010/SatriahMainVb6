VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form account_privligies 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "’·«ÕÌ«  «·Õ”«»« "
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   2100
   ScaleWidth      =   4905
   Begin VB.Frame Frame1 
      Height          =   2040
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "«÷«›…"
         Height          =   255
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   855
         Left            =   840
         TabIndex        =   1
         Top             =   960
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1508
         _Version        =   393216
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   "„"
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
            DataField       =   ""
            Caption         =   "—ﬁ„ «·ﬁ”„"
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·ﬁ”„"
         Height          =   255
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "account_privligies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
End Sub

