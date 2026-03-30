VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMessnger 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·—”«∆· «·œ«Œ·ÌÂ"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13440
   Icon            =   "FrmMessenger.frx":0000
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   13440
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9420
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13440
      _cx             =   23707
      _cy             =   16616
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
      Begin VB.TextBox Text2 
         DataField       =   "id"
         DataSource      =   "Adodc2"
         Height          =   495
         Left            =   4200
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   8640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   0
         Width           =   13425
         Begin VB.TextBox TxtVac_ID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Height          =   240
            Left            =   3030
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   510
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2580
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Text            =   "modflag"
            Top             =   90
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Frame Frmo2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   450
            Visible         =   0   'False
            Width           =   3105
            Begin MSDataListLib.DataCombo DCUser 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   -255
               TabIndex        =   3
               Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
               Top             =   15
               Width           =   2340
               _ExtentX        =   4128
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               BackColor       =   -2147483624
               Text            =   ""
               RightToLeft     =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "«·„” Œœ„"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   13
               Left            =   2160
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   45
               Width           =   855
            End
         End
         Begin MSComctlLib.ImageList GrdImageList 
            Left            =   3120
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   8
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmMessenger.frx":000C
                  Key             =   "CompanyName"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmMessenger.frx":03A6
                  Key             =   "Ser"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmMessenger.frx":0740
                  Key             =   "Vac_Name"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmMessenger.frx":0ADA
                  Key             =   "ShareCount"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmMessenger.frx":0E74
                  Key             =   "Dis_Count"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmMessenger.frx":120E
                  Key             =   "Bouns"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmMessenger.frx":15A8
                  Key             =   "SharesValue"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmMessenger.frx":1B42
                  Key             =   "BuyValue"
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·—”«∆· «·œ«Œ·Ì…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   17640
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   120
            Width           =   3240
         End
      End
      Begin ALLButtonS.ALLButton CMD_language 
         Height          =   495
         Left            =   360
         TabIndex        =   8
         ToolTipText     =   "Language  «··€…"
         Top             =   0
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
         MICON           =   "FrmMessenger.frx":1EDC
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
         Height          =   495
         Left            =   1680
         Top             =   8760
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
         Caption         =   "Adodc1"
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   8775
         Left            =   0
         TabIndex        =   10
         Top             =   600
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   15478
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "«·»—Ìœ «·’«œ—"
         TabPicture(0)   =   "FrmMessenger.frx":1EF8
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Command4"
         Tab(0).Control(1)=   "Text4"
         Tab(0).Control(2)=   "DataGrid2"
         Tab(0).Control(3)=   "DataGrid4"
         Tab(0).Control(4)=   "ALLButton2"
         Tab(0).Control(5)=   "Label4(3)"
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "«·»—Ìœ  «·Ê«—œ"
         TabPicture(1)   =   "FrmMessenger.frx":1F14
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Command3"
         Tab(1).Control(1)=   "Text3"
         Tab(1).Control(2)=   "DataGrid1"
         Tab(1).Control(3)=   "ALLButton1"
         Tab(1).Control(4)=   "DataGrid3"
         Tab(1).Control(5)=   "Label4(2)"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "«‰‘«¡ —”«·Â ÃœÌœ…"
         TabPicture(2)   =   "FrmMessenger.frx":1F30
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Label5"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Label4(0)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label3"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Label2"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Label4(1)"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "DataCombo1"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "TxtSubjects"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "Command1"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "Text1"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "Command2"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "Command5"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).ControlCount=   11
         Begin VB.CommandButton Command5 
            Caption         =   "ÃœÌœ"
            Height          =   345
            Left            =   3750
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   7560
            Width           =   1725
         End
         Begin VB.CommandButton Command4 
            Caption         =   "«·„—›ﬁ« "
            Height          =   375
            Left            =   -73890
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   2670
            Width           =   1935
         End
         Begin VB.CommandButton Command3 
            Caption         =   "«·„—›ﬁ« "
            Height          =   495
            Left            =   -71820
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   2610
            Width           =   1935
         End
         Begin VB.CommandButton Command2 
            Caption         =   "«·„—›ﬁ« "
            Enabled         =   0   'False
            Height          =   345
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   7590
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6255
            Left            =   0
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   1440
            Width           =   12255
         End
         Begin VB.CommandButton Command1 
            Caption         =   "«—”«·"
            Height          =   345
            Left            =   120
            TabIndex        =   18
            Top             =   7590
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            DataField       =   "message"
            DataSource      =   "Adodc3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6135
            Left            =   -74880
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   3120
            Width           =   12255
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            DataField       =   "message"
            DataSource      =   "Adodc4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5115
            Left            =   -74850
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   3420
            Width           =   12615
         End
         Begin VB.TextBox TxtSubjects 
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
            Left            =   1950
            MaxLength       =   40000
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Tag             =   "Ì—Ã∆ ≈œŒ«· „œ… «·„ﬂ«›√…"
            Top             =   1080
            Width           =   10290
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "FrmMessenger.frx":1F4C
            Height          =   2175
            Left            =   -74880
            TabIndex        =   11
            Top             =   480
            Width           =   13215
            _ExtentX        =   23310
            _ExtentY        =   3836
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
            ColumnCount     =   10
            BeginProperty Column00 
               DataField       =   "to"
               Caption         =   "«·„—”· «·Ì…"
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
               DataField       =   "subject"
               Caption         =   "«·„Ê÷Ê⁄"
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
               DataField       =   "date"
               Caption         =   "«· «—ÌŒ"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3073
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "time"
               Caption         =   "«·Êﬁ "
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "H.mm.ss"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3073
                  SubFormatType   =   4
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "from"
               Caption         =   "from"
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
               DataField       =   "to"
               Caption         =   "to"
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
               DataField       =   "message"
               Caption         =   "«·„Ê÷Ê⁄"
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
               DataField       =   "sender_show"
               Caption         =   "sender_show"
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
               DataField       =   "reciever_show"
               Caption         =   "reciever_show"
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
               DataField       =   "recived"
               Caption         =   "recived"
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
                  ColumnWidth     =   3000.189
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   5999.812
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2505.26
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1800
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1800
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1049.953
               EndProperty
               BeginProperty Column09 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   764.787
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "FrmMessenger.frx":1F61
            Height          =   1935
            Left            =   -74880
            TabIndex        =   13
            Top             =   480
            Width           =   13215
            _ExtentX        =   23310
            _ExtentY        =   3413
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
            ColumnCount     =   10
            BeginProperty Column00 
               DataField       =   "from"
               Caption         =   "«·„—”·"
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
               DataField       =   "subject"
               Caption         =   "«·„Ê÷Ê⁄"
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
               DataField       =   "date"
               Caption         =   "«· «—ÌŒ"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3073
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "time"
               Caption         =   "«·Êﬁ "
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "H.mm.ss"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3073
                  SubFormatType   =   4
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "from"
               Caption         =   "from"
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
               DataField       =   "to"
               Caption         =   "to"
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
               DataField       =   "message"
               Caption         =   "«·„Ê÷Ê⁄"
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
               DataField       =   "sender_show"
               Caption         =   "sender_show"
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
               DataField       =   "reciever_show"
               Caption         =   "reciever_show"
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
               DataField       =   "recived"
               Caption         =   "recived"
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
                  ColumnWidth     =   3000.189
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   5999.812
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1995.024
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1800
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1049.953
               EndProperty
               BeginProperty Column09 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   764.787
               EndProperty
            EndProperty
         End
         Begin ALLButtonS.ALLButton ALLButton1 
            Height          =   255
            Left            =   -74880
            TabIndex        =   14
            Top             =   2520
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "Del"
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
            MICON           =   "FrmMessenger.frx":1F76
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataGridLib.DataGrid DataGrid4 
            Bindings        =   "FrmMessenger.frx":1F92
            Height          =   2175
            Left            =   -74880
            TabIndex        =   15
            Top             =   480
            Width           =   13215
            _ExtentX        =   23310
            _ExtentY        =   3836
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
            ColumnCount     =   10
            BeginProperty Column00 
               DataField       =   "to"
               Caption         =   "TO"
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
               DataField       =   "message"
               Caption         =   "subject"
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
               DataField       =   "date"
               Caption         =   "Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3073
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "time"
               Caption         =   "Time"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "H.mm.ss"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3073
                  SubFormatType   =   4
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "from"
               Caption         =   "from"
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
               DataField       =   "to"
               Caption         =   "to"
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
               DataField       =   "message"
               Caption         =   "«·„Ê÷Ê⁄"
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
               DataField       =   "sender_show"
               Caption         =   "sender_show"
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
               DataField       =   "reciever_show"
               Caption         =   "reciever_show"
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
               DataField       =   "recived"
               Caption         =   "recived"
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
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   3000.189
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   6494.74
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1995.024
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1049.953
               EndProperty
               BeginProperty Column09 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   764.787
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "FrmMessenger.frx":1FA7
            Height          =   315
            Left            =   7300
            TabIndex        =   20
            Top             =   720
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "UserName"
            BoundColumn     =   "UserID"
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ALLButtonS.ALLButton ALLButton2 
            Height          =   255
            Left            =   -74760
            TabIndex        =   21
            Top             =   2760
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "DEl"
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
            MICON           =   "FrmMessenger.frx":1FBC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataGridLib.DataGrid DataGrid3 
            Bindings        =   "FrmMessenger.frx":1FD8
            Height          =   1935
            Left            =   -74880
            TabIndex        =   22
            Top             =   480
            Visible         =   0   'False
            Width           =   13215
            _ExtentX        =   23310
            _ExtentY        =   3413
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
            ColumnCount     =   10
            BeginProperty Column00 
               DataField       =   "from"
               Caption         =   "From"
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
               DataField       =   "message"
               Caption         =   "subject"
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
               DataField       =   "date"
               Caption         =   "Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3073
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "time"
               Caption         =   "Time"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "H.mm.ss"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3073
                  SubFormatType   =   4
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "from"
               Caption         =   "from"
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
               DataField       =   "to"
               Caption         =   "to"
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
               DataField       =   "message"
               Caption         =   "«·„Ê÷Ê⁄"
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
               DataField       =   "sender_show"
               Caption         =   "sender_show"
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
               DataField       =   "reciever_show"
               Caption         =   "reciever_show"
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
               DataField       =   "recived"
               Caption         =   "recived"
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
               SizeMode        =   1
               AllowSizing     =   0   'False
               BeginProperty Column00 
                  Alignment       =   1
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   3495.118
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   6105.26
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2294.929
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1049.953
               EndProperty
               BeginProperty Column09 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   764.787
               EndProperty
            EndProperty
         End
         Begin VB.Label Label4 
            Caption         =   " «·—”«·…"
            Height          =   375
            Index           =   3
            Left            =   -62160
            TabIndex        =   29
            Top             =   3000
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   " «·—”«·…"
            Height          =   375
            Index           =   2
            Left            =   -62520
            TabIndex        =   28
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   " «·—”«·…"
            Height          =   375
            Index           =   1
            Left            =   12360
            TabIndex        =   27
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "«·Ï"
            Height          =   375
            Left            =   12720
            TabIndex        =   26
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "«·„Ê÷Ê⁄"
            Height          =   375
            Index           =   0
            Left            =   12360
            TabIndex        =   24
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Subject"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   495
         Left            =   1080
         Top             =   6480
         Width           =   2175
         _ExtentX        =   3836
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
         Caption         =   "Adodc1"
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
         Height          =   495
         Left            =   6240
         Top             =   5400
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
         Caption         =   "Adodc1"
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
         Height          =   495
         Left            =   120
         Top             =   5520
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
         Caption         =   "Adodc1"
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
   End
End
Attribute VB_Name = "FrmMessnger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub ALLButton1_Click()
    On Error Resume Next

    If Adodc3.Recordset.RecordCount = 0 Then Exit Sub
    Adodc3.Recordset.Fields!reciever_show = vbFalse
    Adodc3.Recordset.update
    Adodc3.Refresh

    DataGrid1.Refresh
    DataGrid3.Refresh

    Adodc4.Refresh

    DataGrid2.Refresh
    DataGrid4.Refresh

End Sub

Private Sub ALLButton2_Click()
    On Error Resume Next

    If Adodc4.Recordset.RecordCount = 0 Then Exit Sub
    Adodc4.Recordset.Fields!sender_show = vbFalse
    Adodc4.Recordset.update
    Adodc4.Refresh

    DataGrid2.Refresh
    DataGrid4.Refresh

    Adodc3.Refresh

    DataGrid1.Refresh
    DataGrid3.Refresh

End Sub

'You Must Follow The Letters If They Are Small Write Them In Small Letters If 'They Are Big Write Them In Big Letters So That This Code Works.

Private Sub CMD_language_Click()
    On Error Resume Next

    If CMD_language.Caption = "EN" Then
        my_language = "E"
 
        '''Call Reload(Me)
 
    Else
        my_language = "A"
 
        '''Call Reload(Me)
    End If

End Sub

Private Sub Command1_Click()
    On Error Resume Next

    If DataCombo1.Text = "" Then
        MsgBox "«Œ — «·„—”· «·Ì…" & CHR(13) & "select to ", vbCritical
        Exit Sub

    End If

    If TxtSubjects.Text = "" Then
        MsgBox "  ·«»œ „‰ ﬂ «»… ⁄‰Ê«‰ ··—”«·…  " & CHR(13) & "select to ", vbCritical
        Exit Sub

    End If

'    Adodc2.Recordset.AddNew

    Adodc2.Recordset.Fields![From] = user_name
    Adodc2.Recordset.Fields![To] = DataCombo1.Text
    Adodc2.Recordset.Fields!subject = TxtSubjects.Text

    Adodc2.Recordset.Fields!Message = Text1.Text
    Adodc2.Recordset.Fields!Date = Date
    Adodc2.Recordset.Fields!Time = Time

    Adodc2.Recordset.Fields!sender_show = vbTrue
    Adodc2.Recordset.Fields!reciever_show = vbTrue
    Adodc2.Recordset.Fields!recived = vbFalse

    Adodc2.Recordset.update
    MsgBox " „ «—”«· «·—”«·…" & CHR(13) & "message was sent", vbInformation

    Adodc3.Refresh

    DataGrid1.Refresh
    DataGrid3.Refresh

    Adodc4.Refresh

    DataGrid2.Refresh
    DataGrid4.Refresh
    Command2.Enabled = True
    Command5.Enabled = True
End Sub

Private Sub Command2_Click()
    On Error Resume Next
ShowAttachments Trim(Adodc2.Recordset.Fields!ID & ""), "31780319"
End Sub

Private Sub Command3_Click()
    On Error Resume Next
ShowAttachments Trim(Adodc3.Recordset.Fields!ID & ""), "31780319"

End Sub

Private Sub Command4_Click()
    On Error Resume Next
ShowAttachments Trim(Adodc4.Recordset.Fields!ID & ""), "31780319"

End Sub

Private Sub Command5_Click()
 Adodc2.Recordset.AddNew
   Adodc2.Recordset.Fields![From] = user_name
    Adodc2.Recordset.Fields![To] = DataCombo1.Text
    Adodc2.Recordset.Fields!subject = TxtSubjects.Text

    Adodc2.Recordset.Fields!Message = Text1.Text
    Adodc2.Recordset.Fields!Date = Date
    Adodc2.Recordset.Fields!Time = Time

    Adodc2.Recordset.Fields!sender_show = vbTrue
    Adodc2.Recordset.Fields!reciever_show = vbTrue
    Adodc2.Recordset.Fields!recived = vbFalse

    Adodc2.Recordset.update

Command5.Enabled = False
Command2.Enabled = True
Command1.Enabled = True
End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, _
                             Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF5 Then
        Adodc1.Refresh
        DataCombo1.ReFill
    End If

End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    '   Beep
    'Put This Code on a Command Button Or Anywhere You Want

    'Use Your Own Music File And Put The Inside Te Folder Where The Application is
    'This Is The Correct Code

    'Me.left = (mdifrmmain.Width - Me.Width) / 2
    'Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
        Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
   
      
      
    
    If SystemOptions.UserInterface = EnglishInterface Then
        CMD_language.Caption = "⁄—»Ì"
        ALLButton2.Caption = "Delete"
        ALLButton1.Caption = "Delete"
        SSTab1.TabCaption(0) = "Outbox "
        SSTab1.TabCaption(1) = "Inbox"
        SSTab1.TabCaption(2) = "New message"
        DataGrid3.Visible = True
        DataGrid4.Visible = True
        Label3.Visible = True
        Label2.Visible = False
        Me.Caption = "Message Screen"
        Command1.Caption = "Send"
        DataCombo1.RightToLeft = False
        Text1.Alignment = 0
        Text3.Alignment = 0
        Text4.Alignment = 0
SSTab1.TabCaption(0) = "OutBox"
SSTab1.TabCaption(0) = "InBox"
SSTab1.TabCaption(0) = "New Message"
Label1(2).Caption = "Local Messenger"
    End If
    

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "SELECT * FROM TblUsers  where UserID<> " & user_id
    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT * FROM Messages  "
    Adodc2.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "SELECT  *  FROM  Messages  where   reciever_show=1 and  [to]='" & user_name & "' order by id desc"
    Adodc3.Refresh
    DataGrid1.Refresh

    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText
    Adodc4.RecordSource = "SELECT * FROM Messages  where       sender_show=1 and [from]='" & user_name & "' order by id desc"
    Adodc4.Refresh

    DataGrid2.Refresh

End Sub

Private Sub Text3_Change()
    On Error Resume Next

    If Adodc3.Recordset.RecordCount > 0 Then
        Adodc3.Recordset.Fields!recived = vbTrue
        Adodc3.Recordset.update
    End If

End Sub

Private Sub Text4_Change()
    'If Adodc4.Recordset.RecordCount > 0 Then
    'Adodc4.Recordset.Fields!recived = vbTrue
    'Adodc4.Recordset.Update
    'End If
End Sub
