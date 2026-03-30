VERSION 5.00
Object = "{798A85D3-625A-4512-A9E4-BA96E09CA6A6}#1.0#0"; "ciaXPIML30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3D800911-77E3-43DE-82EA-7FC87C713180}#1.2#0"; "cPopMenu6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "DOCKIN~1.OCX"
Begin VB.MDIForm mdifrmmain 
   BackColor       =   &H00FFFFFF&
   Caption         =   " "
   ClientHeight    =   8295
   ClientLeft      =   5730
   ClientTop       =   3840
   ClientWidth     =   8700
   Icon            =   "MDIFrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerAlret 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.StatusBar XPStusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   7950
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   7080
      Top             =   4080
   End
   Begin cPopMenu6.PopMenu PopMenu1 
      Left            =   4680
      Top             =   2040
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
      HighlightStyle  =   2
      InActiveMenuForeColor=   0
      MenuBackgroundColor=   16777152
      BackgroundPicture=   "MDIFrmMain.frx":16E48
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgLstTree 
      Left            =   4200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   68
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":179A4
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":17D3E
            Key             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":180D8
            Key             =   "Refresh"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":18472
            Key             =   "receipt"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1880C
            Key             =   "Required"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":18DA6
            Key             =   "Balance"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":19140
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":194DA
            Key             =   "Dollar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1AB34
            Key             =   "Item2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1AECE
            Key             =   "Design"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1B268
            Key             =   "Request"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1B802
            Key             =   "Favorites"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1BB9C
            Key             =   "Wizared"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1BF36
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1C2D0
            Key             =   "Excute"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1C66A
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1CC04
            Key             =   "New"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1CF9E
            Key             =   "save"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1D338
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1D6D2
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1DA6C
            Key             =   "Sall"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1DE06
            Key             =   "Clients"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1E1A0
            Key             =   "Groups"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1E53A
            Key             =   "Maintenance"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1E8D4
            Key             =   "Items"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1EC6E
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1F008
            Key             =   "Supplier"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1F3A2
            Key             =   "barcode"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1F73C
            Key             =   "ReturnBack"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1FCD6
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":20070
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2040A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":207A4
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":20B3E
            Key             =   "Purchase"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":20ED8
            Key             =   "store"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":21272
            Key             =   "LIST"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2160C
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":219A6
            Key             =   "DReport"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":21D40
            Key             =   "From"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":220DA
            Key             =   "To"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":22474
            Key             =   "User"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2280E
            Key             =   "Tax"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":22BA8
            Key             =   "Currency"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":22F42
            Key             =   "Discount"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":232DC
            Key             =   "DiscountType"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":23676
            Key             =   "Tick"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":23A10
            Key             =   "Date"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":23DAA
            Key             =   "Ask"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":24344
            Key             =   "number"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":246DE
            Key             =   "qty"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":24A78
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":24E12
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":251AC
            Key             =   "Closed_Node"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":25546
            Key             =   "Open_Node"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":258E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":25E7A
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":26214
            Key             =   "Serial"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":265AE
            Key             =   "code"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":26948
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":26CE2
            Key             =   "Plus"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2707C
            Key             =   "Minus"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":27416
            Key             =   "FillData"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":277B0
            Key             =   "GridOptions"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":27B4A
            Key             =   "Tree"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":27EE4
            Key             =   "Assblied"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2827E
            Key             =   "LinkItem"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":28618
            Key             =   "ItemPart"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":289B2
            Key             =   "Filter"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Cmdlg 
      Left            =   5160
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgLstMenuIcons 
      Left            =   4680
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   127
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":28D4C
            Key             =   "Salles"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":290E6
            Key             =   "Warn"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":29480
            Key             =   "Screen"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2981A
            Key             =   "Execute"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":29BB4
            Key             =   "New"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":29F4E
            Key             =   "Purashes"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2A2E8
            Key             =   "DEV_Preview"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2A682
            Key             =   "OpenAcc"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2AC1C
            Key             =   "AccReports"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2AFB6
            Key             =   "Options"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2B350
            Key             =   "Emp"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2B8EA
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2BC84
            Key             =   "Items"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2CAD6
            Key             =   "store"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2D3B0
            Key             =   "Invoice"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2D74A
            Key             =   "NewAccout"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2DAE4
            Key             =   "NewGroupAccount"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2DE7E
            Key             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2E218
            Key             =   "ToGroup"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2E7B2
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2EB4C
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2EEE6
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2F280
            Key             =   "Screens"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2F61A
            Key             =   "HotKey"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2F934
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2FCCE
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":30068
            Key             =   "Tools"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":30402
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3079C
            Key             =   "PrintSetup"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":30B36
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":30ED0
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3126A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":31604
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3199E
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":31D38
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":320D2
            Key             =   "MoveFirst"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3246C
            Key             =   "MovePrevious"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":32806
            Key             =   "MoveNext"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":32BA0
            Key             =   "MoveLast"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":32F3A
            Key             =   "Money1"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":33814
            Key             =   "ToolTip"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":33BAE
            Key             =   "DEV_Edit"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":33EC8
            Key             =   "Reports"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":34262
            Key             =   "Suppliers"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":347FC
            Key             =   "Customers"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3564E
            Key             =   "Help1"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":37358
            Key             =   "Cal"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":376F2
            Key             =   "OpenStore"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":37B44
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":37EDE
            Key             =   "EditTree"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":38278
            Key             =   "NewItem"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":38612
            Key             =   "Users"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":389AC
            Key             =   "AddUser"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":38D46
            Key             =   "DeleteUser"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":390E0
            Key             =   "UserPass"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3947A
            Key             =   "UserPremis"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":39814
            Key             =   "DataBaseBackup"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":39BAE
            Key             =   "DataBaseRestore"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":39F48
            Key             =   "DataBaseRepaire"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3A2E2
            Key             =   "NewDataBase"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3A67C
            Key             =   "DataBaseReg"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3AA16
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3AE68
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3BB42
            Key             =   "Tick"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3BEDC
            Key             =   "Home"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3C276
            Key             =   "TreeItems"
            Object.Tag             =   "65"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3C610
            Key             =   "NewGroup"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3C9AA
            Key             =   "DataBase"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3CD44
            Key             =   "About"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3D0DE
            Key             =   "WindowMin"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3D478
            Key             =   "WindowMax"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3D812
            Key             =   "City"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3E0EC
            Key             =   "GridDelRow"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3E486
            Key             =   "Bank"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3E7A0
            Key             =   "Pur"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3E8FA
            Key             =   "OutOrder"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3EC94
            Key             =   "InOrder"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3F02E
            Key             =   "Dev_Screen"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3F3C8
            Key             =   "Prop"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3F762
            Key             =   "Money2"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3FAFC
            Key             =   "Money3"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3FE96
            Key             =   "DefColor"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":40230
            Key             =   "CusColor"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":405CA
            Key             =   "Caps"
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":40964
            Key             =   "Clock"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":40CFE
            Key             =   "Num"
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":41098
            Key             =   "Calender"
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":41432
            Key             =   "User"
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":417CC
            Key             =   "KeyBorad"
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":420A6
            Key             =   "LogOFF"
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":42640
            Key             =   "Interface"
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":429DA
            Key             =   "BarCode"
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":42D74
            Key             =   "UserOptions"
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4310E
            Key             =   "InvoiceDesign"
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":434A8
            Key             =   "Unit"
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":43842
            Key             =   "grd"
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":43BDC
            Key             =   "StoreCon"
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":44176
            Key             =   "StoreEx"
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":44510
            Key             =   "StoreIm"
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":448AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":44E44
            Key             =   "Web"
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":453DE
            Key             =   "wazrid"
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":45778
            Key             =   "Vertical"
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":45B12
            Key             =   "Horizental"
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":45EAC
            Key             =   "TabDown"
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":46246
            Key             =   "TabRight"
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":465E0
            Key             =   "TabUp"
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4697A
            Key             =   "TabLeft"
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":46D14
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":470AE
            Key             =   "ItemsPrice"
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":47448
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":477E2
            Key             =   "Unlock"
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":47B7C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":48116
            Key             =   "Help2"
            Object.Tag             =   "113"
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":484B0
            Key             =   "SearchHelp"
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4884A
            Key             =   "Hide"
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":48DE4
            Key             =   "SortASC"
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4917E
            Key             =   "SortDESC"
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":49518
            Key             =   "BrowseFile"
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":49AB2
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":49E4C
            Key             =   "ExportExcel"
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4A1E6
            Key             =   "ExportPDF"
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4A580
            Key             =   "ExportWord"
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4A91A
            Key             =   "ExportHTML"
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4ACB4
            Key             =   "ExportMail"
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4B04E
            Key             =   "Plus"
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4B3E8
            Key             =   "Mins"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5280
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4B782
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4BE5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4C546
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4CC36
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4D30E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4D9EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4E0DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4E7CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4EEAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4F592
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4FC72
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":5035E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":50A4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":5113A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":5181C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":51F01
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711680
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":525EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":52CD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":533A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":53A63
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":5411A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":547DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":54E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":55565
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":55C28
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":562F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":569D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":570A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":57777
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":57E34
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":584F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":58BC6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ciaXPImageList30.XPImageList30 img16 
      Left            =   4680
      Top             =   2670
      _ExtentX        =   1005
      _ExtentY        =   1005
      Size            =   10340
      Images          =   "MDIFrmMain.frx":59294
      KeyCount        =   11
      Keys            =   "ˇˇˇˇˇˇˇˇˇˇ"
   End
   Begin vbalIml6.vbalImageList ilsIcons 
      Left            =   5340
      Top             =   2670
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   32
   End
   Begin XtremeDockingPane.DockingPane DockingPane1 
      Left            =   5880
      Top             =   3960
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   5
      ScaleMode       =   1
   End
   Begin VB.Menu MdiContextMenu 
      Caption         =   "«·ﬁ«∆„… «·«”«”Ì…"
      Begin VB.Menu MnuAccounts 
         Caption         =   "«·Õ”«»« "
         Begin VB.Menu MnuAccCharts 
            Caption         =   "«·œ·Ì· «·„Õ«”»Ì"
            Index           =   0
         End
         Begin VB.Menu MnuAccCharts 
            Caption         =   "«·ﬁÌœ «·«›  «ÕÌ ··Õ”«»« "
            Index           =   1
         End
         Begin VB.Menu MnuAccDEV 
            Caption         =   " Õ—Ì— ﬁÌÊœ «·ÌÊ„Ì…"
            Index           =   0
         End
         Begin VB.Menu MnuAccDEV 
            Caption         =   "ﬁÌÊœ «· ”ÊÌ… «·ÌœÊÌ…"
            Index           =   1
         End
         Begin VB.Menu MnuAccDEV_Post 
            Caption         =   "„—«Ã⁄… Ê —ÕÌ· ﬁÌÊœ«·ÌÊ„Ì…"
            Visible         =   0   'False
         End
         Begin VB.Menu xxx 
            Caption         =   "«‰Ê«⁄ „—«ﬂ“ «· ﬂ·›…"
            Index           =   0
         End
         Begin VB.Menu xxx 
            Caption         =   "„—«ﬂ“ «· ﬂ·›…"
            Index           =   1
         End
         Begin VB.Menu xxx 
            Caption         =   " ﬁ«—Ì— «·Õ”«»« "
            Index           =   12
         End
      End
      Begin VB.Menu Currency 
         Caption         =   "«·„⁄«„·«  «·„«·Ì…"
         Begin VB.Menu ExpensesType 
            Caption         =   "√‰Ê«⁄ «·„’—Ê›« "
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu ExpensesType 
            Caption         =   "√‰Ê«⁄ «·≈Ì—«œ« "
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu ExpensesType 
            Caption         =   "œ›« — «·‘Ìﬂ« "
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu Expenses 
            Caption         =   "›« Ê—… „«·Ì…"
            Index           =   0
         End
         Begin VB.Menu Expenses 
            Caption         =   "›« Ê—… Œœ„Ì…"
            Index           =   1
         End
         Begin VB.Menu Expenses 
            Caption         =   "”‰œ«  «·’—›"
            Index           =   2
            Begin VB.Menu ExpensesSub 
               Caption         =   "«‰Ê«⁄ «·’—›"
               Index           =   0
            End
            Begin VB.Menu ExpensesSub 
               Caption         =   "ÿ» ’—›"
               Index           =   1
            End
            Begin VB.Menu ExpensesSub 
               Caption         =   "”‰œ«  «·’—› -  Õ·Ì·Ì „’—Ê›« "
               Index           =   2
            End
            Begin VB.Menu ExpensesSub 
               Caption         =   "”‰œ«  «·’—› - «·„œ›Ê⁄« "
               Index           =   3
            End
            Begin VB.Menu ExpensesSub 
               Caption         =   "”‰œ ’—› „ ⁄œœ"
               Index           =   4
            End
         End
         Begin VB.Menu Payments 
            Caption         =   "«·„œ›Ê⁄« "
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu Cashing 
            Caption         =   "«·„ﬁ»Ê÷« "
            Index           =   0
         End
         Begin VB.Menu Cashing 
            Caption         =   "”‰œ «·ﬁ»÷ «·⁄«„"
            Index           =   1
         End
         Begin VB.Menu Cashing 
            Caption         =   "ÿ»«⁄… «·‘Ìﬂ« "
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu Cashing 
            Caption         =   "«Ìœ«⁄«  »‰ﬂÌÂ"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu Cashing 
            Caption         =   " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« "
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu Cashing 
            Caption         =   "„–ﬂ—… »‰ﬂ"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu Cashing 
            Caption         =   "«· ”ÊÌ«  «·»ﬂÌ…"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu DelayVal 
            Caption         =   "«·√Ê—«ﬁ «·„«·Ì… «·„” Õﬁ…"
            Index           =   0
         End
         Begin VB.Menu DelayVal 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuFinDiscounts 
            Caption         =   "«·Œ’Ê„«  «·„”„ÊÕ… Ê«·„ﬂ ”»…"
         End
         Begin VB.Menu ReceiptPart 
            Caption         =   " Õ’Ì· Ê”œ«œ √ﬁ”«ÿ"
            Visible         =   0   'False
         End
         Begin VB.Menu RequiredInstallment 
            Caption         =   "«·√ﬁ”«ÿ «·„ÿ·Ê»…"
            Visible         =   0   'False
         End
         Begin VB.Menu MnuCheckOperations 
            Caption         =   " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« "
            Visible         =   0   'False
         End
         Begin VB.Menu MnuCheckBriefcase 
            Caption         =   "Õ«›Ÿ… «·‘Ìﬂ« "
            Visible         =   0   'False
         End
         Begin VB.Menu MnuBoxDeposit 
            Caption         =   "«·«—’œ… «·«›  «ÕÌ…"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu MnuBoxDeposit 
            Caption         =   " „ÊÌ· «·Œ“‰ Ê«” ⁄«÷… «·⁄Âœ"
            Index           =   1
         End
         Begin VB.Menu MnuBoxDeposit 
            Caption         =   " ’›Ì… «·⁄Âœ"
            Index           =   2
         End
         Begin VB.Menu MnuBoxDrawing 
            Caption         =   " ÕÊÌ·«  „«·ÌÂ"
         End
         Begin VB.Menu MnuBoxAccouns 
            Caption         =   "—’Ìœ «·Œ“‰… «·√‰..."
         End
         Begin VB.Menu MnuBoxIncapacity_Increase 
            Caption         =   "“Ì«œ… Ê⁄Ã“ ›Ï ‰ﬁœÌ… «·Œ“‰…"
            Index           =   0
         End
      End
      Begin VB.Menu BankOp 
         Caption         =   "«·„⁄«„·«  «·»‰ﬂÌ…"
         Begin VB.Menu BankOpsub 
            Caption         =   "«·«Ìœ«⁄«  «·»ﬂÌ…"
            Index           =   0
         End
         Begin VB.Menu BankOpsub 
            Caption         =   " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« "
            Index           =   1
         End
         Begin VB.Menu BankOpsub 
            Caption         =   "«· ”ÊÌ«  «·»‰ﬂÌ…"
            Index           =   2
         End
         Begin VB.Menu BankOpsub 
            Caption         =   "„–ﬂ—… »‰ﬂ"
            Index           =   3
         End
         Begin VB.Menu BankOpsub 
            Caption         =   "ÿ»«⁄Â «·‘Ìﬂ« "
            Index           =   4
         End
         Begin VB.Menu BankOpsub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   5
         End
      End
      Begin VB.Menu MNUFixedAssets 
         Caption         =   "«·«’Ê· «·À«» …"
         Begin VB.Menu xxxxx 
            Caption         =   "„Ã„Ê⁄«  «·«’Ê·                  "
            Index           =   0
         End
         Begin VB.Menu xxxxx 
            Caption         =   "»Ì«‰«  «·«’Ê· «·À«» …"
            Index           =   1
         End
         Begin VB.Menu xxxxx 
            Caption         =   "›« Ê—… ‘—«¡ «’·"
            Index           =   2
         End
         Begin VB.Menu xxxxx 
            Caption         =   "«’œ«— «ﬁ”«ÿ «·«Â·«ﬂ"
            Index           =   3
         End
         Begin VB.Menu xxxxx 
            Caption         =   "«· Œ·’ «Ê «” »⁄«œ«  «·«’Ê·"
            Index           =   4
         End
         Begin VB.Menu xxxxx 
            Caption         =   "«÷«›«  «·«’Ê·"
            Index           =   5
         End
         Begin VB.Menu xxxxx 
            Caption         =   "Õ—ﬂ… ‰ﬁ· «·«’Ê·"
            Index           =   6
         End
         Begin VB.Menu xxxxx 
            Caption         =   "Ã—œ «·«’Ê·"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu xxxxx 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   8
         End
      End
      Begin VB.Menu FinAnalysis 
         Caption         =   "«· Õ·Ì· «·„«·Ì"
         Begin VB.Menu xxy 
            Caption         =   "«·„Ê«“‰… «· ﬁœÌ—Ì…"
            Index           =   0
         End
         Begin VB.Menu xxy 
            Caption         =   "ﬁ«∆„… «· œ›ﬁ «·‰ﬁœÌ"
            Index           =   1
         End
         Begin VB.Menu xxy 
            Caption         =   " »ÊÌ» «·„Ì“«‰Ì… "
            Index           =   2
         End
         Begin VB.Menu xxy 
            Caption         =   " Ê“Ì⁄ «·Õ”«»« "
            Index           =   3
         End
         Begin VB.Menu xxy 
            Caption         =   "«⁄œ«œ „⁄«œ·«  «· Õ·Ì· «·„«·Ì"
            Index           =   4
         End
         Begin VB.Menu xxy 
            Caption         =   "ÿ»«⁄Â ‰ «∆Ã „⁄«œ·«  «· Õ·Ì· «·„«·Ì"
            Index           =   5
         End
         Begin VB.Menu xxy 
            Caption         =   "«·Õ”«»«  «·„Ã„⁄Â"
            Index           =   6
         End
         Begin VB.Menu xxy 
            Caption         =   "«·„’«œﬁ« "
            Index           =   7
         End
         Begin VB.Menu xxy 
            Caption         =   "√Ã‰œ… «·⁄„·«¡"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu xxy 
            Caption         =   "«” œ⁄«¡ „Ì“«‰ „—«Ã⁄Â"
            Index           =   9
         End
         Begin VB.Menu xxy 
            Caption         =   "«·„œ›Ê⁄«  «·„ﬁœ„…"
            Index           =   10
            Begin VB.Menu advancedPayment 
               Caption         =   "«‰Ê«⁄ «·„’—Ê›«  «·„ﬁœ„…"
               Index           =   0
            End
            Begin VB.Menu advancedPayment 
               Caption         =   "«À»«  «·„’—Ê›«  «·„ﬁœ„…"
               Index           =   1
            End
            Begin VB.Menu advancedPayment 
               Caption         =   "«ÿ›«¡ «·„’—Ê›«  «·„ﬁœ„…"
               Index           =   2
            End
            Begin VB.Menu advancedPayment 
               Caption         =   "«À»«  «·»œ·«  «·„ﬁœ„…"
               Index           =   3
            End
         End
         Begin VB.Menu xxy 
            Caption         =   "«·Œÿÿ «·«” —« ÌÃÌ…"
            Index           =   11
         End
      End
      Begin VB.Menu taxes 
         Caption         =   "«·ﬁÌ„Â «·„÷«›…"
         Begin VB.Menu TaxexSub 
            Caption         =   "«·«⁄œ«œ« "
            Index           =   0
         End
         Begin VB.Menu TaxexSub 
            Caption         =   " ”ÃÌ· «·„‘ —Ì«  ÌœÊÌ«"
            Index           =   1
         End
         Begin VB.Menu TaxexSub 
            Caption         =   " ”ÃÌ· «·„»Ì⁄«  ÌœÊÌ«"
            Index           =   2
         End
         Begin VB.Menu TaxexSub 
            Caption         =   " ”ÃÌ· „—œÊœ«  «·„‘ —Ì«  ÌœÊÌ«"
            Index           =   3
         End
         Begin VB.Menu TaxexSub 
            Caption         =   " ”ÃÌ· „—œÊœ«  «·„»Ì⁄«  ÌœÊÌ«"
            Index           =   4
         End
         Begin VB.Menu TaxexSub 
            Caption         =   " ”ÃÌ· „‘ —Ì«  «·„⁄œ«  Ê «·«·« "
            Index           =   5
         End
         Begin VB.Menu TaxexSub 
            Caption         =   " ”ÃÌ· „—œÊœ«  „‘ —Ì«  „⁄œ«  Ê«·« "
            Index           =   6
         End
         Begin VB.Menu TaxexSub 
            Caption         =   "«·«‘⁄«—« "
            Index           =   7
         End
         Begin VB.Menu TaxexSub 
            Caption         =   "«·«ﬁ—«—"
            Index           =   8
         End
         Begin VB.Menu TaxexSub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   9
         End
         Begin VB.Menu TaxexSub 
            Caption         =   "⁄„· ﬁÌÊœ «·ﬁÌ„… «·„÷«›… ·‰ﬁ«ÿ «·»Ì⁄"
            Index           =   10
         End
      End
      Begin VB.Menu mangDep 
         Caption         =   "«·‘∆Ê‰ «·«œ«—Ì…"
         Begin VB.Menu mangDepSub 
            Caption         =   "«· ﬁÌÌ„"
            Index           =   0
         End
         Begin VB.Menu mangDepSub 
            Caption         =   "ÿ·»«  «· ÊŸÌ›"
            Index           =   1
         End
         Begin VB.Menu mangDepSub 
            Caption         =   "«·«Õ Ì«Ã«  «·ÊŸÌ›Ì…"
            Index           =   2
         End
      End
      Begin VB.Menu mnuEmployee 
         Caption         =   "‘∆Ê‰ «·„ÊŸ›Ì‰"
         Begin VB.Menu mnuEmployeeBasic 
            Caption         =   "»Ì«‰«  «”«”Ì…                            "
            Index           =   0
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "≈⁄œ«œ „Ê«⁄Ìœ «·Õ÷Ê— Ê«·«‰’—«› ··‘—ﬂ…"
               Index           =   0
               Visible         =   0   'False
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "«‰Ê«⁄ «·œÊ«„ «Ê «·‘Ì› "
               Index           =   1
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "√‰Ê«⁄ «·√Ã«“« "
               Index           =   2
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "«‰Ê«⁄  ⁄«ﬁœ «·„ÊŸ›Ì‰"
               Index           =   3
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "Õ«·«  «·⁄„·"
               Index           =   4
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "»Ì«‰«  √ﬁ”«„ «·⁄„· ›Ï «·‘—ﬂ…"
               Index           =   5
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "»Ì«‰«  √‰Ê«⁄ «·ÊŸ«∆› ›Ï «·‘—ﬂ…"
               Index           =   6
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "»Ì«‰«   Œ’’«  «·⁄„· ›Ï «·‘—ﬂ…"
               Index           =   7
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "«·œ—Ã«  «·ÊŸÌ›Ì…"
               Index           =   8
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "‘—ﬂ«  «· √„Ì‰"
               Index           =   9
               Visible         =   0   'False
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "«‰Ê«⁄ «· √„Ì‰"
               Index           =   10
               Visible         =   0   'False
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "›∆«  «· √„Ì‰"
               Index           =   11
               Visible         =   0   'False
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "⁄‰«’— «· ﬁÌÌ„"
               Index           =   12
               Visible         =   0   'False
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "√‰Ê«⁄ √–Ê‰«  «·Œ—ÊÃ"
               Index           =   13
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "„Ê«ﬁ⁄ «·⁄„·"
               Index           =   14
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "«·Ã‰”Ì« "
               Index           =   15
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "«·œÌ«‰« "
               Index           =   16
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   " ⁄—Ì› «·„ÊÃÊœ«  «·⁄Ì‰Ì…"
               Index           =   17
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "’·Â «· «»⁄Ì‰"
               Index           =   18
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "«·ﬁÿ«⁄« "
               Index           =   19
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "»Ì«‰«  «· √‘Ì—« "
               Index           =   20
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "«‰Ê«⁄ «·Ã“«¡«  «·«œ«—Ì…"
               Index           =   21
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "«⁄œ«œ«  «·«Ã«“… «·„—÷Ì…"
               Index           =   22
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "”Ì«”… «·«Ã«“« "
               Index           =   23
            End
            Begin VB.Menu mnuEmployeeBasicSub 
               Caption         =   "„ﬂ« » «· ›ÊÌ÷"
               Index           =   24
            End
         End
         Begin VB.Menu mnuEmployeeBasic 
            Caption         =   "»Ì«‰«  «·„ÊŸ›Ì‰"
            Index           =   1
            Begin VB.Menu EmployeeDataicSub 
               Caption         =   "„·› «·„ÊŸ›Ì‰"
               Index           =   0
            End
            Begin VB.Menu EmployeeDataicSub 
               Caption         =   "⁄ﬁÊœ «·„ÊŸ›Ì‰"
               Index           =   1
            End
         End
         Begin VB.Menu mnuEmployeeBasic 
            Caption         =   "«· √„Ì‰«  «·«Ã „«⁄Ì… Ê «·ÿ»Ì…"
            Index           =   2
            Begin VB.Menu mnuEmployeInsuranceSub 
               Caption         =   "≈⁄œ«œ«  «· √„Ì‰«  «·«Ã „«⁄Ì…"
               Index           =   0
            End
            Begin VB.Menu mnuEmployeInsuranceSub 
               Caption         =   "‘—ﬂ«  «· √„Ì‰"
               Index           =   1
            End
            Begin VB.Menu mnuEmployeInsuranceSub 
               Caption         =   "«‰Ê«⁄ «· √„Ì‰"
               Index           =   2
               Visible         =   0   'False
            End
            Begin VB.Menu mnuEmployeInsuranceSub 
               Caption         =   "ﬁ∆«  «· √„Ì‰"
               Index           =   3
            End
            Begin VB.Menu mnuEmployeInsuranceSub 
               Caption         =   "«À»«  «” Õﬁ«ﬁ «· √„Ì‰«  «·«Ã „«⁄Ì…"
               Index           =   4
            End
         End
         Begin VB.Menu mnuEmployeeBasic 
            Caption         =   " ﬁÌÌ„ «·„ÊŸ›Ì‰"
            Index           =   3
            Begin VB.Menu mnuEmployeeBasict 
               Caption         =   "«⁄œ«œ ⁄‰«’— «· ﬁÌÌ„"
               Index           =   0
            End
            Begin VB.Menu mnuEmployeeBasict 
               Caption         =   " ﬁœÌ—«  «· ﬁÌÌ„"
               Index           =   1
            End
            Begin VB.Menu mnuEmployeeBasict 
               Caption         =   "«” Õﬁ«ﬁ «· ﬁÌÌ„"
               Index           =   2
            End
         End
         Begin VB.Menu mnuEmployeeBasic 
            Caption         =   "«·Õ÷Ê— Ê«·«‰’—«›"
            Index           =   4
            Begin VB.Menu EmployeeAttendanceSub 
               Caption         =   "«‰Ê«⁄ «·⁄ÿ·« "
               Index           =   0
            End
            Begin VB.Menu EmployeeAttendanceSub 
               Caption         =   "«⁄œ«œ«  «·‘Ì›« "
               Index           =   1
            End
            Begin VB.Menu EmployeeAttendanceSub 
               Caption         =   "«⁄œ«œ«  «·‰ ÌÃ…"
               Index           =   2
            End
            Begin VB.Menu EmployeeAttendanceSub 
               Caption         =   "   ”ÃÌ· »Ì‰«  «·Õ÷Ê— Ê«·«‰’—«› «·Ì«"
               Index           =   3
            End
            Begin VB.Menu EmployeeAttendanceSub 
               Caption         =   " ”ÃÌ· »Ì‰«  «·Õ÷Ê— Ê«·«‰’—«› ÌœÊÌ«"
               Index           =   4
            End
            Begin VB.Menu EmployeeAttendanceSub 
               Caption         =   "«⁄ „«œ «·Õ÷Ê— Ê«·«‰’—«›"
               Index           =   5
            End
         End
         Begin VB.Menu mnuEmployeeBasic 
            Caption         =   "‰„«–Ã «·≈Ã—«¡« "
            Index           =   5
            Begin VB.Menu HRProcedures 
               Caption         =   "ÿ·» ”·›… ‰ﬁœÌ…"
               Index           =   0
               Visible         =   0   'False
            End
            Begin VB.Menu HRProcedures 
               Caption         =   " ’—ÌÕ Œ—ÊÃ „ƒﬁ "
               Index           =   1
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "ÿ·»  ﬂ·Ì› „Â„… ⁄„·"
               Index           =   2
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "ÿ·» ’—› »œ· ”ﬂ‰ „ﬁœ„"
               Index           =   3
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "ÿ·» ‰ﬁ· „ÊŸ›"
               Index           =   4
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "„»«‘—… „ÊŸ›"
               Index           =   5
               Visible         =   0   'False
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "„–ﬂ—… Œ’„"
               Index           =   6
               Visible         =   0   'False
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "«” »Ì«‰ ⁄‰ „ÊŸ›"
               Index           =   7
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "ÿ·» «Ã«“…"
               Index           =   8
               Visible         =   0   'False
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "»Ì«‰«  «·«Ã«“…"
               Index           =   9
               Visible         =   0   'False
            End
            Begin VB.Menu HRProcedures 
               Caption         =   " ”·Ì„ «·⁄Âœ «·⁄Ì‰Ì…"
               Index           =   10
               Visible         =   0   'False
            End
            Begin VB.Menu HRProcedures 
               Caption         =   " ”·Ì„ ÃÊ«“ ”›— ·„ÊŸ›"
               Index           =   11
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "«‰–«— ·„ÊŸ›"
               Index           =   12
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "Œÿ«» ·„‰ ÌÂ„… «·«„—"
               Index           =   13
            End
            Begin VB.Menu HRProcedures 
               Caption         =   " ﬁ—Ì— «’«»… ⁄„·"
               Index           =   14
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "»Ì«‰ «” ·«„ „⁄«„·« "
               Index           =   15
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "„Œ«·’… ‰Â«∆Ì…"
               Index           =   16
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "‰„Ê–Ã «” ·«„ ”Ì«—…"
               Index           =   19
               Visible         =   0   'False
            End
            Begin VB.Menu HRProcedures 
               Caption         =   " ﬁÌÌ„ «·«œ«¡ Œ·«· › —… «·«Œ »«—"
               Index           =   20
               Visible         =   0   'False
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "«· ﬁÌÌ„ «·”‰ÊÌ ·„œ—«¡ «·«œ«—« "
               Index           =   21
               Visible         =   0   'False
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "«· ﬁÌÌ„ «·”‰ÊÌ ··⁄„«· «·⁄«œÌÌ‰"
               Index           =   22
               Visible         =   0   'False
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "«· ﬁÌÌ„ «·”‰ÊÌ ··›‰ÌÌ‰ Ê„‘€·Ï «·„ﬂ«∆‰"
               Index           =   23
               Visible         =   0   'False
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "«‘⁄«— ⁄‰ Õ«·… „ÊŸ›-«” »Ì«‰"
               Index           =   24
               Visible         =   0   'False
            End
            Begin VB.Menu HRProcedures 
               Caption         =   " ÕœÌÀ »Ì«‰«  «·„ÊŸ›Ì‰"
               Index           =   25
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "‘Â«œ… «Œ·«¡ ÿ—›"
               Index           =   26
            End
            Begin VB.Menu HRProcedures 
               Caption         =   " ⁄ﬁÌ» »‘√‰ «Ã—«¡ «œ«—Ì"
               Index           =   27
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "„–ﬂ—… Œ’„"
               Index           =   28
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "‰„Ê–Ã «” ·«„ ”Ì«—…"
               Index           =   29
               Visible         =   0   'False
            End
            Begin VB.Menu HRProcedures 
               Caption         =   "Œÿ«»  ⁄—Ì›"
               Index           =   30
            End
            Begin VB.Menu HRProcedures 
               Caption         =   " ›ÊÌ÷ ﬁÌ«œ…"
               Index           =   31
            End
         End
         Begin VB.Menu mnuEmployeeBasic 
            Caption         =   "«·—Ê« »"
            Index           =   6
            Begin VB.Menu EmployeeSalarySub 
               Caption         =   "«‰Ê«⁄ „›—œ«  «·—« » «·—∆Ì”Ì…"
               Index           =   0
            End
            Begin VB.Menu EmployeeSalarySub 
               Caption         =   "„›—œ«  «·—« »"
               Index           =   1
            End
            Begin VB.Menu EmployeeSalarySub 
               Caption         =   "«·«÷«›Ì "
               Index           =   2
            End
            Begin VB.Menu EmployeeSalarySub 
               Caption         =   "«·Œ’Ê„« "
               Index           =   3
            End
            Begin VB.Menu EmployeeSalarySub 
               Caption         =   " ”ÃÌ· ”·› «·„ÊŸ›Ì‰"
               Index           =   4
               Visible         =   0   'False
            End
            Begin VB.Menu EmployeeSalarySub 
               Caption         =   "—œ ”·›… „ÊŸ›"
               Index           =   5
               Visible         =   0   'False
            End
            Begin VB.Menu EmployeeSalarySub 
               Caption         =   "—Ê« » «·„ÊŸ›Ì‰"
               Index           =   6
            End
            Begin VB.Menu EmployeeSalarySub 
               Caption         =   "Õ”«» „ﬂ«›√… ‰Â«Ì… «·Œœ„…"
               Index           =   7
               Visible         =   0   'False
            End
            Begin VB.Menu EmployeeSalarySub 
               Caption         =   " ”ÃÌ· „›—œ«  «·—« » «·„ €Ì—…"
               Index           =   8
            End
            Begin VB.Menu EmployeeSalarySub 
               Caption         =   " ”ÃÌ· «·„Œ’’«  ··«Ã«“«  Ê‰Â«Ì… «·Œœ„…"
               Index           =   9
            End
            Begin VB.Menu EmployeeSalarySub 
               Caption         =   " ”ÃÌ·  «” Õﬁ«ﬁ «·„›—œ«  «·”‰ÊÌ… "
               Index           =   10
               Visible         =   0   'False
            End
            Begin VB.Menu EmployeeSalarySub 
               Caption         =   "«·“Ì«Ã« "
               Index           =   11
            End
            Begin VB.Menu EmployeeSalarySub 
               Caption         =   " €ÌÌ— „Ì⁄«œ ”·›…"
               Index           =   12
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuEmployeeBasic 
            Caption         =   "«Ã«“«  «·„ÊŸ›Ì‰"
            Index           =   7
            Begin VB.Menu Vscstionsssub 
               Caption         =   " ”ÃÌ· »Ì«‰«  «·«—’œ… «·«›  «ÕÌ…"
               Index           =   0
            End
            Begin VB.Menu Vscstionsssub 
               Caption         =   " ”ÃÌ· »Ì«‰«  «·«Ã«“«  «·”«»ﬁ…"
               Index           =   1
            End
            Begin VB.Menu Vscstionsssub 
               Caption         =   "Œÿ… «·«Ã«“« "
               Index           =   2
            End
            Begin VB.Menu Vscstionsssub 
               Caption         =   "ÿ·» «Ã«“…"
               Index           =   3
            End
            Begin VB.Menu Vscstionsssub 
               Caption         =   "»Ì«‰«  «·«Ã«“…"
               Index           =   4
               Visible         =   0   'False
            End
            Begin VB.Menu Vscstionsssub 
               Caption         =   " ”·Ì„ Ê≈” ·«„ ⁄Âœ ⁄Ì‰Ì…"
               Index           =   5
            End
            Begin VB.Menu Vscstionsssub 
               Caption         =   "„” Õﬁ«  «·«Ã«“…"
               Index           =   6
            End
            Begin VB.Menu Vscstionsssub 
               Caption         =   " √‘Ì—«  «·Œ—ÊÃ Ê«·⁄Êœ…"
               Index           =   7
            End
            Begin VB.Menu Vscstionsssub 
               Caption         =   "‰”ÃÌ·  «·„»«‘—« "
               Index           =   8
            End
            Begin VB.Menu Vscstionsssub 
               Caption         =   "«·«Ã«“«  «·„—÷Ì…"
               Index           =   9
            End
         End
         Begin VB.Menu mnuEmployeeBasic 
            Caption         =   "”·› «·„ÊŸ›Ì‰"
            Index           =   8
            Begin VB.Menu advanceMenu 
               Caption         =   "ÿ·» ”·›…"
               Index           =   0
            End
            Begin VB.Menu advanceMenu 
               Caption         =   " ”ÃÌ· »Ì«‰«  «·”·›"
               Index           =   1
            End
            Begin VB.Menu advanceMenu 
               Caption         =   " ⁄œÌ· /«Ìﬁ«› / —œ  «·”·›"
               Index           =   2
            End
         End
         Begin VB.Menu mnuEmployeeBasic 
            Caption         =   "«‰Â«¡ «·Œœ„…"
            Index           =   9
            Begin VB.Menu FinishSevicersub 
               Caption         =   " ”ÃÌ·  —ﬂ «·Œœ„…"
               Index           =   0
            End
            Begin VB.Menu FinishSevicersub 
               Caption         =   "Õ”«» „ﬂ«›√… ‰Â«Ì… «·Œœ„…"
               Index           =   1
            End
         End
         Begin VB.Menu mnuEmployeeBasic 
            Caption         =   "«À»«  «·»œ·«  «·„ﬁœ„…"
            Index           =   10
         End
         Begin VB.Menu mnuEmployeeBasic 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   11
         End
      End
      Begin VB.Menu StockControl 
         Caption         =   "„—«ﬁ»… «·„Œ“Ê‰"
         Begin VB.Menu StockControlBasic 
            Caption         =   "„·›«  «”«”Ì…       "
            Begin VB.Menu StockControlBasicSub 
               Caption         =   "»Ì«‰«  «·«’‰«›"
               Index           =   0
            End
            Begin VB.Menu StockControlBasicSub 
               Caption         =   "»Ì«‰«  «·„Œ«“‰"
               Index           =   1
            End
            Begin VB.Menu StockControlBasicSub 
               Caption         =   "„Ã„Ê⁄«  «·«’‰«›"
               Index           =   2
            End
            Begin VB.Menu StockControlBasicSub 
               Caption         =   "»Ì«‰«  «·ÊÕœ« "
               Index           =   3
            End
            Begin VB.Menu StockControlBasicSub 
               Caption         =   "»Ì«‰«  «·«·Ê«‰"
               Index           =   4
            End
            Begin VB.Menu StockControlBasicSub 
               Caption         =   "»Ì«‰«  «·„ﬁ«”« "
               Index           =   5
            End
            Begin VB.Menu StockControlBasicSub 
               Caption         =   "«‰Ê«⁄ ›—“ «·«’‰«›"
               Index           =   6
            End
            Begin VB.Menu StockControlBasicSub 
               Caption         =   "«⁄œ«œ «„«ﬂ‰ «· Œ“Ì‰"
               Index           =   7
            End
            Begin VB.Menu StockControlBasicSub 
               Caption         =   "„Ê«’›«  «·«’‰«›"
               Index           =   8
            End
            Begin VB.Menu StockControlBasicSub 
               Caption         =   "⁄‰«’— «· ﬂ«·Ì› «·’‰«⁄ÌÂ"
               Index           =   9
               Visible         =   0   'False
            End
            Begin VB.Menu StockControlBasicSub 
               Caption         =   "«· ﬂ·›… «· ﬁœÌ—Ì… ÿ»ﬁ« ·„Ã„Ê⁄«  «·«’‰«›"
               Index           =   10
               Visible         =   0   'False
            End
            Begin VB.Menu StockControlBasicSub 
               Caption         =   "Œÿ… „»Ì⁄«  «·«’‰«›"
               Index           =   11
               Visible         =   0   'False
            End
            Begin VB.Menu StockControlBasicSub 
               Caption         =   "—»ÿ «·«’‰«› »«·„Œ«“‰"
               Index           =   12
            End
            Begin VB.Menu StockControlBasicSub 
               Caption         =   "«⁄œ«œ«  Õœ «·ÿ·»"
               Index           =   13
            End
         End
         Begin VB.Menu TradingTransaction 
            Caption         =   "«·—’Ìœ «·«›  «ÕÌ"
            Index           =   0
         End
         Begin VB.Menu TradingTransaction 
            Caption         =   "”‰œ«  œ«Œ·Ì…"
            Index           =   1
            Begin VB.Menu XC 
               Caption         =   "ÿ·»«  œ«Œ·Ì…"
               Index           =   0
            End
            Begin VB.Menu XC 
               Caption         =   "”‰œ«  ÕÃ“ »÷«⁄Â œ«Œ·Ì"
               Index           =   1
            End
         End
         Begin VB.Menu TradingTransaction 
            Caption         =   "”‰œ «” ·«„"
            Index           =   2
         End
         Begin VB.Menu TradingTransaction 
            Caption         =   "”‰œ ’—› "
            Index           =   3
            Begin VB.Menu TradingTransactionSub1 
               Caption         =   "ÿ·» œ«Œ·Ì"
               Index           =   0
            End
            Begin VB.Menu TradingTransactionSub1 
               Caption         =   "”‰œ ’—› »÷«⁄Â"
               Index           =   1
            End
            Begin VB.Menu TradingTransactionSub1 
               Caption         =   "”‰œ ’—› Â«·ﬂ «Ê ⁄Ì‰« "
               Index           =   2
            End
         End
         Begin VB.Menu TradingTransaction 
            Caption         =   " ÕÊÌ· «·»÷«⁄… ≈·Ï „Œ“‰ ¬Œ—"
            Index           =   4
         End
         Begin VB.Menu TradingTransaction 
            Caption         =   "Ã—œ «·„Œ«“‰"
            Index           =   5
            Begin VB.Menu TradingTransactionSub 
               Caption         =   "»œ√ «·Ã—œ"
               Index           =   0
            End
            Begin VB.Menu TradingTransactionSub 
               Caption         =   "ÿ»«⁄Â ﬂ‘Ê› «·Ã—œ"
               Index           =   1
            End
            Begin VB.Menu TradingTransactionSub 
               Caption         =   "«œŒ«· «·ﬂ„Ì«  «·›⁄·ÌÂ"
               Index           =   2
            End
            Begin VB.Menu TradingTransactionSub 
               Caption         =   " ‰›Ì– «·Ã—œ"
               Index           =   3
            End
         End
         Begin VB.Menu TradingTransaction 
            Caption         =   " ”ÊÌ… «·„Œ“Ê‰"
            Index           =   6
         End
         Begin VB.Menu TradingTransaction 
            Caption         =   "”‰œ  Ã„Ì⁄"
            Index           =   7
         End
         Begin VB.Menu TradingTransaction 
            Caption         =   "≈” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰›"
            Index           =   8
            Shortcut        =   ^S
         End
         Begin VB.Menu TradingTransaction 
            Caption         =   "»ÕÀ ⁄‰ »Ì«‰«  ”Ì—Ì«·"
            Index           =   9
         End
         Begin VB.Menu TradingTransaction 
            Caption         =   "«·√’‰«› «·„ÿ·Ê»…"
            Index           =   10
         End
         Begin VB.Menu TradingTransaction 
            Caption         =   "„Êﬁ› «·«’‰«› «·Õ«·Ì"
            Index           =   11
            Visible         =   0   'False
         End
         Begin VB.Menu TradingTransaction 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   12
         End
         Begin VB.Menu TradingTransaction 
            Caption         =   "ÿ·» «—Ã«⁄"
            Index           =   13
            Visible         =   0   'False
         End
      End
      Begin VB.Menu AgeingMAster 
         Caption         =   "«⁄„«— «·œÌÊ‰"
         Begin VB.Menu AgeingSub 
            Caption         =   "«⁄œ«œ«  «⁄„«— «·œÌÊ‰ ··„‘ —Ì« "
            Index           =   0
         End
         Begin VB.Menu AgeingSub 
            Caption         =   "«⁄œ«œ«  «⁄„«— «·œÌÊ‰ ··„»Ì⁄« "
            Index           =   1
         End
         Begin VB.Menu AgeingSub 
            Caption         =   " ”ÃÌ· ›Ê« Ì— «·„‘ —Ì«  «·”«»ﬁ…"
            Index           =   2
         End
         Begin VB.Menu AgeingSub 
            Caption         =   " ”ÃÌ· ›Ê« Ì— «·„»Ì⁄«  «·”«»ﬁ…"
            Index           =   3
         End
         Begin VB.Menu AgeingSub 
            Caption         =   " ”ÃÌ· ›Ê« Ì— «·„»Ì⁄«  «·Õ«·Ì…"
            Index           =   4
         End
         Begin VB.Menu AgeingSub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   5
         End
      End
      Begin VB.Menu Purchase 
         Caption         =   "«·„‘ —Ì« "
         Begin VB.Menu PurchaseBasicRoot 
            Caption         =   "„·›«  «”«”ÌÂ"
            Begin VB.Menu PurchaseBasic 
               Caption         =   "»Ì«‰«  «·„Ê—Ì‰"
               Index           =   0
            End
            Begin VB.Menu PurchaseBasic 
               Caption         =   "« ›«ﬁÌ«  «·„Ê—Ì‰"
               Index           =   1
            End
            Begin VB.Menu PurchaseBasic 
               Caption         =   "«⁄œ«œ «⁄„«— «·œÌÊ‰ ··„Ê—œÌ‰"
               Index           =   2
            End
            Begin VB.Menu PurchaseBasic 
               Caption         =   "«‰Ê«⁄ «·‘Õ‰"
               Index           =   3
            End
            Begin VB.Menu PurchaseBasic 
               Caption         =   "«‰Ê«⁄ «·÷„«‰« "
               Index           =   4
            End
            Begin VB.Menu PurchaseBasic 
               Caption         =   "ÿ—ﬁ «·œ›⁄"
               Index           =   5
            End
            Begin VB.Menu PurchaseBasic 
               Caption         =   "„Ã„Ê⁄«  «·„‰«œÌ»"
               Index           =   6
            End
            Begin VB.Menu PurchaseBasic 
               Caption         =   "»Ì«‰«  «·„‰«œÌ»"
               Index           =   7
            End
            Begin VB.Menu PurchaseBasic 
               Caption         =   "ÿ—ﬁ «·‘Õ‰"
               Index           =   8
            End
         End
         Begin VB.Menu PurchaseTransactions 
            Caption         =   "⁄—Ê÷ «·«”⁄«— Êÿ·»«  «·‘—«¡"
            Index           =   0
            Begin VB.Menu PurchaseTransactionssubd 
               Caption         =   "⁄—Ê÷ «·«”⁄«—"
               Index           =   0
               Begin VB.Menu PurchaseTransactionssubs 
                  Caption         =   "ÿ·» ⁄—Ê÷ «·«”⁄«—"
                  Index           =   0
               End
               Begin VB.Menu PurchaseTransactionssubs 
                  Caption         =   "⁄—Ê÷ «·«”⁄«—"
                  Index           =   1
               End
               Begin VB.Menu PurchaseTransactionssubs 
                  Caption         =   "„ﬁ«—‰Â ⁄—Ê÷ «·«”⁄«— "
                  Index           =   2
               End
            End
            Begin VB.Menu PurchaseTransactionssubd 
               Caption         =   "ÿ·»«  «·‘—«¡"
               Index           =   1
               Begin VB.Menu PurchaseTransactionssubs1 
                  Caption         =   "ÿ·»  ‘—«¡"
                  Index           =   0
               End
               Begin VB.Menu PurchaseTransactionssubs1 
                  Caption         =   "«⁄ „«œ «„— ‘—«¡"
                  Index           =   1
                  Visible         =   0   'False
               End
               Begin VB.Menu PurchaseTransactionssubs1 
                  Caption         =   "«„— ‘—«¡"
                  Index           =   2
               End
            End
         End
         Begin VB.Menu PurchaseTransactions 
            Caption         =   "»Ì«‰«  «·‘Õ‰"
            Index           =   1
         End
         Begin VB.Menu PurchaseTransactions 
            Caption         =   "«·«⁄ „«œ«  «·„” ‰œÌÂ"
            Index           =   2
            Begin VB.Menu LCTransactions 
               Caption         =   "«‰Ê«⁄ «·«⁄ „«œ« "
               Index           =   0
            End
            Begin VB.Menu LCTransactions 
               Caption         =   "›« Ê—… „»œ∆ÌÂ"
               Index           =   1
            End
            Begin VB.Menu LCTransactions 
               Caption         =   "› Õ «⁄ „«œ "
               Index           =   2
            End
            Begin VB.Menu LCTransactions 
               Caption         =   " ⁄œÌ· «⁄ „«œ"
               Index           =   3
            End
            Begin VB.Menu LCTransactions 
               Caption         =   "„ «»⁄Â «·‘Õ‰« "
               Index           =   4
            End
            Begin VB.Menu LCTransactions 
               Caption         =   "”‰œ«  «” ·«„ «·‘Õ‰« "
               Index           =   5
            End
            Begin VB.Menu LCTransactions 
               Caption         =   "«·›« Ê—… «·‰Â«∆ÌÂ"
               Index           =   6
            End
            Begin VB.Menu LCTransactions 
               Caption         =   "€·ﬁ «·«⁄ „«œ"
               Index           =   7
            End
            Begin VB.Menu LCTransactions 
               Caption         =   "ÿ·» ÷„«‰ »‰ﬂÌ"
               Index           =   8
            End
            Begin VB.Menu LCTransactions 
               Caption         =   "ÿ·»   „œÌœ ÷„«‰ »‰ﬂÌ"
               Index           =   9
            End
            Begin VB.Menu LCTransactions 
               Caption         =   " ÷„«‰ »‰ﬂÌ ‰Â«∆Ì"
               Index           =   10
            End
            Begin VB.Menu LCTransactions 
               Caption         =   "‘—«¡ «·„‰«›”Â"
               Index           =   11
            End
         End
         Begin VB.Menu PurchaseTransactions 
            Caption         =   "›« Ê—… „‘ —Ì« "
            Index           =   3
         End
         Begin VB.Menu PurchaseTransactions 
            Caption         =   "›« Ê—… „‘ —Ì«  „Ã„⁄Â"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu PurchaseTransactions 
            Caption         =   "„—œÊœ«  «·„‘ —Ì« "
            Index           =   5
         End
         Begin VB.Menu PurchaseTransactions 
            Caption         =   "  ﬁ«—Ì— «⁄„«— œÌÊ‰ «·„Ê—œÌ‰"
            Index           =   6
         End
         Begin VB.Menu PurchaseTransactions 
            Caption         =   " ﬁ«—Ì— «·„‘ —Ì«  Ê «·„Ê—œÌ‰"
            Index           =   7
         End
      End
      Begin VB.Menu MarketingMnu 
         Caption         =   "«· ”ÊÌﬁ"
         Begin VB.Menu MarketingMnusub 
            Caption         =   "«·«⁄œ«œ«  «·⁄«„…"
            Index           =   0
            Begin VB.Menu mnuSalesBasic 
               Caption         =   "«⁄œ«œ«  «Êﬁ«  «·“Ì«—« "
               Index           =   0
            End
            Begin VB.Menu mnuSalesBasic 
               Caption         =   "«‰Ê«⁄ «·“Ì«—« "
               Index           =   1
            End
            Begin VB.Menu mnuSalesBasic 
               Caption         =   " ﬁÌÌ„ «·⁄„·«¡"
               Index           =   2
            End
            Begin VB.Menu mnuSalesBasic 
               Caption         =   " ⁄—Ì› „ ÿ·»«  «·“Ì«—« "
               Index           =   3
            End
         End
         Begin VB.Menu MarketingMnusub 
            Caption         =   "⁄—Ê÷ «·«’‰«›"
            Index           =   1
         End
         Begin VB.Menu MarketingMnusub 
            Caption         =   "„ «»⁄Â «·⁄„·«¡"
            Index           =   2
            Begin VB.Menu MarketingMnusubsub 
               Caption         =   " ”ÃÌ· „Ê«⁄Ìœ «·⁄„·«¡"
               Index           =   0
            End
            Begin VB.Menu MarketingMnusubsub 
               Caption         =   " ”ÃÌ· “Ì«—«  «·⁄„·«¡"
               Index           =   1
            End
            Begin VB.Menu MarketingMnusubsub 
               Caption         =   "„ «»⁄Â “Ì«—«  «·⁄„·«¡"
               Index           =   2
            End
            Begin VB.Menu MarketingMnusubsub 
               Caption         =   "«” ÿ·«⁄ —√Ì «·⁄„·«¡"
               Index           =   3
            End
            Begin VB.Menu MarketingMnusubsub 
               Caption         =   " ”ÃÌ· ‘ﬂÊÏ «·⁄„·«¡"
               Index           =   4
            End
            Begin VB.Menu MarketingMnusubsub 
               Caption         =   "„ «»⁄Â ‘ﬂÊÏ «·⁄„·«¡"
               Index           =   5
            End
            Begin VB.Menu MarketingMnusubsub 
               Caption         =   "œ·Ì· «·Â« ›"
               Index           =   6
            End
            Begin VB.Menu MarketingMnusubsub 
               Caption         =   "‘«‘… „ «»⁄Â «·„‰«œÌ»"
               Index           =   7
            End
            Begin VB.Menu MarketingMnusubsub 
               Caption         =   "‘«‘… «·« ’«·« "
               Index           =   8
            End
         End
         Begin VB.Menu MarketingMnusub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   3
         End
         Begin VB.Menu MarketingMnusub 
            Caption         =   " ﬁ«—Ì— «·« ’«·« "
            Index           =   4
         End
      End
      Begin VB.Menu Sales 
         Caption         =   "«·„»Ì⁄« "
         Begin VB.Menu SalesBasic 
            Caption         =   "«·»Ì«‰«  «·«”«”ÌÂ"
            Begin VB.Menu SalesBasicSub 
               Caption         =   "«‰Ê«⁄ «·⁄„·«¡  Ê «·„Ê—œÌ‰"
               Index           =   0
            End
            Begin VB.Menu SalesBasicSub 
               Caption         =   "»Ì«‰«  «·⁄„·«¡"
               Index           =   1
               Begin VB.Menu SalesBasicSubsub 
                  Caption         =   "„Ã„Ê⁄«  «·⁄„·«¡"
                  Index           =   0
               End
               Begin VB.Menu SalesBasicSubsub 
                  Caption         =   " ’‰Ì›«  «·⁄„·«¡"
                  Index           =   1
               End
               Begin VB.Menu SalesBasicSubsub 
                  Caption         =   "ÿ·» › Õ Õ”«» ⁄„Ì·"
                  Index           =   2
               End
               Begin VB.Menu SalesBasicSubsub 
                  Caption         =   "„·› «·⁄„·«¡"
                  Index           =   3
               End
               Begin VB.Menu SalesBasicSubsub 
                  Caption         =   "«·⁄„·«¡ «·‰ﬁœÌÌ‰"
                  Index           =   4
               End
            End
            Begin VB.Menu SalesBasicSub 
               Caption         =   "« ›«ﬁÌ«  «·⁄„·«¡"
               Index           =   2
            End
            Begin VB.Menu SalesBasicSub 
               Caption         =   "«⁄œ«œ «⁄„«— «·œÌÊ‰ ··⁄„·«¡"
               Index           =   3
            End
            Begin VB.Menu SalesBasicSub 
               Caption         =   " ⁄—Ì› «”⁄«— «·»Ì⁄"
               Index           =   4
            End
            Begin VB.Menu SalesBasicSub 
               Caption         =   "«⁄œ·œ«  «·«’‰«› «·—ﬂœ…"
               Index           =   5
            End
            Begin VB.Menu SalesBasicSub 
               Caption         =   "«⁄œ«œ  ‰”» «Âœ› «·„»Ì⁄«  Ê «· Õ’Ì·« "
               Index           =   6
            End
            Begin VB.Menu SalesBasicSub 
               Caption         =   "„Ã„Ê⁄«  «·„‰«œÌ»"
               Index           =   7
            End
            Begin VB.Menu SalesBasicSub 
               Caption         =   "„·› «·„‰œÊ»"
               Index           =   8
            End
            Begin VB.Menu SalesBasicSub 
               Caption         =   "«‰Ê«⁄ ÷„«‰«  «· ﬁ”Ìÿ"
               Index           =   9
            End
            Begin VB.Menu SalesBasicSub 
               Caption         =   "«‰Ê«⁄ «·„—œÊœ« "
               Index           =   10
            End
            Begin VB.Menu SalesBasicSub 
               Caption         =   "«‰Ê«⁄ «·÷„«‰« "
               Index           =   11
            End
         End
         Begin VB.Menu SalesTransactions 
            Caption         =   "⁄—Ê÷ «·√”⁄«— Ê √Ê«„— «·»Ì⁄"
            Index           =   0
            Begin VB.Menu SalesTransactionssubss0 
               Caption         =   "⁄—Ê÷ «·«”⁄«—"
               Index           =   0
               Begin VB.Menu SalesTransactionssubss00 
                  Caption         =   "ÿ·» ⁄—Ê÷ «”⁄«— „»œ∆Ì…  „‰ «·⁄„·«¡"
                  Index           =   0
               End
               Begin VB.Menu SalesTransactionssubss00 
                  Caption         =   "⁄—Ê÷ «”⁄«— ‰Â«∆Ì… "
                  Index           =   1
               End
               Begin VB.Menu SalesTransactionssubss00 
                  Caption         =   "⁄—÷ ”⁄— „ Œ’’"
                  Index           =   2
                  Visible         =   0   'False
               End
            End
            Begin VB.Menu SalesTransactionssubss0 
               Caption         =   "√Ê«„— «·»Ì⁄"
               Index           =   1
               Begin VB.Menu SalesTransactionssubss000 
                  Caption         =   "ÿ·» «„— »Ì⁄"
                  Index           =   0
               End
               Begin VB.Menu SalesTransactionssubss000 
                  Caption         =   "√„— »Ì⁄"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu SalesTransactions 
            Caption         =   "«Ê«„— «·»Ì⁄"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu SalesTransactions 
            Caption         =   "›« Ê—…  „»Ì⁄« "
            Index           =   2
         End
         Begin VB.Menu SalesTransactions 
            Caption         =   "„—œÊœ«  «·„»Ì⁄« "
            Index           =   3
         End
         Begin VB.Menu SalesTransactions 
            Caption         =   "›« Ê—… „Ã„⁄Â"
            Index           =   4
         End
         Begin VB.Menu SalesTransactions 
            Caption         =   "⁄—Ê÷ «·«’‰«›"
            Index           =   5
         End
         Begin VB.Menu SalesTransactions 
            Caption         =   "ŒÿÂ  ”⁄Ì— «·«’‰«›"
            Index           =   6
         End
         Begin VB.Menu SalesTransactions 
            Caption         =   "ﬁ«∆„… «·«”⁄«—"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu SalesTransactions 
            Caption         =   "„ «»⁄Â «·„‰«œÌ»"
            Index           =   8
            Begin VB.Menu SalesTransactionsEmp 
               Caption         =   "«⁄œ«œ ⁄„Ê·«  «·„»Ì⁄«   Ê «· Õ’Ì·« "
               Index           =   0
            End
            Begin VB.Menu SalesTransactionsEmp 
               Caption         =   "ŒÿÂ «·„»Ì⁄«  Ê «· Õ’Ì·« "
               Index           =   1
            End
            Begin VB.Menu SalesTransactionsEmp 
               Caption         =   "‰”»  Õﬁﬁ ŒÿÂ «·„»Ì⁄«  Ê «· Õ’Ì·« "
               Index           =   2
            End
            Begin VB.Menu SalesTransactionsEmp 
               Caption         =   "«·⁄„Ê·«  «·„” Õﬁ… ··„‰«œÌ»"
               Index           =   3
            End
            Begin VB.Menu SalesTransactionsEmp 
               Caption         =   "”Ì«”Â/⁄—Ê÷  ⁄ÃÌ· «·œ›⁄"
               Index           =   4
            End
         End
         Begin VB.Menu SalesTransactions 
            Caption         =   " ﬁ—Ì— «⁄„«— œÌÊ‰ «·⁄„·«¡"
            Index           =   9
         End
         Begin VB.Menu SalesTransactions 
            Caption         =   " ﬁ«—Ì— «·„»Ì⁄«  Ê«·⁄„·«¡"
            Index           =   10
         End
         Begin VB.Menu SalesTransactions 
            Caption         =   " ﬁ—Ì— «·⁄„·«¡ «·‰ﬁœÌÌ‰"
            Index           =   11
         End
      End
      Begin VB.Menu Container 
         Caption         =   "«·Õ«ÊÌ« "
         Begin VB.Menu ContainerSub 
            Caption         =   "„Ã„Ê⁄«  «·Õ«ÊÌ« "
            Index           =   0
         End
         Begin VB.Menu ContainerSub 
            Caption         =   "»Ì«‰«  «·Õ«ÊÌ« "
            Index           =   1
         End
         Begin VB.Menu ContainerSub 
            Caption         =   " ⁄—Ì› «·„‰«ÿﬁ"
            Index           =   2
         End
         Begin VB.Menu ContainerSub 
            Caption         =   "«‰Ê«⁄ «·‘«Õ‰« "
            Index           =   3
         End
         Begin VB.Menu ContainerSub 
            Caption         =   "ÿ—«“ «·‘«Õ‰« "
            Index           =   4
         End
         Begin VB.Menu ContainerSub 
            Caption         =   " ⁄—Ì› «·‘«Õ‰« "
            Index           =   5
         End
         Begin VB.Menu ContainerSub 
            Caption         =   " ⁄—Ì› «·”«∆ﬁÌ‰"
            Index           =   6
         End
         Begin VB.Menu ContainerSub 
            Caption         =   "«‰Ê«⁄ «·⁄„·«¡ "
            Index           =   7
         End
         Begin VB.Menu ContainerSub 
            Caption         =   "«·⁄„·«¡ "
            Index           =   8
         End
         Begin VB.Menu ContainerSub 
            Caption         =   "«·⁄ﬁÊœ"
            Index           =   9
         End
         Begin VB.Menu ContainerSub 
            Caption         =   "«·Õ—ﬂ« "
            Index           =   10
         End
         Begin VB.Menu ContainerSub 
            Caption         =   "«· ›—Ì€« "
            Index           =   11
         End
         Begin VB.Menu ContainerSub 
            Caption         =   "„Êﬁ› «·Õ«ÊÌ« "
            Index           =   12
         End
         Begin VB.Menu ContainerSub 
            Caption         =   "«· »ÌÂ« "
            Index           =   13
         End
         Begin VB.Menu ContainerSub 
            Caption         =   "”‰œ«  «·ﬁ»÷"
            Index           =   14
         End
         Begin VB.Menu ContainerSub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   15
         End
      End
      Begin VB.Menu COLLECTIONS 
         Caption         =   "«· Õ’Ì·« "
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   "„Ã„Ê⁄Â «·„‰«œÌ»"
            Index           =   0
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   "»Ì«‰«  «·„‰«œÌ»"
            Index           =   1
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   "«‰Ê«⁄ «·⁄„·«¡"
            Index           =   2
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   "œ·Ì· «·Â« ›"
            Index           =   3
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   " ”ÃÌ· «·« ’«·« "
            Index           =   4
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   "ÿ·» ﬁ Õ Õ”«» ⁄„Ì·"
            Index           =   5
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   "»Ì«‰«  «·⁄„·«¡"
            Index           =   6
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   " ”ÃÌ· „Ê«⁄Ìœ «·“Ì«—« "
            Index           =   7
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   "„ «»⁄Â «·„‰«œÌ»"
            Index           =   8
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   "„Êﬁ› “Ì«—… «·⁄„·«¡"
            Index           =   9
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   "‘«‘Â «· Õ’Ì·« "
            Index           =   10
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   "‘ﬂ«ÊÌ «·⁄„·«¡"
            Index           =   11
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   " ﬁ«—Ì— «⁄„«— «·œÌÊ‰"
            Index           =   12
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   " ﬁ«—Ì— «·„ﬁ»Ê÷« "
            Index           =   13
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   " ﬁ«—Ì— «· Õ’Ì·« "
            Index           =   14
         End
         Begin VB.Menu COLLECTIONSUB 
            Caption         =   "«· ‰»ÌÂ« "
            Index           =   15
         End
      End
      Begin VB.Menu SalesIns 
         Caption         =   "„»Ì⁄«  «· ﬁ”Ìÿ"
         Begin VB.Menu SalesInsSub 
            Caption         =   "ÿ·» ‘—«¡ »«· ﬁ”Ìÿ"
            Index           =   0
         End
         Begin VB.Menu SalesInsSub 
            Caption         =   "ÿ·» › Õ Õ”«» ⁄„Ì·"
            Index           =   1
         End
         Begin VB.Menu SalesInsSub 
            Caption         =   "«·⁄„·«¡"
            Index           =   2
         End
         Begin VB.Menu SalesInsSub 
            Caption         =   "›« Ê—… „»Ì⁄«  «· ﬁ”Ìÿ"
            Index           =   3
         End
         Begin VB.Menu SalesInsSub 
            Caption         =   " Õ’Ì· «·«ﬁ”«ÿ"
            Index           =   4
         End
         Begin VB.Menu SalesInsSub 
            Caption         =   "«· ‰»ÌÂ« "
            Index           =   5
         End
         Begin VB.Menu SalesInsSub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   6
         End
      End
      Begin VB.Menu POSTRansactiosG 
         Caption         =   "‰ﬁ«ÿ «·»Ì⁄"
         Begin VB.Menu POSTRansactios 
            Caption         =   "»Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄"
            Index           =   0
         End
         Begin VB.Menu POSTRansactios 
            Caption         =   "»Ì«‰«  ﬂ«‘Ì—"
            Index           =   1
         End
         Begin VB.Menu POSTRansactios 
            Caption         =   "»Ì«‰«  «·‘Ì› "
            Index           =   2
         End
         Begin VB.Menu POSTRansactios 
            Caption         =   "»Ì«‰«  «·„Ê«ﬁ⁄"
            Index           =   3
         End
         Begin VB.Menu POSTRansactios 
            Caption         =   "«⁄œ«œ«  ‰ﬁ«ÿ «·⁄„·«¡"
            Index           =   4
         End
         Begin VB.Menu POSTRansactios 
            Caption         =   " ”ÃÌ· «·œŒÊ·"
            Index           =   5
         End
         Begin VB.Menu POSTRansactios 
            Caption         =   "’—› „ﬂÊ‰«  «·«’‰«›"
            Index           =   6
         End
         Begin VB.Menu POSTRansactios 
            Caption         =   " ﬁ»÷ ⁄«„  ‰ﬁ«ÿ «·»Ì⁄"
            Index           =   7
         End
         Begin VB.Menu POSTRansactios 
            Caption         =   " ﬁ«—Ì— ‰ﬁ«ÿ «·»Ì⁄"
            Index           =   8
         End
         Begin VB.Menu POSTRansactios 
            Caption         =   "ÿ»«⁄Â ﬂ—Ê  «·⁄„·«¡"
            Index           =   9
         End
         Begin VB.Menu POSTRansactios 
            Caption         =   "«·ﬁ”«∆„ «·„Ã«‰Ì…"
            Index           =   10
         End
         Begin VB.Menu POSTRansactios 
            Caption         =   "„Õÿ«  «·ÊﬁÊœ"
            Index           =   11
         End
      End
      Begin VB.Menu shipmentMnu 
         Caption         =   "«·‘Õ‰ Ê «· Ê“Ì⁄"
         Begin VB.Menu ShpmentBasicdata 
            Caption         =   "«·»Ì«‰«  «·”«”Ì…"
            Index           =   0
            Begin VB.Menu ShpmentBasicdatasub 
               Caption         =   "»Ì«‰«  «·œÊ·"
               Index           =   0
            End
            Begin VB.Menu ShpmentBasicdatasub 
               Caption         =   "»Ì«‰«  «·„‰«ÿﬁ «·„Õ«›Ÿ« "
               Index           =   1
            End
            Begin VB.Menu ShpmentBasicdatasub 
               Caption         =   "«·„”«›«  »Ì‰ «·„œ‰"
               Index           =   2
            End
            Begin VB.Menu ShpmentBasicdatasub 
               Caption         =   "»Ì«‰«  «·√ÕÌ«¡"
               Index           =   3
            End
            Begin VB.Menu ShpmentBasicdatasub 
               Caption         =   "»Ì«‰«  «·‘Ê«—⁄"
               Index           =   4
            End
            Begin VB.Menu ShpmentBasicdatasub 
               Caption         =   "«‰Ê«⁄ «·„—ﬂ»« "
               Index           =   5
            End
            Begin VB.Menu ShpmentBasicdatasub 
               Caption         =   "»Ì«‰«  «·„—ﬂ»« "
               Index           =   6
            End
            Begin VB.Menu ShpmentBasicdatasub 
               Caption         =   "»Ì«‰«  «·”«∆ﬁÌ‰"
               Index           =   7
            End
            Begin VB.Menu ShpmentBasicdatasub 
               Caption         =   "«‰Ê«⁄ «·‘Õ‰"
               Index           =   8
            End
            Begin VB.Menu ShpmentBasicdatasub 
               Caption         =   "«‰Ê«⁄ «·’Ì«‰…"
               Index           =   9
            End
         End
         Begin VB.Menu ShpmentBasicdata 
            Caption         =   "«·»÷«∆⁄ ﬁÌœ «· ”·Ì„"
            Index           =   1
         End
         Begin VB.Menu ShpmentBasicdata 
            Caption         =   "  Œ’Ì’ «·‘«Õ‰« "
            Index           =   2
         End
         Begin VB.Menu ShpmentBasicdata 
            Caption         =   " ”ÃÌ·  ÊﬁÌ «   «· ”·Ì„"
            Index           =   3
         End
         Begin VB.Menu ShpmentBasicdata 
            Caption         =   "Œÿ… «·‘Õ‰"
            Index           =   4
         End
         Begin VB.Menu ShpmentBasicdata 
            Caption         =   "ÿ·» ‘Õ‰"
            Index           =   5
         End
         Begin VB.Menu ShpmentBasicdata 
            Caption         =   " ”ÃÌ· «·‘Õ‰"
            Index           =   6
         End
         Begin VB.Menu ShpmentBasicdata 
            Caption         =   "«” ·«„ «·‘Õ‰…"
            Index           =   7
         End
         Begin VB.Menu ShpmentBasicdata 
            Caption         =   " ﬁ«—Ì— «·‘Õ‰"
            Index           =   8
         End
      End
      Begin VB.Menu prdo 
         Caption         =   "«·«‰ «Ã Ê√Ê«„— «·‘€·"
         Begin VB.Menu prdo1 
            Caption         =   "«·»Ì«‰«  «·«”«”Ì…"
            Index           =   0
            Begin VB.Menu prdo1sub 
               Caption         =   "»Ì«‰«  «·„⁄œ«  / «·„«ﬂÌ‰« "
               Index           =   0
            End
            Begin VB.Menu prdo1sub 
               Caption         =   "⁄‰«’— «· ﬂ«·Ì› «·’‰«⁄Ì…"
               Index           =   1
            End
            Begin VB.Menu prdo1sub 
               Caption         =   "«· ﬂ·›… «· ﬁœÌ—Ì… ÿ»ﬁ« ·„Ã„Ê⁄Â «·«’‰«›"
               Index           =   2
            End
            Begin VB.Menu prdo1sub 
               Caption         =   "»Ì«‰«  «·ﬁÊ«·»"
               Index           =   3
            End
            Begin VB.Menu prdo1sub 
               Caption         =   "«‰Ê«⁄ «·«‰ «Ã"
               Index           =   4
            End
            Begin VB.Menu prdo1sub 
               Caption         =   "«· ﬂ«·Ì› «· ﬁœÌ—Ì… ÿ»ﬁ« ··«’‰«›"
               Index           =   5
            End
         End
         Begin VB.Menu prdo1 
            Caption         =   "ŒÿÊÿ «·«‰ «Ã"
            Index           =   4
            Begin VB.Menu prosub1 
               Caption         =   " ⁄—Ì› ŒÿÊÿ «·«‰ «Ã"
               Index           =   0
            End
            Begin VB.Menu prosub1 
               Caption         =   " Œ’Ì’  Ê‰ﬁ· «·⁄„«·"
               Index           =   1
            End
         End
         Begin VB.Menu prdo1 
            Caption         =   "„—«Õ· «·«‰ «Ã"
            Index           =   5
            Begin VB.Menu PrbH 
               Caption         =   "”‰œ ’—› „—«Õ· «‰ «Ã"
               Index           =   0
            End
            Begin VB.Menu PrbH 
               Caption         =   "«„— «‰ «Ã ‰’› „’‰⁄"
               Index           =   1
            End
            Begin VB.Menu PrbH 
               Caption         =   "”‰œ «” ·«„ «‰ «Ã ‰’› „’‰⁄"
               Index           =   2
            End
         End
         Begin VB.Menu prdo1 
            Caption         =   "”‰œ ÕÃ“ «‰ «Ã"
            Index           =   6
         End
         Begin VB.Menu prdo1 
            Caption         =   "«„— «·«‰ «Ã/«·‘€·"
            Index           =   7
         End
         Begin VB.Menu prdo1 
            Caption         =   "”‰œ ’—› „Ê«œ Œ«„"
            Index           =   8
         End
         Begin VB.Menu prdo1 
            Caption         =   "”‰œ «” ·«„ «‰ «Ã  «„"
            Index           =   9
         End
         Begin VB.Menu prdo1 
            Caption         =   "Õ”«»  ﬂ«·Ì› «·«‰ «Ã «·‰„ÿÌ"
            Index           =   10
         End
         Begin VB.Menu prdo1 
            Caption         =   " Ê“Ì⁄ «· ﬂ«·Ì› €Ì— «·„Ì«‘—…"
            Index           =   11
            Visible         =   0   'False
         End
         Begin VB.Menu prdo1 
            Caption         =   " Œ’Ì’ ŒÿÊÿ «·«‰ «Ã ·√Ê«„— «·‘€·"
            Index           =   12
         End
         Begin VB.Menu prdo1 
            Caption         =   "«÷«›… √„ «— „‘€·Ì‰ Ê—œÊœ"
            Index           =   13
         End
         Begin VB.Menu prdo1 
            Caption         =   "”‰œ «· Ã„Ì⁄"
            Index           =   14
         End
         Begin VB.Menu prdo1 
            Caption         =   " ﬁ«—Ì— «·«‰ «Ã"
            Index           =   15
         End
      End
      Begin VB.Menu ProductionPlan 
         Caption         =   " «· ŒÿÌÿ Ê„—«ﬁ»Â «·ÃÊœ…"
         Begin VB.Menu ProductionPlansub 
            Caption         =   "ŒÿÂ «·«‰ «Ã"
            Index           =   0
         End
         Begin VB.Menu ProductionPlansub 
            Caption         =   " ⁄—Ì› ⁄‰«’— „—«ﬁ»Â «·ÃÊœ…"
            Index           =   1
         End
         Begin VB.Menu ProductionPlansub 
            Caption         =   "  ’‰Ì› «·„‰ Ã« "
            Index           =   2
         End
         Begin VB.Menu ProductionPlansub 
            Caption         =   " ⁄—Ì› «·«Ã—«¡«  «· ’ÕÌÕÌÂ"
            Index           =   3
         End
         Begin VB.Menu ProductionPlansub 
            Caption         =   "›Õ’ ÃÊœ… «·„‰ Ã «· «„"
            Index           =   4
         End
         Begin VB.Menu ProductionPlansub 
            Caption         =   "„ «»⁄Â Ê ”ÃÌ· «’·«Õ «·„‰ Ã«  «·„⁄Ì»Â"
            Index           =   5
         End
         Begin VB.Menu ProductionPlansub 
            Caption         =   "›—“ «·ÃÊœ…"
            Index           =   6
         End
         Begin VB.Menu ProductionPlansub 
            Caption         =   "„·«ÕŸ… «·„⁄œ« "
            Index           =   7
         End
      End
      Begin VB.Menu MnuElevators 
         Caption         =   "«œ«—… «·„’«⁄œ"
         Begin VB.Menu MnuElevatorssUB 
            Caption         =   " ⁄—Ì› „Õœœ«  «·⁄—Ê÷"
            Index           =   0
         End
         Begin VB.Menu MnuElevatorssUB 
            Caption         =   "—Ìÿ „Õœœ«  «·⁄—Ê÷"
            Index           =   1
         End
         Begin VB.Menu MnuElevatorssUB 
            Caption         =   "⁄—Ê÷ «·«”⁄«— «·„ Œ’’…"
            Index           =   2
         End
         Begin VB.Menu MnuElevatorssUB 
            Caption         =   "«·⁄—Ê÷ «·›‰Ì…"
            Index           =   3
         End
         Begin VB.Menu MnuElevatorssUB 
            Caption         =   " «·’Ì«‰… Ê «·÷„«‰"
            Index           =   4
            Begin VB.Menu Elevatorsmaintenance 
               Caption         =   "«·÷„«‰"
               Index           =   0
            End
            Begin VB.Menu Elevatorsmaintenance 
               Caption         =   "’—› ﬁÿ⁄ «·€Ì«—"
               Index           =   1
            End
            Begin VB.Menu Elevatorsmaintenance 
               Caption         =   " ‰»ÌÂ«  «·’Ì«‰Â «·œÊ—Ì…"
               Index           =   2
            End
            Begin VB.Menu Elevatorsmaintenance 
               Caption         =   " ‰»ÌÂ«  ⁄ﬁÊœ «·’Ì«‰…"
               Index           =   3
            End
            Begin VB.Menu Elevatorsmaintenance 
               Caption         =   " ‰»ÌÂ«  «·÷„«‰«  "
               Index           =   4
            End
            Begin VB.Menu Elevatorsmaintenance 
               Caption         =   " ﬁ«—Ì— «·’Ì«‰…"
               Index           =   5
            End
         End
         Begin VB.Menu MnuElevatorssUB 
            Caption         =   " ‰»ÌÂ«  «·’Ì«‰… «·œÊ—Ì…"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu MnuElevatorssUB 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   6
         End
      End
      Begin VB.Menu CeramicEstimation 
         Caption         =   "«·„ﬁ«Ì”« "
         Begin VB.Menu CeramicEstimationsub 
            Caption         =   "ÊÕœ«  «·⁄„·Ì« "
            Index           =   0
         End
         Begin VB.Menu CeramicEstimationsub 
            Caption         =   " ⁄—Ì› «·⁄„·Ì« "
            Index           =   1
         End
         Begin VB.Menu CeramicEstimationsub 
            Caption         =   "ÿ·» —›⁄ „ﬁ«”« "
            Index           =   2
         End
         Begin VB.Menu CeramicEstimationsub 
            Caption         =   " Ê“Ì⁄ «·ÿ·»« "
            Index           =   3
         End
         Begin VB.Menu CeramicEstimationsub 
            Caption         =   "«·« ›«ﬁÌ« "
            Index           =   4
         End
         Begin VB.Menu CeramicEstimationsub 
            Caption         =   "«·„‘«—Ì⁄"
            Index           =   5
         End
         Begin VB.Menu CeramicEstimationsub 
            Caption         =   " ”ÃÌ· «·«⁄„«· «·ÌÊ„Ì…"
            Index           =   6
         End
         Begin VB.Menu CeramicEstimationsub 
            Caption         =   "«·›Ê« Ì—"
            Index           =   7
         End
         Begin VB.Menu CeramicEstimationsub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   8
         End
      End
      Begin VB.Menu MnuProjects 
         Caption         =   "«œ«—… «·„‘«—Ì⁄"
         Begin VB.Menu MnuProjectsBasic 
            Caption         =   "»Ì«‰«  «”«”Ì…"
            Begin VB.Menu MnuProjectsBasicSub 
               Caption         =   "Õ«·«  «·„‘«—Ì⁄"
               Index           =   0
            End
            Begin VB.Menu MnuProjectsBasicSub 
               Caption         =   "«‰Ê«⁄ ⁄ﬁÊœ «·„‘«—Ì⁄"
               Index           =   1
            End
            Begin VB.Menu MnuProjectsBasicSub 
               Caption         =   "»Ì«‰«  „ﬁ«Ê·Ì «·»«ÿ‰"
               Index           =   2
            End
            Begin VB.Menu MnuProjectsBasicSub 
               Caption         =   " ⁄—Ì› «·»‰Êœ"
               Index           =   3
            End
            Begin VB.Menu MnuProjectsBasicSub 
               Caption         =   "ÊÕœ«  «·⁄„·Ì« "
               Index           =   4
            End
            Begin VB.Menu MnuProjectsBasicSub 
               Caption         =   "  ⁄—Ì› «·⁄„·Ì«  "
               Index           =   5
            End
            Begin VB.Menu MnuProjectsBasicSub 
               Caption         =   "»Ì«‰«  «·„⁄œ«  Ê«··«·« "
               Index           =   6
            End
            Begin VB.Menu MnuProjectsBasicSub 
               Caption         =   "⁄ﬁÊœ «·„ﬁ«Ê·Ì‰"
               Index           =   7
            End
         End
         Begin VB.Menu MnuProjectsTransactions 
            Caption         =   "»Ì«‰«  «·„‘«—Ì⁄"
            Index           =   0
         End
         Begin VB.Menu MnuProjectsTransactions 
            Caption         =   "’—› „Ê«œ ⁄·Ï „‘—Ê⁄"
            Index           =   1
         End
         Begin VB.Menu MnuProjectsTransactions 
            Caption         =   "„—œÊœ „‘«—Ì⁄"
            Index           =   2
         End
         Begin VB.Menu MnuProjectsTransactions 
            Caption         =   " Œ’Ì’ ⁄„«·Â ·„‘—Ê⁄"
            Index           =   3
         End
         Begin VB.Menu MnuProjectsTransactions 
            Caption         =   "‰ﬁ· ⁄„«·Â »Ì‰ «·„‘«—Ì⁄"
            Index           =   4
         End
         Begin VB.Menu MnuProjectsTransactions 
            Caption         =   " Œ’Ì’ „⁄œ«  ··„‘—Ê⁄"
            Index           =   5
         End
         Begin VB.Menu MnuProjectsTransactions 
            Caption         =   "‰ﬁ· „⁄œ«  »Ì‰ «·„‘«—Ì⁄"
            Index           =   6
         End
         Begin VB.Menu MnuProjectsTransactions 
            Caption         =   "„ «»⁄Â «·⁄„·Ì« "
            Index           =   7
         End
         Begin VB.Menu MnuProjectsTransactions 
            Caption         =   "›« Ê—… „‘—Ê⁄"
            Index           =   8
         End
         Begin VB.Menu MnuProjectsTransactions 
            Caption         =   "«’œ«— «·›Ê« Ì— «·‘Â—Ì…"
            Index           =   9
         End
         Begin VB.Menu MnuProjectsTransactions 
            Caption         =   " ﬁ«—Ì— «·„‘«—Ì⁄"
            Index           =   10
         End
      End
      Begin VB.Menu rentcar 
         Caption         =   "Œœ„«  VIP"
         Begin VB.Menu rentcarSub 
            Caption         =   "«·„Ê«ﬁ⁄"
            Index           =   0
         End
         Begin VB.Menu rentcarSub 
            Caption         =   "«·›∆« "
            Index           =   1
         End
         Begin VB.Menu rentcarSub 
            Caption         =   " ”ÃÌ· œŒÊ· «·„⁄œ« /«·”Ì«—« "
            Index           =   2
         End
         Begin VB.Menu rentcarSub 
            Caption         =   " ⁄—Ì› «·„ÊŸ›Ì‰"
            Index           =   3
         End
         Begin VB.Menu rentcarSub 
            Caption         =   " ”ÃÌ· «·Õ÷Ê— Ê«·«‰’—«›"
            Index           =   4
         End
         Begin VB.Menu rentcarSub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   5
            Begin VB.Menu rentcarSubReport 
               Caption         =   " ﬁ«—Ì— «·„⁄œ« /«·”Ì«—« "
               Index           =   0
            End
            Begin VB.Menu rentcarSubReport 
               Caption         =   " ﬁ«—Ì— «·„ÊŸ›Ì‰"
               Index           =   1
            End
         End
      End
      Begin VB.Menu rsInvestment 
         Caption         =   "«·«” À„«— «·⁄ﬁ«—Ì"
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   "«·»Ì«‰«  «·«”«”Ì…"
            Index           =   0
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   "»Ì«‰«  «·„”«Â„Ì‰"
            Index           =   1
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   "› Õ «·„”«Â„…"
            Index           =   2
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   "› Õ «·«ﬂ  «» ›Ì „”«Â„…"
            Index           =   3
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   "«ﬂ  «» «·„”«Â„Ì‰"
            Index           =   4
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   "‘—«¡ «·«—«÷Ì"
            Index           =   5
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   " ›⁄Ì· «·„”«Â„…"
            Index           =   6
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   "„’—Ê›«  «· ÿÊÌ—"
            Index           =   7
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   "„—œÊœ«  «· ÿÊÌ—"
            Index           =   8
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   " ﬁ”Ì„ «·«—«÷Ì"
            Index           =   9
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   "› Õ «·»Ì⁄"
            Index           =   10
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   "›« Ê—… «·„»Ì⁄« "
            Index           =   11
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   " ’›Ì… «·„”«Â„…"
            Index           =   12
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   " Ê“Ì⁄ «·«—»«Õ"
            Index           =   13
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   "«· ‰«“·"
            Index           =   14
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   "«·«›—«€"
            Index           =   15
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   "„‘«—Ì⁄ «·„”«Â„« "
            Index           =   16
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   "ÕÃ“ «·ÊÕœ« "
            Index           =   17
         End
         Begin VB.Menu rsInvestmentsUB 
            Caption         =   " ﬁ«—Ì— «·«” À„«—"
            Index           =   18
         End
      End
      Begin VB.Menu RealEstateMarketing 
         Caption         =   " «· ”ÊÌﬁ «·⁄ﬁ«—Ì"
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«·œÊ·"
            Index           =   0
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«·„œ‰"
            Index           =   1
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«·«ÕÌ«¡"
            Index           =   2
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«·‘Ê«—⁄"
            Index           =   3
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "Õ«·Â «·⁄ﬁ«—"
            Index           =   4
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«‰Ê«⁄ «·⁄ﬁ«—"
            Index           =   5
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«‰Ê«⁄  «·ÊÕœ« "
            Index           =   6
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "„Ã„Ê⁄«  «·„‰«œÌ»"
            Index           =   7
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«·„‰«œÌ»"
            Index           =   8
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«‰Ê«⁄ «· ‘ÿÌ»"
            Index           =   9
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«·«ÿ·«·« "
            Index           =   10
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«‰Ê«⁄ «·⁄„·«¡"
            Index           =   11
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«·⁄„·«¡"
            Index           =   12
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«·⁄—Ê÷ Ê«·ÿ·»« "
            Index           =   13
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "„ﬁ«—‰… «·⁄—Ê÷ Ê«·ÿ·»« "
            Index           =   14
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«ﬁ›«· «·ÿ·»« "
            Index           =   15
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«·ﬁ«∆„Â «·”Êœ«¡"
            Index           =   16
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "œ·Ì· «·Â« ›"
            Index           =   17
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   " ”ÃÌ· «·« ’«·« "
            Index           =   18
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«·‰»ÌÂ« "
            Index           =   19
         End
         Begin VB.Menu RealEstateMarketingSub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   20
         End
      End
      Begin VB.Menu AssetsMngBase 
         Caption         =   "«œ«—… «·«„·«ﬂ"
         Begin VB.Menu AssetsMng 
            Caption         =   "„·›«  «”«”Ì…       "
            Index           =   0
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   "«‰Ê«⁄ «·⁄„·«¡"
               Index           =   0
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   "«‰Ê«⁄ «·⁄ﬁ«—« "
               Index           =   1
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   "«‰Ê«⁄ «·ÊÕœ« "
               Index           =   2
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   "„Ã„Ê⁄«  «·„‰«œÌ»"
               Index           =   3
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   "»Ì«‰«  «·„‰«œÌ»"
               Index           =   4
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   " ⁄—Ì› «·œÊ·"
               Index           =   5
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   " ⁄—Ì›  «·„œ‰"
               Index           =   6
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   " ⁄—Ì›  «·«ÕÌ«¡"
               Index           =   7
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   " ⁄—Ì›  «·„Œÿÿ« "
               Index           =   8
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   "œ·Ì· «·Â« ›"
               Index           =   9
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   " ⁄—Ì› «·„·«ﬂ"
               Index           =   10
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   "  ⁄—Ì›  «·„” √Ã—Ì‰ "
               Index           =   11
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   " ⁄—Ì› «·„’—Ê›« "
               Index           =   12
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   "Õ«·«  «·ÊÕœ« "
               Index           =   13
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   "„ﬂÊ‰«  «·ÊÕœ« "
               Index           =   14
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   " ⁄—Ì› «·⁄ﬁ«—« "
               Index           =   15
            End
            Begin VB.Menu AssetsMngBasicFiles 
               Caption         =   "«‰Ê«⁄ «·«‘⁄«—« "
               Index           =   16
            End
         End
         Begin VB.Menu AssetsMng 
            Caption         =   "«·Õ—ﬂ« "
            Index           =   1
            Begin VB.Menu AssetsMngTrans 
               Caption         =   " ”ÃÌ· ÿ·»«  «·»Ì⁄ Ê «·‘—«¡ Ê «·«ÌÃ«—"
               Index           =   0
               Visible         =   0   'False
            End
            Begin VB.Menu AssetsMngTrans 
               Caption         =   " ”ÃÌ·  ⁄—Ê÷   «·»Ì⁄ Ê «·‘—«¡  Ê «·«ÌÃ«—"
               Index           =   1
               Visible         =   0   'False
            End
            Begin VB.Menu AssetsMngTrans 
               Caption         =   "„ﬁ«—‰… «·⁄—Ê÷ Ê «·ÿ·»« "
               Index           =   2
               Visible         =   0   'False
            End
            Begin VB.Menu AssetsMngTrans 
               Caption         =   "«·⁄ﬁÊœ"
               Index           =   4
               Begin VB.Menu AssetsMngContrac 
                  Caption         =   "⁄ﬁÊœ «ÌÃ«—"
                  Index           =   0
               End
               Begin VB.Menu AssetsMngContrac 
                  Caption         =   "⁄ﬁÊœ »Ì⁄"
                  Index           =   1
                  Visible         =   0   'False
               End
            End
            Begin VB.Menu AssetsMngTrans 
               Caption         =   "”‰«œ«  ﬁ»÷"
               Index           =   5
            End
            Begin VB.Menu AssetsMngTrans 
               Caption         =   "”‰œ«  ’—›"
               Index           =   6
            End
            Begin VB.Menu AssetsMngTrans 
               Caption         =   "”‰œ ’—› «·„œ›Ê⁄« "
               Index           =   7
            End
            Begin VB.Menu AssetsMngTrans 
               Caption         =   "ÿ»«⁄Â «·‘Ìﬂ« "
               Index           =   8
            End
            Begin VB.Menu AssetsMngTrans 
               Caption         =   "«· ’›Ì…"
               Index           =   9
            End
            Begin VB.Menu AssetsMngTrans 
               Caption         =   " ’›Ì… «·⁄Âœ…"
               Index           =   10
            End
            Begin VB.Menu AssetsMngTrans 
               Caption         =   "«·ﬁ«∆„Â «·”Êœ«¡"
               Index           =   11
            End
            Begin VB.Menu AssetsMngTrans 
               Caption         =   "«’œ«— «‘⁄«—  ”œÌœ - «‰–«—"
               Index           =   12
            End
            Begin VB.Menu AssetsMngTrans 
               Caption         =   "„’—Ê›«  «·ﬂÂ—»«¡ Ê«· ’›Ì« "
               Index           =   13
            End
            Begin VB.Menu AssetsMngTrans 
               Caption         =   "«·’Ì«‰Â"
               Index           =   14
               Begin VB.Menu estateMain 
                  Caption         =   "ÿ·» ’Ì«‰…"
                  Index           =   0
               End
               Begin VB.Menu estateMain 
                  Caption         =   "«ﬁ›«· ÿ·» ’Ì«‰…"
                  Index           =   1
               End
            End
            Begin VB.Menu AssetsMngTrans 
               Caption         =   "«·Œ’Ê„« "
               Index           =   15
            End
         End
         Begin VB.Menu AssetsMng 
            Caption         =   "«·«” Õﬁ«ﬁ«  "
            Index           =   2
            Begin VB.Menu AssetsMngsub 
               Caption         =   "«À»«  «·«” Õﬁ«ﬁ« "
               Index           =   0
            End
            Begin VB.Menu AssetsMngsub 
               Caption         =   "«À»«  «·«Ì—«œ"
               Index           =   1
            End
         End
         Begin VB.Menu AssetsMng 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   3
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ«—Ì— «·⁄ﬁ«—« "
               Index           =   0
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ«—Ì— «·⁄„Ê·« "
               Index           =   1
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ«—Ì— «·⁄ﬁÊœ «·„‰ ÂÌ…"
               Index           =   2
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ«—Ì— «·’Ì«‰…"
               Index           =   3
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ«—Ì— «· ’›Ì« "
               Index           =   4
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ«—Ì— «· Õ’Ì·« "
               Index           =   5
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ«—Ì— «·„·«ﬂ"
               Index           =   6
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ—Ì— «·«‘⁄«—«  Ê«·Œÿ«»« "
               Index           =   7
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ«—Ì— «Ã„«·Ì…"
               Index           =   8
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   "‰ﬁ«—Ì— «·⁄—»Ê‰"
               Index           =   9
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ«—Ì— «·”⁄Ì"
               Index           =   10
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ«—Ì— «·«ÌÃ«—«  «·„” Õﬁ…"
               Index           =   11
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ«—Ì— «·„” √Ã—Ì‰"
               Index           =   12
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ—Ì— Õ«·… «·ÊÕœ« "
               Index           =   13
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ«—Ì— «·„’—Ê›«  Ê«·«Ì—«œ« "
               Index           =   14
            End
            Begin VB.Menu AssetsMngReport 
               Caption         =   " ﬁ«—Ì— «·⁄ﬁÊœ «·„’›«…"
               Index           =   15
            End
         End
         Begin VB.Menu AssetsMng 
            Caption         =   "—”«∆· ··⁄„·«¡"
            Index           =   4
         End
         Begin VB.Menu AssetsMng 
            Caption         =   "«·»ÕÀ ⁄‰ «·ÊÕœ«  «·‘«€—…"
            Index           =   5
         End
      End
      Begin VB.Menu TransporterMain 
         Caption         =   "«·‰ﬁ·Ì« "
         Begin VB.Menu TransporterSub 
            Caption         =   "»Ì«‰«  «·„œ‰"
            Index           =   0
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "«·„”«›«  »Ì‰ «·„œ‰"
            Index           =   1
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "«·„Ê«‰Ì¡"
            Index           =   2
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "«·”›‰"
            Index           =   3
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "«‰Ê«⁄ «·‰ﬁ·"
            Index           =   4
         End
         Begin VB.Menu TransporterSub 
            Caption         =   " ⁄—Ì› «·—œÊœ"
            Index           =   5
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "»Ì«‰«  «·⁄„·«¡"
            Index           =   6
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "»Ì«‰«  «·„Ê—œÌ‰"
            Index           =   7
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "»Ì«‰«  «·”«∆ﬁÌ‰"
            Index           =   8
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "«‰Ê«⁄ «·„—ﬂ»« "
            Index           =   9
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "ÿ—«“«  «·„—ﬂ»« "
            Index           =   10
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "‘—ﬂ«  «· √„Ì‰"
            Index           =   11
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "«‰Ê«⁄ «·’Ì«‰… «·œÊ—Ì…"
            Index           =   12
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "»Ì«‰«  «·„—ﬂ»« "
            Index           =   13
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "Œÿ… «·’Ì«‰Â"
            Index           =   14
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "« ›«ﬁÌ«  ⁄„·«¡ «·‰ﬁ·"
            Index           =   15
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "√Ê«„— «· Õ„Ì·"
            Index           =   16
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "»Ì«‰«  «·—Õ·« "
            Index           =   17
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "›Ê« Ì— «·⁄„·«¡"
            Index           =   18
         End
         Begin VB.Menu TransporterSub 
            Caption         =   " ’›ÌÂ «·⁄Âœ… ··”«∆ﬁÌ‰"
            Index           =   19
         End
         Begin VB.Menu TransporterSub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   20
         End
      End
      Begin VB.Menu hajMnu 
         Caption         =   "«·ÕÃ Ê«·⁄„—…"
         Begin VB.Menu hajMnuSub 
            Caption         =   "«·»Ì«‰«  «·«”«”Ì…"
            Index           =   0
            Begin VB.Menu hajMnuSub1 
               Caption         =   "«·»Ì«‰«  «·«”«”Ì…"
               Index           =   0
            End
            Begin VB.Menu hajMnuSub1 
               Caption         =   "»Ì«‰«  «·„œ‰"
               Index           =   1
            End
            Begin VB.Menu hajMnuSub1 
               Caption         =   "«·„”«›«  »Ì‰ «·„œ‰"
               Index           =   2
               Visible         =   0   'False
            End
            Begin VB.Menu hajMnuSub1 
               Caption         =   "»Ì«‰«  «·”«∆ﬁÌ‰ "
               Index           =   3
            End
            Begin VB.Menu hajMnuSub1 
               Caption         =   "«‰Ê«⁄ «·„—ﬂ»« "
               Index           =   4
            End
            Begin VB.Menu hajMnuSub1 
               Caption         =   "»Ì«‰«  «·„—ﬂ»« "
               Index           =   5
            End
            Begin VB.Menu hajMnuSub1 
               Caption         =   "«‰Ê«⁄ «·⁄„·«¡"
               Index           =   6
               Visible         =   0   'False
            End
            Begin VB.Menu hajMnuSub1 
               Caption         =   "ÿ·» › Õ Õ”«» ⁄„Ì·"
               Index           =   7
            End
            Begin VB.Menu hajMnuSub1 
               Caption         =   "‘—ﬂ«  „‰ «·œ«Œ·"
               Index           =   8
               Visible         =   0   'False
            End
            Begin VB.Menu hajMnuSub1 
               Caption         =   "«·⁄„·«¡"
               Index           =   9
            End
            Begin VB.Menu hajMnuSub1 
               Caption         =   "« ›«ﬁÌ«  «·⁄„·«¡"
               Index           =   10
            End
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   "«⁄ „«œ ‰ﬁ· «·ÕÃ«Ã Ê «·„⁄ „—Ì‰"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   "«·„ÿ«·»« "
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   "«·„Œ«·’« "
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   "ÿ·»«  «·ÕÃ“  «·⁄„—…"
            Index           =   4
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   " √ﬂÌœ «·ÕÃ“"
            Index           =   5
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   "«„—  ‘€Ì· Õ«›·… «·⁄„—…"
            Index           =   6
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   "ÃœÊ· «· —ÕÌ·«  «·⁄„—… "
            Index           =   7
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   "«·„”«—«  «·„Œ’Ê„… ··⁄„—…"
            Index           =   8
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   "«⁄ „«œ «—ﬂ«» «·ÕÃ«Ã"
            Index           =   9
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   "ÃœÊ· «· —ÕÌ·«  ·«—ﬂ«» «·ÕÃ«Ã"
            Index           =   10
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   "«⁄ „«œ «—ﬂ«» «·„‘«⁄—"
            Index           =   11
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   " Ê“Ì⁄ Õ«›·«  «·„‘«⁄—"
            Index           =   12
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   "«Œ·«¡ «·ÿ—›"
            Index           =   13
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   "«·„ÿ«·»« "
            Index           =   14
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   "«·Õ”„Ì« "
            Index           =   15
         End
         Begin VB.Menu hajMnuSub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   16
         End
      End
      Begin VB.Menu CarMaintenance 
         Caption         =   "Ê—‘ ’Ì«‰Â «·„⁄œ« /«·”Ì«—« "
         Begin VB.Menu CarMaintenancesub 
            Caption         =   "«·»Ì«‰«  «·«”«”Ì…"
            Index           =   0
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "«‰Ê«⁄ «·„—ﬂ»« "
               Index           =   0
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "ÿ—«“«  «·„—ﬂ»« "
               Index           =   1
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "»Ì«‰«  «·„—ﬂ»« "
               Index           =   2
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "«‰Ê«⁄ «·«’·«Õ« "
               Index           =   3
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "«‰Ê«⁄ «·„‘ —Ì«  Ê «·«⁄„«· ·Œ«—ÃÌ…"
               Index           =   4
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "«·„‘ —Ì«  Ê «·«⁄„«· ·Œ«—ÃÌ…"
               Index           =   5
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "«‰Ê«⁄ «⁄ÿ«· ›Õ’ «·ﬂ„»ÌÊ —"
               Index           =   6
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "«·Ê«‰ «·„—ﬂ»« "
               Index           =   7
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "»Ì«‰«  «·„Œ«“‰"
               Index           =   8
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "„Ã„Ê⁄«  «·«’‰«›"
               Index           =   9
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "«·ÊÕœ« "
               Index           =   10
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "«·«’‰«›"
               Index           =   11
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "«·⁄„·«¡ Ê «·„Ê—œÌ‰"
               Index           =   12
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "»Ì«‰«  «·„ÊŸ›Ì‰"
               Index           =   13
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "„Ã„Ê⁄«  «·⁄„· »«·Ê—‘…"
               Index           =   14
               Visible         =   0   'False
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "»Ì«‰«  «·„‘—›Ì‰"
               Index           =   15
               Visible         =   0   'False
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "»Ì«‰«  «·„Â‰œ”Ì‰   Ê «·›‰ÌÌ‰"
               Index           =   16
               Visible         =   0   'False
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   "»Ì«‰«  «ﬁ”«„ «·Ê—‘…"
               Index           =   17
            End
            Begin VB.Menu CarMaintenancesub1 
               Caption         =   " «·„‘—›Ì‰ Ê«·›‰ÌÌ‰"
               Index           =   18
            End
         End
         Begin VB.Menu CarMaintenancesub 
            Caption         =   "«·Õ—ﬂ« "
            Index           =   1
            Begin VB.Menu CarMaintenancesub2 
               Caption         =   "«–‰ œŒÊ· ’Ì«‰…"
               Index           =   0
            End
            Begin VB.Menu CarMaintenancesub2 
               Caption         =   " ›« Ê—… ›Õ’ ﬂ„»ÌÊ —"
               Index           =   1
            End
            Begin VB.Menu CarMaintenancesub2 
               Caption         =   "”‰œ ’—› ﬁÿ⁄ €Ì«—"
               Index           =   2
            End
            Begin VB.Menu CarMaintenancesub2 
               Caption         =   "√Ê«„— «·‘—«¡"
               Index           =   3
            End
            Begin VB.Menu CarMaintenancesub2 
               Caption         =   "›« Ê—…  ··’Ì«‰…"
               Index           =   4
            End
            Begin VB.Menu CarMaintenancesub2 
               Caption         =   "«·⁄„Ê·«  «·„” Õﬁ…"
               Index           =   5
            End
            Begin VB.Menu CarMaintenancesub2 
               Caption         =   "«ÃÊ— «·Ìœ"
               Index           =   6
            End
            Begin VB.Menu CarMaintenancesub2 
               Caption         =   "«·ﬁÿ⁄ «·„ﬁœ—…"
               Index           =   7
            End
         End
         Begin VB.Menu CarMaintenancesub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   2
         End
      End
      Begin VB.Menu MnuMaintnance 
         Caption         =   " «·’Ì«‰…"
         Begin VB.Menu MnuMaintnanceBasic 
            Caption         =   "»Ì«‰«  «”«”ÌÂ       "
            Begin VB.Menu MnuMaintnanceBasicSub 
               Caption         =   "«‰Ê«⁄ «·’Ì«‰…"
               Index           =   0
            End
            Begin VB.Menu MnuMaintnanceBasicSub 
               Caption         =   "«‰Ê«⁄ «·„—ﬂ»« "
               Index           =   1
            End
            Begin VB.Menu MnuMaintnanceBasicSub 
               Caption         =   "ÿ—«“«  «·„—ﬂ»« "
               Index           =   2
            End
            Begin VB.Menu MnuMaintnanceBasicSub 
               Caption         =   "«·Ê«‰ «·„—ﬂ»« "
               Index           =   3
            End
            Begin VB.Menu MnuMaintnanceBasicSub 
               Caption         =   "»Ì«‰«  «·„—ﬂ»« "
               Index           =   4
            End
            Begin VB.Menu MnuMaintnanceBasicSub 
               Caption         =   "«ﬁ”«„ «·Ê—‘…"
               Index           =   5
            End
            Begin VB.Menu MnuMaintnanceBasicSub 
               Caption         =   "«·›‰ÌÌ‰ Ê«·„‘—›Ì‰"
               Index           =   6
            End
            Begin VB.Menu MnuMaintnanceBasicSub 
               Caption         =   "‘—ﬂ«  «·’Ì«‰Â"
               Index           =   7
            End
            Begin VB.Menu MnuMaintnanceBasicSub 
               Caption         =   " ⁄—Ì› «·„’—Ê›« "
               Index           =   8
            End
            Begin VB.Menu MnuMaintnanceBasicSub 
               Caption         =   "«·„Œ«“‰"
               Index           =   9
            End
            Begin VB.Menu MnuMaintnanceBasicSub 
               Caption         =   "„Ã„Ê⁄«  «·«’‰«›"
               Index           =   10
            End
            Begin VB.Menu MnuMaintnanceBasicSub 
               Caption         =   "»Ì«‰«  «·«’‰ ›"
               Index           =   11
            End
            Begin VB.Menu MnuMaintnanceBasicSub 
               Caption         =   " ⁄—Ì› «·Ê—œÌ« "
               Index           =   12
            End
         End
         Begin VB.Menu MnuMaintnanceTransactions 
            Caption         =   "ŒÿÂ «·’Ì«‰…"
            Index           =   0
         End
         Begin VB.Menu MnuMaintnanceTransactions 
            Caption         =   "ÿ·» ’Ì«‰…"
            Index           =   1
         End
         Begin VB.Menu MnuMaintnanceTransactions 
            Caption         =   "«„— ‘€·"
            Index           =   2
         End
         Begin VB.Menu MnuMaintnanceTransactions 
            Caption         =   "«·ÿ·»«  «·œ«Œ·Ì…"
            Index           =   3
         End
         Begin VB.Menu MnuMaintnanceTransactions 
            Caption         =   "”‰œ «” ·«„ „Ê«œ  "
            Index           =   4
         End
         Begin VB.Menu MnuMaintnanceTransactions 
            Caption         =   "’—› ﬁÿ⁄ €Ì«— ··’Ì«‰…"
            Index           =   5
         End
         Begin VB.Menu MnuMaintnanceTransactions 
            Caption         =   " ”ÃÌ· «·Ê—œÌ…"
            Index           =   6
         End
         Begin VB.Menu MnuMaintnanceTransactions 
            Caption         =   "—’Ìœ «›  «ÕÌ ·„Œ“‰ «·’Ì«‰…"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu MnuMaintnanceTransactions 
            Caption         =   " ”·Ì„ Ê≈” ·«„ ⁄Âœ ⁄Ì‰Ì…"
            Index           =   8
         End
         Begin VB.Menu MnuMaintnanceTransactions 
            Caption         =   " ›ÊÌ÷ «·ﬁÌ«œ…"
            Index           =   9
         End
         Begin VB.Menu MnuMaintnanceTransactions 
            Caption         =   "«·÷„«‰"
            Index           =   10
         End
         Begin VB.Menu MnuMaintnanceTransactions 
            Caption         =   "‰ﬁ—Ì— «·Õ«œÀ"
            Index           =   11
         End
         Begin VB.Menu MnuMaintnanceTransactions 
            Caption         =   " ﬁ«—Ì— «·’Ì«‰Â"
            Index           =   12
         End
      End
      Begin VB.Menu Strategy 
         Caption         =   "«·‰ﬁ· «·„œ—”Ì"
         Begin VB.Menu StrategyBasicdata 
            Caption         =   "«·»Ì«‰«  «·«”«”Ì…"
            Index           =   0
            Begin VB.Menu StrategyBasicdatasub 
               Caption         =   "«·„Õ«›Ÿ« "
               Index           =   0
            End
            Begin VB.Menu StrategyBasicdatasub 
               Caption         =   "«·„‰«ÿﬁ «·«œ«—Ì…"
               Index           =   1
            End
            Begin VB.Menu StrategyBasicdatasub 
               Caption         =   "«·„ ⁄ÂœÌ‰"
               Index           =   2
            End
            Begin VB.Menu StrategyBasicdatasub 
               Caption         =   "«‰Ê«⁄ «·Õ«›·« "
               Index           =   3
            End
            Begin VB.Menu StrategyBasicdatasub 
               Caption         =   "»Ì«‰«  «·”«∆ﬁÌ‰"
               Index           =   4
               Visible         =   0   'False
            End
            Begin VB.Menu StrategyBasicdatasub 
               Caption         =   "»Ì«‰«  «·Õ«›·« "
               Index           =   5
            End
            Begin VB.Menu StrategyBasicdatasub 
               Caption         =   "»Ì«‰«  «·„œ«—”"
               Index           =   6
            End
            Begin VB.Menu StrategyBasicdatasub 
               Caption         =   "«·⁄«„ «·œ—«”Ì Ê«·› —« "
               Index           =   7
            End
            Begin VB.Menu StrategyBasicdatasub 
               Caption         =   "„Ã„Ê⁄«  «·„Œ«·›« "
               Index           =   8
            End
            Begin VB.Menu StrategyBasicdatasub 
               Caption         =   "«‰Ê«⁄ «·„Œ«·›« "
               Index           =   9
            End
            Begin VB.Menu StrategyBasicdatasub 
               Caption         =   "«‰Ê«⁄ «·⁄ÿ·« "
               Index           =   10
            End
            Begin VB.Menu StrategyBasicdatasub 
               Caption         =   "«·”«∆ﬁÌ‰"
               Index           =   11
            End
         End
         Begin VB.Menu StrategyBasicdata 
            Caption         =   "⁄ﬁœ Ê“«—…"
            Index           =   1
         End
         Begin VB.Menu StrategyBasicdata 
            Caption         =   " Œ’Ì’ „‘—›Ì‰ ··„œ«—”"
            Index           =   2
         End
         Begin VB.Menu StrategyBasicdata 
            Caption         =   " Œ’Ì’ «·”«∆ﬁÌ‰ ··Õ«›·« "
            Index           =   3
         End
         Begin VB.Menu StrategyBasicdata 
            Caption         =   " Œ’Ì’ «·Õ«›·« "
            Index           =   4
         End
         Begin VB.Menu StrategyBasicdata 
            Caption         =   "⁄ﬁœ «·«”‰«œ"
            Index           =   5
         End
         Begin VB.Menu StrategyBasicdata 
            Caption         =   "«À»«  «· ⁄ÿ· ··„‰«ÿﬁ"
            Index           =   6
         End
         Begin VB.Menu StrategyBasicdata 
            Caption         =   "«À»«  «·„Œ«·›« "
            Index           =   7
         End
         Begin VB.Menu StrategyBasicdata 
            Caption         =   "«À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ⁄·Ì «·Ê“«—…"
            Index           =   8
         End
         Begin VB.Menu StrategyBasicdata 
            Caption         =   "«À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ··„ ⁄ÂœÌ‰"
            Index           =   9
         End
         Begin VB.Menu StrategyBasicdata 
            Caption         =   "ÿ·» ’—› ··„ ⁄ÂœÌ‰"
            Index           =   10
         End
         Begin VB.Menu StrategyBasicdata 
            Caption         =   "”‰œ ’—› „ ⁄ÂœÌ‰"
            Index           =   11
         End
         Begin VB.Menu StrategyBasicdata 
            Caption         =   "«Ìﬁ«› ”Ì«—…"
            Index           =   12
         End
         Begin VB.Menu StrategyBasicdata 
            Caption         =   "«÷«›«  «·«Ì«„"
            Index           =   13
         End
         Begin VB.Menu StrategyBasicdata 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   14
         End
      End
      Begin VB.Menu StudentMenue 
         Caption         =   "«·„⁄«Âœ «· ⁄·Ì„Ì…"
         Begin VB.Menu StudentMenueSub 
            Caption         =   "«·»Ì«‰«  «·«”«”Ì…"
            Index           =   0
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   "«·„œ—»Ì‰"
            Index           =   1
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   "»Ì«‰«  «·‘—ﬂ« "
            Index           =   2
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   "ÿ·»  œ—Ì»"
            Index           =   3
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   "»Ì‰«  «·ÿ·«»"
            Index           =   4
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   "«·⁄ﬁÊœ"
            Index           =   5
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   " —‘ÌÕ «·ÿ·«» "
            Index           =   6
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   "„Ê«›ﬁÂ «· —‘ÌÕ"
            Index           =   7
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   "„Ã„Ê⁄«  «·ÿ·»…"
            Index           =   8
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   "«·Õ÷Ê— Ê «·«‰’—«›"
            Index           =   9
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   "«·« ’«·« "
            Index           =   10
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   "«·›’·"
            Index           =   11
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   " „œÌœ Ê«‰Â«¡ «·„Ã„Ê⁄« "
            Index           =   12
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   "«‰Â«¡ ⁄ﬁÊœ «·‘—ﬂ« "
            Index           =   13
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   "≈’œ«— «·›Ê« Ì—"
            Index           =   14
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   "‰ﬁ· Ê«÷«›Â ÊÕ–› «·ÿ·«» „‰ «·„Ã„Ê⁄« "
            Index           =   15
         End
         Begin VB.Menu StudentMenueSub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   16
         End
      End
      Begin VB.Menu Archiving 
         Caption         =   "«·«—‘Ì› «·«·ﬂ —Ê‰Ì"
         Begin VB.Menu ArchivingSub 
            Caption         =   "«·«œ«—«  …«·«ﬁ”«„"
            Index           =   0
         End
         Begin VB.Menu ArchivingSub 
            Caption         =   "»Ì«‰«  «·«—‘Ì› ›Ì «·«ﬁ”«„"
            Index           =   1
         End
         Begin VB.Menu ArchivingSub 
            Caption         =   " €—› «·Õ›Ÿ ›Ì ﬂ· «—‘Ì›"
            Index           =   2
         End
         Begin VB.Menu ArchivingSub 
            Caption         =   "’‰«œÌﬁ/œÊ·«Ì» «·Õ›Ÿ ›Ì «·€—›"
            Index           =   3
         End
         Begin VB.Menu ArchivingSub 
            Caption         =   "»Ì«‰«  «·«—›› ›Ì ﬂ· ’‰œÊﬁ/œÊ·«»"
            Index           =   4
         End
         Begin VB.Menu ArchivingSub 
            Caption         =   "«‰Ê«⁄ «·„⁄«„·« "
            Index           =   5
         End
         Begin VB.Menu ArchivingSub 
            Caption         =   "«÷«›… «·‰„«–Ã"
            Index           =   6
         End
         Begin VB.Menu ArchivingSub 
            Caption         =   " ”ÃÌ· «·„⁄«„·« "
            Index           =   7
         End
         Begin VB.Menu ArchivingSub 
            Caption         =   "⁄—÷ «·„⁄«„·« "
            Index           =   8
         End
         Begin VB.Menu ArchivingSub 
            Caption         =   " ‰»Ì… «·„⁄«„·«  "
            Index           =   9
         End
         Begin VB.Menu ArchivingSub 
            Caption         =   " ﬁ«—Ì— «·„⁄«„·« "
            Index           =   10
         End
      End
      Begin VB.Menu LegalIssue 
         Caption         =   "«·‘∆Ê‰ «·ﬁ«‰Ê‰Ì…"
         Visible         =   0   'False
         Begin VB.Menu LegalIssueSub 
            Caption         =   "«”„«¡ «·„Õ«ﬂ„"
            Index           =   0
         End
         Begin VB.Menu LegalIssueSub 
            Caption         =   "«‰Ê«⁄ «·ﬁ÷«Ì«"
            Index           =   1
         End
         Begin VB.Menu LegalIssueSub 
            Caption         =   "»Ì«‰«  «·ﬁ÷«Ì«"
            Index           =   2
         End
         Begin VB.Menu LegalIssueSub 
            Caption         =   " ”ÃÌ· „Ê«⁄Ìœ «·Ã·”« "
            Index           =   3
         End
         Begin VB.Menu LegalIssueSub 
            Caption         =   " ”ÃÌ· ”Ì— «·ﬁ÷Ì…"
            Index           =   4
         End
         Begin VB.Menu LegalIssueSub 
            Caption         =   "«· ‰»ÌÂ« "
            Index           =   5
         End
         Begin VB.Menu LegalIssueSub 
            Caption         =   "LegalIssueSub"
            Index           =   6
            Visible         =   0   'False
         End
      End
      Begin VB.Menu dev 
         Caption         =   "„ «»⁄Â «·«œ«¡"
         Begin VB.Menu devsub 
            Caption         =   " ⁄—Ì› «·„Â«„"
            Index           =   0
         End
         Begin VB.Menu devsub 
            Caption         =   "„ «»⁄Â «·„Â«„"
            Index           =   1
         End
         Begin VB.Menu devsub 
            Caption         =   " ﬁ—Ì— ”Ì— «·⁄„· «·ÌÊ„Ì"
            Index           =   2
         End
         Begin VB.Menu devsub 
            Caption         =   " ‰»ÌÂ«  «·„Â«„"
            Index           =   3
         End
         Begin VB.Menu devsub 
            Caption         =   " „—«Ã⁄Â Ê  ﬁÌÌ„ ”Ì— «·⁄„·"
            Index           =   4
         End
         Begin VB.Menu devsub 
            Caption         =   " ﬁ«—Ì— «·„Â«„"
            Index           =   5
         End
      End
      Begin VB.Menu Tailor 
         Caption         =   "«·ŒÌ«ÿ…"
         Begin VB.Menu Tailorsub 
            Caption         =   " ⁄—Ì› «·„Â«„"
            Index           =   0
         End
         Begin VB.Menu Tailorsub 
            Caption         =   " ⁄—Ì› «·„ﬁ«”« "
            Index           =   1
         End
         Begin VB.Menu Tailorsub 
            Caption         =   "»Ì«‰«  «·«’‰«›"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu Tailorsub 
            Caption         =   "»Ì«‰«  «·„ÊŸ›Ì‰"
            Index           =   3
         End
         Begin VB.Menu Tailorsub 
            Caption         =   "»Ì«‰«  «·⁄„·«¡"
            Index           =   4
         End
         Begin VB.Menu Tailorsub 
            Caption         =   "√Ê«„— «·‘€·"
            Index           =   5
         End
         Begin VB.Menu Tailorsub 
            Caption         =   "›Ê« Ì— «·„»Ì⁄« "
            Index           =   6
         End
         Begin VB.Menu Tailorsub 
            Caption         =   "”‰œ«  «·ﬁ»÷"
            Index           =   7
         End
         Begin VB.Menu Tailorsub 
            Caption         =   " ”ÃÌ· «‰ «ÃÌ… «·„ÊŸ›Ì‰"
            Index           =   8
         End
         Begin VB.Menu Tailorsub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   9
         End
      End
      Begin VB.Menu Beauty 
         Caption         =   "«· Ã„Ì·"
         Begin VB.Menu Beautysub 
            Caption         =   "„Ã„Ê⁄«  «·Œœ„« "
            Index           =   0
         End
         Begin VB.Menu Beautysub 
            Caption         =   " ⁄—Ì› «·Œœ„« "
            Index           =   1
         End
         Begin VB.Menu Beautysub 
            Caption         =   "»Ì«‰«  «·⁄«„·« "
            Index           =   2
         End
         Begin VB.Menu Beautysub 
            Caption         =   "„Ã„Ê⁄«  «·⁄„·/«·‘Ì› « "
            Index           =   3
         End
         Begin VB.Menu Beautysub 
            Caption         =   "«‰Ê«⁄ «·—«Õ« "
            Index           =   4
         End
         Begin VB.Menu Beautysub 
            Caption         =   "—»ÿ «·„ÊŸ›Ì‰ »«·Œœ„« "
            Index           =   5
         End
         Begin VB.Menu Beautysub 
            Caption         =   " ⁄—Ì› «·⁄„Ì·« "
            Index           =   6
         End
         Begin VB.Menu Beautysub 
            Caption         =   "«‰Ê«⁄ «·ÕÃ“"
            Index           =   7
         End
         Begin VB.Menu Beautysub 
            Caption         =   "ŒÿÂ «·—«Õ« "
            Index           =   8
         End
         Begin VB.Menu Beautysub 
            Caption         =   "ÕÃ“ «·„Ê«⁄Ìœ"
            Index           =   9
         End
         Begin VB.Menu Beautysub 
            Caption         =   "⁄—÷ «·ÕÃÊ“« "
            Index           =   10
         End
         Begin VB.Menu Beautysub 
            Caption         =   "»Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄"
            Index           =   11
         End
         Begin VB.Menu Beautysub 
            Caption         =   "»Ì«‰«  «·ﬂ«‘Ì—"
            Index           =   12
         End
         Begin VB.Menu Beautysub 
            Caption         =   " ”ÃÌ· «·œŒÊ· ··‰ﬁÿÂ"
            Index           =   13
         End
         Begin VB.Menu Beautysub 
            Caption         =   "”‰œ ﬁ»÷ ⁄—»Ê‰"
            Index           =   14
         End
         Begin VB.Menu Beautysub 
            Caption         =   "ﬁ»÷ ⁄«„ ‰ﬁ«ÿ «·»Ì⁄"
            Index           =   15
         End
         Begin VB.Menu Beautysub 
            Caption         =   " ﬁ«—Ì— ‰ﬁ«ÿ «·»Ì⁄"
            Index           =   16
         End
         Begin VB.Menu Beautysub 
            Caption         =   " ﬁ«—Ì— «·⁄„·«¡ «·‰ﬁœÌÌ‰"
            Index           =   17
         End
         Begin VB.Menu Beautysub 
            Caption         =   "."
            Index           =   18
            Visible         =   0   'False
         End
      End
      Begin VB.Menu eye 
         Caption         =   "«·»’—Ì« "
         Begin VB.Menu eyeSub 
            Caption         =   "»Ì«‰«  «·„Œ«“‰"
            Index           =   0
         End
         Begin VB.Menu eyeSub 
            Caption         =   "„Ã„Ê⁄«  «·«’‰«›"
            Index           =   1
         End
         Begin VB.Menu eyeSub 
            Caption         =   "«·ÊÕœ« "
            Index           =   2
         End
         Begin VB.Menu eyeSub 
            Caption         =   "»Ì«‰«  «·«’‰«›"
            Index           =   3
         End
         Begin VB.Menu eyeSub 
            Caption         =   " ⁄—Ì› «·⁄œ”« "
            Index           =   4
         End
         Begin VB.Menu eyeSub 
            Caption         =   "»Ì«‰«  «·⁄„·«¡"
            Index           =   5
         End
         Begin VB.Menu eyeSub 
            Caption         =   "»Ì«‰«  «·„‰«œÌ»"
            Index           =   6
         End
         Begin VB.Menu eyeSub 
            Caption         =   "»Ì«‰«  «·«ÿ»«¡"
            Index           =   7
         End
         Begin VB.Menu eyeSub 
            Caption         =   "«· ⁄«ﬁœ«  / ‘—ﬂ«  «· √„Ì‰"
            Index           =   8
         End
         Begin VB.Menu eyeSub 
            Caption         =   "»Ì«‰«  «·œÊ·"
            Index           =   9
         End
         Begin VB.Menu eyeSub 
            Caption         =   "»Ì«‰«  «·«·Ê«‰"
            Index           =   10
         End
         Begin VB.Menu eyeSub 
            Caption         =   "«· ’‰Ì›« "
            Index           =   11
         End
         Begin VB.Menu eyeSub 
            Caption         =   "›Ê« Ì— «·„‘ —Ì« "
            Index           =   12
         End
         Begin VB.Menu eyeSub 
            Caption         =   "„—œÊœ«  «·„‘ —Ì« "
            Index           =   13
         End
         Begin VB.Menu eyeSub 
            Caption         =   "›Ê« Ì— «·„»Ì⁄« "
            Index           =   14
         End
         Begin VB.Menu eyeSub 
            Caption         =   "„—œÊœ«  «·„»Ì⁄« "
            Index           =   15
         End
         Begin VB.Menu eyeSub 
            Caption         =   "”‰œ«  «·ﬁ»÷"
            Index           =   16
         End
         Begin VB.Menu eyeSub 
            Caption         =   "”‰œ«  «·’—› "
            Index           =   17
         End
         Begin VB.Menu eyeSub 
            Caption         =   " ’›ÌÂ «·⁄Âœ…"
            Index           =   18
         End
         Begin VB.Menu eyeSub 
            Caption         =   "«·„œ›Ê⁄« "
            Index           =   19
         End
         Begin VB.Menu eyeSub 
            Caption         =   "«·«‘⁄«—« "
            Index           =   20
         End
         Begin VB.Menu eyeSub 
            Caption         =   "«· ﬁ«—Ì— «·⁄«„Â"
            Index           =   21
         End
         Begin VB.Menu eyeSub 
            Caption         =   "«· ﬁ«—Ì— «·„Õ«”»Ì…"
            Index           =   22
         End
         Begin VB.Menu eyeSub 
            Caption         =   " ‰»ÌÂ«  «·ÿ·»«  «·œ«Œ·Ì…"
            Index           =   23
         End
         Begin VB.Menu eyeSub 
            Caption         =   " ‰»ÌÂ«  «·„⁄—÷"
            Index           =   24
         End
         Begin VB.Menu eyeSub 
            Caption         =   " ‰»ÌÂ«  «·„⁄„·"
            Index           =   25
         End
      End
      Begin VB.Menu gobus 
         Caption         =   "‰ﬁ· «·—ﬂ«»"
         Begin VB.Menu gobusSub 
            Caption         =   "»Ì«‰«  «·œÊ·"
            Index           =   0
         End
         Begin VB.Menu gobusSub 
            Caption         =   "»Ì«‰«  «·„Õ«›Ÿ« "
            Index           =   1
         End
         Begin VB.Menu gobusSub 
            Caption         =   "«·”«›«  »Ì‰ «·„œ‰"
            Index           =   2
         End
         Begin VB.Menu gobusSub 
            Caption         =   "«‰Ê«⁄ «·„—ﬂ»« "
            Index           =   3
         End
         Begin VB.Menu gobusSub 
            Caption         =   "ÿ—«“«  «·„—ﬂ»« "
            Index           =   4
         End
         Begin VB.Menu gobusSub 
            Caption         =   "«·Ê«‰ «·„—ﬂ»« "
            Index           =   5
         End
         Begin VB.Menu gobusSub 
            Caption         =   "»Ì«‰«  «·„—ﬂ»« "
            Index           =   6
         End
         Begin VB.Menu gobusSub 
            Caption         =   "«·”«∆ﬁÌ‰"
            Index           =   7
         End
         Begin VB.Menu gobusSub 
            Caption         =   "»Ì«‰«  «·⁄„·«¡"
            Index           =   8
         End
         Begin VB.Menu gobusSub 
            Caption         =   " Œ’Ì’ «·”«∆ﬁÌ‰ ··Õ«›·« "
            Index           =   9
         End
         Begin VB.Menu gobusSub 
            Caption         =   " ”ÃÌ· «·—Õ·« "
            Index           =   10
         End
         Begin VB.Menu gobusSub 
            Caption         =   " ”ÃÌ· «·ÕÃ“"
            Index           =   11
         End
         Begin VB.Menu gobusSub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   12
         End
      End
      Begin VB.Menu xyz 
         Caption         =   "«·«” ﬁœ«„ Ê ‘€Ì· «·⁄„«·Â"
         Begin VB.Menu xyzSub 
            Caption         =   "»Ì«‰«  «·‘—ﬂ«  "
            Index           =   0
         End
         Begin VB.Menu xyzSub 
            Caption         =   "«·⁄ﬁÊœ  "
            Index           =   1
         End
         Begin VB.Menu xyzSub 
            Caption         =   "»Ì«‰«  «· √‘Ì—«  "
            Index           =   2
         End
         Begin VB.Menu xyzSub 
            Caption         =   "«· —‘ÌÕ"
            Index           =   3
         End
         Begin VB.Menu xyzSub 
            Caption         =   "«·„‘«—Ì⁄"
            Index           =   4
         End
         Begin VB.Menu xyzSub 
            Caption         =   " Œ’Ì’ «·⁄„«·Â ··„‘«—Ì⁄"
            Index           =   5
         End
         Begin VB.Menu xyzSub 
            Caption         =   "«·„” Œ·’« "
            Index           =   6
         End
         Begin VB.Menu xyzSub 
            Caption         =   "«·›Ê« Ì— «·‘Â—ÌÂ"
            Index           =   7
         End
         Begin VB.Menu xyzSub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   8
         End
      End
      Begin VB.Menu Farm 
         Caption         =   "«·„“—⁄Â"
         Begin VB.Menu FarmSub 
            Caption         =   "«·‰Ê⁄"
            Index           =   0
         End
         Begin VB.Menu FarmSub 
            Caption         =   "«· ’‰Ì›"
            Index           =   1
         End
         Begin VB.Menu FarmSub 
            Caption         =   "«·”·«·Â"
            Index           =   2
         End
         Begin VB.Menu FarmSub 
            Caption         =   "«·’›…"
            Index           =   3
         End
         Begin VB.Menu FarmSub 
            Caption         =   "«·„«·ﬂ"
            Index           =   4
         End
         Begin VB.Menu FarmSub 
            Caption         =   "«·Õ«·« "
            Index           =   5
         End
         Begin VB.Menu FarmSub 
            Caption         =   "«·„Êﬁ⁄"
            Index           =   6
         End
         Begin VB.Menu FarmSub 
            Caption         =   "«·«ÿ»«¡"
            Index           =   7
         End
         Begin VB.Menu FarmSub 
            Caption         =   " ”ÃÌ· »Ì«‰«  «·ÕÌÊ«‰« "
            Index           =   8
         End
         Begin VB.Menu FarmSub 
            Caption         =   "ÿ·» ⁄·«Ã"
            Index           =   9
         End
         Begin VB.Menu FarmSub 
            Caption         =   " ”ÃÌ· ‰›Êﬁ"
            Index           =   10
         End
         Begin VB.Menu FarmSub 
            Caption         =   "«· ﬁ«—Ì—"
            Index           =   11
         End
      End
      Begin VB.Menu Reports 
         Caption         =   "«· ﬁ«—Ì—"
         Begin VB.Menu Report 
            Caption         =   "«· ﬁ«—Ì— «·⁄«„…"
         End
         Begin VB.Menu DailyReport 
            Caption         =   "«· ﬁ—Ì— «·ÌÊ„Ì"
         End
         Begin VB.Menu MnuReports_Assblied 
            Caption         =   "«· ﬁ—Ì— «·„Ã„⁄ ⁄‰ › —…"
         End
         Begin VB.Menu ReportDesign 
            Caption         =   "„’„„ «· ﬁ«—Ì—"
         End
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "„œÌ— «·‰Ÿ«„"
      Begin VB.Menu Options 
         Caption         =   "«⁄œ«œ«  «·‰Ÿ«„"
      End
      Begin VB.Menu UsersData 
         Caption         =   "„” Œœ„Ì «·‰Ÿ«„"
         Begin VB.Menu UsersGroup 
            Caption         =   "„Ã„Ê⁄«  «·„” Œœ„Ì‰"
         End
         Begin VB.Menu AddUser 
            Caption         =   "≈÷«›… „” Œœ„..."
         End
         Begin VB.Menu EditPw 
            Caption         =   " ⁄œÌ· ﬂ·„… «·„—Ê—..."
         End
         Begin VB.Menu UserAbility 
            Caption         =   "’·«ÕÌ«  «·„” Œœ„Ì‰"
         End
         Begin VB.Menu UserRpt 
            Caption         =   " ﬁ«—Ì— «·„” Œœ„Ì‰"
         End
      End
      Begin VB.Menu ScreenSetting 
         Caption         =   "«⁄œ«œ«  «·‘«‘« "
         Begin VB.Menu MnuLevels 
            Caption         =   "«⁄ „«œ «·œÊ—… «·„” ‰œÌ…"
            Index           =   0
            Begin VB.Menu MnuLevelsSub 
               Caption         =   " ⁄—Ì› „” ÊÌ«  «·«⁄ „«œ"
               Index           =   0
            End
            Begin VB.Menu MnuLevelsSub 
               Caption         =   " ⁄—Ì› «⁄ „«œ«  «·„” œ« "
               Index           =   1
            End
         End
         Begin VB.Menu MnuLevels 
            Caption         =   "„Õœœ«  «·‘«‘« "
            Index           =   1
            Begin VB.Menu MnuLevelsSub2 
               Caption         =   " ⁄—Ì› „Õœœ«  «·‘«‘« "
               Index           =   0
            End
            Begin VB.Menu MnuLevelsSub2 
               Caption         =   "«⁄œ«œ  «·‘«‘« "
               Index           =   1
            End
         End
      End
      Begin VB.Menu ShortCuts 
         Caption         =   "„›« ÌÕ «·«Œ ’«—"
      End
      Begin VB.Menu MnuToolsSetPrinters 
         Caption         =   "«⁄œ«œ œ·Ì· «·Õ”«»« "
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu MnuToolsSetPrinters 
         Caption         =   "«‰Ê«⁄ «·”‰œ« "
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu MnuToolsSetPrinters 
         Caption         =   "«·«ÿ·«⁄ ⁄·Ï «· ‰»ÌÂ« "
         Index           =   3
      End
      Begin VB.Menu MnuToolsSetPrinters 
         Caption         =   " ﬂÊÌœ «·”‰œ« "
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu MnuToolsSetPrinters 
         Caption         =   " ﬂÊÌœ «·ÕﬁÊ·"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu MnuToolsSetPrinters 
         Caption         =   "«·—”«∆· «·œ«Œ·Ì…"
         Index           =   6
      End
      Begin VB.Menu MnuToolsSetPrinters 
         Caption         =   "«·ﬁ«„Ê”"
         Index           =   7
      End
   End
   Begin VB.Menu Basicdata 
      Caption         =   "«·»Ì«‰«  «·√”«”Ì…"
      Begin VB.Menu BasicDataM 
         Caption         =   "«‰Ê«⁄ «·„’—Ê›« "
         Index           =   0
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "«‰Ê«⁄ «·«Ì—«œ« "
         Index           =   1
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·»‰Êﬂ   "
         Index           =   2
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·Œ“‰ Ê  «·⁄Âœ"
         Index           =   3
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  ÿ—ﬁ «·œ›⁄ »«·‘»ﬂÂ"
         Index           =   4
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·„Ê—œÌ‰"
         Index           =   5
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·⁄„·«¡"
         Index           =   6
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·„ÊŸ›Ì‰"
         Index           =   7
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·«’‰«›"
         Index           =   8
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·⁄„·« "
         Index           =   9
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "«·Ã‰”Ì« "
         Index           =   10
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "«·œÌ«‰« "
         Index           =   11
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·œÊ·"
         Index           =   12
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·„Õ«›Ÿ«  Ê«·„‰«ÿﬁ"
         Index           =   13
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·«ÕÌ«¡"
         Index           =   14
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·‘Ê«—⁄"
         Index           =   15
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "«·„‘«—Ì⁄"
         Index           =   16
      End
      Begin VB.Menu BasicDataM 
         Caption         =   " ﬁ«—Ì—"
         Index           =   17
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·«’‰«›"
         Index           =   18
         Visible         =   0   'False
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "„·› «·„ÊŸ›Ì‰"
         Index           =   19
         Visible         =   0   'False
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "Œ—ÊÃ"
         Index           =   20
      End
   End
   Begin VB.Menu tech 
      Caption         =   "«·«œÊ«  «·›‰Ì…"
      Begin VB.Menu MnuToolsSetPrinters0 
         Caption         =   "«·œ⁄„ «·›‰Ì"
         Index           =   0
         Begin VB.Menu MnuToolsSetPrinters0sub 
            Caption         =   "ÿ·» œ⁄„ ›‰Ì"
            Index           =   0
         End
         Begin VB.Menu MnuToolsSetPrinters0sub 
            Caption         =   "„ «»⁄Â «·ﬂ«„Ì—« "
            Index           =   1
         End
         Begin VB.Menu MnuToolsSetPrinters0sub 
            Caption         =   "œ⁄„ ›‰Ì „ Œ’’"
            Index           =   2
         End
         Begin VB.Menu MnuToolsSetPrinters0sub 
            Caption         =   "«·«ﬁ›«·"
            Index           =   3
         End
         Begin VB.Menu MnuToolsSetPrinters0sub 
            Caption         =   "„“«„‰Â «·„«ﬂÌ‰« "
            Index           =   4
         End
         Begin VB.Menu MnuToolsSetPrinters0sub 
            Caption         =   "«·≈”‰«œ"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu MnuToolsSetPrinters0sub 
            Caption         =   "„Êﬁ› «·“Ì«—« "
            Index           =   6
         End
         Begin VB.Menu MnuToolsSetPrinters0sub 
            Caption         =   "„ÊÀﬁ «· ÃÂÌ“"
            Index           =   7
         End
         Begin VB.Menu MnuToolsSetPrinters0sub 
            Caption         =   "«⁄«œ… «Õ ”«» «· ﬂ·›…"
            Index           =   8
         End
         Begin VB.Menu MnuToolsSetPrinters0sub 
            Caption         =   "«⁄«œ…  ﬂ·›… ”‰œ«  «·’—›"
            Index           =   9
         End
         Begin VB.Menu MnuToolsSetPrinters0sub 
            Caption         =   "«·« ’«· ⁄‰ »⁄œ"
            Index           =   10
         End
      End
      Begin VB.Menu MnuToolsSetPrinters0 
         Caption         =   "≈⁄œ«œ «·ÿ«»⁄… ›Ï «·ÃÂ«“ «·Õ«·Ì"
         Index           =   1
      End
      Begin VB.Menu Barcode 
         Caption         =   " ’„Ì„ «·»«—ﬂÊœ"
         Shortcut        =   ^W
      End
      Begin VB.Menu MnuPrintItemsCodes 
         Caption         =   "ÿ»«⁄… »«—ﬂÊœ  ·√ﬂÊ«œ «·√’‰«›"
      End
      Begin VB.Menu MnuToolsSetPrinters7 
         Caption         =   " ≈⁄œ«œ«  —”«∆· «·ÃÊ«· Ê «·«Ì„Ì·« "
         Begin VB.Menu Texh 
            Caption         =   " ≈⁄œ«œ«  ›‰Ì… ··—”«∆·   «·‰’Ì…  Ê«·«Ì„Ì·« "
            Index           =   0
         End
         Begin VB.Menu Texh 
            Caption         =   "‰„«–Ã «·—”«∆·"
            Index           =   1
         End
         Begin VB.Menu Texh 
            Caption         =   " ⁄—Ì› «·—”«∆· ··‘«‘« "
            Index           =   2
         End
         Begin VB.Menu Texh 
            Caption         =   "—”«∆· «·⁄„·«¡ "
            Index           =   3
         End
         Begin VB.Menu Texh 
            Caption         =   "«⁄œ«œ«  «·«Ì„Ì·« "
            Index           =   4
            Visible         =   0   'False
         End
      End
      Begin VB.Menu MnuToolCustomers 
         Caption         =   "Ÿ»ÿ ›Ê« Ì— «·⁄„·«¡"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuToolsDataBase 
         Caption         =   " ‰‘Ìÿ «·√ ’«· »ﬁ«⁄œ… «·»Ì«‰« "
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu MnuToolsDataBase 
         Caption         =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Index           =   1
      End
      Begin VB.Menu MnuDataBaseTools 
         Caption         =   "√œÊ«  ﬁ«⁄œ… «·»Ì«‰« "
      End
   End
   Begin VB.Menu LIFEINDICATORMNU 
      Caption         =   "«·„ƒ‘—«  «·ÕÌ…"
   End
   Begin VB.Menu Help 
      Caption         =   "„”«⁄œ…"
      Begin VB.Menu HelpFileSub 
         Caption         =   "„·›«  «·„”«⁄œ…"
         Index           =   0
      End
      Begin VB.Menu HelpFileSub 
         Caption         =   "›Â—” „·›«  «·„”«⁄œ…"
         Index           =   1
      End
      Begin VB.Menu HelpFileSub 
         Caption         =   "«·»ÕÀ ›Ì „·›«  «·„”«⁄œ…"
         Index           =   2
      End
      Begin VB.Menu HelpFileSub 
         Caption         =   "«· ·„ÌÕ «·ÌÊ„Ì"
         Index           =   3
      End
      Begin VB.Menu HelpFileSub 
         Caption         =   "⁄‰ «·»—‰«„Ã..."
         Index           =   4
      End
      Begin VB.Menu HelpFileSub 
         Caption         =   " ”ÃÌ· «·»—‰«„Ã..."
         Index           =   5
      End
      Begin VB.Menu HelpFileSub 
         Caption         =   "ﬁ«∆„… «·„Â«„"
         Index           =   6
      End
      Begin VB.Menu HelpFileSub 
         Caption         =   "« ’· »‰«"
         Index           =   7
      End
   End
   Begin VB.Menu MnuInterface 
      Caption         =   "Ê«ÃÂ… «·»—‰«„Ã"
      Begin VB.Menu MnuInterfaceSub 
         Caption         =   "Ê«ÃÂ… ⁄—»ÌÌ…"
         Index           =   0
      End
      Begin VB.Menu MnuInterfaceSub 
         Caption         =   "Ê«ÃÂ… «‰Ã·Ì“Ì…"
         Index           =   1
      End
   End
   Begin VB.Menu FavoritesMenue 
      Caption         =   "«·ﬁ«∆„… «·„›÷·…"
      Begin VB.Menu help_list 
         Caption         =   " ⁄œÌ· «·ﬁ«∆„…"
         Index           =   0
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   11
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   12
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   13
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   14
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   15
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   16
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   17
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   18
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   19
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   20
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   21
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   22
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   23
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   24
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   25
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   26
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   27
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   28
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   29
      End
      Begin VB.Menu help_list 
         Caption         =   ""
         Index           =   30
      End
   End
   Begin VB.Menu PriceListPop 
      Caption         =   ""
      Enabled         =   0   'False
      Begin VB.Menu ShowItems 
         Caption         =   "≈” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰›"
      End
      Begin VB.Menu ItemsPrice 
         Caption         =   "√”⁄«— «·√’‰«›"
      End
   End
   Begin VB.Menu SupList 
      Caption         =   ""
      Enabled         =   0   'False
      Begin VB.Menu AddItem 
         Caption         =   "≈÷«›… ’‰›"
      End
      Begin VB.Menu DelItem 
         Caption         =   "Õ–› ’‰›"
      End
   End
   Begin VB.Menu MdiContextMenu1 
      Caption         =   ""
      Enabled         =   0   'False
      Begin VB.Menu PopPriceList 
         Caption         =   "ﬁ«∆„… «·√”⁄«— "
      End
      Begin VB.Menu PopSallBill 
         Caption         =   "›« Ê—… »Ì⁄"
      End
      Begin VB.Menu PopPurchaseBill 
         Caption         =   "›« Ê—… ‘—«¡"
      End
      Begin VB.Menu PopReturn 
         Caption         =   "„— Ã⁄ «·„‘ —Ì« "
      End
      Begin VB.Menu PopMaintanence 
         Caption         =   "’Ì«‰…"
      End
      Begin VB.Menu PopBalance 
         Caption         =   "«·—’Ìœ «·«›  «ÕÌ"
      End
      Begin VB.Menu PopGard 
         Caption         =   "Ã—œ «·„Œ«“‰"
      End
      Begin VB.Menu PopAvailable 
         Caption         =   "«·√ÃÂ“… «·„ «Õ…"
      End
      Begin VB.Menu PopSerialData 
         Caption         =   "»ÕÀ ⁄‰ »Ì«‰«  ”Ì—Ì«·"
      End
      Begin VB.Menu PpBarcode 
         Caption         =   " ’„Ì„ «·»«—ﬂÊœ"
      End
   End
   Begin VB.Menu MnuPops 
      Caption         =   ""
      Enabled         =   0   'False
      Begin VB.Menu MnuOutBarOptions 
         Caption         =   "ŒÌ«—«  ‘—Ìÿ «·√Œ ’«—« "
         Begin VB.Menu MnuOutBarItemsStyle 
            Caption         =   "⁄—÷ √”„«¡ «·√Œ ’«—« "
            Begin VB.Menu MnuOutBarStyle 
               Caption         =   "⁄—÷ «·√”„«¡ ›Ï «·Ã‰»"
               Index           =   0
            End
            Begin VB.Menu MnuOutBarStyle 
               Caption         =   "⁄—÷ «·√”„«¡ ›Ï «·√”›·"
               Index           =   1
            End
         End
         Begin VB.Menu MnuOutBarGroup 
            Caption         =   "≈÷«›… „Ã„Ê⁄… ÃœÌœ…"
            Index           =   0
         End
         Begin VB.Menu MnuOutBarGroup 
            Caption         =   " ⁄œÌ· «”„ «·„Ã„Ê⁄…"
            Index           =   1
         End
         Begin VB.Menu MnuOutBarGroup 
            Caption         =   "Õ–› «·„Ã„Ê⁄…"
            Index           =   2
         End
         Begin VB.Menu MnuOutBarGroup 
            Caption         =   "≈÷«›… ≈Œ ’«— ›Ï «·„Ã„Ê⁄…"
            Index           =   3
         End
         Begin VB.Menu MnuOutBarGroup 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu MnuOutBarGroup 
            Caption         =   " ⁄œÌ· «”„ «·√Œ ’«— "
            Index           =   5
         End
         Begin VB.Menu MnuOutBarGroup 
            Caption         =   "Õ–› «·√Œ ’«—  «·„Õœœ"
            Index           =   6
         End
         Begin VB.Menu MnuInvSales_Mnu6 
            Caption         =   ""
         End
         Begin VB.Menu MnuInvSales_Mnu7 
            Caption         =   ""
         End
         Begin VB.Menu MnuInvViewList 
            Caption         =   "⁄—÷ ﬁ«∆„… ..."
         End
         Begin VB.Menu MnuInvInsertTemp 
            Caption         =   " ≈œ—«Ã ⁄—÷ Ã«Â“..."
         End
         Begin VB.Menu MnuInvSales_Mnu1 
            Caption         =   "ﬂ‘› Õ”«» ⁄„Ì· «·›« Ê—…"
         End
         Begin VB.Menu MnuInvSales_Refresh 
            Caption         =   " ÕœÌÀ «·»Ì«‰« "
         End
         Begin VB.Menu MnuPopPane 
            Caption         =   "«·„⁄«„·«  «·„«·Ì…"
         End
      End
      Begin VB.Menu MnuInvPurchase 
         Caption         =   "ﬁ«∆„… ›« Ê—… «·‘—«¡"
         Begin VB.Menu MnuInvPurchaseMnu1 
            Caption         =   ""
         End
         Begin VB.Menu MnuInvPurchaseMnu2 
            Caption         =   ""
         End
         Begin VB.Menu MnuInvPurchaseMnu3 
            Caption         =   ""
         End
         Begin VB.Menu MnuInvPurchaseMnu4 
            Caption         =   ""
         End
      End
      Begin VB.Menu MnuManTools 
         Caption         =   "ﬁ«∆„… √œÊ«  «·’Ì«‰…"
         Begin VB.Menu MnuManToolsSub5 
            Caption         =   "ﬂ «»…  ﬁ—Ì— „ «»⁄… «·’Ì«‰…"
         End
      End
      Begin VB.Menu MnuManTools2 
         Caption         =   "ﬁ«∆„… √œÊ«  «·’Ì«‰…"
         Begin VB.Menu MnuManTools2Sub1 
            Caption         =   " „ «· Ã„Ì⁄"
         End
         Begin VB.Menu MnuManTools2Sub2 
            Caption         =   " ”·Ì„ «·ÃÂ«“"
         End
      End
      Begin VB.Menu MnuCusTools 
         Caption         =   "ﬁ«∆„… «·⁄„Ì·"
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   " ﬁ—Ì— ﬂ‘› Õ”«»"
            Index           =   0
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "›Ê« Ì— „»Ì⁄«  «·⁄„Ì·"
            Index           =   2
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "›Ê« Ì— „— Ã⁄«  «·⁄„Ì·"
            Index           =   3
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "›Ê« Ì— „‘ —Ì«  «·⁄„Ì· («·„Ê—œ)"
            Index           =   5
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "›Ê« Ì— „— Ã⁄ „‘ —Ì«  «·⁄„Ì· («·„Ê—œ)"
            Index           =   6
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "«·ﬁÌ„ «·„«·Ì… «·√Ã·… ··⁄„Ì·"
            Index           =   8
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "«·ﬁÌ„ «·„«·Ì… «·√Ã·… ⁄·Ï «·⁄„Ì·"
            Index           =   9
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "«·„ﬁ»Ê÷«  «· Ï Õ’·  „‰ «·⁄„Ì·"
            Index           =   11
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "«·„œ›Ê⁄«  «· Ï ”œœ  ≈·Ï «·⁄„Ì·"
            Index           =   12
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "-"
            Index           =   13
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "√ﬁ”«ÿ „” Õﬁ… ⁄·Ï «·⁄„Ì·"
            Index           =   14
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "√ﬁ”«ÿ „” Õﬁ… ··⁄„Ì·"
            Index           =   15
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "-"
            Index           =   16
         End
         Begin VB.Menu MnuCusTools_Item 
            Caption         =   "⁄—÷ »Ì«‰«  «·⁄„Ì· ( «·„Ê—œ)"
            Index           =   17
         End
      End
      Begin VB.Menu MnuItemTools 
         Caption         =   "ﬁ«∆„… «·’‰›"
         Begin VB.Menu MnuItemTools_ItemCart 
            Caption         =   "⁄—÷  ﬁ—Ì— ﬂ«—  «·’‰›"
         End
         Begin VB.Menu MnuItemTools_ItemQty 
            Caption         =   "≈” ⁄·«„ ⁄‰ ﬂ„Ì… «·’‰›"
         End
         Begin VB.Menu MnuItemTools_ItemSerial 
            Caption         =   "≈” ⁄·«„ ⁄‰ ”Ì—Ì«· «·’‰›"
         End
         Begin VB.Menu MnuItemTools_ItemCostTrans 
            Caption         =   "⁄—÷ „ Ê”ÿ  ﬂ·›… «·’‰›"
            Visible         =   0   'False
         End
         Begin VB.Menu MnuItemTools_Sep 
            Caption         =   "«” ⁄·«„ ⁄‰ «·«’‰«› «·»œÌ·…"
         End
         Begin VB.Menu MnuItemTools_Reports 
            Caption         =   ""
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu MnuItemTools_Reports 
            Caption         =   ""
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu MnuItemTools_Reports 
            Caption         =   "-"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu MnuItemTools_Reports 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu MnuItemTools_Reports 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu MnuItemTools_ItemData 
            Caption         =   "»Ì«‰«  «·’‰› ›Ï ‘«”… «·√’‰«›"
         End
         Begin VB.Menu MnuPopItemsTreePane_Array 
            Caption         =   ""
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu MnuPopItemsTreePane_Array 
            Caption         =   "-"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu MnuPopItemsTreePane_Array 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu MnuPopItemsTreePane_Array 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu MnuPopItemsTreePane_Array 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
      End
   End
End
Attribute VB_Name = "mdifrmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit
 
Private Declare Function sndPlay _
                Lib "winmm.dll" _
                Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
                                       ByVal uFlags As Long) As Long

Private Const SND_ASYNC = &H1

Dim messengerTime As Integer
Dim AlarmAutoTime As Integer


Private Const SND_SYNC = &H0

Private Const SND_LOOP = &H8

Private Const SND_NODEFAULT = &H1

Private Const SND_VALID = &H1F

Private Const SND_MEMORY = &H4

Private Const SND_PURGE = &H40

Dim formx As Integer
Dim formy As Integer
Const ID_THEME_OFFICE2000 = 140
Const ID_THEME_OFFICE2003 = 141
Const ID_THEME_NATIVE = 142
Const ID_THEME_OFFICE2000_PLAIN = 143
Const ID_THEME_OFFICEXP_PLAIN = 144
Const ID_THEME_OFFICE2003_PLAIN = 145
Const ID_THEME_NATIVE_PLAIN = 146

Const ID_TASKITEM_HIDECONTENTS = 1
Const ID_TASKITEM_ADDORREMOVE = 2
Const ID_TASKITEM_SEARCH = 3
Const ID_TASKITEM_NEWFOLDER = 4
Const ID_TASKITEM_PUBLISH = 5
Const ID_TASKITEM_SHARE = 6
Const ID_TASKITEM_MYCOMPUTER = 7
Const ID_TASKITEM_MYDOCUMENTS = 8
Const ID_TASKITEM_SHAREDDOCUMENTS = 9
Const ID_TASKITEM_MYNETWORKPLACES = 10

Const FCONTROL = 8

Private Type PaneRecorde
    PaneID As Integer
    PaneTitle As String * 50
    PanePositon As Integer
    PaneCx As Single
    PaneCy As Single
    PaneClosed As Boolean
    PaneEnabled As Boolean
    PaneFloated As Boolean
    PaneHidden As Boolean
    PaneLeft As Single
    PaneTop As Single
    PaneWidth As Single
    PaneHeight As Single
End Type

Private Sub About_Click()

End Sub

Private Sub AddItem_Click()
'    FrmMainPriceList.XPBtnAdd_Click
End Sub

Private Sub AddUser_Click()
    Dim Msg As String

    If user_id <> 1 And SystemOptions.PlaywithAuthorityMatrix = False Then
        '   MsgBox ""
        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
        '    Msg = Msg & Chr(13) & "Õ ‰Â“— Ê·««ÌÂ "
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

 

    If checkApility("FrmEditUsers") = False Then
        Exit Sub
    End If

   ' FrmAddUser.show vbModal
    
    FrmEditUsers.show
    
End Sub

Private Sub Asset_Click(Index As Integer)
End Sub

Private Sub advancedPayment_Click(Index As Integer)
Select Case Index
Case 0
     If checkApility("FrmPripaidExpenses") = False Then
                Exit Sub
            End If
'FrmExpensesAdvanced.show

FrmPripaidExpenses.show
Case 1
     If checkApility("FrmProofExpenses") = False Then
               Exit Sub
            End If
' FrmExpensespaidAdvanced.show
 FrmProofExpenses.show
 
 Case 2
      If checkApility("FrmPaytAmortization") = False Then
               Exit Sub
            End If
 FrmPaytAmortization.show
 
  Case 3
      If checkApility("FrmAproveComponYear") = False Then
               Exit Sub
            End If
 FrmAproveComponYear.show
 
End Select
End Sub

Private Sub advanceMenu_Click(Index As Integer)
Select Case Index
Case 0

    If checkApility("FrmEmpsAdvanceRequest") = False Then
        Exit Sub
    End If

FrmEmpsAdvanceRequest.show

Case 1

    If checkApility("FrmEmpsAdvance") = False Then
                Exit Sub
            End If

            FrmEmpsAdvance.show
            FrmEmpsAdvance.ZOrder 0

Case 2
  If checkApility("FrmEmpsAdvancePayed1") = False Then
                Exit Sub
            End If

            FrmEmpsAdvancePayed1.show

End Select
End Sub

Private Sub AgeingSub_Click(Index As Integer)
Select Case Index
Case 0
            If checkApility("Ageng") = False Then
                Exit Sub
            End If

            Ageng.show

Case 1
            If checkApility("Ageng1") = False Then
                Exit Sub
            End If

            Ageng.show

Case 2
           If checkApility("FrmOldContract") = False Then
                Exit Sub
            End If
 Unload FrmOldContract
 
 FrmOldContract.ScrenFlg = 1
 FrmOldContract.show
 
 Case 3
           If checkApility("FrmOldContract") = False Then
                Exit Sub
            End If
            
 Unload FrmOldContract
FrmOldContract.ScrenFlg = 0
FrmOldContract.show


Case 4
        If checkApility("ClientsInv") = False Then
                Exit Sub
            End If
ClientsInv.show

Case 5

FrmVizitScreen.mIndex = 8
FrmVizitScreen.show
Exit Sub
            If checkApility("Ageng_all1") = False Then
                Exit Sub
            End If
               Unload Ageng_all
Ageng_all.Indx = 0
            Ageng_all.show
            


End Select
End Sub

Private Sub ArchivingSub_Click(Index As Integer)

    Select Case Index

        Case 0
        
        
                    If checkApility("FrmEmpDepartments") = False Then
                Exit Sub
            End If
            
            FrmEmpDepartments.show
            
     Case 1
     
              If checkApility("FrmBasicDataINvArch") = False Then
                Exit Sub
            End If
            
     FrmBasicDataINvArch.Indx = 0
     FrmBasicDataINvArch.show
     
     Case 2
                If checkApility("FrmBasicDataINvArch") = False Then
                Exit Sub
            End If
            
     FrmBasicDataINvArch.Indx = 1
     FrmBasicDataINvArch.show
     
     Case 3
             If checkApility("FrmBasicDataINvArch") = False Then
                Exit Sub
            End If
            
     FrmBasicDataINvArch.Indx = 2
     FrmBasicDataINvArch.show
     
          Case 4
                     If checkApility("FrmBasicDataINvArch") = False Then
                Exit Sub
            End If
            
     FrmBasicDataINvArch.Indx = 3
     FrmBasicDataINvArch.show
          Case 5
                     If checkApility("FrmBasicDataINvArch") = False Then
                Exit Sub
            End If
            
     FrmBasicDataINvArch.Indx = 4
     FrmBasicDataINvArch.show
     Case 6
                If checkApility("loading_temolates") = False Then
                Exit Sub
            End If
            
            loading_temolates.show
'                 If Dir(App.path & "\checklist\Checklist.exe") <> "" Then
'         Shell App.path & "\Archive\Archive.exe", vbNormalFocus
'     End If
Case 7
                If checkApility("FrmTransacRegistr") = False Then
                Exit Sub
            End If
            FrmTransacRegistr.show

Case 8
                If checkApility("FrmTransacRegAlarm") = False Then
                Exit Sub
            End If
            FrmTransacRegAlarm.show
Case 9
             If checkApility("FrmTransacRegAlarm") = False Then
                Exit Sub
            End If
FrmTransacRegAlarm.show

Case 10
          If checkApility("FrmArchReports") = False Then
                Exit Sub
            End If
FrmArchReports.show
    End Select

End Sub

Private Sub ArrowsFollow_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("ArrowsFinancialMarkets") = False Then
                Exit Sub
            End If

     '       ArrowsFinancialMarkets.show

        Case 1

            If checkApility("ArrowsGroup") = False Then
                Exit Sub
            End If

     '       ArrowsGroup.show

        Case 2

            If checkApility("ArrowsAllCompanyilstDetails1") = False Then
                Exit Sub
            End If

     '       ArrowsAllCompanyilstDetails1.show

        Case 3

            If checkApility("Arrows") = False Then
                Exit Sub
            End If

     '       Arrows.show

        Case 4

            If checkApility("ArrowsHistory") = False Then
                Exit Sub
            End If

     '       ArrowsHistory.show
            'ArrowsAllCompanyilstDetails.Show

    End Select

End Sub

Private Sub ArrowsFollowa_Click(Index As Integer)

    Select Case Index

        Case 0
    '        ArrowsAccount.show
    End Select

End Sub

Private Sub ArrowsFollowBocket_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("ArrowsAccount") = False Then
                Exit Sub
            End If

    '        ArrowsAccount.show

        Case 1

            If checkApility("ArrowsPurchase") = False Then
                Exit Sub
            End If

    '        ArrowsPurchase.show

        Case 2

            'ArrowsSale.Show
            If checkApility("ArrowsSale1") = False Then
                Exit Sub
            End If

    '        ArrowsSale1.show

        Case 3

            If checkApility("ArrowsCurrentValue") = False Then
                Exit Sub
            End If

    '        ArrowsCurrentValue.show
    End Select

End Sub

Private Sub AssetsMng_Click(Index As Integer)

    Select Case Index

        Case 4

         '   If checkApility("messages_frm") = False Then
         '       Exit Sub
         '   End If
'
'            messages_frm.show


            If checkApility("FrmCustomerBalances1") = False Then
                Exit Sub
            End If

            FrmCustomerBalances1.show
            
        Case 5
              If checkApility("FrmSerachUnitEmpty") = False Then
                Exit Sub
            End If

            FrmSerachUnitEmpty.show
            
        
    End Select

End Sub

Private Sub AssetsMngBasicFiles_Click(Index As Integer)

    Select Case Index
Case 0

        If checkApility("FrmCustomerType") = False Then
                Exit Sub
            End If
FrmCustomerType.Indx = 0
            FrmCustomerType.show
            
      Case 1
      
            If checkApility("FrmAkarType") = False Then
                Exit Sub
            End If

            FrmAkarType.show
      Case 2
            If checkApility("FrmAkarUnit") = False Then
                Exit Sub
            End If

            FrmAkarUnit.show
      
      Case 3
      

            If checkApility("FrmSalesRePGroups") = False Then
                Exit Sub
            End If

            FrmSalesRePGroups.show

        Case 4

'            If checkApility("FrmSalesRepData") = False Then
'                Exit Sub
'            End If

'            FrmSalesRepData.show
    If checkApility("FrmPay_Garanty_Shipment") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment.SendForm = 7
FrmPay_Garanty_Shipment.show
            
            
  

        Case 5

            If checkApility("FrmCountriesData") = False Then
                Exit Sub
            End If

            FrmCountriesData.show

        Case 6

            If checkApility("FrmGovernmentData") = False Then
                Exit Sub
            End If

            FrmGovernmentData.show

        Case 7

            If checkApility("FrmGovernCitiesData") = False Then
                Exit Sub
            End If

            FrmGovernCitiesData.show
 
        Case 8 '„Œÿÿ« 

            If checkApility("frmCustomerType") = False Then
                Exit Sub
            End If
FrmCustomerType.Indx = 1
        FrmCustomerType.show
        
 
        Case 9

            If checkApility("RSPhoneBook") = False Then
                Exit Sub
            End If

            RSPhoneBook.show
            
           Case 10

            If checkApility("RSOwner") = False Then
                Exit Sub
            End If

            RSOwner.show

        Case 11

            If checkApility("RsCustomers") = False Then
                Exit Sub
            End If

            RsCustomers.show

        Case 12

   
         If checkApility("FrmExpensesType") = False Then
                Exit Sub
            End If

            OpenScreen ExpensesTypes

            
             Case 13

            If checkApility("FrmAkarStatus") = False Then
                Exit Sub
            End If

            FrmAkarStatus.show
           Case 14
              If checkApility("FrmIqarCompnent") = False Then
                Exit Sub
            End If

            FrmIqarCompnent.show
            
            
       Case 15

            If checkApility("RSAkar") = False Then
                Exit Sub
            End If

            RSAkar.show
            
     Case 16
     
            If checkApility("FrmAlarMType") = False Then
                Exit Sub
            End If

            FrmAlarMType.show
            
    End Select

End Sub

Private Sub AssetsMngBasicFiles2_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("RSAkar") = False Then
                Exit Sub
            End If

            RSAkar.show

        Case 1

            If checkApility("RsApartement") = False Then
                Exit Sub
            End If

          '  RsApartement.show

        Case 2

            If checkApility("RsRoom") = False Then
                Exit Sub
            End If

            'RsRoom.show

        Case 3

            If checkApility("RsStore") = False Then
                Exit Sub
            End If

          '  RsStore.show

    End Select

End Sub

Private Sub AssetsMngBasicFilesR_Click(Index As Integer)

    Select Case Index

        Case 1
            'RsVila.show

        Case 2
            'RSland.show

        Case 3
         '   RsStores.show

        Case 4
         '   RSWorkShop.show

        Case 5
     '       RSTradingCenter.show

    End Select

End Sub

Private Sub AssetsMngContrac_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("RSContract") = False Then
                Exit Sub
            End If

            RSContract.show

        Case 1

         '   If checkApility("RSContract1") = False Then
         '       Exit Sub
         '   End If
'
'            RSContract.show
    End Select

End Sub

Private Sub AssetsMngReport_Click(Index As Integer)
Select Case Index
Case 0
     If checkApility("FrmAqarReport") = False Then
                Exit Sub
            End If
FrmAqarReport.show
Case 1



''     If checkApility("FrmAqarReport1") = False Then
 '               Exit Sub
 '           End If
'FrmAqarReport1.show




    If checkApility("FrmAqarReport1") = False Then
                Exit Sub
            End If
FrmAmolatReports.show





Case 2
     If checkApility("FrmExpiredContract") = False Then
                Exit Sub
            End If
FrmExpiredContract.show
 Case 3
      If checkApility("FrmMaintnanceReport") = False Then
                Exit Sub
            End If
            
 FrmMaintnanceReport.show
 
 Case 4
       If checkApility("FrmWaiverReport") = False Then
                Exit Sub
            End If
            
  FrmWaiverReport.show
 Case 5
        If checkApility("FrmConttractTotalService") = False Then
                Exit Sub
            End If
            
 FrmConttractTotalService.show

Case 6
    If checkApility("FrmOwnerAqarReport") = False Then
                Exit Sub
            End If
FrmOwnerAqarReport.show

Case 7
    If checkApility("FrmAlrmReports") = False Then
                Exit Sub
            End If
FrmAlrmReports.show
 Case 8
     If checkApility("FrmTotalsReport") = False Then
                Exit Sub
            End If
 FrmTotalsReport.show
 
 
  Case 9
     If checkApility("FrmOrboon") = False Then
                Exit Sub
            End If
 FrmOrboon.show
 
 
  Case 10
     If checkApility("FrmCommissionReports") = False Then
                Exit Sub
            End If
 FrmCommissionReports.show
 Case 11
 
      If checkApility("FrmRentsOwendReports") = False Then
                Exit Sub
            End If
     FrmRentsOwendReports.show
     
 
 
  Case 12
 
      If checkApility("FrmCustomerAqarReport") = False Then
                Exit Sub
            End If
     FrmCustomerAqarReport.show
     
     
     Case 13
         If checkApility("FrmUnitInfoReports") = False Then
                Exit Sub
            End If
     FrmUnitInfoReports.show
     
      
     
 Case 14
         If checkApility("FrmIncomAndExpenReports") = False Then
                Exit Sub
            End If
     FrmIncomAndExpenReports.show
     
 Case 15
         If checkApility("FrmContractReport") = False Then
                Exit Sub
            End If
     FrmContractReport.show
          
          
     
End Select
End Sub

Private Sub AssetsMngsub_Click(Index As Integer)
Select Case Index
Case 0
        If checkApility("FrmAllocationToContract") = False Then
                Exit Sub
            End If
FrmAllocationToContract.show
Case 1
        If checkApility("FrmAllocationToContract1") = False Then
                Exit Sub
            End If
FrmAllocationToContract1.show

End Select
End Sub

Private Sub AssetsMngTrans_Click(Index As Integer)

    Select Case Index

        Case 0

        
             

        Case 5

 
           If checkApility("FrmCashing1") = False Then
                Exit Sub
            End If
            
FrmCashing1.show

        Case 6

            If checkApility("RsExpenses") = False Then
                Exit Sub
            End If

            RsExpenses.show
            
            
Case 7
      
            If checkApility("FrmPayments2") = False Then
                Exit Sub
            End If

     
            
            
    If SystemOptions.SpecialVersion = True Then
     FrmPayments.show
     Else
     FrmPayments2.show
   End If
   

   

Case 8
      If checkApility("PrintCheque") = False Then
                Exit Sub
            End If

            PrintCheque.show
        
        Case 9
 

      If checkApility("FrmWaiverSettlement") = False Then
                Exit Sub
            End If

            FrmWaiverSettlement.show

        Case 10
 

      If checkApility("FrmExpenses301") = False Then
                Exit Sub
            End If

            FrmExpenses301.show

     
        Case 11

       
  If checkApility("Frmblacklist") = False Then
             Exit Sub
        End If
'
'
frmblacklist.show
        Case 12

            If checkApility("FrmRsCustomerAlarm") = False Then
                Exit Sub
            End If

            FrmRsCustomerAlarm.show
            
            Case 13
          If checkApility("FrmOtheExpensAqar") = False Then
                Exit Sub
            End If
FrmOtheExpensAqar.show


            Case 15
          If checkApility("dean") = False Then
                Exit Sub
            End If
            dean.mIndex = 12
dean.show

    End Select

End Sub

Private Sub balancsheet_Click(Index As Integer)

    Select Case Index

        Case 0
           ' BaklanceSheet.show

        Case 1
            'BaklanceSheetvIEW.show
    End Select

    'FrmAccountingReport1.Show

End Sub

Private Sub BankAdM_Click()

End Sub

Private Sub BankOpSub_Click(Index As Integer)
Select Case Index
Case 0
            If checkApility("FrmBankDeposite") = False Then
                Exit Sub
            End If

            FrmBankDeposite.show

        Case 1

            If checkApility("FrmBankDeposite1") = False Then
                Exit Sub
            End If
FrmBankDeposite1.show
Case 2


            If checkApility("BankSettlementt") = False Then
                Exit Sub
            End If

            BankSettlementt.show

Case 3
           If checkApility("FrmBankAdj") = False Then
                Exit Sub
            End If

            FrmBankAdj.show
Case 4

            If checkApility("PrintCheque") = False Then
                Exit Sub
            End If

            PrintCheque.show


Case 5



         If checkApility("ReporBanks") = False Then
                Exit Sub
            End If

            FrmReports.show
            FrmReports.C1TabMain.CurrTab = 19



End Select

End Sub

Private Sub Barcode_Click()

    If checkApility("FrmBarcode") = False Then
        Exit Sub
    End If

    FrmBarcode.show
    FrmBarcode.ZOrder 0
    Exit Sub
ErrTrap:
End Sub

Private Sub case_Click()
 
End Sub

Private Sub BasicDataM_Click(Index As Integer)
'Debug.Print Mid(GetallChilddata(2), 2, Len(GetallChilddata(2)))
' FrmGoldDetaiks.show
 
    Select Case Index

        Case 0

            'FrmExpensesType
            If checkApility("FrmExpensesType") = False Then
                Exit Sub
            End If

            OpenScreen ExpensesTypes

        Case 1

            'FrmRevenuesTypes
            If checkApility("FrmRevenuesTypes") = False Then
                Exit Sub
            End If

            OpenScreen RevenuesTypes
            
            

        Case 2

            If checkApility("FrmBanksData") = False Then
                Exit Sub
            End If

            OpenScreen BanksDataScreen

        Case 3

            If checkApility("FrmBoxesData") = False Then
                Exit Sub
            End If

            OpenScreen BoxesDataScreen

        Case 4

            If checkApility("FrmPaymentType") = False Then
                Exit Sub
            End If
FrmPaymentType.mIndex = 0 'ÿ—ﬁ «·œ›⁄

            FrmPaymentType.show

        Case 5

            If checkApility("FrmCompany") = False Then
                Exit Sub
            End If
FrmCompany.mIndex = 0
            FrmCompany.show

        Case 6

            If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If

            OpenScreen CustomersScreen '






Case 7
Unload FrmEmployee
FrmEmployee.DriverOnly = 0
            'FrmEmployee
            If checkApility("FrmEmployee") = False Then
                Exit Sub
            End If

            OpenScreen EmployeesScreen
FrmEmployee.WorkShop_Job = 0
 
         Case 8

            If checkApility("FrmItems") = False Then
                Exit Sub
            End If
FrmItems.show
        Case 9

            If checkApility("FRMcurrency") = False Then
                Exit Sub
            End If
FRMcurrency.mIndex = 0
            FRMcurrency.show

        Case 10

            If checkApility("Nationality") = False Then
                Exit Sub
            End If

            Nationality.show

        Case 11

                   If checkApility("dean") = False Then
                Exit Sub
            End If
dean.mIndex = 2
dean.show
 
        Case 12

            If checkApility("FrmCountriesData") = False Then
                Exit Sub
            End If

            FrmCountriesData.show

        Case 13

            If checkApility("FrmGovernmentData") = False Then
                Exit Sub
            End If

            FrmGovernmentData.show

        Case 14

            If checkApility("FrmGovernCitiesData") = False Then
                Exit Sub
            End If

            FrmGovernCitiesData.show

        Case 15

            If checkApility("streets") = False Then
                Exit Sub
            End If

            streets.show
         Case 16
             If checkApility("Projects") = False Then
                Exit Sub
            End If
  Projects.show
         
         
        Case 17
            ' FrmDocType.Show


'            OpenScreen ItemsDataScreen
  If checkApility("FrmTotals2Report") = False Then
                Exit Sub
            End If

FrmTotals2Report.show


        Case 20
            AskForExit

    End Select

End Sub

Private Sub Beautysub_Click(Index As Integer)
Select Case Index
Case 0
    If checkApility("FrmGroups") = False Then
                Exit Sub
            End If

            'FrmGroups
            OpenScreen ItemsGroupsScreen
            
   Case 1
   'FrmItems
            If checkApility("FrmItems") = False Then
                Exit Sub
            End If
FrmItems.EleHeader.Caption = ""
            OpenScreen ItemsDataScreen
            
Case 2

Unload FrmEmployee

            'FrmEmployee
            If checkApility("FrmEmployee") = False Then
                Exit Sub
            End If

            OpenScreen EmployeesScreen
FrmEmployee.WorkShop_Job = 0

Case 3

                 If checkApility("frm_sheft") = False Then
                Exit Sub
            End If
frm_sheft.show

Case 4 '«‰Ê«⁄ «·—« Õ« 


        If checkApility("FrmItemsClass") = False Then
                Exit Sub
            End If
FrmItemsClass.mIndex = 4
FrmItemsClass.show



Case 5 '—»ÿ «·„ÊŸ›Ì‰
            If checkApility("dean") = False Then
                Exit Sub
            End If
dean.mIndex = 7
dean.show


Case 6
  
   If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If

            'FrmCustemers
            OpenScreen CustomersScreen '

Case 7  'RESERVE TYPE
            If checkApility("dean") = False Then
                Exit Sub
            End If
dean.mIndex = 5
dean.show

Case 8 'ŒÿÂ «·—«Õ« 

        If checkApility("FrmItemsClass") = False Then
                Exit Sub
            End If
FrmItemsClass.mIndex = 5
FrmItemsClass.show


Case 9 'RESERVATI
       If checkApility("FrmStudentCalling") = False Then
                Exit Sub
            End If
            
FrmStudentCalling.show

Case 10 ' SHOW RESER
            If checkApility("dean") = False Then
                Exit Sub
            End If
dean.mIndex = 6
dean.show
Case 11
   
 
            If checkApility("FrmPOSDATA") = False Then
                Exit Sub
            End If

            FrmPOSDATA.show


       Case 12

            If checkApility("cachierData") = False Then
                Exit Sub
            End If

            cachierData.show


 

 
        Case 13

            If checkApility("CashierLogin") = False Then
                Exit Sub
            End If
 
            CashierLogin.show
            'frmsalebill1.Show
 
 
  
            Case 14
         'FrmCashing
            If checkApility("FrmCashing") = False Then
                Exit Sub
            End If

            OpenScreen CashingDataScreen
            
 
 Case 15
 
              If checkApility("FrmBankDeposite3") = False Then
                Exit Sub
            End If

            FrmBankDeposite3.show

 
 
        Case 16

            If checkApility("ReportSales") = False Then
                Exit Sub
            End If

            FrmReports.show
            FrmReports.C1TabMain.CurrTab = 0
 'Case 9
 
 '           If checkApility("FrmAnalysItems") = False Then
 ''               Exit Sub
  '          End If
'
          '  FrmReports.show
          '  FrmReports.C1TabMain.CurrTab = 0
' FrmAnalysItems.show
 
 Case 17
          If checkApility("FrmCustCash") = False Then
                Exit Sub
            End If
      FrmCustCash.show
       

End Select
End Sub

Private Sub CarMaintenancesub_Click(Index As Integer)
Select Case Index

Case 2
    '  If checkApility("FrmCarReports") = False Then
    '            Exit Sub
    '        End If
'
'FrmCarReports.show

Load FrmCarReportsRequerNo
FrmCarReportsRequerNo.show

End Select
End Sub

Private Sub CarMaintenancesub1_Click(Index As Integer)
Select Case Index

Case 0
       If checkApility("CarTypes") = False Then
                Exit Sub
            End If

            CarTypes.show
            
Case 1
  If checkApility("FrmCarModels") = False Then
                Exit Sub
            End If

FrmCarModels.show


Case 2


            If checkApility("FrmCars") = False Then
              Exit Sub
           End If
            FrmCars.show
            
 Case 3
 
      If checkApility("FrmMaintenensWork") = False Then
              Exit Sub
           End If
 FrmMaintenensWork.show
 Case 4
       If checkApility("FrmTypeExtraExpenses") = False Then
              Exit Sub
           End If
 FrmTypeExtraExpenses.show
 
 Case 5
 
 
      If checkApility("FrmExtraExpenses") = False Then
              Exit Sub
           End If
 FrmExtraExpenses.show
 
 
 Case 6
       If checkApility("FrmComputerChek") = False Then
              Exit Sub
           End If
 FrmComputerChek.show
 Case 7
  If checkApility("FrmColor") = False Then
              Exit Sub
           End If
 FrmColor.show
 
 Case 8
      If checkApility("FrmStoreData") = False Then
                Exit Sub
            End If

            'FrmStoreData
            OpenScreen StoresDataScreen

 Case 9
  
          If checkApility("FrmGroups") = False Then
                Exit Sub
            End If

            'FrmGroups
            OpenScreen ItemsGroupsScreen

  Case 10
   '        If checkApility("FrmSystemUnites") = False Then
   '             Exit Sub
   '         End If
'
'            FrmSystemUnites.show

Unload FrmPay_Garanty_Shipment3M
             If checkApility("FrmPay_Garanty_Shipment3M") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment3M.SendForm = 0
FrmPay_Garanty_Shipment3M.show
  Case 11
 If checkApility("FrmItems") = False Then
                Exit Sub
            End If

            OpenScreen ItemsDataScreen
  
  
Case 12
  If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If

            'FrmCustemers
            OpenScreen CustomersScreen '
            
 Case 13
        If checkApility("FrmEmployee") = False Then
                Exit Sub
            End If

         '   OpenScreen EmployeesScreen
Unload FrmEmployee
 FrmEmployee.WorkShop_Job = 0
 FrmEmployee.show
 
' Case 13   '„Ã„Ê⁄«  «·Ê—‘…
' FrmSalesRePGroups3.show


'WorkShop_Job
Case 15   '««·„‘—›Ì‰
        If checkApility("FrmEmployee") = False Then
        
                Exit Sub
            End If
Unload FrmEmployee
FrmEmployee.WorkShop_Job = 1
FrmEmployee.show
FrmEmployee.EleHeader.Caption = "»Ì«‰«  «·„‘—›Ì‰"
    '        OpenScreen EmployeesScreen


'Unload FrmSalesRepData3
'Workshopgroupid = 1
'FrmSalesRepData3.show
'FrmSalesRepData3.Label1(2).Caption = "»Ì«‰«  «·„Â‰œ”Ì‰"
'FrmSalesRepData3.Caption = FrmSalesRepData3.Label1(2).Caption
'FrmSalesRepData3.DCSalesRepGroups.BoundText = Workshopgroupid
Case 16 '·›‰ÌÌ‰
    If checkApility("FrmEmployee") = False Then
                Exit Sub
            End If
            Unload FrmEmployee
FrmEmployee.WorkShop_Job = 2
FrmEmployee.show
FrmEmployee.EleHeader.Caption = "»Ì«‰«  «·„Â‰œ”‰ Ê «·› ÌÌ‰"
'Unload FrmSalesRepData3
'Workshopgroupid = 2
'FrmSalesRepData3.show
'FrmSalesRepData3.Label1(2).Caption = "»Ì«‰«  «·„‘—›Ì‰"

'FrmSalesRepData3.Caption = FrmSalesRepData3.Label1(2).Caption
'FrmSalesRepData3.DCSalesRepGroups.BoundText = Workshopgroupid
 'Case 17 '›‰ÌÌ‰
' Unload FrmSalesRepData3
' Workshopgroupid = 3
'FrmSalesRepData3.show
'FrmSalesRepData3.Label1(2).Caption = "»Ì«‰«  «·›‰ÌÌ‰"
'FrmSalesRepData3.Caption = FrmSalesRepData3.Label1(2).Caption
'FrmSalesRepData3.DCSalesRepGroups.BoundText = Workshopgroupid
'

 
Case 17
             If checkApility("FrmcarEmpDepartments") = False Then
                Exit Sub
            End If

       '     FrmEmpDepartments.show
            
            FrmcarEmpDepartments.show
Case 18
    If checkApility("FrmSuperVisor") = False Then
                Exit Sub
            End If
          FrmSuperVisor.show
 
End Select

End Sub
'
Private Sub CarMaintenancesub2_Click(Index As Integer)

Select Case Index
Case 0

      If checkApility("FrmCarAuthontication") = False Then
                Exit Sub
            End If
 FrmCarAuthontication.show
 
Case 1

      If checkApility("FrmBillComputerChek") = False Then
                Exit Sub
            End If
FrmBillComputerChek.show
Case 2

      If checkApility("FrmOut") = False Then
                Exit Sub
            End If
            
            FrmOut.show
            FrmOut.TxtTicketNO.Visible = True
            FrmOut.lbl(32).Visible = True
              

Case 3
       '     GeneralPriceType = 1

            If checkApility("FrmPO10") = False Then
                Exit Sub
            End If
FrmPO10.show

Case 4
      If checkApility("FrmBillCarMaintExtra") = False Then
                Exit Sub
            End If
   'FrmBillCarMaintExtra.show
  
 ' FrmManCusRecive.show
 
Load FrmBillCarMaintExtra
FrmBillCarMaintExtra.show

Case 5
      If checkApility("FrmCommisRece") = False Then
                Exit Sub
            End If
 FrmCommisRece.show


Case 6
      If checkApility("FrmVizitScreen") = False Then
                Exit Sub
            End If
            
 
FrmVizitScreen.mIndex = 1

 FrmVizitScreen.show
 
 Case 7
 FrmItemsClass.mIndex = 9
 FrmItemsClass.show
  
End Select

End Sub

Private Sub Cashing_Click(Index As Integer)

    Select Case Index

        Case 0

            'FrmCashing
            If checkApility("FrmCashing") = False Then
                Exit Sub
            End If

            OpenScreen CashingDataScreen

        Case 1 '”‰œ «·ﬁÌ÷ «·⁄„Ê„Ì
        
              If checkApility("FrmGeneralFundReceipt") = False Then
                Exit Sub
            End If

          FrmGeneralFundReceipt.show
'FrmBankDeposite2
            'projectsbill.Show
        Case 2

            If checkApility("PrintCheque") = False Then
                Exit Sub
            End If

            PrintCheque.show

        Case 3

            If checkApility("FrmBankDeposite") = False Then
                Exit Sub
            End If

            FrmBankDeposite.show

        Case 4

            If checkApility("FrmBankDeposite1") = False Then
                Exit Sub
            End If

            'FrmChiqueRelease.Show

            FrmBankDeposite1.show

        Case 5

            If checkApility("FrmBankAdj") = False Then
                Exit Sub
            End If

            FrmBankAdj.show


        Case 6


            If checkApility("BankSettlementt") = False Then
                Exit Sub
            End If

            BankSettlementt.show
            
            Case 7
          '  If checkApility("FrmADVPaymentsAlloc") = False Then
          '      Exit Sub
          '  End If

          '  FrmADVPaymentsAlloc.show
            
    End Select

End Sub

Private Sub ComingTimes_Click()
 

End Sub

Private Sub ConnectUs_Click()

End Sub
 
Private Sub CeramicEstimationsub_Click(Index As Integer)
Select Case Index
Case 0
            If checkApility("FrmProcessUnit") = False Then
                Exit Sub
            End If
            FrmProcessUnit.show

        Case 1
           If checkApility("FrmProcessDef") = False Then
                Exit Sub
            End If
         If SystemOptions.UserInterface = ArabicInterface Then
            FrmProcessDef.Caption = " ⁄—Ì› «·⁄„·Ì« "
         Else
         FrmProcessDef.Caption = "  Operations Define "
         End If
            FrmProcessDef.Ele(5).Caption = FrmProcessDef.Caption
            FrmProcessDef.show


 


Case 2
            If checkApility("Frm_NewMeasure") = False Then
                Exit Sub
            End If

 Frm_NewMeasure.show
 

Case 3
            If checkApility("Frm_TRansOrder") = False Then
                Exit Sub
            End If

 Frm_TRansOrder.show
 
Case 4
            If checkApility("Frm_TradingContract") = False Then
                Exit Sub
            End If

 Frm_TradingContract.show
  
 
 Case 5
 
          If checkApility("Projects") = False Then
                Exit Sub
            End If
 
       Projects.show
 
 
 Case 6 '«·«„ «— «·„‰Ã“Â ÌÊ„Ì«
 
           If checkApility("Frm_BusinessDialy") = False Then
                Exit Sub
            End If
   Frm_BusinessDialy.show
   
   Case 7
   
       If checkApility("emp_CONTRACT_TYPE") = False Then
                Exit Sub
            End If
            Unload emp_CONTRACT_TYPE
emp_CONTRACT_TYPE.mIndex = 3
emp_CONTRACT_TYPE.show
     
     
  Case 8 '«· ﬁ«—Ì—
 
           If checkApility("FrmReportsStudent") = False Then
                Exit Sub
            End If
             FrmReportsStudent.Indx = 1
   FrmReportsStudent.show
   
   
End Select
End Sub

Private Sub COLLECTIONSUB_Click(Index As Integer)
Select Case Index
                
                Case 0
                   If checkApility("FrmSalesRePGroups") = False Then
                                Exit Sub
                            End If
                
                            FrmSalesRePGroups.show
                
                Case 1
                
                             If checkApility("FrmPay_Garanty_Shipment") = False Then
                                 Exit Sub
                             End If
                FrmPay_Garanty_Shipment.SendForm = 7
                FrmPay_Garanty_Shipment.show
                
                
                Case 2
                FrmCustomerType.Indx = 0
                FrmCustomerType.show
                
                Case 3
                            If checkApility("RSPhoneBook") = False Then
                                Exit Sub
                            End If
                
                            RSPhoneBook.show
                Case 4
                       If checkApility("FrmStudentCalling") = False Then
                                Exit Sub
                            End If
                FrmStudentCalling.show
                
                Case 5
                        If checkApility("FrmCreditFacicity") = False Then
                                Exit Sub
                            End If
                
                     FrmCreditFacicity.show
                
                Case 6
                   If checkApility("FrmCustemers") = False Then
                                Exit Sub
                            End If
                
                            'FrmCustemers
                            OpenScreen CustomersScreen '
                Case 7
                    If checkApility("FrmRegDateDelgate") = False Then
                                Exit Sub
                            End If
                FrmRegDateDelgate.show
                
                Case 8
                       If checkApility("FrmShowRegDateDelegate") = False Then
                                Exit Sub
                            End If
                FrmShowRegDateDelegate.show
                
                Case 9
                            If checkApility("FrmCustomerssFollow") = False Then
                                Exit Sub
                            End If
                
                             FrmCustomerssFollow.show
                Case 10
                    If checkApility("FrmReceiptPart") = False Then
                        Exit Sub
                    End If
                
                    OpenScreen ReceiptPartScreen
                
                Case 11
                         If checkApility("FrmCustomerssComplaint") = False Then
                                Exit Sub
                            End If
                
                     FrmCustomerssComplaint.show
                
                Case 12
                            If checkApility("Ageng_all1") = False Then
                                Exit Sub
                            End If
                               Unload Ageng_all
                Ageng_all.Indx = 0
                            Ageng_all.show
                            
                Case 13
                            If checkApility("ReportPurchase") = False Then
                                Exit Sub
                            End If
                
                            FrmReports.show
                            FrmReports.C1TabMain.CurrTab = 10
                Case 14

           If checkApility("Ageng_all") = False Then
                Exit Sub
            End If
            Unload Ageng_all
            Ageng_all.Indx = 3
Ageng_all.show


Case 15
    If checkApility("FrmPaymentTime") = False Then
        Exit Sub
    End If

    FrmPaymentTime.show
    FrmPaymentTime.ZOrder 0
End Select
End Sub

Private Sub ContainerSub_Click(Index As Integer)
Select Case Index
        Case 0

            If checkApility("FixedAssetsGroup") = False Then
                Exit Sub
            End If

            FixedAssetsGroup.show

        Case 1

            If checkApility("FixedAssets") = False Then
                Exit Sub
            End If

            FixedAssets.show
            Case 2
                      If checkApility("FrmGovernCitiesData") = False Then
                Exit Sub
            End If

            FrmGovernCitiesData.show



     Case 3

            If checkApility("CarTypes") = False Then
                Exit Sub
            End If

            CarTypes.show
            
            
      Case 4 '«·ÿ—«“« 
  If checkApility("FrmCarModels") = False Then
                Exit Sub
            End If
            FrmCarModels.show
            
 
 
        Case 5

            If checkApility("FrmCars") = False Then
                Exit Sub
            End If
FrmCars.Caption = " «·‘«Õ‰«  "
FrmCars.Ele(0) = FrmCars.Caption
            FrmCars.show


Case 6
            If checkApility("FrmDrivers") = False Then
                Exit Sub
            End If

            FrmDrivers.show
            
Case 7


FrmCustomerType.Indx = 0
            FrmCustomerType.show


Case 8
   If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If

            'FrmCustemers
            OpenScreen CustomersScreen '


Case 9

    If checkApility("emp_CONTRACT_TYPE") = False Then
                Exit Sub
            End If
            Unload emp_CONTRACT_TYPE
emp_CONTRACT_TYPE.mIndex = 1
emp_CONTRACT_TYPE.show
     

Case 10
    If checkApility("emp_CONTRACT_TYPE") = False Then
                Exit Sub
            End If
            Unload emp_CONTRACT_TYPE
emp_CONTRACT_TYPE.mIndex = 2
emp_CONTRACT_TYPE.show

Case 11
    If checkApility("emp_CONTRACT_TYPE") = False Then
                Exit Sub
            End If
            Unload emp_CONTRACT_TYPE
emp_CONTRACT_TYPE.mIndex = 4
emp_CONTRACT_TYPE.show


Case 12
            If checkApility("FrmTables") = False Then
                Exit Sub
            End If
FrmTables.mIndex = 1
            FrmTables.show
            
Case 13

             If checkApility("System_alarms") = False Then
               Exit Sub
            End If

            System_alarms.show
            
            
         Case 14
            If checkApility("FrmCashing") = False Then
                Exit Sub
            End If

            OpenScreen CashingDataScreen


Case 15
If checkApility("Ageng_all") = False Then
                Exit Sub
            End If
            Unload Ageng_all
            Ageng_all.Indx = 4
Ageng_all.show


            End Select
End Sub

Private Sub DailyReport_Click()
    Dim Msg As String

    If checkApility("FrmDailtyReport") = False Then
        Exit Sub
    End If

    FrmDailtyReport.show
    FrmDailtyReport.ZOrder 0
   
    'If SystemOptions.usertype = UserAdminAll Or SystemOptions.usertype = UserNourCo Then
    '    FrmDailtyReport.Show
    '    FrmDailtyReport.ZOrder 0
    'Else
    '    Msg = "·«Ì„ﬂ‰ﬂ «· ⁄«„· „⁄ Â–Â «·‘«‘… ...."
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    'End If

End Sub

Private Sub DailyToolTip_Click()

End Sub
 
Private Sub DelayVal_Click(Index As Integer)

    Select Case Index

        Case 0

            'FrmPaymentTime
            If checkApility("FrmPaymentTime") = False Then
                Exit Sub
            End If

            OpenScreen PopUpShowPaymentTime

        Case 1
            Ageng.show

        Case 2
            Ageng_all.show

    End Select

End Sub

Private Sub DelItem_Click()
'    FrmMainPriceList.XPBtnRemove_Click
End Sub

Private Sub DelUser_Click()
    Dim Msg As String
    ''If user_id <> 1 Then
    ''    Msg = "€Ì— „”„ÊÕ ·ﬂ »«· ⁄«„· „⁄ Â–Â «·‰«›–…"
    '    MsgBox Msg, vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "’·«ÕÌ«  «·„” Œœ„Ì‰"
    '    Exit Sub
    'End If

    If user_id <> 1 Then
        '   MsgBox ""
        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
        '    Msg = Msg & Chr(13) & "Õ ‰Â“— Ê·««ÌÂ "
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If checkApility("FrmDelUser") = False Then
        Exit Sub
    End If

'    FrmDelUser.show vbModal
End Sub

Private Sub Destruction_Click()
    OpenScreen DestructionScreen
End Sub

Private Sub devsub_Click(Index As Integer)
Select Case Index
Case 0
            If checkApility("FrmDailyWorkflow") = False Then
                Exit Sub
            End If

FrmDailyWorkflow.show

Case 1
          If checkApility("FrmShowDailyWorkflow") = False Then
                Exit Sub
            End If
FrmShowDailyWorkflow.show

Case 2
            If checkApility("FrmOpDevelopment1") = False Then
                Exit Sub
            End If

FrmOpDevelopment1.show
Case 3
            If checkApility("FrmRegDevelopment") = False Then
                Exit Sub
            End If

FrmRegDevelopment.show


 

Case 4
           If checkApility("FrmAlarmDevelopmen") = False Then
                Exit Sub
            End If
FrmAlarmDevelopmen.show
Case 5
            If checkApility("FrmReportDevelopment") = False Then
                Exit Sub
            End If

FrmReportDevelopment.show


End Select
End Sub

Private Sub DockingPane1_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, _
                                ByVal Pane As XtremeDockingPane.IPane, _
                                ByVal Container As XtremeDockingPane.IPaneActionContainer, _
                                Cancel As Boolean)
 Exit Sub
    Dim Frm As Form
    Dim i  As Integer
    Dim Msg As String

    On Error GoTo hErr

    If Pane.ID = DockingPanesIDs.NewsBarPaneID Then
       ' If Not FrmNewsBarPane Is Nothing Then
       '     If Action = PaneActionClosed Then
       '         FrmNewsBarPane.TimerData.Enabled = False
       '     ElseIf Action = PaneActionCollapsed Then
       '         FrmNewsBarPane.TimerData.Enabled = False
       '     ElseIf Action = PaneActionCollapsing Then
       '         FrmNewsBarPane.TimerData.Enabled = False
       '     ElseIf Action = PaneActionExpanding Then
       '         FrmNewsBarPane.TimerData.Enabled = True
       '     ElseIf Action = PaneActionExpanded Then
       '         FrmNewsBarPane.TimerData.Enabled = True
       '     End If
       ' End If

    ElseIf Pane.ID = DockingPanesIDs.MantainceID Then

       ' If Not FrmMantaincePane Is Nothing Then
       '     If Action = PaneActionExpanded Or Action = PaneActionExpanding Then
       '         FrmMantaincePane.SetDcboSearch
       '     End If
       ' End If
    End If

    'For i = 0 To Forms.count - 1
    '    If Forms(i).Name <> "MDIFrmMain" Then
    '        If Forms(i).MDIChild = True Then
    '            Resize_Form Forms(i)
    '        End If
    '    End If
    'Next i
    
    'If Action = PaneActionPinned Or Me.DockingPane1.ActivePane Is Nothing Then
    '    For I = 0 To Forms.count - 1
    '        If Forms(I).Name <> "MDIFrmMain" Then
    '            If Forms(I).MDIChild = True Then
    '                Resize_Form Forms(I)
    '            End If
    '        End If
    '    Next I
    'End If
    Exit Sub
hErr:
    Msg = Err.Number
    Msg = Msg + Chr(13) & Err.description
    Msg = Msg + Chr(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub DockingPane1_AttachPane(ByVal Item As XtremeDockingPane.IPane)

    If Not Item Is Nothing Then
        If Item.ID = DockingPanesIDs.NewsBarPaneID Then
            'Set FrmNewsBarPane = New FrmPane
            'FrmNewsBarPane.PanelType = 1
            'Item.Handle = FrmNewsBarPane.hWnd
            'FrmNewsBarPane.backcolor = &HE2E9E9
        ElseIf Item.ID = DockingPanesIDs.OutBarPaneID Then
         '   Set FrmOutBarPane = New FrmOurBarPane
         '   Item.Handle = FrmOutBarPane.hwnd
         '   FrmOutBarPane.backcolor = &HE2E9E9
        ElseIf Item.ID = DockingPanesIDs.ItemsTreeID Then
            'Set ItemsTreePane = New FrmPaneTree
            'Item.Handle = ItemsTreePane.hWnd
            'ItemsTreePane.backcolor = &HE2E9E9
        ElseIf Item.ID = DockingPanesIDs.MantainceID Then
            'Set FrmMantaincePane = New FrmPane
            'FrmMantaincePane.PanelType = 3
            'Item.Handle = FrmMantaincePane.hWnd
            'FrmMantaincePane.backcolor = &HE2E9E9
        ElseIf Item.ID = DockingPanesIDs.InternetNews Then
            'Set FrmInternetNews = New FrmPane
            'FrmInternetNews.PanelType = 2
            'Item.Handle = FrmInternetNews.hWnd
            'FrmInternetNews.backcolor = &HE2E9E9
        ElseIf Item.ID = DockingPanesIDs.DynamicHelp Then
            Set FrmDynamicHelpPane = New FrmPaneHelp
            Item.Handle = FrmDynamicHelpPane.hwnd
            FrmDynamicHelpPane.backcolor = &HE2E9E9
            FrmDynamicHelpPane.Width = 100
            
        ElseIf Item.ID = DockingPanesIDs.CalendarPaneID Then
            'Set FrmCalendarPane = New FrmPaneCalendar
            'Item.Handle = FrmCalendarPane.hWnd 'salim found
            'FrmCalendarPane.backcolor = &HE2E9E9
        End If
    End If

End Sub

Private Sub DockingPane1_PanePopupMenu(ByVal Pane As XtremeDockingPane.IPane, _
                                       ByVal x As Long, _
                                       ByVal Y As Long, _
                                       Handled As Boolean)

    Select Case Pane.ID

        Case DockingPanesIDs.ItemsTreeID
            Me.MnuPopPane.Tag = DockingPanesIDs.ItemsTreeID
            MnuPopItemsTreePane_Array(2).Checked = Not Me.DockingPane1(DockingPanesIDs.ItemsTreeID).Hidden
            Me.PopupMenu Me.MnuPopPane
    End Select

End Sub

Private Sub EditPw_Click()

    If checkApility("FrmEditPW") = False Then
        Exit Sub
    End If

    FrmEditPW.show vbModal
    
    
    
    
End Sub

Private Sub Employee_Click(Index As Integer)

End Sub

Private Sub exit_Click()

End Sub
 
Private Sub EmployeSalary_Click()

End Sub

Private Sub Elevatorsmaintenance_Click(Index As Integer)
Select Case Index
Case 0
              If checkApility("FrmWarrantyOffer") = False Then
              Exit Sub
           End If
            FrmWarrantyOffer.show
Case 1
           If checkApility("FrmOut") = False Then
              Exit Sub
           End If
            FrmOut.show

Case 2
           If checkApility("FrmPerfMantAlaram") = False Then
              Exit Sub
           End If
            FrmPerfMantAlaram.show
            
            
   Case 3
           If checkApility("FrmMaintainanceAlarm") = False Then
              Exit Sub
           End If
           Unload FrmMaintainanceAlarm
           FrmMaintainanceAlarm.indexx = 0
            FrmMaintainanceAlarm.show
            
   Case 4
           If checkApility("FrmMaintainanceAlarm") = False Then
              Exit Sub
           End If
                  Unload FrmMaintainanceAlarm
           FrmMaintainanceAlarm.indexx = 1
            FrmMaintainanceAlarm.show
            
  Case 5
           If checkApility("FrmReports") = False Then
                Exit Sub
            End If

            FrmReports.show
            FrmReports.C1TabMain.CurrTab = 4
            FrmReports.Command2(9).Visible = False
            
End Select
End Sub

Private Sub EmployeeAttendanceSub_Click(Index As Integer)

    Select Case Index
            Case 0
              If checkApility("FrmVactionTypes") = False Then
              Exit Sub
           End If
            FrmVactionTypes.show
            
            FrmVactionTypes.WindowState = 0
            
 
 
        Case 1
                 If checkApility("frm_sheft") = False Then
                Exit Sub
            End If
frm_sheft.show
 '           If checkApility("FrmTimeSetting1") = False Then
 '               Exit Sub
 '           End If
'
'            Dim Frm As New FrmTimeSetting
'            Frm.WorkType = 1
'            Frm.show
'            Frm.ZOrder 0
'
        Case 2
                 If checkApility("FrmYearDurations2") = False Then
                Exit Sub
            End If
FrmYearDurations2.show
'            If checkApility("FrmPresentTime") = False Then
'                Exit Sub
'            End If
'
'            FrmPresentTime.show
'            FrmPresentTime.ZOrder 0
 
        Case 3
   
                 If checkApility("FrmImportShifts") = False Then
                Exit Sub
            End If
            FrmImportShifts.Auto_Man = 1
FrmImportShifts.show

        Case 4
   
                 If checkApility("FrmImportShifts") = False Then
                Exit Sub
            End If
            FrmImportShifts.Auto_Man = 0
FrmImportShifts.show


        Case 5
                 If checkApility("FrmApproveShift") = False Then
                Exit Sub
            End If
FrmApproveShift.show
'

          '  If checkApility("FrmAbsent") = False Then
          '      Exit Sub
          '  End If
'
'            FrmAbsent.show
'            FrmAbsent.ZOrder 0
'
        Case 6

'            If checkApility("FrmEmpMonthShow") = False Then
'                Exit Sub
'            End If

'            FrmEmpMonthShow.show
    
    End Select

End Sub

Private Sub EmployeeDataicSub_Click(Index As Integer)

    Select Case Index

        Case 0

Unload FrmEmployee
FrmEmployee.DriverOnly = 0
            'FrmEmployee
            If checkApility("FrmEmployee") = False Then
                Exit Sub
            End If

            OpenScreen EmployeesScreen
FrmEmployee.WorkShop_Job = 0
 

        Case 1

            If checkApility("frmEmpContract") = False Then
                Exit Sub
            End If

            frmEmpContract.show

    End Select

End Sub

Private Sub EmployeeSalarySub_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("MOFRAD") = False Then
                Exit Sub
            End If

            MOFRAD.show

        Case 1

            If checkApility("MOFRAD") = False Then
                Exit Sub
            End If

            If checkApility("mofradat2") = False Then
                Exit Sub
            End If

            mofradat2.show

        Case 2

            If checkApility("FrmMkafea") = False Then
                Exit Sub
            End If

            FrmMkafea.show
            FrmMkafea.ZOrder 0

        Case 3

            If checkApility("FrmKhsm") = False Then
                Exit Sub
            End If

            FrmKhsm.show
            FrmKhsm.ZOrder 0

        Case 4

            If checkApility("FrmEmpsAdvance") = False Then
                Exit Sub
            End If

            FrmEmpsAdvance.show
            FrmEmpsAdvance.ZOrder 0

        Case 5

         '   If checkApility("FrmEmpsAdvancePayed") = False Then
         '       Exit Sub
         '   End If
'
'            FrmEmpsAdvancePayed.show

        Case 6

            If checkApility("FrmEmpSalary") = False Then
                Exit Sub
            End If

            FrmEmpSalary5.show
            FrmEmpSalary5.ZOrder 0

        Case 7

        Case 8

            If checkApility("FrmChangedComponentData") = False Then
                Exit Sub
            End If

            FrmChangedComponentData.show

        Case 9

            If checkApility("FrmChangedComponentData1") = False Then
                Exit Sub
            End If

            FrmChangedComponentData1.show

        Case 10

        '    If checkApility("FrmChangedComponentData3") = False Then
        '        Exit Sub
        '    End If
'
'            FrmChangedComponentData3.show
'
        Case 11

  If checkApility("FrmEmpIncreaseSalaries") = False Then
                Exit Sub
            End If

FrmEmpIncreaseSalaries.show
'        Case 12
'
'            If checkApility("FrmEmpsAdvancePayed1") = False Then
'                Exit Sub
'            End If
'
'            FrmEmpsAdvancePayed1.show

    End Select

End Sub

 

Private Sub estateMain_Click(Index As Integer)
Select Case Index

Case 0
          If checkApility("FrmOrderMaintenance") = False Then
                Exit Sub
            End If
FrmOrderMaintenance.show
Case 1
          If checkApility("FrmLockedOrderMaintenance") = False Then
                Exit Sub
            End If
FrmLockedOrderMaintenance.show






End Select
End Sub

Private Sub Expenses_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("FrmExpenses3") = False Then
                Exit Sub
            End If

            FrmExpenses3.show
Case 1

  If checkApility("frmserviceInvoice") = False Then
        Exit Sub
    End If

    frmserviceInvoice.show


        Case 1

    End Select

End Sub

Private Sub ExpensesSub_Click(Index As Integer)

    Select Case Index
         Case 0

            '           OpenScreen ExpensesDataScreen
            If checkApility("FrmDataTypeExchange") = False Then
                Exit Sub
            End If

            FrmDataTypeExchange.show
            
        Case 1

            '           OpenScreen ExpensesDataScreen
            If checkApility("FrmTypeExchange") = False Then
                Exit Sub
            End If

            FrmTypeExchange.show
            
        Case 2

            '           OpenScreen ExpensesDataScreen
            If checkApility("FrmExpenses5") = False Then
                Exit Sub
            End If

            FrmExpenses5.show

        Case 3

            'Frm
            'Payments
            If checkApility("FrmPayments") = False Then
                Exit Sub
            End If

            OpenScreen PaymentsDataScreen

Case 4
            If checkApility("FrmAccEditJournal3") = False Then
                Exit Sub
            End If


FrmAccEditJournal3.show

    End Select
 
End Sub

Private Sub ExpensesType_Click(Index As Integer)

    Select Case Index

        Case 0

            'FrmExpensesType
            If checkApility("FrmExpensesType") = False Then
                Exit Sub
            End If

            OpenScreen ExpensesTypes

        Case 1

            'FrmRevenuesTypes
            If checkApility("FrmRevenuesTypes") = False Then
                Exit Sub
            End If

            OpenScreen RevenuesTypes
            
           Case 2
                   If checkApility("FrmBanksCheck") = False Then
                Exit Sub
            End If
FrmBanksCheck.show

    End Select

End Sub

Private Sub eyeSub_Click(Index As Integer)
Select Case Index

Case 0 '„Œ«“‰
            If checkApility("FrmStoreData") = False Then
                Exit Sub
            End If

            'FrmStoreData
            OpenScreen StoresDataScreen


Case 1 '„Ã„Ê⁄« 

            If checkApility("FrmGroups") = False Then
                Exit Sub
            End If

            'FrmGroups
            OpenScreen ItemsGroupsScreen

Case 2 'ÊÕœ« 

Unload FrmPay_Garanty_Shipment3M
             If checkApility("FrmPay_Garanty_Shipment3M") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment3M.SendForm = 0
FrmPay_Garanty_Shipment3M.show
            
Case 3 '"«’‰«›
If checkApility("FrmItems") = False Then
                Exit Sub
            End If

            OpenScreen ItemsDataScreen

Case 4 '«·⁄œ”« 
      If checkApility("FrmVizitScreen") = False Then
                Exit Sub
            End If
            
 
FrmVizitScreen.mIndex = 3

 FrmVizitScreen.show
 

Case 5 '⁄„·«¡




            If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If

            OpenScreen CustomersScreen '
       Case 6 '„‰œÌ»
                    If checkApility("FrmPay_Garanty_Shipment") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment.SendForm = 7
FrmPay_Garanty_Shipment.show


Case 7 '«ÿ»«¡

        If checkApility("project_status") = False Then
                Exit Sub
            End If

          project_status.mIndex = 4
                project_status.show
                
Case 8 '‘—ﬂ«   √„Ì‰

If checkApility("insurancecompanies") = False Then
                Exit Sub
            End If
            
            insurancecompanies.show
          
Case 9

            If checkApility("Nationality") = False Then
                Exit Sub
            End If

            Nationality.show


Case 10



Unload FrmPay_Garanty_Shipment3M
             If checkApility("FrmPay_Garanty_Shipment3M") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment3M.SendForm = 1
FrmPay_Garanty_Shipment3M.show









Case 11 '»Ì«‰«   «Œ—Ì

If checkApility("FrmBasicDataINv") = False Then
                Exit Sub
            End If
            FrmBasicDataINv.mIndex = 2
            FrmBasicDataINv.show
            

Case 12 '„‘ —Ì« 
          If checkApility("FrmBillBuy") = False Then
                Exit Sub
            End If

            OpenScreen PurchaseScreen

Case 13 '„—œÊœ« 
    If checkApility("FrmReturnpurchases") = False Then
                Exit Sub
            End If

            OpenScreen RetrunPurchse

Case 14 '„»Ì⁄« 
If checkApility("FrmSaleBill4") = False Then
                Exit Sub
            End If

     frmsalebill4.show
            'OpenScreen InvoiceScreen
Case 15 '„—œÊœ« 
 If checkApility("FrmReturnSalling") = False Then
                Exit Sub
            End If

            'FrmReturnSalling
            OpenScreen RetrunSalles
Case 16 '”‰œ«  «·ﬁ»÷
   If checkApility("FrmCashing") = False Then
                Exit Sub
            End If

            OpenScreen CashingDataScreen

Case 17 '’—›

            If checkApility("FrmExpenses5") = False Then
                Exit Sub
            End If

            FrmExpenses5.show

Case 18 ' ’›ÌÂ ⁄Âœ…
   If checkApility("FrmExpenses30") = False Then
                Exit Sub
            End If

            FrmExpenses30.show


Case 19 '„œ›Ê⁄« 
  If checkApility("FrmPayments") = False Then
                Exit Sub
            End If

            OpenScreen PaymentsDataScreen


Case 20 '«‘⁄«—« 
  If checkApility("FrmDiscounts") = False Then
        Exit Sub
    End If

    OpenScreen AllowsDiscountsScreen

Case 21 '‰ﬁ«—Ì— ⁄«„Â

            If checkApility("ReportPurchase") = False Then
                Exit Sub
            End If

            FrmReports.show
            FrmReports.C1TabMain.CurrTab = 1


Case 22 ' ﬁ«—Ì— „Õ«”»ÌÂ


            If checkApility("FrmAccountingReport") = False Then
                Exit Sub
            End If

            FrmAccountingReport.show

Case 23 '  «·ÿ·»«  «·œ«Œ·Ì… ' ‰»ÌÂ« 
            
      If checkApility("FrmVizitScreen") = False Then
                Exit Sub
            End If
            
 
FrmVizitScreen.mIndex = 4

 FrmVizitScreen.show
 
 Case 24 ' ‰»ÌÂ« 

 
            
      If checkApility("FrmVizitScreen") = False Then
                Exit Sub
            End If
 
FrmVizitScreen.mIndex = 5

 FrmVizitScreen.show
 
 Case 25 ' ‰»ÌÂ« 

 
            
      If checkApility("FrmVizitScreen") = False Then
                Exit Sub
            End If
            
 
FrmVizitScreen.mIndex = 6

 FrmVizitScreen.show
 
 
 
End Select
End Sub

Private Sub FarmSub_Click(Index As Integer)
Select Case Index
Case 0
        If checkApility("FrmBasicDataINv") = False Then
                Exit Sub
            End If

FrmBasicDataINv.mIndex = 5
FrmBasicDataINv.show

Case 1
        If checkApility("FrmBasicDataINv") = False Then
                Exit Sub
            End If

FrmBasicDataINv.mIndex = 5
FrmBasicDataINv.show

Case 2
        If checkApility("FrmBasicDataINv") = False Then
                Exit Sub
            End If

FrmBasicDataINv.mIndex = 5
FrmBasicDataINv.show

Case 3
        If checkApility("FrmBasicDataINv") = False Then
                Exit Sub
            End If

FrmBasicDataINv.mIndex = 5
FrmBasicDataINv.show

Case 4
        If checkApility("FrmBasicDataINv") = False Then
                Exit Sub
            End If

FrmBasicDataINv.mIndex = 5
FrmBasicDataINv.show

Case 5

        If checkApility("FrmBasicDataINv") = False Then
                Exit Sub
            End If

FrmBasicDataINv.mIndex = 5
FrmBasicDataINv.show

Case 6
        If checkApility("FrmBasicDataINv") = False Then
                Exit Sub
            End If

FrmBasicDataINv.mIndex = 5
FrmBasicDataINv.show

Case 7
        If checkApility("FrmBasicDataINv") = False Then
                Exit Sub
            End If

FrmBasicDataINv.mIndex = 5
FrmBasicDataINv.show

Case 8 ' ”ÃÌ· ÃÌÊ«‰
        If checkApility("FrmItemsClass") = False Then
                Exit Sub
            End If

FrmItemsClass.mIndex = 12
FrmItemsClass.show



Case 9 ' ”ÃÌ·  ÿ·» ⁄·«Ã
        If checkApility("FrmItemsClass") = False Then
                Exit Sub
            End If
FrmItemsClass.mIndex = 13
FrmItemsClass.show

Case 10 ' ”ÃÌ· ‰›Êﬁ
        If checkApility("FrmItemsClass") = False Then
                Exit Sub
            End If
FrmItemsClass.mIndex = 14
FrmItemsClass.show

Case 11 ' ﬁ«—Ì—
        If checkApility("FrmItemsClass") = False Then
                Exit Sub
            End If
FrmItemsClass.mIndex = 15
FrmItemsClass.show


End Select
End Sub

Private Sub FinishSevicersub_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("FrmRegisterHoliday") = False Then
                Exit Sub
            End If

            FrmRegisterHoliday.show

        Case 1

            If checkApility("End_oF_service") = False Then
                Exit Sub
            End If

            End_oF_service.show

    End Select

End Sub

Private Sub Groups_Click()

End Sub

Private Sub HelpFile_Click()
    
End Sub

Private Sub HelpIndex_Click()

End Sub

Private Sub insurance_type_Click()

End Sub

Private Sub Items_Click(Index As Integer)

End Sub

 

Private Sub gobusSub_Click(Index As Integer)
Select Case Index

Case 0 '
       If checkApility("FrmCountriesData") = False Then
                Exit Sub
            End If

            FrmCountriesData.show
Case 1 '
         If checkApility("FrmGovernmentData") = False Then
                Exit Sub
            End If

            FrmGovernmentData.show


Case 2 '
            If checkApility("FrmCitiesDistance") = False Then
                Exit Sub
            End If
FrmCitiesDistance.Indx = 0
            FrmCitiesDistance.show



Case 3 '
       If checkApility("CarTypes") = False Then
                Exit Sub
            End If

            CarTypes.show
Case 4 '
  If checkApility("FrmCarModels") = False Then
                Exit Sub
            End If

FrmCarModels.show

Case 5 '
  If checkApility("FrmColor") = False Then
              Exit Sub
           End If
 FrmColor.show
Case 6 '
    If checkApility("FrmCars") = False Then
              Exit Sub
           End If
            FrmCars.show
Case 7 '
       If checkApility("FrmDrivers") = False Then
                Exit Sub
            End If

            FrmDrivers.show

Case 8 '
       If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If

            OpenScreen CustomersScreen '


Case 9 '
     If checkApility("FrmDriverAllocation") = False Then
                Exit Sub
            End If

 
FrmDriverAllocation.show
Case 10 '
FrmItemsClass.mIndex = 7
FrmItemsClass.show
Case 11 '
FrmItemsClass.mIndex = 8
FrmItemsClass.show


Case 0 '
Case 0 '
Case 0 '
Case 0 '


End Select
End Sub

Private Sub hajMnuSub_Click(Index As Integer)
Select Case Index
Case 4 'ÿ·» ÕÃ“
      If checkApility("FrmBookingRequest") = False Then
                Exit Sub
            End If
FrmBookingRequest.show


Case 5 '  √ﬂÌœ ÕÃ“
   If checkApility("FrmApproveRequset") = False Then
                Exit Sub
            End If
FrmApproveRequset.show
Case 6 '«„—  ‘€Ì·
            If checkApility("FrmBookingRequest2") = False Then
                Exit Sub
            End If
FrmBookingRequest2.show

Case 7 ' 7 ÃœÊ· «· —ÕÌ·« 
      If checkApility("FrmDeported") = False Then
                Exit Sub
            End If
FrmDeported.show

Case 8 '««·„”«—«  «·„Œ’Ê„… ··⁄„—…
      If checkApility("FrmExtinAccounts") = False Then
                Exit Sub
            End If
FrmExtinAccounts.show

Case 9 '«⁄ „«œ «—ﬂ«» «·ÕÃ«Ã
      If checkApility("FrmEndorseTrans") = False Then
                Exit Sub
            End If
FrmEndorseTrans.show

Case 10 '9 ÃœÊ· «· —ÕÌ·«  ·«—ﬂ«»
      If checkApility("FrmPilgrimsService") = False Then
                Exit Sub
            End If
FrmPilgrimsService.show ' ÿÃœÊ· «· —ÕÌ·«  ··ÕÃ



Case 11 '10 «⁄ „«œ «·„‘«⁄—
    If checkApility("FrmEndorseTransMashar") = False Then
                Exit Sub
            End If
FrmEndorseTransMashar.show
 Case 12 '11  Ê“Ì⁄ Õ«›·«  «·„‘«⁄—
     If checkApility("FrmBusesDistribution") = False Then
                Exit Sub
            End If
FrmBusesDistribution.show


Case 13  ' «Œ·«¡ ÿ—›
      If checkApility("FrmEvacation") = False Then
                Exit Sub
            End If
FrmEvacation.show
  
Case 14  '«·„ÿ«·»« 
      If checkApility("frmDetailsAdoption") = False Then
                Exit Sub
            End If
    frmDetailsAdoption.show

Case 15  '«·Õ”„Ì« 
      If checkApility("FrmDeduction") = False Then
                Exit Sub
          End If
   FrmDeduction.show




Case 16
      If checkApility("FrmHajjReports") = False Then
                Exit Sub
            End If
    FrmHajjReports.show

End Select
End Sub

Private Sub hajMnuSub1_Click(Index As Integer)
Select Case Index

     Case 0

            If checkApility("FrmBasicDataHajj") = False Then
                Exit Sub
            End If

            FrmBasicDataHajj.show

  Case 1

            If checkApility("FrmGovernmentData") = False Then
                Exit Sub
            End If

            FrmGovernmentData.show

        Case 2

            If checkApility("FrmCitiesDistance") = False Then
                Exit Sub
            End If

            FrmCitiesDistance.show

   
         Case 3

Unload FrmEmployee
 
            If checkApility("FrmEmployee") = False Then
                Exit Sub
            End If
FrmEmployee.show


  Case 4
  
       If checkApility("CarTypes") = False Then
                Exit Sub
            End If

            CarTypes.show
            CarTypes.Caption = "«‰Ê«⁄ «·Õ«›·« "
            CarTypes.Label1(2).Caption = CarTypes.Caption
             
     Case 5
            If checkApility("FrmCars") = False Then
              Exit Sub
           End If
            FrmCars.show
            FrmCars.Caption = "»Ì«‰«  «·Õ«›·« "
          '  FrmCars.Ele.Caption = FrmCars.Caption
                 FrmCars.Image2.Visible = True
                 FrmCars.lbl(7).Visible = False
                 FrmCars.TxtEquQty.Visible = False
                 FrmCars.Label4.Visible = False
                 
               FrmCars.WindowState = 0

     Case 6
FrmCustomerType.Indx = 0
   
FrmCustomerType.show

   Case 7
   If checkApility("FrmCreditFacicity") = False Then
                Exit Sub
            End If
            
      '    FormRequestOpenAccount.show
    FrmCreditFacicity.show
    

    Case 8

    '     If checkApility("FrmCompany") = False Then
    '            Exit Sub
    '        End If
'
'            FrmCompany.show
'
        Case 9

            If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If

            OpenScreen CustomersScreen '

Case 10
       If checkApility("FrmCompanyContract") = False Then
                Exit Sub
            End If

FrmCompanyContract.show

End Select
End Sub

Private Sub help_list_Click(Index As Integer)
 
If Index = 0 Then
FrmFavorites.show
  
Else

showFavoritesSelectedMenue help_list(Index).Caption
End If

 
End Sub

Private Sub HelpFileSub_Click(Index As Integer)


'FRMTRansferData.show
'FrmAssignCarDuration.show

 

'FrmGroupsx.show
'FrmPaneHelp.show
'FrmPaneTree.show
 'mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID).Closed = 1   'Not MDIFrmMain.DockingPane1.FindPane(DockingPanesIDs.DynamicHelp).Closed
  '      Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.CalendarPaneID)

'mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.CalendarPaneID).Closed = 0

'mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.DynamicHelp).Closed = 0 'Not MDIFrmMain.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID).Closed
'FrmAllocationToContract.show
'Splish.show
'Exit Sub
Select Case Index
Case 0
FrmVizitScreen.mIndex = 7
FrmVizitScreen.show
 'mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID).Closed = 1   'Not MDIFrmMain.DockingPane1.FindPane(DockingPanesIDs.DynamicHelp).Closed=-1
   ' SystemOptions.SysHelp.HHDisplayContents Me.hWnd
Case 1
   ' SystemOptions.SysHelp.HHDisplayIndex Me.hWnd
 Case 2
   ' SystemOptions.SysHelp.HHDisplaySearch Me.hWnd
Case 3

    'FrmDailyToolTip.show
Case 4
    'frmabout.show vbModal

Case 5
    
    Dim Msg As String

FrmActivation.show

Exit Sub


    If SystemOptions.SysRegisterState = DemoRun Or SystemOptions.SysRegisterState = DemoStop Then
  FrmActivation.show
     '   FrmRegisteration.show vbModal
    Else
        Msg = "‰”Œ… „”Ã·… "
        Msg = Msg & Chr(13) & "‘ﬂ—« .. .·≈” Œœ«„ﬂ„ »—‰«„Ã ‰Ÿ«„ œÌ‰«„Ìﬂ »«Ì "
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If


Case 6
     If Dir(App.path & "\checklist\Checklist.exe") <> "" Then
         Shell App.path & "\checklist\Checklist.exe", vbNormalFocus
     End If
    
Case 7
'    OpenWebSite "http://www.sattaryah.com"

 'frmsalebill5.show

End Select
End Sub

Private Sub HRProcedures_Click(Index As Integer)
Select Case Index
Case 0

'    If checkApility("FrmEmpsAdvanceRequest") = False Then
'        Exit Sub
'    End If
'
'FrmEmpsAdvanceRequest.show


Case 1
'FrmPassover.show
    If checkApility("FrmMovingEmp") = False Then
        Exit Sub
    End If
FrmMovingEmp.show
Case 2
    If checkApility("FrmBusinessJob") = False Then
        Exit Sub
    End If
FrmBusinessJob.show
Case 3
    If checkApility("FrmAdvancedHousingpayments") = False Then
        Exit Sub
    End If
    
FrmAdvancedHousingpayments.show
Case 4
    If checkApility("FormEmpMoveDepartment") = False Then
        Exit Sub
    End If
    
FormEmpMoveDepartment.show

Case 5
    If checkApility("FrmEmbarkation") = False Then
        Exit Sub
    End If
FrmEmbarkation.show
Case 7


    If checkApility("FrmQUesEmp") = False Then
        Exit Sub
    End If
FrmQUesEmp.show ''«” »Ì«‰


Case 8
    If checkApility("formvocatinl") = False Then
        Exit Sub
    End If
formvocatinl.show
Case 9
   If checkApility("FrmHolidayData") = False Then
        Exit Sub
    End If
    
'FrmHolidayData.show
Case 10
   'If checkApility("frmdriveassest") = False Then
   '     Exit Sub
   ' End If
   '
    
'frmdriveassest.show

   If checkApility("frmdriveassestMove") = False Then
        Exit Sub
    End If
    
    
frmdriveassestMove.show


Case 11
   If checkApility("FrmPassports") = False Then
        Exit Sub
    End If
FrmPassports.show
Case 12
   If checkApility("FRmEmployeeWarning") = False Then
        Exit Sub
    End If
FRmEmployeeWarning.show
Case 13
   If checkApility("FrmTreament") = False Then
        Exit Sub
    End If
FrmTreament.show 'ÿ ÌÂ„… «·«„— ·„

Case 14
   If checkApility("FrmRepInjuy") = False Then
        Exit Sub
    End If
FrmRepInjuy.show ' ﬁ—Ì— «’«»Â ⁄„·

Case 15
   If checkApility("FrmReceivingTreatment") = False Then
        Exit Sub
    End If
FrmReceivingTreatment.show ' ﬁ—Ì—   «” ·«„ „⁄«„·« 

Case 16
   If checkApility("FrmFinalSettlement") = False Then
        Exit Sub
    End If
FrmFinalSettlement.show ' ﬁ—Ì—   „Œ«·’… ‰Â«∆Ì…


Case 25
   If checkApility("FrmChangeEmployeedata") = False Then
        Exit Sub
    End If
FrmChangeEmployeedata.show


Case 26
   If checkApility("FrmClearanceCerTifcate") = False Then
        Exit Sub
    End If
FrmClearanceCerTifcate.show

Case 27
   If checkApility("FrmFolloAdminMeasure") = False Then
        Exit Sub
    End If
FrmFolloAdminMeasure.show

Case 28
   If checkApility("FrmDeductionNote") = False Then
        Exit Sub
    End If
FrmDeductionNote.show

Case 29
   If checkApility("FrmCarReceipt") = False Then
        Exit Sub
    End If
FrmCarReceipt.show

Case 30
   If checkApility("FrmDefineEmp") = False Then
        Exit Sub
    End If
FrmDefineEmp.show

Case 31
   If checkApility("FrmMovingEmp2") = False Then
        Exit Sub
    End If
FrmMovingEmp2.show




End Select
End Sub

Private Sub LeavingRecord_Click()

    If checkApility("FrmGoTime") = False Then
        Exit Sub
    End If

'    FrmGoTime.show
'    FrmGoTime.ZOrder 0
End Sub

Private Sub m3_Click(Index As Integer)

End Sub

Private Sub LCTransactions_Click(Index As Integer)
    Dim rsOut As ADODB.Recordset
    Dim RsOptions As ADODB.Recordset
    Dim Msg As String

    Select Case Index

        Case 0

            If checkApility("FrmLCTypes") = False Then
                Exit Sub
            End If

            FrmLCTypes.show

        Case 1

            If checkApility("FrmShowPrice2") = False Then
                Exit Sub
            End If

            GeneralPriceType = 2
            FrmShowPrice.show

        Case 2

            If checkApility("FrmLC") = False Then
                Exit Sub
            End If

            FrmLC.show

        Case 3

            If checkApility("FrmLC1") = False Then
                Exit Sub
            End If

            FrmLC.show

        Case 4

            If checkApility("shipmentA") = False Then
                Exit Sub
            End If

            If SystemOptions.UserInterface = ArabicInterface Then
                shipmentA.show
            Else
                shipmentA.show
            End If

        Case 5

            If checkApility("FrmInpout1") = False Then
                Exit Sub
            End If

            Set rsOut = New ADODB.Recordset
            rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

            If Not (rsOut.EOF Or rsOut.BOF) Then
                If rsOut!checkinpo = True Then
                    FrmInpout.show
                ElseIf rsOut!checkbey = True Then
                    Msg = "⁄›Ê«  „ «Œ Ì«— ›« Ê—… «·‘—«¡ ··«÷«›…  ... ·«Ì„ﬂ‰ «·«÷«›…  „‰ «–‰ «·«÷«›… "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                Else
                End If
            End If

        Case 6
            Set RsOptions = New ADODB.Recordset
            RsOptions.Open "tbloptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
            OpenScreen PurchaseScreen

            If checkApility("FrmBillBuy1") = False Then
                Exit Sub
            End If

            'Purchase Invoices
        Case 7

            If checkApility("FrmLC2") = False Then
                Exit Sub
            End If

            FrmLC.show


        Case 8

            If checkApility("FrmBankpledge1") = False Then
                Exit Sub
            End If

            FrmBankPledge1.show
            
       Case 9

            If checkApility("FrmBankPledge2") = False Then
                Exit Sub
            End If

            FrmBankPledge2.show
            
       Case 10

            If checkApility("FrmBankPledge3") = False Then
                Exit Sub
            End If

            FrmBankPledge3.show
Case 11

            If checkApility("FrmBankPledge4") = False Then
                Exit Sub
            End If
FrmBankPledge4.show


    End Select

 
End Sub

Private Sub m2_Click()
    'xx.show
    'xx.top = 0
    'xx.left = 11500
    ' xx.SmartMenuXP1_Click (0)
 
End Sub

Private Sub LegalIssueSub_Click(Index As Integer)
Select Case Index
Case 0
FRMcurrency.mIndex = 2
FRMcurrency.show

Case 1
FRMcurrency.mIndex = 1
FRMcurrency.show
Case 2

FRMcurrency.mIndex = 3
FRMcurrency.show
Case 3
FRMcurrency.mIndex = 4
FRMcurrency.show
Case 4
FRMcurrency.mIndex = 5
FRMcurrency.show

Case 5
Nationality.mIndex = 1
Nationality.show
End Select
End Sub

Private Sub LIFEINDICATORMNU_Click()
        If checkApility("ProjectsBillAlarm") = False Then
                Exit Sub
            End If


ProjectsBillAlarm.SendForm = "Dash"
ProjectsBillAlarm.show
End Sub

Private Sub mangDepSub_Click(Index As Integer)
Select Case Index
Case 0
 If checkApility("frmtakeem") = False Then
                Exit Sub
            End If

            frmtakeem.show
            
Case 1
          If checkApility("FRmEmployMentModell") = False Then
                Exit Sub
            End If
FRmEmployMentModell.show
Case 2
          If checkApility("NotifyJobNeeded") = False Then
                Exit Sub
            End If
NotifyJobNeeded.show


End Select
End Sub

Private Sub MarketingMnusub_Click(Index As Integer)
Select Case Index
 
Case 1
          If checkApility("Frmovers") = False Then
                Exit Sub
            End If

            Frmovers.show
 

Case 2

Case 3
       If checkApility("FrmRegDateDelgateREport") = False Then
                Exit Sub
            End If
FrmRegDateDelgateREport.show
Case 4
                         If checkApility("FrmReportsStudent") = False Then
                Exit Sub
            End If
               FrmReportsStudent.show
               FrmReportsStudent.XPTab301.TabVisible(1) = False
               FrmReportsStudent.AttRB.Visible = False
               FrmReportsStudent.ComRep.Visible = False
               FrmReportsStudent.StuInfoRB.Visible = False
 
               
               
End Select
End Sub

Private Sub MarketingMnusubsub_Click(Index As Integer)
Select Case Index
Case 0
    If checkApility("FrmRegDateDelgate") = False Then
                Exit Sub
            End If
FrmRegDateDelgate.show
Case 1

            If checkApility("FrmCustomerssFollow") = False Then
                Exit Sub
            End If

             FrmCustomerssFollow.show
            
    Case 2

            If checkApility("FrmCustomerssFollow") = False Then
                Exit Sub
            End If

             FrmCustomerssFollow.show
             
       Case 3
       
         If checkApility("FrmCustomerssComplaint") = False Then
                Exit Sub
            End If

     FrmCustomerssComplaint.show
           
     Case 4
         If checkApility("FrmCustomerssComplaint") = False Then
                Exit Sub
            End If

     FrmCustomerssComplaint.show
          Case 5
         If checkApility("FrmCustomerssComplaint") = False Then
                Exit Sub
            End If

     FrmCustomerssComplaint.show



Case 6
            If checkApility("RSPhoneBook") = False Then
                Exit Sub
            End If

            RSPhoneBook.show


Case 7
       If checkApility("FrmShowRegDateDelegate") = False Then
                Exit Sub
            End If
FrmShowRegDateDelegate.show

Case 8
       If checkApility("FrmStudentCalling") = False Then
                Exit Sub
            End If
FrmStudentCalling.show

End Select
End Sub

Public Function showmnue()
        PopupMenu mdifrmmain.MdiContextMenu, , Me.Width, 0 ', vbPopupMenuRightAlign, X, Y + 200

End Function

Private Sub MDIForm_Click()
  'loadmyModule
'showmnue
'Unload WebForm
'Load WebForm

End Sub



Private Sub MDIForm_DblClick()

    With Cmdlg
        '*.jpg,*.jpeg,*.jpe,*.jfif
        .CancelError = False
        .DialogTitle = " ≈Œ Ì«— ’Ê—…"
        'Set The Filter to show pictures only
        .filter = "Bitmap (*.bmp)|*.bmp|JPEG(*.JPG,*.JPEG,*.JPE,*.JFIF)|*.jpg;*.jpeg;*.jpe;*.jfif|" & "GIF (*.gif)|*.gif|All Files|*.*" ' choose formats to include
        .ShowOpen
    
    
        If .filename <> "" Then
            'Set Me.ImgPic.Picture = LoadPicture(.FileName)
      ' Me.Picture = LoadPicture(.FileName)
            WebForm.Image1.Picture = LoadPicture(.filename)
            SaveSetting StrAppRegPath, "View_Type", "BackGroundImag", .filename
        Else

            If Dir(App.path & "\Garphics\wallpaper_Main.jpg") <> "" Then
     '           Me.Picture = LoadPicture(App.path & "\Garphics\wallpaper_Main.jpg")
     '           WebForm.Picture = LoadPicture(.FileName)
                SaveSetting StrAppRegPath, "View_Type", "BackGroundImag", App.path & "\Garphics\wallpaper_Main.jpg"
                                
            End If

        End If

    End With

    ' €ÌÌ— «·Œ·›Ì…

End Sub
Function showFavoritesSelectedMenue(Optional Displayname As String = "")
On Error Resume Next
Dim i As Integer
 
 If Displayname = "" Then Exit Function
 Dim formname As String
     Dim sql As String
    Dim rsMenue As New ADODB.Recordset
    Dim NoOfMinute As Double
    'Dim TimeCateg As Double
 
    sql = "SELECT   FORMNAME"
sql = sql & " from dbo.TblMyMenue "
sql = sql & " WHERE   Displayname='" & Displayname & "'"
     rsMenue.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsMenue.RecordCount > 0 Then
 formname = IIf(IsNull(rsMenue("FORMNAME").value), 0, rsMenue("FORMNAME").value)
    
      '           ShowForm (formname)
      
     
'Debug.Print formname
      


      If formname = "FrmVizitScreen" Then
      If checkApility("FrmVizitScreen") = False Then
                Exit Function
            End If
            
 
FrmVizitScreen.mIndex = 1

 FrmVizitScreen.show
 
      End If


      If formname = "FrmRsContractAlarm" Then
      FrmRsContractAlarm.show
      End If
            If formname = "Ageng_all" Then
                        Unload Ageng_all
            Ageng_all.Indx = 1
Ageng_all.show

 
      End If
      
                 If formname = "FrmTotalsReport" Then
      FrmTotalsReport.show
      End If
      
      
                 If formname = "FrmIncomAndExpenReports" Then
      FrmIncomAndExpenReports.show
      End If
      
      
      
      
      
          If formname = "FrmShowDailyWorkflow" Then
      FrmShowDailyWorkflow.show
      End If
      
      
          If formname = "FrmOpDevelopment1" Then
      FrmOpDevelopment1.show
      End If
      
             If formname = "FrmRegDevelopment" Then
      FrmRegDevelopment.show
      End If
      

             If formname = "FrmAlarmDevelopmen" Then
      FrmAlarmDevelopmen.show
      End If
      
            If formname = "FrmReportDevelopment" Then
      FrmReportDevelopment.show
      End If
      
           If formname = "FrmDailyWorkflow" Then
      FrmDailyWorkflow.show
      End If
      
      
      If formname = "FrmCashing" Then
      FrmCashing.show
      End If
      
      
          If formname = "FrmEmpSalary5" Then
      FrmEmpSalary5.show
      End If
      
      
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmExpensesType" Then
      FrmExpensesType.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmRevenuesTypes" Then
      FrmRevenuesTypes.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBanksData" Then
      FrmBanksData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBoxesData" Then
      FrmBoxesData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPaymentType" Then
      FrmPaymentType.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCompany" Then
  FrmCompany.mIndex = 0
      FrmCompany.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustemers" Then
      FrmCustemers.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustemers" Then
      FrmCustemers.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmItems" Then
      FrmItems.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FRMcurrency" Then
      FRMcurrency.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "Nationality" Then
      Nationality.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "dean" Then
      dean.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCountriesData" Then
      FrmCountriesData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmGovernmentData" Then
      FrmGovernmentData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmGovernCitiesData" Then
      FrmGovernCitiesData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "streets" Then
      streets.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmTotals2Report" Then
      FrmTotals2Report.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "project_status" Then
      project_status.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "Contract_type" Then
      contract_type.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmOtherCustomers" Then
      FrmOtherCustomers.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmProcessUnit" Then
      FrmProcessUnit.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmProcessDef" Then
      FrmProcessDef.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmequipment" Then
      frmequipment.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCashing" Then
      FrmCashing.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "Projects" Then
      Projects.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmDestruction" Then
      FrmDestruction.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmpSalary3" Then
      FrmEmpSalary3.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmpSalary4" Then
      FrmEmpSalary4.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmpSalary3A" Then
      FrmEmpSalary3A.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmpSalary4A" Then
      FrmEmpSalary4A.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmOperationsFollow" Then
      FrmOperationsFollow.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "projectsbill" Then
      projectsbill.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmProjectsReports" Then
      frmProjectsReports.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmequipment" Then
      frmequipment.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "UnitsIndustrialCost" Then
      UnitsIndustrialCost.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmProductionElements" Then
      'FrmProductionElements.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmequipment1" Then
      frmequipment1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmProductionType" Then
      FrmProductionType.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmDistriExpensItems" Then
      FrmDistriExpensItems.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmProductLine" Then
      frmProductLine.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmTransferEmployee" Then
      FrmTransferEmployee.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmProductionOrder1" Then
      FrmProductionOrder1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmOutProductionOrder1" Then
      FrmOutProductionOrder1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmInpoutWorkOrder1" Then
      FrmInpoutWorkOrder1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPO9" Then
      FrmPO9.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmProductionOrder" Then
      FrmProductionOrder.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmOutProductionOrder" Then
      FrmOutProductionOrder.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmInpoutWorkOrder" Then
      FrmInpoutWorkOrder.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCalcCostPrice" Then
      FrmCalcCostPrice.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmProductionAllocation" Then
      FrmProductionAllocation.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmProductionReport" Then
      frmProductionreport.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmItems" Then
      FrmItems.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmStoreData" Then
      FrmStoreData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmGroups" Then
      FrmGroups.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSystemUnites" Then
'      FrmSystemUnites.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmItemsColor" Then
      'FrmItemsColor.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmItemsSize" Then
      'FrmItemsSize.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmItemsClass" Then
      FrmItemsClass.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmStoresLocation" Then
      'FrmStoresLocation.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmStoresLocation" Then
      'FrmStoresLocation.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmLinkItemToStore" Then
      FrmLinkItemToStore.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmOpeningBalance" Then
      FrmOpeningBalance.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPO6" Then
      FrmPO6.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPO7" Then
      FrmPO7.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmInpout" Then
      FrmInpout.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmOut" Then
      FrmOut.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmOut1" Then
      FrmOut1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmMoving" Then
      FrmMoving.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmStartGard" Then
      FrmStartGard.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmGardReport" Then
      FrmGardReport.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmNewGard" Then
      FrmNewGard.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmNewGard1" Then
      FrmNewGard1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmStockSettlement" Then
      FrmStockSettlement.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmDefinCompItem" Then
      FrmDefinCompItem.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSearchSerial" Then
      FrmSearchSerial.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSerialData" Then
      FrmSerialData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmRequest" Then
      FrmRequest.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmReports" Then
      FrmReports.show
    End If
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmVendorContract" Then
      FrmVendorContract.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "Ageng" Then
      Ageng.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmShipment_mode" Then
      FrmShipment_mode.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmGaranty_type" Then
      FrmGaranty_type.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPaymentData" Then
      'FrmPaymentData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSalesRepData1" Then
'      FrmSalesRepData1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmShipingData" Then
'      FrmShipingData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPO4" Then
      FrmPO4.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPO5" Then
      FrmPO5.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmComparePrices" Then
      FrmComparePrices.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPO8" Then
      FrmPO8.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPO10" Then
      FrmPO10.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "shipment" Then
       
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmLCTypes" Then
      FrmLCTypes.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmLC" Then
      FrmLC.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "shipmentA" Then
      shipmentA.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBillBuy" Then
      FrmBillBuy.show
    End If




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmReturnpurchases" Then
      FrmReturnpurchases.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "Ageng_all" Then
      Ageng_all.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustomerType" Then
  FrmCustomerType.Indx = 0
      FrmCustomerType.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCreditFacicity" Then
      FrmCreditFacicity.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustCash" Then
      FrmCustCash.show
    End If
Print
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustomerContract" Then
      FrmCustomerContract.show
    End If
Print
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "Ageng1" Then
      Ageng.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSalePriceNames" Then
      FrmSalePriceNames.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "AgengItem" Then
      AgengItem.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "SalesTargetSettings" Then
      SalesTargetSettings.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSalesRePGroups" Then
      FrmSalesRePGroups.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSalesRepData" Then
'      FrmSalesRepData.show
    End If

Print
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmTypeDiscards" Then
      FrmTypeDiscards.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPO" Then
      FrmPO.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPO1" Then
      FrmPO1.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPO2" Then
      FrmPO2.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPO3" Then
      FrmPO3.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmQotation" Then
      FrmQotation.show
    End If
Print
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmsalebill" Then
      frmsalebill.show
    End If
 
  If formname = "frmsalebill1" Then
      frmsalebill1.show
    End If
 
   If formname = "frmsalebill2" Then
      frmsalebill2.show
    End If
   If formname = "frmsalebill3" Then
      frmsalebill3.show
    End If
   If formname = "frmsalebill4" Then
      frmsalebill4.show
    End If
 
    If formname = "frmsalebill5" Then
      frmsalebill5.show
    End If
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmReturnSalling" Then
      FrmReturnSalling.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmsalebillCompose" Then
      frmsalebillCompose.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "Frmovers" Then
      Frmovers.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSallingPlan" Then
      FrmSallingPlan.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "shipmentA" Then
      shipmentA.show
    End If
Print
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSalesRepComm" Then
      FrmSalesRepComm.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSalesRepCommtarget" Then
      FrmSalesRepCommtarget.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSalesRepCommtargetPercentage" Then
      FrmSalesRepCommtargetPercentage.show
    End If
Print
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSalesRepCommValues" Then
      FrmSalesRepCommValues.show
    End If
Print
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "Ageng_all1" Then
      Ageng_all.show
    End If
Print
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustomerReports" Then
      FrmCustomerReports.show
    End If
Print
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBanksCheck" Then
      FrmBanksCheck.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmExpenses3" Then
      FrmExpenses3.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmserviceInvoice" Then
      frmserviceInvoice.show
    End If
Print
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmDataTypeExchange" Then
      FrmDataTypeExchange.show
    End If
Print
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmTypeExchange" Then
      FrmTypeExchange.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmExpenses5" Then
      FrmExpenses5.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPayments" Then
      FrmPayments.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAccEditJournal3" Then
      FrmAccEditJournal3.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCashing" Then
      FrmCashing.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmGeneralFundReceipt" Then
      FrmGeneralFundReceipt.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "PrintCheque" Then
      PrintCheque.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBankDeposite" Then
      FrmBankDeposite.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBankDeposite1" Then
      FrmBankDeposite1.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBankAdj" Then
      FrmBankAdj.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "BankSettlementt" Then
      BankSettlementt.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPaymentTime" Then
      FrmPaymentTime.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmDiscounts" Then
      FrmDiscounts.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmReceiptPart" Then
      FrmReceiptPart.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmInstallmentMustPay" Then
      FrmInstallmentMustPay.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPayments1" Then
      FrmPayments1.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmExpenses30" Then
      FrmExpenses30.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBoxDrawing" Then
      FrmBoxDrawing.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBoxesAccounts" Then
      FrmBoxesAccounts.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBoxStock" Then
   '   FrmBoxStock.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBoxIncapacity" Then
      FrmBoxIncapacity.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmTimeSetting" Then
'      FrmTimeSetting.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmVacancy" Then
      FrmVacancy.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "emp_CONTRACT_TYPE" Then
      emp_CONTRACT_TYPE.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "jobstatus" Then
      jobstatus.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmpDepartments" Then
      FrmEmpDepartments.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmpJobsTypes" Then
      FrmEmpJobsTypes.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmpSpecifications" Then
      FrmEmpSpecifications.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmpGrade" Then
      FrmEmpGrade.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "insurancecompanies" Then
      insurancecompanies.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "insurancetype" Then
  '    insurancetype.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "Insurance_class" Then
      Insurance_class.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmOutType" Then
      FrmOutType.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmGroupsDEp" Then
      FrmGroupsDEp.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "dean" Then
      dean.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAsest" Then
      FrmAsest.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmRelations" Then
      FrmRelations.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSection" Then
      FrmSection.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmVisa" Then
      FrmVisa.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmployee" Then
  FrmEmployee.mIndex = 1
      FrmEmployee.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmEmpContract" Then
      frmEmpContract.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmtakeem" Then
      frmtakeem.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmRating" Then
'      FrmRating.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPresentTime" Then
'      FrmPresentTime.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmpSalary2" Then
   '   FrmEmpSalary2.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAbsent" Then
'      FrmAbsent.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmpMonthShow" Then
   '   FrmEmpMonthShow.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "MOFRAD" Then
      MOFRAD.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "mofradat2" Then
      mofradat2.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmMkafea" Then
      FrmMkafea.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmKhsm" Then
      FrmKhsm.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmChangedComponentData" Then
      FrmChangedComponentData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmChangedComponentData1" Then
      FrmChangedComponentData1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmpIncreaseSalaries" Then
      FrmEmpIncreaseSalaries.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmHolidayPlan" Then
      FrmHolidayPlan.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "formvocatinl" Then
      formvocatinl.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmHolidayData" Then
'      FrmHolidayData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmdriveassest" Then
     ' frmdriveassest.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmVocationEntitlements" Then
      FrmVocationEntitlements.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmbarkation" Then
      FrmEmbarkation.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmpsAdvanceRequest" Then
      FrmEmpsAdvanceRequest.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmpsAdvance" Then
      FrmEmpsAdvance.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmpsAdvancePayed1" Then
      FrmEmpsAdvancePayed1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmRegisterHoliday" Then
      FrmRegisterHoliday.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "End_oF_service" Then
      End_oF_service.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmMovingEmp" Then
      FrmMovingEmp.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBusinessJob" Then
      FrmBusinessJob.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAdvancedHousingpayments" Then
      FrmAdvancedHousingpayments.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FormEmpMoveDepartment" Then
      FormEmpMoveDepartment.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmbarkation" Then
      FrmEmbarkation.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmQUesEmp" Then
      FrmQUesEmp.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPassports" Then
      FrmPassports.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FRmEmployeeWarning" Then
      FRmEmployeeWarning.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmTreament" Then
      FrmTreament.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmRepInjuy" Then
      FrmRepInjuy.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmReceivingTreatment" Then
      FrmReceivingTreatment.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmFinalSettlement" Then
      FrmFinalSettlement.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmChangeEmployeedata" Then
      FrmChangeEmployeedata.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmClearanceCerTifcate" Then
      FrmClearanceCerTifcate.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmFolloAdminMeasure" Then
      FrmFolloAdminMeasure.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmDeductionNote" Then
      FrmDeductionNote.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCarReceipt" Then
      FrmCarReceipt.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmDefineEmp" Then
      FrmDefineEmp.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAccountCharts" Then
      FrmAccountCharts.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAccEditJournal1" Then
      FrmAccEditJournal1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAccEditJournal" Then
      FrmAccEditJournal.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAccEditJournal4" Then
      FrmAccEditJournal4.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCostCenterType1" Then
      FrmCostCenterType1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "CostCenter" Then
      CostCenter.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FixedAssetsGroup" Then
      FixedAssetsGroup.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FixedAssets" Then
      FixedAssets.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmExpenses4" Then
      FrmExpenses4.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCase1" Then
      FrmCase1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmExpenses40E" Then
'      FrmExpenses40E.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmExpenses40A" Then
      FrmExpenses40A.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmNewGard10" Then
      FrmNewGard10.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmFixedAsseteports" Then
      frmFixedAsseteports.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustomerType" Then
  FrmCustomerType.Indx = 0
      FrmCustomerType.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAkarType" Then
      FrmAkarType.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAkarUnit" Then
      FrmAkarUnit.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSalesRePGroups" Then
      FrmSalesRePGroups.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSalesRepData" Then
'      FrmSalesRepData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCountriesData" Then
      FrmCountriesData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmGovernmentData" Then
      FrmGovernmentData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmGovernCitiesData" Then
      FrmGovernCitiesData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmSchemes" Then
  '    frmSchemes.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "RSPhoneBook" Then
      RSPhoneBook.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "RSOwner" Then
      RSOwner.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "RsCustomers" Then
      RsCustomers.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmExpensesType" Then
      FrmExpensesType.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAkarStatus" Then
      FrmAkarStatus.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmIqarCompnent" Then
      FrmIqarCompnent.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "RSAkar" Then
      RSAkar.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAlarMType" Then
      FrmAlarMType.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "RsOrders" Then
'      RsOrders.show
    End If
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmComparePrices" Then
      FrmComparePrices.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "RSContract" Then
      RSContract.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCashing1" Then
      FrmCashing1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "RsExpenses" Then
      RsExpenses.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPayments2" Then
      FrmPayments2.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "PrintCheque" Then
      PrintCheque.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmWaiverSettlement" Then
      FrmWaiverSettlement.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "Frmblacklist" Then
      frmblacklist.show
    End If
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmOrderMaintenance" Then
      FrmOrderMaintenance.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmLockedOrderMaintenance" Then
      FrmLockedOrderMaintenance.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAllocationToContract" Then
      FrmAllocationToContract.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAllocationToContract1" Then
      FrmAllocationToContract1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAqarReport" Then
      FrmAqarReport.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAmolatReports" Then
      FrmAmolatReports.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmExpiredContract" Then
      FrmExpiredContract.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmMaintnanceReport" Then
      FrmMaintnanceReport.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmWaiverReport" Then
      FrmWaiverReport.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmConttractTotalService" Then
      FrmConttractTotalService.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmOwnerAqarReport" Then
      FrmOwnerAqarReport.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAlrmReports" Then
      FrmAlrmReports.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmTotalsReport" Then
      FrmTotalsReport.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmOrboon" Then
      FrmOrboon.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCommissionReports" Then
      FrmCommissionReports.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmRentsOwendReports" Then
      FrmRentsOwendReports.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustomerAqarReport" Then
      FrmCustomerAqarReport.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmUnitInfoReports" Then
      FrmUnitInfoReports.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmIncomAndExpenReports" Then
      FrmIncomAndExpenReports.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmContractReport" Then
      FrmContractReport.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustomerBalances1" Then
      FrmCustomerBalances1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSerachUnitEmpty" Then
      FrmSerachUnitEmpty.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmProductionreport" Then
      frmProductionreport.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmTravelRports" Then
      frmTravelRports.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAccountingReport" Then
      FrmAccountingReport.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmDailtyReport" Then
      FrmDailtyReport.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAssbliedInterval" Then
      FrmAssbliedInterval.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmOptions" Then
      FrmOptions.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmDelUser" Then
    '  FrmDelUser.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEditPW" Then
      FrmEditPW.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "System_alarms" Then
      System_alarms.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmMessnger" Then
      FrmMessnger.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
  If formname = "cachierData" Then
      cachierData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frm_sheft" Then
      frm_sheft.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmTables" Then
      FrmTables.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "CashierLogin" Then
      CashierLogin.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBankDeposite3" Then
      FrmBankDeposite3.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmGovernmentData" Then
      FrmGovernmentData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCitiesDistance" Then
      FrmCitiesDistance.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmDrivers" Then
      FrmDrivers.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "CarTypes" Then
      CarTypes.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FRMMaintenanceTypes" Then
  '    FRMMaintenanceTypes.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCars" Then
      FrmCars.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmTravelTransactions" Then
      FrmTravelTransactions.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEstimations" Then
      FrmEstimations.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "Cash_flow" Then
      Cash_flow.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBalanceSheet" Then
      FrmBalanceSheet.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAccountDestribution" Then
      FrmAccountDestribution.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FinancialAnalysis" Then
      FinancialAnalysis.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FinancialAnalysisView" Then
      FinancialAnalysisView.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCompositeAccounts" Then
      FrmCompositeAccounts.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmStatistics" Then
 
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustomersAgenda" Then
     ' FrmCustomersAgenda.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmLoadTrialBalance" Then
      FrmLoadTrialBalance.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmExpensesAdvanced" Then
      'FrmExpensesAdvanced.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmExpensespaidAdvanced" Then
    '  FrmExpensespaidAdvanced.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmItServiceTicket" Then
      FrmItServiceTicket.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "Frmcameralocation" Then
      Frmcameralocation.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBarcode" Then
      FrmBarcode.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPrintItemsBarcodes" Then
      'FrmPrintItemsBarcodes.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "SMSSeTTings" Then
      SMSSeTTings.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPlainMessage" Then
      FrmPlainMessage.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmDEfineMessage" Then
      FrmDEfineMessage.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustomerBalances1" Then
      FrmCustomerBalances1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmToolsSerials" Then
      'FrmToolsSerials.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmToolsCustomers" Then
     ' FrmToolsCustomers.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmToolsRepireItemsCost" Then
     ' FrmToolsRepireItemsCost.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "AdminLogin" Then
      AdminLogin.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmDataBaseTools" Then
      FrmDataBaseTools.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmyaersData" Then
      FrmyaersData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBranchesData" Then
      FrmBranchesData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "baranches" Then
      baranches.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAccountsSeetting" Then
      FrmAccountsSeetting.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmDocType" Then
      FrmDocType.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "System_manger2" Then
  '    System_manger2.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "coding" Then
      Coding.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "CarTypes" Then
      CarTypes.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCarModels" Then
      FrmCarModels.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCars" Then
      FrmCars.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmMaintenensWork" Then
      FrmMaintenensWork.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmTypeExtraExpenses" Then
      FrmTypeExtraExpenses.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmExtraExpenses" Then
      FrmExtraExpenses.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmColor" Then
      FrmColor.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmStoreData" Then
      FrmStoreData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmGroups" Then
      FrmGroups.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSystemUnites" Then
'      FrmSystemUnites.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmItems" Then
      FrmItems.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustemers" Then
      FrmCustemers.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmEmployee" Then
      FrmEmployee.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmcarEmpDepartments" Then
      FrmcarEmpDepartments.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmSuperVisor" Then
      FrmSuperVisor.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCarAuthontication" Then
      FrmCarAuthontication.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBillComputerChek" Then
      FrmBillComputerChek.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmOut" Then
      FrmOut.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPO10" Then
      FrmPO10.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmBillCarMaintExtra" Then
      FrmBillCarMaintExtra.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCommisRece" Then
      FrmCommisRece.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustomerssFollow" Then
      FrmCustomerssFollow.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustomerssComplaint" Then
      FrmCustomerssComplaint.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCountriesData" Then
      FrmCountriesData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmGovernmentData" Then
      FrmGovernmentData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCitiesDistance" Then
      FrmCitiesDistance.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmGovernCitiesData" Then
      FrmGovernCitiesData.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "streets" Then
      streets.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "CarTypes" Then
      CarTypes.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCars" Then
      FrmCars.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmDrivers" Then
      FrmDrivers.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmTypesofshipping" Then
      FrmTypesofshipping.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FRMMaintenanceTypes" Then
   '   FRMMaintenanceTypes.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmShipmentFollow" Then
      frmShipmentFollow.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "frmSipmentAllocation" Then
      frmSipmentAllocation.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmProductionPlan" Then
      FrmProductionPlan.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmShipmentOrder" Then
      FrmShipmentOrder.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmShipmentRegestration" Then
      FrmShipmentRegestration.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmShipmentRegestration1" Then
      FrmShipmentRegestration1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmShipmentRegestration1" Then
      FrmShipmentRegestration1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "all_alarms" Then
      all_alarms.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "ProjectsAlarm1" Then
      ProjectsAlarm1.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "ProjectsBillAlarm" Then
      ProjectsBillAlarm.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "Ageng_all" Then
      Ageng_all.show
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "RSRentAlarm" Then
      RSRentAlarm.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAccountDestributionView" Then
      FrmAccountDestributionView.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPaymentTime" Then
      FrmPaymentTime.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmAccreditOrder" Then
      frmaccreditOrder.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmInstallmentMustPay" Then
      FrmInstallmentMustPay.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmCustomerBalances" Then
      FrmCustomerBalances.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmRequest" Then
      FrmRequest.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmStagnantItems" Then
      FrmStagnantItems.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmGuaranteeAlram" Then
      FrmGuaranteeAlram.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmPerfMantAlaram" Then
      FrmPerfMantAlaram.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmMantinanceReport" Then
      FrmMantinanceReport.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "Car_alarms" Then

    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If formname = "FrmApprovalTransactions" Then
      FrmApprovalTransactions.show
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
      
         
rsMenue.Close
Set rsMenue = Nothing
      
    End If




End Function

 
Public Function showFavoritesMenue()
On Error Resume Next
Dim i As Integer


        
For i = 1 To 30
help_list(i).Visible = False
Next i

     Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim NoOfMinute As Double
    'Dim TimeCateg As Double
 
    sql = "SELECT   Displayname "
sql = sql & " from dbo.TblMyMenue "
sql = sql & " WHERE   USERID=" & user_id & " order by id"
     rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
Dim torecord As Integer
    If rs.RecordCount > 0 Then
    torecord = rs.RecordCount
        If torecord > 30 Then torecord = 30
        
   For i = 1 To torecord
                If Not IsNull(rs("Displayname").value) Then
                            help_list(i).Visible = True
                            help_list(i).Caption = IIf(IsNull(rs("Displayname").value), 0, rs("Displayname").value)
                 End If
 rs.MoveNext
   Next i
         
         
    Else
     
    End If




End Function
Private Sub MDIForm_Load()
    Dim BGround As ClsBackGroundPic
    Dim BolShowRequest As Boolean
    'On Local Error GoTo ErrTrap
    On Error Resume Next
    Me.backcolor = vbWhite
    Me.Caption = GetAppTitle  'App.Title
   CreateDocks
    loadmyModule
    
   messengerTime = 0
      AlarmAutoTime = 0
    
'RemoveMenus Me, True, True, True, False, True, True, False
  
  
    LoadInterface SystemOptions.UserInterface

'    If Messnger = True Then mdifrmmain.Timer1.Enabled = True: FrmMessnger.show
'«·Õ Â œÌ » ÂÌ” ›Ì Œ·›Ì… «·ﬁÊ«Ì„ »”  €Ì— «”„ «·’Ê—Â
    BackGroundImag = GetSetting(StrAppRegPath, "View_Type", "BackGroundImag", App.path & "\Garphics\logoMain.jpg")
If onLineMOde = True Then
'BackGroundImag = App.path & "Garphics\" & onlinebackground
BackGroundImag = App.path & onlinebackground
End If

    If Dir(BackGroundImag) <> "" Then
          Me.Picture = LoadPicture(App.path & "\Garphics\wallpaper_Main.jpg")
        Me.Picture = LoadPicture(BackGroundImag)
        '  WebForm.Picture = LoadPicture(BackGroundImag)
        'AskOption
        Set Me.PopMenu1.BackgroundPicture = LoadPicture(App.path & "\Garphics\wallpaper_Main.jpg")
Else
'Me.Picture = LoadPicture(App.path & "\Garphics\wallpaper_Main.jpg")
    
            If Dir(App.path & "\Garphics\wallpaper_Main.jpg") <> "" Then
                     ' Image1.Picture = LoadPicture(App.path & "\Garphics\wallpaper_Main.jpg")
                 '     Me.Picture = LoadPicture(App.path & "\Garphics\wallpaper_Main.jpg")
                 End If
                 
    
    
    End If

    'Grid_WallPaper.jpg
   If Dir(App.path & "\Garphics\Grid_WallPaper.jpg") <> "" Then
          Set Me.PopMenu1.BackgroundPicture = LoadPicture(App.path & "\Garphics\Grid_WallPaper.jpg")
    End If

    'If Dir(App.Path & "\ReportDesign.exe") = "" Then
    '    ReportDesigner.Visible = False
    '    Sep30.Visible = False
    'End If
showFavoritesMenue
LoadMainSystemOptions
         mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.OutBarPaneID).Closed = 1
    mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID).Closed = 1

    mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.CalendarPaneID).Closed = 1
    mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.NewsBarPaneID).Closed = 1
          If Dir(App.path & "\Garphics\Grid_WallPaper.jpg") <> "" Then
          Set Me.PopMenu1.BackgroundPicture = LoadPicture(App.path & "\Garphics\Grid_WallPaper.jpg")
    End If
    
    Exit Sub
ErrTrap:

    If SystemOptions.SysRegisterState = DevelopVersion Then
        Stop
        Resume
    End If

    connection_string = Cn.ConnectionString
    



End Sub

Private Sub MDIForm_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              Y As Single)
    'xx.Hide
 '   showmnue


End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, _
                                UnloadMode As Integer)

    If UnloadMode <> VBRUN.QueryUnloadConstants.vbFormCode Then
         If AskForExit = False Then
            Cancel = True
            Exit Sub
        Else

        End If
    End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    Dim FreeF As Integer, sFile As String, sLayout As String
    sFile = App.path & "\Layout.000"
    FreeF = FreeFile

    If Dir(sFile, vbNormal) <> "" Then
        Kill sFile
    End If

    Open sFile For Binary As #FreeF
    Put #FreeF, , Me.DockingPane1.SaveStateToString
    Close #FreeF
End Sub

 

Private Sub MnuAccCharts_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("FrmAccountCharts") = False Then
                Exit Sub
            End If

            FrmAccountCharts.show

        Case 1

            If checkApility("FrmAccEditJournal1") = False Then
                Exit Sub
            End If

            FrmAccEditJournal1.show
    End Select

End Sub

Private Sub MnuAccDEV_Click(Index As Integer)

    Select Case Index

        Case 0

            'frmsandat_ked2.Show
            'frmsandat_ked.Show
            If checkApility("FrmAccEditJournal") = False Then
                Exit Sub
            End If

            FrmAccEditJournal.show

        Case 1
           If checkApility("FrmAccEditJournal4") = False Then
                Exit Sub
            End If

            FrmAccEditJournal4.show
            
          '  keddawrym.show

    End Select

End Sub

Private Sub MnuAccDEV_Post_Click()
  ' Frm_General_Journal.show
End Sub

Private Sub MnuAccIntervals_Click()
   ' FrmAccountIntervals.show
End Sub

Private Sub MnuAccReports_Click()

End Sub

Private Sub MnuBasicCitiesData_Click(Index As Integer)

End Sub

Private Sub MnuBoxDeposit_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("FrmBoxDeposit") = False Then
                Exit Sub
            End If

            'FrmBoxDeposit.show
            'FrmBoxDeposit.ZOrder 0

        Case 1

            If checkApility("FrmPayments1") = False Then
                Exit Sub
            End If

            FrmPayments1.show

        Case 2
 
            If checkApility("FrmExpenses30") = False Then
                Exit Sub
            End If

            FrmExpenses30.show

    End Select

End Sub

Private Sub MnuBoxDetectErrors_Click()

    If checkApility("FrmBoxDetetErrors") = False Then
        Exit Sub
    End If

    'FrmBoxDetetErrors.show
End Sub

Private Sub MnuBoxIncapacity_Increase_Click(Index As Integer)
Select Case Index
 Case 0
    If checkApility("FrmBoxIncapacity") = False Then
        Exit Sub
    End If

    FrmBoxIncapacity.show
    
 Case 1
  
End Select
End Sub

Private Sub MnuBoxStock_Click()

    If checkApility("FrmBoxStock") = False Then
        Exit Sub
    End If

    OpenScreen BoxesStockScreen
End Sub

Private Sub MnuCheckBriefcase_Click()
    'FrmChecksBriefcase.show
End Sub

Private Sub MNUCloseYear_Click()
'    FrmClose.show
End Sub

Private Sub MnuCorrectSerial_Click()

    If checkApility("FrmToolsSerials") = False Then
        Exit Sub
    End If

    'FrmToolsSerials.show
End Sub

Private Sub MnuCurrencyData_Click()

End Sub

Private Sub MnuCusTools_Item_Click(Index As Integer)
    Dim LngCusID As Long
    Dim IntDealerType As Integer

    LngCusID = val(Me.MnuCusTools.Tag)

    If LngCusID = 0 Then Exit Sub

    Select Case Index

        Case 0
            'ﬂ‘› Õ”«» «·⁄„Ì·
            ShowCusBalDailog LngCusID, 0

        Case 1

            'Menu Sep
        Case 2
            '›Ê« Ì— „»Ì⁄«  «·⁄„Ì·
            ShowCusBalDailog LngCusID, 1

        Case 3
            ShowCusBalDailog LngCusID, 2

        Case 4

            'Menu Sep
        Case 5
            ShowCusBalDailog LngCusID, 3

        Case 6
            ShowCusBalDailog LngCusID, 4

        Case 7

            'Menu Sep
        Case 8
            ShowCusBalDailog LngCusID, 5

        Case 9
            ShowCusBalDailog LngCusID, 6
        
        Case Me.MnuCusTools_Item.UBound
            IntDealerType = GetDealerType(LngCusID)

            If IntDealerType = 1 Then
                OpenScreen CustomersScreen, LngCusID
            ElseIf IntDealerType = 2 Then
                OpenScreen SuppliersScreen, LngCusID
            End If

    End Select

End Sub

Private Sub MnuDataBaseTools_Click()
    Dim Msg As String
    'If user_id <> 1 Then
        '   MsgBox ""
'        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
        '    Msg = Msg & Chr(13) & "Õ ‰Â“— Ê·««ÌÂ "
'        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        Exit Sub
'    End If
    
    If checkApility("FrmDataBaseTools") = False Then
        Exit Sub
    End If

    If Me.ActiveForm Is Nothing Then
        FrmDataBaseTools.show vbModal
    Else
        Msg = "ÌÃ» €·ﬁ «Ï ‘«‘… „‰ ‘«‘«  «·»—‰«„Ã ﬁ»·"
        Msg = Msg & Chr(13) & "«‰  ” Œœ„ Â–« «·‘«‘…....!!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

End Sub

Private Sub MnuEmpDepartmentData_Click()

End Sub

Private Sub MnuEmpJobsData_Click()

End Sub

Private Sub MnuEmpsEmpTimeSeeting_Click()

End Sub

Private Sub MnueHouseMainSub_Click(Index As Integer)
Select Case Index
 
End Select
End Sub

Private Sub MnuElevatorssUB_Click(Index As Integer)
    Dim Msg As String
 
    Select Case Index

        Case 0
      If checkApility("FrmScreenCriteria") = False Then
                 Exit Sub
             End If
            FrmScreenCriteria.show

        Case 1
             If checkApility("frmScreenCreteriaSettings") = False Then
                 Exit Sub
             End If
               frmScreenCreteriaSettings.show
           Case 2
                 If checkApility("FrmQotation") = False Then
                 Exit Sub
             End If
FrmQotation.show

 

           Case 3
                       If checkApility("FrmShowTech") = False Then
                 Exit Sub
             End If
             
                      Load FrmShowTech
                      
                      
              Case 4
              
                
'    If checkApility("FrmPerfMantAlaram") = False Then
'        Exit Sub
'    End If

'    FrmPerfMantAlaram.show
    
    Case 6
 

           If checkApility("FrmReports") = False Then
                Exit Sub
            End If

            FrmReports.show
            FrmReports.C1TabMain.CurrTab = 0
            


                
    End Select
    
    
    
End Sub

Private Sub mnuEmployeeBasic_Click(Index As Integer)
Select Case Index

Case 10
     If checkApility("FrmComponentYear") = False Then
                Exit Sub
            End If
            FrmComponentYear.show
            
Case 11
     If checkApility("ReportEmployees") = False Then
                Exit Sub
            End If

            FrmReports.show
            FrmReports.C1TabMain.CurrTab = 16
End Select
End Sub

Private Sub mnuEmployeeBasicSub_Click(Index As Integer)

    Select Case Index

        Case 0
          '  Dim Frm As FrmTimeSetting

          '  If checkApility("FrmTimeSetting") = False Then
          '      Exit Sub
          '  End If
'
'            Set Frm = New FrmTimeSetting

           ' Frm.WorkType = 0
'            Frm.show
           ' Frm.ZOrder 0

        Case 1

            If checkApility("frm_sheft") = False Then
                Exit Sub
            End If

            frm_sheft.show

        Case 2

            If checkApility("FrmVacancy") = False Then
                Exit Sub
            End If

            FrmVacancy.show
            FrmVacancy.ZOrder 0

        Case 3

            If checkApility("emp_CONTRACT_TYPE") = False Then
                Exit Sub
            End If
            emp_CONTRACT_TYPE.mIndex = 0
            emp_CONTRACT_TYPE.show

        Case 4

            If checkApility("jobstatus") = False Then
                Exit Sub
            End If
 
            jobstatus.show

        Case 5

            If checkApility("FrmEmpDepartments") = False Then
                Exit Sub
            End If
            
            FrmEmpDepartments.show

        Case 6

            If checkApility("FrmEmpJobsTypes") = False Then
                Exit Sub
            End If
            
            FrmEmpJobsTypes.show

        Case 7

            If checkApility("FrmEmpSpecifications") = False Then
                Exit Sub
            End If
            
            FrmEmpSpecifications.show

Case 8

   If checkApility("FrmEmpGrade") = False Then
                Exit Sub
            End If

FrmEmpGrade.show
        Case 9

            If checkApility("insurancecompanies") = False Then
                Exit Sub
            End If
            
            insurancecompanies.show

        Case 10

            If checkApility("insurancetype") = False Then
                Exit Sub
            End If
            
'            insurancetype.show

        Case 11

            If checkApility("Insurance_class") = False Then
                Exit Sub
            End If
            
            Insurance_class.show

        Case 12

'            If checkApility("frmtakeem") = False Then
'                Exit Sub
'            End If
'
'            frmtakeem.show

Case 13

       If checkApility("FrmOutType") = False Then
                Exit Sub
            End If
            
FrmOutType.show
Case 14
       If checkApility("FrmGroupsDEp") = False Then
                Exit Sub
            End If
FrmGroupsDEp.show


        Case 15

            If checkApility("Nationality") = False Then
                Exit Sub
            End If

            Nationality.show

        Case 16

            If checkApility("dean") = False Then
                Exit Sub
            End If

            dean.show
            
        Case 17
                If checkApility("FrmAsest") = False Then
                Exit Sub
            End If
       FrmAsest.show
        
           Case 18
                      If checkApility("FrmRelations") = False Then
                Exit Sub
            End If
           FrmRelations.show
           
    Case 19
                      If checkApility("FrmSection") = False Then
                Exit Sub
            End If
           FrmSection.show
           
           Case 20
                            If checkApility("FrmVisa") = False Then
                Exit Sub
            End If
           FrmVisa.show
           
           Case 21
                            If checkApility("FrmAdminSanction") = False Then
                Exit Sub
            End If
           FrmAdminSanction.show
           Case 22
                          If checkApility("FrmSickleave") = False Then
                Exit Sub
            End If
            
           FrmSickleave.show
           
                 Case 23
                          If checkApility("FrmVacationSettings") = False Then
                Exit Sub
            End If
            
           FrmVacationSettings.show
           
           
           Case 24
            
           FrmVizitScreen.mIndex = 2
FrmVizitScreen.show
    End Select

End Sub

Private Sub mnuEmployeeBasict_Click(Index As Integer)
Select Case Index

Case 0
 If checkApility("FrmEvaluation_Standerd") = False Then
                Exit Sub
            End If

            FrmEvaluation_Standerd.show
Case 1
 If checkApility("FrmEvaluation") = False Then
                Exit Sub
            End If
FrmEvaluation.show
Case 2
  If checkApility("FrmEvaluaEntit") = False Then
                Exit Sub
            End If
FrmEvaluaEntit.show
' If checkApility("FrmRating") = False Then
'                Exit Sub
'            End If
'FrmRating.show

' If checkApility("FrmChangedComponentData4") = False Then
'                Exit Sub
'            End If
'FrmChangedComponentData4.show
Case 2

End Select

End Sub

Private Sub mnuEmployeInsuranceSub_Click(Index As Integer)
Select Case Index

Case 0
            If checkApility("FrmSocialInsurance") = False Then
                Exit Sub
            End If
            
            
FrmSocialInsurance.show
        Case 1

            If checkApility("insurancecompanies") = False Then
                Exit Sub
            End If
            
            insurancecompanies.show

        Case 2

  '          If checkApility("insurancetype") = False Then
  '              Exit Sub
  '          End If
  '
  '          insurancetype.show
'
        Case 3

            If checkApility("Insurance_class") = False Then
                Exit Sub
            End If
            
            Insurance_class.show

       Case 4

            If checkApility("FrmInsurances") = False Then
                Exit Sub
            End If
            
            FrmInsurances.show
            


End Select
End Sub



 
Private Sub MnuInvSalesOptions_Click()
 
End Sub

Private Sub MnuItemTools_ItemCart_Click()
    Dim VarTemp As Variant
    Dim StrTemp As String
    Dim LngItemID As Long
    Dim LngStoreID As Long
    Dim d_StartDate As Variant
    Dim d_EndDate As Variant

    If MnuItemTools_ItemCart.Tag <> "" Then
        StrTemp = MnuItemTools_ItemCart.Tag
        VarTemp = Split(StrTemp, "-", , vbTextCompare)
        LngItemID = val(VarTemp(0))
        LngStoreID = val(VarTemp(1))

        If UBound(VarTemp) > 2 Then
            If IsDate(VarTemp(2)) Then
                d_StartDate = CDate(VarTemp(2))
            Else
                d_StartDate = Null
            End If
        End If

        If UBound(VarTemp) > 2 Then
            If IsDate(VarTemp(3)) Then
                d_EndDate = CDate(VarTemp(3))
            Else
                d_EndDate = Null
            End If
        End If

        OpenScreen PopUpShowItemCardScreen, LngItemID, LngStoreID, , d_StartDate, d_EndDate, 0
    End If

End Sub

Private Sub MnuItemTools_ItemCostTrans_Click()
    Dim Msg As String

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        OpenScreen PopUpShowItemCostScreen, val(Me.MnuItemTools_ItemCostTrans.Tag)
    Else
        Msg = "⁄›Ê« ...Â–Â «·≈„ﬂ«‰Ì… €Ì— „ «Õ… ›Ï ‰”Œ… «·√ﬂ””....!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

End Sub

Private Sub MnuItemTools_ItemData_Click()
    Dim LngCurrentItemID As Long
    LngCurrentItemID = val(MnuItemTools_ItemData.Tag)

    If LngCurrentItemID <> 0 Then
        OpenScreen ItemsDataScreen, LngCurrentItemID
                 FrmItems.CALLEDFPRM = True
    End If

End Sub
'MnuItemTools_Sep

Private Sub MnuItemTools_Sep_Click()
    Dim LngCurrentItemID As Long
    LngCurrentItemID = val(MnuItemTools_ItemQty.Tag)

    If LngCurrentItemID <> 0 Then
        OpenScreen CheckItemswaped, LngCurrentItemID
    End If

End Sub

Private Sub MnuItemTools_ItemQty_Click()
    Dim LngCurrentItemID As Long
    LngCurrentItemID = val(MnuItemTools_ItemQty.Tag)

    If LngCurrentItemID <> 0 Then
        OpenScreen CheckItemQty, LngCurrentItemID
    End If

End Sub

Private Sub MnuItemTools_ItemSerial_Click()
    Dim VarTemp As Variant

    If MnuItemTools_ItemSerial.Tag <> "" Then
        VarTemp = Split(Me.MnuItemTools_ItemSerial.Tag, "-", , vbTextCompare)
        OpenScreen CheckItemSerial, val(VarTemp(0)), Trim(VarTemp(1))
    End If

End Sub

Private Sub MnuItemTools_Reports_Click(Index As Integer)
    Dim VarTemp As Variant
    Dim StrTemp As String
    Dim LngItemID As Long
    Dim LngStoreID As Long
    Dim d_StartDate As Variant
    Dim d_EndDate As Variant

    If MnuItemTools.Tag <> "" Then
        StrTemp = MnuItemTools.Tag
        VarTemp = Split(StrTemp, "-", , vbTextCompare)
        LngItemID = val(VarTemp(0))

        '    LngStoreID = Val(VarTemp(1))
        '    If UBound(VarTemp) > 2 Then
        '        If IsDate(VarTemp(2)) Then
        '            d_StartDate = CDate(VarTemp(2))
        '        Else
        '            d_StartDate = Null
        '        End If
        '    End If
        '    If UBound(VarTemp) > 2 Then
        '        If IsDate(VarTemp(3)) Then
        '            d_EndDate = CDate(VarTemp(3))
        '        Else
        '            d_EndDate = Null
        '        End If
        '    End If
        Select Case Index

            Case 0
                OpenScreen PopUpShowItemCardScreen, LngItemID, , , Null, Null, 2

            Case 1
                OpenScreen PopUpShowItemCardScreen, LngItemID, , , Null, Null, 3

            Case 2

                'Mnu Sep
            Case 3
                OpenScreen PopUpShowItemCardScreen, LngItemID, , , Null, Null, 5

            Case 4
                OpenScreen PopUpShowItemCardScreen, LngItemID, , , Null, Null, 6
        End Select

    End If

End Sub

Private Sub MnuManCompanies_Click(Index As Integer)

End Sub

Private Sub MnuLevelsSub_Click(Index As Integer)


    Select Case Index

        Case 0
            frm_Levels.show

        Case 1
            frmDocApproval.show
    End Select



End Sub

Private Sub MnuLevelsSub2_Click(Index As Integer)
    Select Case Index

        Case 0
            FrmScreenCriteria.show

        Case 1
               frmScreenCreteriaSettings.show
    End Select
End Sub

Private Sub MnuMaintnanceBasicSub_Click(Index As Integer)

    Select Case Index

        Case 0
             If checkApility("FrmMaintenTypes") = False Then
                Exit Sub
            End If
            
FrmMaintenTypes.show

 
 Case 1
       If checkApility("CarTypes") = False Then
                Exit Sub
            End If

            CarTypes.show
            
Case 2
  If checkApility("FrmCarModels") = False Then
                Exit Sub
            End If

FrmCarModels.show


 Case 3


  If checkApility("FrmColor") = False Then
              Exit Sub
           End If
 FrmColor.show

Case 4


            If checkApility("FrmCars") = False Then
              Exit Sub
           End If
            FrmCars.show
            


 
 
 Case 5
             If checkApility("FrmcarEmpDepartments") = False Then
                Exit Sub
            End If
    
            FrmcarEmpDepartments.show
Case 6
    If checkApility("FrmSuperVisor") = False Then
                Exit Sub
            End If
            FrmSuperVisor.xid = 1
          FrmSuperVisor.show
 


        Case 7
        
           If checkApility("FrmCompany") = False Then
        Exit Sub
    End If
FrmCompany.mIndex = 0
    FrmCompany.show

      Case 8
 
     If checkApility("FrmDataTypeExchange") = False Then
                Exit Sub
            End If

            FrmDataTypeExchange.show
            
            
            
      Case 9
        If checkApility("FrmStoreData") = False Then
                Exit Sub
            End If

          
            OpenScreen StoresDataScreen
            
  Case 10
    If checkApility("FrmGroups") = False Then
                Exit Sub
            End If
                        OpenScreen ItemsGroupsScreen
                        
   Case 11
       If checkApility("FrmItems") = False Then
                Exit Sub
            End If

            OpenScreen ItemsDataScreen
            
    
   Case 12
       If checkApility("project_status") = False Then
                Exit Sub
            End If

          project_status.mIndex = 1
                project_status.show
    End Select

End Sub

Private Sub MnuMaintnanceBasicSub1_Click()


End Sub

Private Sub MnuMaintnanceTransactions_Click(Index As Integer)

    Select Case Index
    
    Case 0
         'publicCarId = val(Me.XPTxtID.text)
              If checkApility("FrmCarsPlan") = False Then
                Exit Sub
            End If
            FrmCarsPlan.show
            
    Case 1
     If checkApility("FrmRequerMainten") = False Then
                Exit Sub
            End If
FrmRequerMainten.show
        Case 2
             If checkApility("FrmOrderMaintin") = False Then
                Exit Sub
            End If
FrmOrderMaintin.show

     '       Load FrmManAddNew
     '       FrmManAddNew.TxtModFlg.text = "N"
     '       FrmManAddNew.show
            
        Case 3

'            If checkApility("FrmManStore") = False Then
'                Exit Sub
'            End If
'
'            FrmManStore.show
'            FrmManStore.ZOrder 0
        If checkApility("FrmPO6") = False Then
                Exit Sub
            End If

            FrmPO6.show
            
        Case 4

            If checkApility("FrmInpout") = False Then
                Exit Sub
                
                 
            End If

            FrmInpout.show
            
           ' FrmOut.TxtTicketNO.Visible = True
           ' FrmOut.lbl(32).Visible = True
              
        Case 5
                    If checkApility("FrmOut") = False Then
                Exit Sub
                
                 
            End If

            FrmOut.show
            
'            FrmManCusRecive.show

        Case 6
       If checkApility("project_status") = False Then
                Exit Sub
            End If

          project_status.mIndex = 2
                project_status.show
        Case 7
'            FrmManOpenBalance.show

        Case 8
                    If checkApility("frmdriveassestMove") = False Then
                Exit Sub
            End If

           ' FrmFixedAssetMoving.show
frmdriveassestMove.show
          '  FrmManStoreStock.show
 Case 9
                  If checkApility("FrmMovingEmp2") = False Then
                Exit Sub
                  
            End If
         FrmMovingEmp2.show
         
        Case 10
                  If checkApility("FrmWarrantyOffer") = False Then
                Exit Sub
                  
            End If
         FrmWarrantyOffer.show
         '   FrmManAlram.show

            'FrmItemTip.Show
            ' √ÀÌ— ›« Ê—… ‘—«¡ «Ê —’Ìœ ≈›  «ÕÏ ›Ï √—»«Õ ›Ê« Ì— «·„»Ì⁄« 
            'FrmItemPurCostEffect.Show
            'FrmReportControl.Show
            '⁄—÷ „ Ê”ÿ «· ﬂ·›… ·’‰›
            'FrmItemCostShow.Show

            'FrmItemsCostUpdate.Show
            '«Õ’«∆Ì«  ÃÌœ…
            'FrmStatistics.Show
 
            '«Ã‰œÂ «·⁄„·«¡
            ' FrmCustomersAgenda.Show

            ' CALENDERCONVERT.Show
            '‰ﬁ·  «·⁄„·«¡ Ê«‰‘«¡ Õ”«»« Â„
            'Form1.Show
Case 11
            If checkApility("FrmAccidentReport") = False Then
                Exit Sub
            End If
FrmAccidentReport.show
        Case 12

            If checkApility("FrmMantinanceReport") = False Then
                Exit Sub
            End If
FrmMantinanceReport.show

            '    FrmManStore.Show
            '    FrmManStore.ZOrder 0
    '        FrmReports.show
    '        FrmReports.C1TabMain.CurrTab = 4

    End Select

End Sub

Private Sub MnuMaintnanceTransactionssub_Click(Index As Integer)
Select Case Index
'Case 0
'     If checkApility("FrmRequerMainten") = False Then
'                Exit Sub
'            End If
'FrmRequerMainten.show
'Case 1
'     If checkApility("FrmOrderMaintin") = False Then
'                Exit Sub
'            End If
'FrmOrderMaintin.show
End Select
End Sub

Private Sub MnuManTools2Sub1_Click()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim LngTableID As Long

    LngTableID = val(Me.MnuManTools2.Tag)

    If LngTableID = 0 Then
        Exit Sub
    End If

    StrSQL = "Select * From TblManAlram Where TableID=" & LngTableID
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs("State").value = 2
        rs("DoneDate").value = Now
        rs("DoneUserID").value = user_id
        rs("DoneMsg").value = " „ «· Ã„Ì⁄"
        rs.update
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub MnuManTools2Sub2_Click()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim LngTableID As Long

    LngTableID = val(Me.MnuManTools2.Tag)

    If LngTableID = 0 Then
        Exit Sub
    End If

    StrSQL = "Select * From TblManAlram Where TableID=" & LngTableID
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs("State").value = 3
        rs("ReleaseDate").value = Now
        rs("ReleaseUserID").value = user_id
        rs.update
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub MnuManToolsSub5_Click()
    Dim VarTemp As Variant
    Dim StrTemp  As String

    StrTemp = Me.MnuManTools.Tag

    If StrTemp = "" Then Exit Sub
    VarTemp = Split(StrTemp, "-", , vbTextCompare)

   ' Load FrmManEmpReport
   ' FrmManEmpReport.TxtOrgManID.text = val(VarTemp(0))
   ' FrmManEmpReport.TxtTicketNO.text = val(VarTemp(1))
   ' FrmManEmpReport.lblReciptNumber.Caption = val(VarTemp(2))
   ' FrmManEmpReport.show vbModal

End Sub

Private Sub MnuManToolsSub6_Click()
    Dim StrTemp As String
    Dim VarTemp As Variant
    Dim LngItemID As Long
    Dim StrItemSerial  As String

    'If mdifrmmain.MnuManToolsSub6.Tag <> "" Then
        
    '    StrTemp = mdifrmmain.MnuManToolsSub6.Tag
    '    VarTemp = Split(StrTemp, ";", , vbTextCompare)
    '    LngItemID = val(VarTemp(0))
    '    StrItemSerial = Trim$(VarTemp(1))
    '    OpenScreen CheckItemSerial, LngItemID, StrItemSerial
    'End If

End Sub

Private Sub MnuOutBarGroup_Click(Index As Integer)
    Dim YTemp As dxItemLink
    Dim xTemp As dxItem
    Dim IntGroupLinks As Integer

    Dim i As Integer

    Select Case Index

        Case 0
            ModOutBar.AddNewGroup

        Case 1
            ModOutBar.EditGroup

        Case 2
            ModOutBar.DeleteGroup

        Case 3
            ModOutBar.AddItem_Link

        Case 4

        Case 5
            ModOutBar.EditItemLink

        Case 6
            ModOutBar.RemoveItemLink
    End Select

End Sub

Private Sub MnuOutBarStyle_Click(Index As Integer)
    Dim i As Integer
    Dim x As DXSIDEBARLibCtl.IconStyle

    Select Case Index

        Case 0
            x = SmallIcon

        Case 1
            x = LargeIcon
    End Select

   ' For i = 0 To FrmOutBarPane.OutBar.Groups.count - 1
   '     FrmOutBarPane.OutBar.Groups(i).ItemsStyle = X
   ' Next i

    SaveSetting StrAppRegPath, "OutBarOptions", "ItemsStyle", x
End Sub

Private Sub MnuPopItemsTreePane_Array_Click(Index As Integer)
    Dim xPane As XtremeDockingPane.Pane
    Dim IntPaneIndex As Integer
    IntPaneIndex = val(Me.MnuPopPane.Tag)

    If IntPaneIndex = 0 Then
        Exit Sub
    End If

    Select Case Index

        Case 0

            'If Not ItemsTreePane Is Nothing Then
            '    ItemsTreePane.LoadData ItemsTreePane.GroupsSort, ItemsTreePane.ItemsSort
            'End If

        Case 1

            'Sep
        Case 2
            'Hidden
            MnuPopItemsTreePane_Array(Index).Checked = Not (MnuPopItemsTreePane_Array(Index).Checked)
            Me.DockingPane1(IntPaneIndex).Hidden = Not MnuPopItemsTreePane_Array(Index).Checked

        Case 3
            'Close
            Me.DockingPane1(IntPaneIndex).Close
    End Select

End Sub

Private Sub MnuPrintItemsCodes_Click()

    If checkApility("FrmBarcodePrinting") = False Then
        Exit Sub
    End If
 FrmBarcodePrinting.show
End Sub

Private Sub MnuProjectsBasicSub_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("project_status") = False Then
                Exit Sub
            End If
    project_status.mIndex = 3
            project_status.show

        Case 1

            If checkApility("Contract_type") = False Then
                Exit Sub
            End If

            contract_type.show

        Case 2

            If checkApility("FrmOtherCustomers") = False Then
                Exit Sub
            End If

            OpenScreen OtherCustomersScreen '

    Case 3
               If checkApility("FrmPands") = False Then
                Exit Sub
            End If
            FrmPands.show
            
        Case 4
           If checkApility("FrmProcessUnit") = False Then
                Exit Sub
            End If
            FrmProcessUnit.show

        Case 5
           If checkApility("FrmProcessDef") = False Then
                Exit Sub
            End If
            FrmProcessDef.show

        Case 6
            If checkApility("frmequipment") = False Then
                Exit Sub
            End If


            frmequipment.show

Case 7
            If checkApility("frmSubcontractorContractl") = False Then
               Exit Sub
            End If
 
       frmSubcontractorContract.show

    End Select

End Sub

Private Sub MnuProjectsTransactions_Click(Index As Integer)

    Select Case Index

Case 0
            If checkApility("Projects") = False Then
                Exit Sub
            End If

       '     Projects1.show
         Projects.show
        Case 1

            'FrmDestruction
            If checkApility("FrmDestruction") = False Then
                Exit Sub
            End If

            OpenScreen DestructionScreen
Case 2
If checkApility("FrmDestructionRet") = False Then
                Exit Sub
            End If
FrmDestructionRet.show
        Case 3

            If checkApility("FrmEmpSalary3") = False Then
                 Exit Sub
            End If

           
            FrmEmpSalary3.show

        Case 4

            If checkApility("FrmEmpSalary4") = False Then
                Exit Sub
            End If

            FrmEmpSalary4.show


Case 5
   If checkApility("FrmEmpSalary3A") = False Then
                Exit Sub
            End If
FrmEmpSalary3A.show
Case 6
   If checkApility("FrmEmpSalary4A") = False Then
                Exit Sub
            End If
FrmEmpSalary4A.show

 


        Case 7

            If checkApility("FrmOperationsFollow") = False Then
                Exit Sub
            End If

            FrmOperationsFollow.show
 
        Case 8

            If checkApility("projectsbill") = False Then
                Exit Sub
            End If
 
            projectsbill.show

        Case 9

            If checkApility("FrmProjectMonthBill") = False Then
                Exit Sub
            End If
FrmProjectMonthBill.show

        Case 10

            If checkApility("projectsReports") = False Then
                Exit Sub
            End If
frmProjectsReports.show
'         Projects.ShowReports
    
    End Select

End Sub

Private Sub MnuReports_Assblied_Click()
    Dim Msg As String

    If checkApility("FrmAssbliedInterval") = False Then
        Exit Sub
    End If

    FrmAssbliedInterval.show
    FrmAssbliedInterval.ZOrder 0
    
    'If SystemOptions.usertype = UserAdminAll Or SystemOptions.usertype = UserNourCo Then
    '    FrmAssbliedInterval.Show
    '    FrmAssbliedInterval.ZOrder 0
    'Else
    '    Msg = "·«Ì„ﬂ‰ﬂ «· ⁄«„· „⁄ Â–Â «·‘«‘… ...."
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    'End If

End Sub

Private Sub mnuSalesBasic_Click(Index As Integer)
Select Case Index

Case 0

    If checkApility("FrmRegDateDelegateTime") = False Then
        Exit Sub
    End If
    
FrmRegDateDelegateTime.show

Case 1
    If checkApility("FrmTypeVisit") = False Then
        Exit Sub
    End If
FrmTypeVisit.show
Case 2
    If checkApility("FrmSpecialAsement") = False Then
        Exit Sub
    End If
FrmSpecialAsement.show
Case 3
    If checkApility("FrmComponent") = False Then
        Exit Sub
    End If
FrmComponent.show

End Select
End Sub

Private Sub MnuToolCustomers_Click()
    Dim Msg As String

    If checkApility("FrmToolsCustomers") = False Then
        Exit Sub
    End If

    'If SystemOptions.usertype = UserNormal Then
    '    Msg = "ÌÃ» «‰  ﬂÊ‰ ·ﬂ ’·«ÕÌ… „œÌ— Õ Ï  ” ÿÌ⁄ ≈” Œœ«„ Â–Â «·√œ«…"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    Exit Sub
    'End If
    'FrmToolsCustomers.Show
End Sub

Private Sub MnuToolRepaireItemsCost_Click()

    'Dim Msg As String
    'If SystemOptions.SysMainStockCostMethod <> ModernWeightAverage Then
    '    Msg = "«·‰”Œ… «·„Œ’’… ·ﬂ...·« ” Œœ„ Â–Â «·√„ﬂ«‰Ì…"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    Exit Sub
    'End If
    If checkApility("FrmToolsRepireItemsCost") = False Then
        Exit Sub
    End If

'   FrmToolsRepireItemsCost.show
End Sub

Private Sub MnuToolsDataBase_Click(Index As Integer)
    Dim Msg As String

    Select Case Index

        Case 0

            If checkApility("open_my_connection") = False Then
                Exit Sub
            End If

         '   open_my_connection
FrmSQLConData.show
        Case 1
    If user_id <> 1 Then
        '   MsgBox ""
        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
        '    Msg = Msg & Chr(13) & "Õ ‰Â“— Ê·««ÌÂ "
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    
            If checkApility("AdminLogin") = False Then
                Exit Sub
            End If

            AdminLogin.show

        Case 2
            Unload WebForm

            If Me.ActiveForm Is Nothing Then

     '           FrmNEWlOGIN.show
            Else
     '           Msg = "ÌÃ» €·ﬁ «Ï ‘«‘… „‰ ‘«‘«  «·»—‰«„Ã ﬁ»·"
     '           Msg = Msg & Chr(13) & "«‰  ” Œœ„ Â–« «·‘«‘…....!!!!"
     '           MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
           End If

    End Select

End Sub

Private Sub MnuToolsSetPrinters_Click(Index As Integer)

    Select Case Index

        Case 0
            Dim Msg As String

            On Error GoTo hErr
            Me.Cmdlg.CancelError = False
            Me.Cmdlg.ShowPrinter
            Exit Sub
hErr:
            Msg = "ÕœÀ Œÿ« √À‰«¡ ≈⁄œ«œ «·ÿ«»⁄… ..."
            Msg = Msg & Chr(13) & Err.description
            Msg = Msg & Chr(13) & Err.Number
            Msg = Msg & Chr(13) & Err.Source
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

        Case 1

            If checkApility("FrmAccountsSeetting") = False Then
                Exit Sub
            End If

            FrmAccountsSeetting.show

        Case 2
            'If checkApility("FrmDocType") = False Then
            '    Exit Sub
            'End If

            FrmDocType.show

        Case 3

            
            
        '    If checkApility("System_alarms") = False Then
        '        Exit Sub
        '    End If

            System_alarms.show

        Case 4

            If checkApility("System_manger2") = False Then
                Exit Sub
            End If

        '    System_manger2.show

        Case 5

            If checkApility("coding") = False Then
                Exit Sub
            
            End If

            Coding.show

        Case 6

            If checkApility("FrmMessnger") = False Then
                Exit Sub
            End If

            FrmMessnger.show

        Case 7
        FrmDictionary.show
  '      FrmOlapShow.show
'FrmADDToDictionary.show
    '        If checkApility("SMSSeTTings") = False Then
    '            Exit Sub
    '        End If
'
'            SMSSeTTings.show
            'WebForm.Show
    End Select

End Sub

Private Sub MnuToolsSetPrinters0_Click(Index As Integer)
  On Error GoTo hErr
 Dim Msg As String
Select Case Index
Case 0
   '    FrmItServiceTicket.show
  
  
    Case 1
   Me.Cmdlg.CancelError = False
    Me.Cmdlg.ShowPrinter
    Exit Sub
    
hErr:
    Msg = "ÕœÀ Œÿ« √À‰«¡ ≈⁄œ«œ «·ÿ«»⁄… ..."
    Msg = Msg & Chr(13) & Err.description
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

    End Select
    

End Sub

Private Sub MnuToolsSetPrinters0sub_Click(Index As Integer)
'frmsalebill5.show

             Dim Msg As String
        Select Case Index
        Case 0
            If checkApility("FrmItServiceTicket") = False Then
                Exit Sub
            End If
             FrmItServiceTicket.show
           Case 1
                If checkApility("Frmcameralocation") = False Then
                Exit Sub
            End If
             Frmcameralocation.show
             
             Case 2

                 If SystemOptions.usertype = UserNormal Then
 
        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
         
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    Dim x As Integer
 
             FrmAccountRecreation.show
             
             
        Case 3
        
  If SystemOptions.usertype = UserNormal Then
        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
        FrmOpenClosPeriod.show
        
  Case 4
    If SystemOptions.usertype = UserNormal Then
        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
'  Frmvending.show
    'XML  XML
    '   FrmXmlRet.show
  Case 5
  '  If SystemOptions.usertype = UserNormal Then
  '      Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
  '      MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
  '      Exit Sub
  '  End If
  'FRMSolver.show 'Õ· „‘«ﬂ· «·«”‰«œ
    Case 6
  '     If SystemOptions.usertype = UserNormal Then
  '      Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
  '      MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
  '      Exit Sub
  '  End If
    FrmVizits.show
      Case 7
  '          If SystemOptions.usertype = UserNormal Then
  '      Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
  '      MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
  '      Exit Sub
   ' End If
   FrmVizitScreen.mIndex = 0
      FrmVizitScreen.show
      
      
            Case 8
            'If SystemOptions.usertype = UserNormal Then
       ' Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
       ' MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       ' Exit Sub
    'End If
         If checkApility("frmEditCost") = False Then
                Exit Sub
            End If
            
      frmEditCost.show
      
      Case 9
             
'            If SystemOptions.usertype = UserNormal Then
'        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
'        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        Exit Sub
'    End If
         If checkApility("FrmReCost") = False Then
                Exit Sub
            End If
                   
'      FrmReCost.show
FrmReCost.mIndex = 1
      FrmReCost.show
      
      Case 10
     If Dir(App.path & "\team.exe") <> "" Then
         Shell App.path & "\team.exe", vbNormalFocus
     End If
          
          
        End Select
End Sub

Private Sub MnuUsersScreensPremission_Click()
    Dim Msg As String
    
    If SystemOptions.usertype = UserNormal Then
    
        '   MsgBox ""
        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
        '    Msg = Msg & Chr(13) & "Õ ‰Â“— Ê·««ÌÂ "
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If Not mdifrmmain.ActiveForm Is Nothing Then
        ModPremis.ShowScreenPermission Me.ActiveForm.Name
    Else
        ModPremis.ShowScreenPermission ""
    End If

End Sub

Private Sub MnuView_Click()
    Exit Sub
    Dim Msg As String

    On Error Resume Next

    If Me.DockingPane1.PanesCount <= 0 Then
        Me.PopMenu1.Checked("MnuView_Item(0)") = False
        Me.PopMenu1.Checked("MnuView_Item(1)") = False
        Me.PopMenu1.Checked("MnuView_Item(2)") = False
        Me.PopMenu1.Checked("MnuView_Item(3)") = False
        Me.PopMenu1.Checked("MnuView_Item(4)") = False
        Me.PopMenu1.Checked("MnuView_Item(5)") = False
        Me.PopMenu1.Checked("MnuView_Item(6)") = False
        Exit Sub
    End If

    If Not Me.DockingPane1.FindPane(DockingPanesIDs.OutBarPaneID) Is Nothing Then
        'Me.MnuView_Item(0).Checked = Not Me.DockingPane1.FindPane(DockingPanesIDs.OutBarPaneID).Closed
        Me.PopMenu1.Checked("MnuView_Item(0)") = Not Me.DockingPane1.FindPane(DockingPanesIDs.OutBarPaneID).Closed
    Else
        'Me.MnuView_Item(0).Checked = False
        Me.PopMenu1.Checked("MnuView_Item(0)") = False
    End If

    If Not Me.DockingPane1.FindPane(DockingPanesIDs.NewsBarPaneID) Is Nothing Then
        Me.PopMenu1.Checked("MnuView_Item(1)") = Not Me.DockingPane1.FindPane(DockingPanesIDs.NewsBarPaneID).Closed
        'Me.MnuView_Item(1).Checked = Not Me.DockingPane1.FindPane(DockingPanesIDs.NewsBarPaneID).Closed
    Else
        Me.PopMenu1.Checked("MnuView_Item(1)") = False
        ' Me.MnuView_Item(1).Checked = False
    End If

    If Not Me.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID) Is Nothing Then
        '"‘—Ìÿ ‘Ã—… «·√’‰«›"
        Me.PopMenu1.Checked("MnuView_Item(2)") = Not Me.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID).Closed
        '  Me.MnuView_Item(2).Checked = Not Me.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID).Closed
    Else
        Me.PopMenu1.Checked("MnuView_Item(2)") = False
        '  Me.MnuView_Item(2).Checked = False
    End If

    If Not Me.DockingPane1.FindPane(DockingPanesIDs.MantainceID) Is Nothing Then
        '"‘—Ìÿ „⁄·Ê„«  «·’Ì«‰…"
        Me.PopMenu1.Checked("MnuView_Item(3)") = Not Me.DockingPane1.FindPane(DockingPanesIDs.MantainceID).Closed
        '  Me.MnuView_Item(3).Checked = Not Me.DockingPane1.FindPane(DockingPanesIDs.MantainceID).Closed
    Else
        Me.PopMenu1.Checked("MnuView_Item(3)") = False
        '  Me.MnuView_Item(3).Checked = False
    End If

    If Not Me.DockingPane1.FindPane(DockingPanesIDs.InternetNews) Is Nothing Then
        '"‘—Ìÿ √Œ»«— «·√‰ —‰ "
        Me.PopMenu1.Checked("MnuView_Item(4)") = Not Me.DockingPane1.FindPane(DockingPanesIDs.InternetNews).Closed
        '  Me.MnuView_Item(4).Checked = Not Me.DockingPane1.FindPane(DockingPanesIDs.InternetNews).Closed
    Else
        Me.PopMenu1.Checked("MnuView_Item(4)") = False
        '  Me.MnuView_Item(4).Checked = False
    End If

    If Not Me.DockingPane1.FindPane(DockingPanesIDs.DynamicHelp) Is Nothing Then
        Me.PopMenu1.Checked("MnuView_Item(5)") = Not Me.DockingPane1.FindPane(DockingPanesIDs.DynamicHelp).Closed
        '  Me.MnuView_Item(5).Checked = Not Me.DockingPane1.FindPane(DockingPanesIDs.DynamicHelp).Closed
    Else
        Me.PopMenu1.Checked("MnuView_Item(5)") = False
        '    Me.MnuView_Item(5).Checked = False
    End If

    If Not Me.DockingPane1.FindPane(DockingPanesIDs.CalendarPaneID) Is Nothing Then
        Me.PopMenu1.Checked("MnuView_Item(6)") = Not Me.DockingPane1.FindPane(DockingPanesIDs.CalendarPaneID).Closed
        '  Me.MnuView_Item(6).Checked = Not Me.DockingPane1.FindPane(DockingPanesIDs.CalendarPaneID).Closed
    Else
        Me.PopMenu1.Checked("MnuView_Item(6)") = False
        '    Me.MnuView_Item(6).Checked = False
    End If

    Exit Sub
    '-------
hErr:

    'Dim xPane As XtremeDockingPane.Pane
    'Select Case Index
    '    Case 0
    '        Me.MnuView_Item(Index).Checked = Not MnuView_Item(Index).Checked
    '        Me.DockingPane1.FindPane(DockingPanesIDs.OutBarPaneID).Closed = Not _
    '            Me.DockingPane1.FindPane(DockingPanesIDs.OutBarPaneID).Closed
    '    Case 1
    '        Me.MnuView_Item(Index).Checked = Not MnuView_Item(Index).Checked
    '        Me.DockingPane1.FindPane(DockingPanesIDs.NewsBarPaneID).Closed = Not _
    '            Me.DockingPane1.FindPane(DockingPanesIDs.NewsBarPaneID).Closed
    '    Case 2
    '        Me.MnuView_Item(Index).Checked = Not MnuView_Item(Index).Checked
    '        Me.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID).Closed = Not _
    '            Me.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID).Closed
    '    Case 3
    '        Me.MnuView_Item(Index).Checked = Not MnuView_Item(Index).Checked
    '        Me.DockingPane1.FindPane(DockingPanesIDs.MantainceID).Closed = Not _
    '            Me.DockingPane1.FindPane(DockingPanesIDs.MantainceID).Closed
    '    Case 4
    '        Me.MnuView_Item(Index).Checked = Not MnuView_Item(Index).Checked
    '
    '        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.InternetNews)
    '        If Not xPane Is Nothing Then
    '            Me.DockingPane1.FindPane(DockingPanesIDs.InternetNews).Closed = Not _
    '                Me.DockingPane1.FindPane(DockingPanesIDs.InternetNews).Closed
    '        Else
    '
    '        End If
    '    Case 5
    '        Me.MnuView_Item(Index).Checked = Not MnuView_Item(Index).Checked
    '        Me.DockingPane1.FindPane(DockingPanesIDs.DynamicHelp).Closed = Not _
    '            Me.DockingPane1.FindPane(DockingPanesIDs.DynamicHelp).Closed
    '    Case 6
    '        Me.MnuView_Item(Index).Checked = Not MnuView_Item(Index).Checked
    '        Me.DockingPane1.FindPane(DockingPanesIDs.CalendarPaneID).Closed = Not _
    '            Me.DockingPane1.FindPane(DockingPanesIDs.CalendarPaneID).Closed
    'End Select
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            Y As Single)
    On Error GoTo ErrTrap

    If Button = vbRightButton Then
'        PopupMenu mdifrmmain.MdiContextMenu  ', vbPopupMenuRightAlign, X, Y + 200
      If DoPremis(Do_Search, "FrmAccEditJournal", True) = False Then
                Exit Sub
            End If
            
     Unload Voucher_search
            Voucher_search.show
            
    End If

ErrTrap:
End Sub

Private Sub MDIForm_Resize()

    Dim i As Integer
    On Error Resume Next
If Me.WindowState <> 1 Then
'Me.WindowState = vbMaximized
Exit Sub
End If

    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then

        For i = 0 To Forms.count - 1

            If Forms(i).Name <> "MDIFrmMain" Then
                If Forms(i).MDIChild = True Then
                    Resize_Form Forms(i)
                End If
            End If

        Next i

    End If

End Sub

Private Sub MnuBoxAccouns_Click()

    If checkApility("FrmBoxesAccounts") = False Then
        Exit Sub
    End If

    OpenScreen PopUpShowBoxesAccounts
End Sub

Private Sub MnuBoxDrawing_Click()



    If checkApility("FrmBoxDrawing") = False Then
        Exit Sub
    End If

    FrmBoxDrawing.show
    FrmBoxDrawing.ZOrder 0
End Sub

Private Sub MnuEmpsAdvance_Click()
    FrmEmpsAdvance.show
End Sub



Private Sub MnuFinDiscounts_Click()

    'FrmDiscounts
    If checkApility("FrmDiscounts") = False Then
        Exit Sub
    End If

    OpenScreen AllowsDiscountsScreen
End Sub

Private Sub MnuInterface_Click()

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.MnuInterfaceSub(0).Enabled = False
        Me.MnuInterfaceSub(1).Enabled = True
    Else
        Me.MnuInterfaceSub(0).Enabled = True
        Me.MnuInterfaceSub(1).Enabled = False
    End If

End Sub

Private Sub MnuInterfaceSub_Click(Index As Integer)

Dim Msg As String
          If Me.ActiveForm Is Nothing Then

              
            Else
                Msg = "ÌÃ» €·ﬁ «Ï ‘«‘… „‰ ‘«‘«  «·»—‰«„Ã ﬁ»·"
                Msg = Msg & Chr(13) & "«‰  ” Œœ„ Â–« «·‘«‘…....!!!!"
                
                 Msg = Msg & Chr(13) & "Close All Screen Firstly"
                
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If
            
            
            
    If Not mdifrmmain.ActiveForm Is Nothing Then
        'GetMsgs 156, vbExclamation
        ' Exit Sub
    End If

    Unload System_alarms
' Unload WebForm
 
 
    Select Case Index

        Case 0 'Load Arabic Interface
'        Reload Me
     '  Unload Me
     '   Load Me
        'reload
            LoadInterface ArabicInterface

        Case 1 'Load English Interface
            LoadInterface EnglishInterface
    End Select
  
' Load WebForm
'   Load System_alarms
'   System_alarms.SetFocus
'    System_alarms.show

    ClosePanes
    CreateDocks True
End Sub

Private Function ImgInImgList(sKey As String) As Integer
    On Error GoTo ErrTrap
    ImgInImgList = Me.ImgLstMenuIcons.ListImages(sKey).Index
    Exit Function
ErrTrap:

    If Err.Number = 35601 Then
        ImgInImgList = -1
    End If

End Function

Private Sub MPITP_GSort_Option_Click(Index As Integer)
    Dim StrTemp As String
    Dim i As Integer

    Select Case Index

        Case 0
            StrTemp = " GroupID ASC"

        Case 1
            StrTemp = " GroupID DESC"

        Case 2

        Case 3
            StrTemp = " GroupCode ASC"

        Case 4
            StrTemp = " GroupCode DESC"

        Case 5

        Case 6
            StrTemp = " GroupName ASC"

        Case 7
            StrTemp = " GroupName DESC"
    End Select

'    For i = MPITP_GSort_Option.LBound To MPITP_GSort_Option.UBound
'        MPITP_GSort_Option(i).Checked = False
'    Next i

'    MPITP_GSort_Option(Index).Checked = True

    'If Not ItemsTreePane Is Nothing Then
    '    ItemsTreePane.GroupsSort = StrTemp
    '    ItemsTreePane.LoadData StrTemp, ItemsTreePane.ItemsSort
    'End If

End Sub

Private Sub MPITP_ISort_Option_Click(Index As Integer)
    Dim i As Integer

    Dim StrTemp As String

    Select Case Index

        Case 0
            StrTemp = " ItemID ASC"

        Case 1
            StrTemp = " ItemID DESC"

        Case 2

        Case 3
            StrTemp = " ItemCode ASC"

        Case 4
            StrTemp = " ItemCode DESC"

        Case 5

        Case 6
            StrTemp = " ItemName ASC"

        Case 7
            StrTemp = " ItemName DESC"
    End Select

    'For i = MPITP_ISort_Option.LBound To MPITP_ISort_Option.UBound
    '    MPITP_ISort_Option(i).Checked = False
    'Next i

    'MPITP_ISort_Option(Index).Checked = True

    'If Not ItemsTreePane Is Nothing Then
    '    ItemsTreePane.ItemsSort = StrTemp
    '    ItemsTreePane.LoadData ItemsTreePane.GroupsSort, StrTemp
    'End If

End Sub

Private Sub Options_Click()

    If checkApility("FrmOptions") = False Then
        Exit Sub
    End If

    OpenScreen OptionsScreen
End Sub
 
Private Sub planningMnuSub_Click(Index As Integer)
Select Case Index
Case 0
           If checkApility("FrmProductionPlan") = False Then
                Exit Sub
            End If
            
            FrmProductionPlan.show
'FrmProductionPlan.Caption = "Œÿ… «·«‰ «Ã"
'FrmProductionPlan.Ele(5).Caption = FrmProductionPlan.Caption
'        FrmProductionPlan.lblPlantype.Caption = 0
        

End Select
End Sub

Private Sub PopAvailable_Click()
    'Trading_Click (17)
End Sub

Private Sub PopBalance_Click()
    'Trading_Click (12)
End Sub

Private Sub PopBanks_Click()
    'Stores_Click (1)
End Sub

Private Sub PopClients_Click()
    'Employee_Click (3)
End Sub

Private Sub PopEmployee_Click()
    'Employee_Click (0)
End Sub

Private Sub PopGard_Click()
    'Trading_Click (13)
End Sub

Private Sub PopGroups_Click()
    'Groups_Click
End Sub

Private Sub PopItems_Click()
    'Items_Click (0)
End Sub

Private Sub PopMaintanence_Click()
    'Trading_Click (9)
End Sub

Private Sub PopMenu1_Click(ItemNumber As Long)
    On Error Resume Next

    If ItemNumber = 108 Then Exit Sub
    Dim Lparent As Long
    Dim temp As String
    Dim TempArry As Variant
    Dim i As Integer

    With Me.PopMenu1
        Lparent = .MenuIndex("MnuWindowsList")
        temp = .HierarchyPath(.MenuKey(ItemNumber), 1, "-")

        If temp <> "" Then
            TempArry = Split(temp, "-", , vbTextCompare)

            If CStr(TempArry(1)) Like .Caption("MnuWindowsList") Then

                For i = 0 To Forms.count - 1

                    If Forms(i).Name = .MenuKey(ItemNumber) Then

                        Forms(i).ZOrder 0
                        Exit For
                    End If

                Next i

            End If
        End If

    End With

End Sub

Private Sub PopMenu1_InitPopupMenu(ParentItemNumber As Long)
    Debug.Print Me.PopMenu1.MenuKey(ParentItemNumber)

    If Me.PopMenu1.MenuKey(ParentItemNumber) = "MnuWindowsList" Then
        'CreateWindowList
    End If

    CreateWindowList
End Sub

Private Sub PopMenu1_ItemHighlight(ItemNumber As Long, _
                                   bEnabled As Boolean, _
                                   bSeparator As Boolean)
    'Me.PopMenu1.Checked("MnuView_Item(0)") = Not Me.DockingPane1.Panes(DockingPanesIDs.OutBarPaneID).Closed
End Sub

Private Sub PopPriceList_Click()
    'PriceList_Click
End Sub

Private Sub PopPurchaseBill_Click()
    'Trading_Click (6)
End Sub

Private Sub PopReturn_Click()
    'Trading_Click (8)
End Sub

Private Sub PopSallBill_Click()
    'Trading_Click (5)
End Sub

Private Sub PopSerialData_Click()
    'Trading_Click (17)
End Sub

Private Sub PopStore_Click()
    'Stores_Click (0)
End Sub

Private Sub PopSupliers_Click()
    'Employee_Click (4)
End Sub

Private Sub POSTRansactios_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("FrmPOSDATA") = False Then
                Exit Sub
            End If

            FrmPOSDATA.show


       Case 1

            If checkApility("cachierData") = False Then
                Exit Sub
            End If

            cachierData.show



        Case 2
 
            If checkApility("frm_sheft") = False Then
                Exit Sub
            End If

            frm_sheft.show
 
        Case 3
 
            If checkApility("FrmTables") = False Then
                Exit Sub
            End If

            FrmTables.show


Case 4
           If checkApility("FrmPoints") = False Then
                Exit Sub
            End If
FrmPoints.show

 
        Case 5

            If checkApility("CashierLogin") = False Then
                Exit Sub
            End If
 
            CashierLogin.show
            'frmsalebill1.Show
 
 
  Case 6
 
              If checkApility("FrmProductionOrder4") = False Then
                Exit Sub
            End If

            FrmProductionOrder4.show
            
 
 Case 7
 
              If checkApility("FrmBankDeposite3") = False Then
                Exit Sub
            End If

            FrmBankDeposite3.show

 
 
        Case 8

            If checkApility("ReportSales") = False Then
                Exit Sub
            End If

            FrmReports.show
            FrmReports.C1TabMain.CurrTab = 0
 'Case 9
 
 '           If checkApility("FrmAnalysItems") = False Then
 ''               Exit Sub
  '          End If
'
          '  FrmReports.show
          '  FrmReports.C1TabMain.CurrTab = 0
' FrmAnalysItems.show
 
 Case 9
          If checkApility("FrmCustCash") = False Then
                Exit Sub
            End If
      FrmCustCash.show
      
       Case 10
          If checkApility("FrmCoupons") = False Then
                Exit Sub
            End If
      FrmCoupons.show
      
        Case 11 'ÿ—„»« 
        FrmItemsClass.mIndex = 11
FrmItemsClass.show

        '  If checkApility("FrmCoupons") = False Then
        '        Exit Sub
        '    End If
     ' FrmCoupons.show
      
    End Select

End Sub

Private Sub PpBarcode_Click()
    'Barcode_Click
End Sub

Private Sub PrbH_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("FrmProductionOrder1") = False Then
                Exit Sub
            End If

            FrmProductionOrder1.show


        Case 1

        
    If checkApility("FrmOutProductionOrder1") = False Then
                Exit Sub
            End If

            FrmOutProductionOrder1.show


        Case 2

            If checkApility("FrmInpoutWorkOrder1") = False Then
                Exit Sub
            End If

            FrmInpoutWorkOrder1.show
    End Select

End Sub

Private Sub prdo1_Click(Index As Integer)

    Select Case Index

        Case 0

  '          If checkApility("frmequipment1") = False Then
  '              Exit Sub
  '          End If
'
'            frm_sheft.show
'frmequipment1.show

        Case 1

            If checkApility("frmequipment") = False Then
                Exit Sub
            End If

            frmequipment.show
            'Case 2
            'If checkApility("frmProductLine") = False Then
            '    Exit Sub
            'End If

            'frmProductLine.Show


Case 2
       '     If checkApility("FrmProductionElements") = False Then
       '         Exit Sub
       '     End If
'
'            FrmProductionElements.show
'
Unload FrmPay_Garanty_Shipment3M
             If checkApility("FrmPay_Garanty_Shipment3M") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment3M.SendForm = 5
FrmPay_Garanty_Shipment3M.show
Case 3

  If checkApility("UnitsIndustrialCost") = False Then
                Exit Sub
            End If

            UnitsIndustrialCost.show






        Case 6

      '      If checkApility("FrmShowPrice1") = False Then
      '          Exit Sub
      '      End If

            'FrmCustomerOrder.Show
      '      GeneralPriceType = 1
      '      FrmShowPrice.show

   If checkApility("FrmPO9") = False Then
                Exit Sub
            End If
FrmPO9.show
        Case 7

            If checkApility("FrmProductionOrder") = False Then
                Exit Sub
            End If

            FrmProductionOrder.show
 
        Case 8

            If checkApility("FrmOutProductionOrder") = False Then
                Exit Sub
            End If

            FrmOutProductionOrder.show

            'FrmOut.Show
            'FrmOutForOrder.Show
        Case 9

            If checkApility("FrmInpoutWorkOrder") = False Then
                Exit Sub
            End If
 
            FrmInpoutWorkOrder.show

        Case 10

            If checkApility("FrmCalcCostPrice") = False Then
                Exit Sub
            End If

            FrmCalcCostPrice.show

        Case 11

            If checkApility("FrmCalcCostPrice1") = False Then
                Exit Sub
            End If

            FrmCalcCostPrice2.show

        Case 12

            If checkApility("FrmProductionAllocation") = False Then
                Exit Sub
            End If

            FrmProductionAllocation.show

Case 13
            If checkApility("FrmDriverTrip") = False Then
                Exit Sub
            End If
FrmDriverTrip.show

Case 14
       If checkApility("FrmDefinCompItem") = False Then
                Exit Sub
            End If

            FrmDefinCompItem.show
            
        Case 15

            If checkApility("FrmProductionReport") = False Then
                '    Exit Sub
            End If

            frmProductionreport.show

    End Select

End Sub

Private Sub prdo1sub_Click(Index As Integer)
Select Case Index
Case 0
    If checkApility("frmequipment") = False Then
                Exit Sub
            End If
frmequipment.show
Case 1
 
  ' If checkApility("FrmProductionElements") = False Then
  '              Exit Sub
  '          End If
'
'            FrmProductionElements.show
'
            Unload FrmPay_Garanty_Shipment3M
             If checkApility("FrmPay_Garanty_Shipment3M") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment3M.SendForm = 5
FrmPay_Garanty_Shipment3M.show

Case 2
  If checkApility("UnitsIndustrialCost") = False Then
                Exit Sub
            End If

            UnitsIndustrialCost.show
Case 3
           If checkApility("frmequipment1") = False Then
                Exit Sub
            End If
'
 
frmequipment1.show


Case 4
           If checkApility("FrmProductionType") = False Then
                Exit Sub
            End If
'
 
FrmProductionType.show


Case 5

 If checkApility("FrmDistriExpensItems") = False Then
                Exit Sub
            End If
            
FrmDistriExpensItems.show
End Select

End Sub

Private Sub PriceChips_Click()
    'FrmMainPriceList.FgMain_DblClick
End Sub

Private Sub PriceOffer_Click()
    
End Sub

Private Sub ProductionPlansub_Click(Index As Integer)

    Select Case Index

        Case 0
         If checkApility("FrmProductionPlan") = False Then
                Exit Sub
            End If
            FrmProductionPlan.show
If SystemOptions.UserInterface = ArabicInterface Then
 FrmProductionPlan.Caption = "Œÿ… «·«‰ «Ã"
Else
FrmProductionPlan.Caption = "Production Plan"
End If

FrmProductionPlan.Ele(5).Caption = FrmProductionPlan.Caption
        FrmProductionPlan.lblPlantype.Caption = 0
        Case 1
                 If checkApility("FrmQCitems") = False Then
                Exit Sub
            End If
            FrmQCitems.show

        Case 2
                    If checkApility("FrmItemsClass") = False Then
                Exit Sub
            End If
            Unload FrmItemsClass
            FrmItemsClass.show
         '   FrmItemsClass.Caption = " ’‰Ì› «·„‰ Ã« "
    '        FrmItemsClass.EleHeader.Caption = FrmItemsClass.Caption

        Case 3
                            If checkApility("frmcorrectaction") = False Then
                Exit Sub
            End If
            frmcorrectaction.show

        Case 4
                                    If checkApility("FrmBatchSheet") = False Then
                Exit Sub
            End If
            FrmBatchSheet.show

     '       FrmInpoutWorkOrder.show
     '       If SystemOptions.UserInterface = ArabicInterface Then
     '       FrmInpoutWorkOrder.Caption = "›Õ’  ÃÊœ… «·„‰ Ã «· «„"
     '       Else
     '       FrmInpoutWorkOrder.Caption = "Items Quality Test"
     '       End If
     '       FrmInpoutWorkOrder.Ele(6).Caption = FrmInpoutWorkOrder.Caption

        Case 5
                                            If checkApility("FrmTestCertificate") = False Then
                Exit Sub
            End If
            FrmTestCertificate.show
     '       FrmProductionOrder.show
     '       If SystemOptions.UserInterface = ArabicInterface Then
     '       FrmProductionOrder.Caption = "«„— ‘€· «’·«Õ «·„‰ Ã«  «·„⁄Ì»…"
     '       Else
     '       FrmProductionOrder.Caption = "Repair Failled Items"
     '       End If
     '       FrmProductionOrder.Ele(6).Caption = FrmProductionOrder.Caption
     
     
     Case 6
                                            If checkApility("FrmQuality") = False Then
                Exit Sub
            End If
     FrmQuality.show
Case 7
                                          If checkApility("FrmProcessRep") = False Then
                Exit Sub
            End If
FrmProcessRep.show
    End Select

End Sub

Private Sub prosub1_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("frmProductLine") = False Then
                Exit Sub
            End If

            frmProductLine.show

        Case 1

            If checkApility("FrmTransferEmployee") = False Then
                Exit Sub
            End If

            FrmTransferEmployee.show

    End Select

End Sub

Private Sub PurchaseBasic_Click(Index As Integer)

    Select Case Index

        Case 0

            'FrmCompany
            If checkApility("FrmCompany") = False Then
                Exit Sub
            End If

            OpenScreen SuppliersScreen

        Case 1

            If checkApility("FrmVendorContract") = False Then
                Exit Sub
            End If

            FrmVendorContract.show

        Case 2

            If checkApility("Ageng") = False Then
                Exit Sub
            End If

            Ageng.show

        Case 3

'            If checkApility("FrmShipment_mode") = False Then
'                Exit Sub
'            End If
'
'            FrmShipment_mode.show
'
Unload FrmPay_Garanty_Shipment
             If checkApility("FrmPay_Garanty_Shipment") = False Then
                 Exit Sub
             End If

FrmPay_Garanty_Shipment.SendForm = 2
FrmPay_Garanty_Shipment.show

        Case 4
Unload FrmPay_Garanty_Shipment
             If checkApility("FrmPay_Garanty_Shipment") = False Then
                 Exit Sub
             End If

FrmPay_Garanty_Shipment.SendForm = 1
FrmPay_Garanty_Shipment.show


'            If checkApility("FrmGaranty_type") = False Then
'                Exit Sub
'            End If
'
'            FrmGaranty_type.show
'
        Case 5
'         If checkApility("FrmPaymentData") = False Then
'                Exit Sub
'            End If

'            FrmPaymentData.show
            
 Unload FrmPay_Garanty_Shipment
             If checkApility("FrmPay_Garanty_Shipment") = False Then
                 Exit Sub
             End If

FrmPay_Garanty_Shipment.SendForm = 0
FrmPay_Garanty_Shipment.show



    Case 6

       '      If checkApility("FrmSalesRePGroups1") = False Then
       '          Exit Sub
       '     End If
'
'            FrmSalesRePGroups1.show
'

 Unload FrmPay_Garanty_Shipment
             If checkApility("FrmPay_Garanty_Shipment") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment.SendForm = 3
FrmPay_Garanty_Shipment.show


    Case 7

     '       If checkApility("FrmSalesRepData1") = False Then
     '           Exit Sub
     '       End If
'
'            FrmSalesRepData1.show

 Unload FrmPay_Garanty_Shipment
             If checkApility("FrmPay_Garanty_Shipment") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment.SendForm = 6
FrmPay_Garanty_Shipment.show

    Case 8

      '      If checkApility("FrmShipingData") = False Then
      '          Exit Sub
      '      End If
'
'            FrmShipingData.show

 Unload FrmPay_Garanty_Shipment
             If checkApility("FrmPay_Garanty_Shipment") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment.SendForm = 4
FrmPay_Garanty_Shipment.show




    End Select

End Sub

Private Sub PurchaseTransactions_Click(Index As Integer)
    Dim RsOptions As New ADODB.Recordset

    Select Case Index

        Case 0
            'FrmShowPrice
            'GeneralPriceType = 1
            'If checkApility("FrmShowPrice1") = False Then
            '    Exit Sub
            'End If

            'OpenScreen ScreensName.ShowPriceScreen

        Case 1

            If checkApility("shipment") = False Then
                Exit Sub
            End If

            If SystemOptions.UserInterface = ArabicInterface Then
                shipmentA.show
            Else
                'shipment.show
            End If

        Case 3
            Set RsOptions = New ADODB.Recordset
            RsOptions.Open "tbloptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

            If checkApility("FrmBillBuy") = False Then
                Exit Sub
            End If

            OpenScreen PurchaseScreen

'        Case 4
'
'            If checkApility("FrmBillBuyComposite") = False Then
'                Exit Sub
'            End If
'
'FrmBillBuyComposite.show
            'FrmBillBuy
        Case 5

            If checkApility("FrmReturnpurchases") = False Then
                Exit Sub
            End If

            OpenScreen RetrunPurchse

            'FrmReturnpurchases
        Case 6

            If checkApility("Ageng_all") = False Then
                Exit Sub
            End If

            Ageng_all.show

        Case 7

            If checkApility("ReportPurchase") = False Then
                Exit Sub
            End If

            FrmReports.show
            FrmReports.C1TabMain.CurrTab = 1

    End Select

End Sub

Private Sub PurchaseTransactionssubd_Click(Index As Integer)

    Select Case Index

        Case 0
            'FrmShowPrice
            'GeneralPriceType = 1
            'If checkApility("FrmShowPrice1") = False Then
            '    Exit Sub
            'End If
            '
            'OpenScreen ScreensName.ShowPriceScreen

        Case 1
            'GeneralPriceType = 6
            'If checkApility("FrmShowPrice1") = False Then
            '    Exit Sub
            'End If
            '
            'OpenScreen ScreensName.ShowPriceScreen

        Case 2

    End Select

End Sub

Private Sub PurchaseTransactionssubs_Click(Index As Integer)
Select Case Index
Case 0
       If checkApility("FrmPO4") = False Then
                Exit Sub
            End If
GeneralPriceType = 0
FrmPO4.show

Case 1
       If checkApility("FrmPO5") = False Then
                Exit Sub
            End If
GeneralPriceType = 0
FrmPO5.show
 
Case 2
      If checkApility("FrmComparePrices") = False Then
                Exit Sub
            End If
 
 FrmComparePrices.show
 
End Select
End Sub

Private Sub PurchaseTransactionssubs1_Click(Index As Integer)

    Select Case Index

        Case 0
                       If checkApility("FrmPO8") = False Then
                Exit Sub
            End If
FrmPO8.show

      
        Case 1

        Case 2
       '     GeneralPriceType = 1

            If checkApility("FrmPO10") = False Then
                Exit Sub
            End If
FrmPO10.show
            

    End Select

End Sub

Private Sub RealEstateMarketingSub_Click(Index As Integer)
Select Case Index

Case 0
            If checkApility("FrmCountriesData") = False Then
                Exit Sub
            End If

            FrmCountriesData.show

        Case 1

            If checkApility("FrmGovernmentData") = False Then
                Exit Sub
            End If

            FrmGovernmentData.show

        Case 2

            If checkApility("FrmGovernCitiesData") = False Then
                Exit Sub
            End If

            FrmGovernCitiesData.show

        Case 3

            If checkApility("streets") = False Then
                Exit Sub
            End If

            streets.show
Case 4
      If checkApility("FrmAkarStatus") = False Then
                Exit Sub
            End If
FrmAkarStatus.show
      Case 5
      
            If checkApility("FrmAkarType") = False Then
                Exit Sub
            End If

            FrmAkarType.show
      Case 6
            If checkApility("FrmAkarUnit") = False Then
                Exit Sub
            End If
FrmAkarUnit.mIndex = 0
            FrmAkarUnit.show

Case 7

     If checkApility("FrmSalesRePGroups") = False Then
                Exit Sub
            End If

            FrmSalesRePGroups.show

        Case 8

 
             If checkApility("FrmPay_Garanty_Shipment") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment.SendForm = 7
FrmPay_Garanty_Shipment.show

Case 9
     If checkApility("streets") = False Then
                Exit Sub
            End If
streets.mIndex = 1
            streets.show

Case 10
     If checkApility("streets") = False Then
                Exit Sub
            End If
streets.mIndex = 2
            streets.show
 
Case 11
          If checkApility("FrmCustomerType") = False Then
                Exit Sub
            End If
FrmCustomerType.Indx = 0
            FrmCustomerType.show


Case 12
          If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If
            FrmCustemers.show
Case 13
          If checkApility("FrmCustomerType") = False Then
                Exit Sub
            End If
FrmCustomerType.Indx = 2
FrmCustomerType.show


Case 14
          If checkApility("FrmAkarUnit") = False Then
                Exit Sub
            End If
FrmAkarUnit.mIndex = 1

FrmAkarUnit.show
Case 15

          If checkApility("streets") = False Then
                Exit Sub
            End If
streets.mIndex = 3
streets.show

 Case 16
   If checkApility("Frmblacklist") = False Then
             Exit Sub
        End If
'
'
frmblacklist.show
Case 17
            If checkApility("RSPhoneBook") = False Then
                Exit Sub
            End If

            RSPhoneBook.show
Case 18
             If checkApility("FrmStudentCalling") = False Then
                Exit Sub
            End If
FrmStudentCalling.show


End Select
End Sub

Private Sub ReceiptPart_Click()

    'FrmReceiptPart
    If checkApility("FrmReceiptPart") = False Then
        Exit Sub
    End If

    OpenScreen ReceiptPartScreen
End Sub

Private Sub rentcarSub_Click(Index As Integer)
Unload dean
Select Case Index

Case 0

       If checkApility("FrmBranchesData") = False Then
            Exit Sub
        End If
             
        If bigUser = False Then
         If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "€Ì— „”„ÊÕ ·ﬂ »«· ⁄«„· „⁄ Â–Â «·‰«›–…", vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "’·«ÕÌ«  «·„” Œœ„Ì‰"
                 Else
                    MsgBox "Not Allowed", vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "  Users Privligies"
                 End If
             
             
             Exit Sub
        End If

         FrmBranchesData.show
         
Case 1
            If checkApility("dean") = False Then
                Exit Sub
            End If

dean.mIndex = 8
dean.show

Case 2

            If checkApility("dean") = False Then
                Exit Sub
            End If
dean.mIndex = 9
dean.show

Case 3
            If checkApility("dean") = False Then
                Exit Sub
            End If
dean.mIndex = 10
dean.show

Case 4
            If checkApility("dean") = False Then
                Exit Sub
            End If
dean.mIndex = 11
dean.show



End Select

End Sub

Private Sub rentcarSubReport_Click(Index As Integer)
Select Case Index
Case 0


           If checkApility("Ageng_all") = False Then
                Exit Sub
            End If
            Unload Ageng_all
            Ageng_all.Indx = 5
Ageng_all.show

Case 1

           If checkApility("FrmItemsClass") = False Then
                Exit Sub
            End If
            Unload FrmItemsClass
FrmItemsClass.mIndex = 6
FrmItemsClass.show



End Select
End Sub

Private Sub Report_Click()
    'If checkApility("FrmReports3") = False Then
    '    Exit Sub
    'End If
    'FrmReportsNew.Show
    FrmReports.show
    FrmReports.ZOrder 0
End Sub

Private Sub ReportDesigner_Click()
    On Error GoTo ErrTrap
    ''If checkApility("FrmReportDesigner") = False Then
    '    Exit Sub
    ''End If
    'If Dir(App.Path & "\ReportDesign.exe") <> "" Then
    '    Shell App.Path & "\ReportDesign.exe"
    'End If
    Exit Sub
ErrTrap:
End Sub

Private Sub ReportDesign_Click()
FrmReportsDesign.show
End Sub

Private Sub RequiredInstallment_Click()

    'FrmInstallmentMustPay
    If checkApility("FrmInstallmentMustPay") = False Then
        Exit Sub
    End If

    OpenScreen PopUpShowInstallmentMustPay
End Sub

Private Sub rsInvestmentsUB_Click(Index As Integer)
Select Case Index
Case 0
            If checkApility("FrmBasicDataINv") = False Then
                Exit Sub
            End If
            FrmBasicDataINv.show
    Case 1
    
    
                If checkApility("FrmShareholders") = False Then
                Exit Sub
            End If
    FrmShareholders.show
    
    
    Case 2
    
    
                If checkApility("Frminvestment") = False Then
                Exit Sub
            End If
    Frminvestment.show
    
    
    Case 3
    
    
                If checkApility("FrmIPO") = False Then
                Exit Sub
            End If
    FrmIPO.show
    
    
    
    Case 4
    
     
                If checkApility("FrmIPOSharer") = False Then
                Exit Sub
            End If
    FrmIPOSharer.show
    
    
  Case 5
   
                If checkApility("FrmBuylandRealEstate") = False Then
                Exit Sub
            End If
   FrmBuylandRealEstate.show
    
    
   Case 6
   
                If checkApility("FrmActiveInvestment") = False Then
                Exit Sub
            End If
   FrmActiveInvestment.show
   
      Case 7
   
                If checkApility("FrmExpensesInvestment") = False Then
                Exit Sub
            End If
   FrmExpensesInvestment.show
   
  Case 8
                 If checkApility("FrmReturnExpensInves") = False Then
                Exit Sub
            End If
  FrmReturnExpensInves.show
   
      Case 9
   
                If checkApility("FrmDiviInvestment") = False Then
                Exit Sub
            End If
   FrmDiviInvestment.show
Case 10
         If checkApility("FrmInvesSales") = False Then
                Exit Sub
            End If
FrmInvesSales.show
         Case 11
   
                If checkApility("FrmSaleBillInvestment") = False Then
                Exit Sub
            End If
   FrmSaleBillInvestment.show
   
Case 12
         If checkApility("FrmInvestliquidation") = False Then
                Exit Sub
            End If
FrmInvestliquidation.show
            Case 13
   
                If checkApility("FrmInvestProfitDistribution") = False Then
                Exit Sub
            End If
   FrmInvestProfitDistribution.show
   
   
         Case 14
   
                If checkApility("FrmBuyBillInvestment") = False Then
                Exit Sub
            End If
   FrmBuyBillInvestment.show
   
            Case 15
   
                If checkApility("FrmOrderedEmptying") = False Then
                Exit Sub
            End If
   FrmOrderedEmptying.show
   
   Case 16
              If checkApility("FrmProjecInvestment") = False Then
                Exit Sub
            End If
   FrmProjecInvestment.show
   
  Case 17
              If checkApility("FrmBookingBondsInvs") = False Then
                Exit Sub
            End If
   FrmBookingBondsInvs.show
             
             
            Case 18
   
                If checkApility("FrmInvestmentsReports") = False Then
                Exit Sub
            End If
   FrmInvestmentsReports.show
   
   
      
   
   
   
   
            
  End Select
End Sub

Private Sub SalesBasicSub_Click(Index As Integer)

    Select Case Index

        Case 0

FrmCustomerType.Indx = 0
            FrmCustomerType.show

        Case 1

      '      If checkApility("FrmCustemers") = False Then
      '          Exit Sub
      '      End If

            'FrmCustemers
      '      OpenScreen CustomersScreen '

        Case 2

            If checkApility("FrmCustomerContract") = False Then
                Exit Sub
            End If

            FrmCustomerContract.show

        Case 3

            If checkApility("Ageng1") = False Then
                Exit Sub
            End If

            Ageng.show

        Case 4
      If checkApility("FrmSalePriceNames") = False Then
                Exit Sub
            End If

            FrmSalePriceNames.show


        Case 5

            If checkApility("AgengItem") = False Then
                Exit Sub
            End If

            AgengItem.show

        Case 6

            If checkApility("SalesTargetSettings") = False Then
                Exit Sub
            End If

            SalesTargetSettings.show

        Case 7

            If checkApility("FrmSalesRePGroups") = False Then
                Exit Sub
            End If

            FrmSalesRePGroups.show

        Case 8

       '     If checkApility("FrmSalesRepData") = False Then
       '         Exit Sub
       '     End If
'
'            FrmSalesRepData.show
' Unload FrmPay_Garanty_Shipment
             If checkApility("FrmPay_Garanty_Shipment") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment.SendForm = 7
FrmPay_Garanty_Shipment.show


           Case 9

      '      If checkApility("Gbasic") = False Then
      '          Exit Sub
      '      End If
'
'            Gbasic.show
' Unload FrmPay_Garanty_Shipment
             If checkApility("FrmPay_Garanty_Shipment") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment.SendForm = 5
FrmPay_Garanty_Shipment.show

            Case 10
                    If checkApility("FrmTypeDiscards") = False Then
                Exit Sub
            End If
            FrmTypeDiscards.show
            
       Case 11
           

            If checkApility("FrmPaymentType") = False Then
                Exit Sub
            End If
FrmPaymentType.mIndex = 1 'ÿ—ﬁ «·œ›⁄

            FrmPaymentType.show


            
        
    End Select

End Sub

Private Sub SalesBasicSubsub_Click(Index As Integer)
Select Case Index

Case 0


   If checkApility("FrmItemsClass") = False Then
                Exit Sub
            End If
            Unload FrmItemsClass
  FrmItemsClass.mIndex = 0
  FrmItemsClass.show
    
    
Case 1


   If checkApility("FrmItemsClass") = False Then
                Exit Sub
            End If
            Unload FrmItemsClass
  FrmItemsClass.mIndex = 1
  FrmItemsClass.show
        
        
Case 2
   If checkApility("FrmCreditFacicity") = False Then
                Exit Sub
            End If
            
      '    FormRequestOpenAccount.show
    FrmCreditFacicity.show
    
    
    
    Case 3
  
   If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If

            'FrmCustemers
            OpenScreen CustomersScreen '
            
      Case 4
      
         If checkApility("FrmCustCash") = False Then
                Exit Sub
            End If
      FrmCustCash.show
      
            
End Select
End Sub

Private Sub SalesInsSub_Click(Index As Integer)
Select Case Index
    Case 0
          If checkApility("FrmBuyGoodsInst") = False Then
                Exit Sub
            End If
      FrmBuyGoodsInst.show
      
     Case 1
        If checkApility("FrmCreditFacicity") = False Then
                Exit Sub
            End If

     FrmCreditFacicity.show
      
Case 2
   If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If

 
            OpenScreen CustomersScreen '
             
     Case 3
     
            If checkApility("FrmSaleBill") = False Then
                Exit Sub
            End If
            OpenScreen InvoiceScreen

Case 4
    If checkApility("FrmReceiptPart") = False Then
        Exit Sub
    End If

    OpenScreen ReceiptPartScreen

Case 5
         If checkApility("System_alarms") = False Then
               Exit Sub
            End If

            System_alarms.show

Case 6
          If checkApility("ReportPurchase") = False Then
                Exit Sub
            End If

            FrmReports.show
            FrmReports.C1TabMain.CurrTab = 15
            


            
End Select


End Sub

Private Sub SalesTransactions_Click(Index As Integer)

    Select Case Index

        Case 0
            'If checkApility("FrmTemplate") = False Then
            '    Exit Sub
            'End If

            'FrmTemplate
            'OpenScreen TemplateScreen

        Case 1
            'FrmShowPrice

            'GeneralPriceType = 0
            'If checkApility("FrmShowPrice") = False Then
            '    Exit Sub
            'End If

            'OpenScreen ScreensName.ShowPriceScreen
        Case 2

            If checkApility("FrmSaleBill") = False Then
                Exit Sub
            End If

       '     Dim RsOptions As New ADODB.Recordset
       '     Set RsOptions = New ADODB.Recordset
        '    RsOptions.Open "tbloptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
            'If RsOptions("out") = 1 Then
            'FrmOut.Show
            'Else
            'FrmSaleBill
            OpenScreen InvoiceScreen

            'End If
        Case 3

            If checkApility("FrmReturnSalling") = False Then
                Exit Sub
            End If

            'FrmReturnSalling
            OpenScreen RetrunSalles

        Case 4
        
              If checkApility("frmsalebillCompose") = False Then
                Exit Sub
            End If

            frmsalebillCompose.show

        Case 5

            If checkApility("Frmovers") = False Then
                Exit Sub
            End If

            Frmovers.show

        Case 6

            If checkApility("FrmSallingPlan") = False Then
                Exit Sub
            End If

            'OpenScreen ItemsPricePlane
            FrmSallingPlan.show

        Case 7

            If checkApility("FrmSallingPlan") = False Then
                Exit Sub
            End If

            OpenScreen ItemsMainPriceLise

        Case 9

            If checkApility("Ageng_all1") = False Then
                Exit Sub
            End If

            Ageng_all.show

        Case 10

            If checkApility("ReportSales") = False Then
                Exit Sub
            End If

            FrmReports.show
            FrmReports.C1TabMain.CurrTab = 0
            'FrmReports.EleMain(0).Enabled = True
            
        
        Case 11

            If checkApility("FrmCustomerReports") = False Then
                Exit Sub
            End If

            FrmCustomerReports.show
          
    End Select

End Sub

Private Sub SalesTransactionsEmp_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("FrmSalesRepComm") = False Then
                Exit Sub
            End If

            FrmSalesRepComm.show

        Case 1

            If checkApility("FrmSalesRepCommtarget") = False Then
                Exit Sub
            End If

            FrmSalesRepCommtarget.show

        Case 2

            If checkApility("FrmSalesRepCommtargetPercentage") = False Then
                Exit Sub
            End If

            FrmSalesRepCommtargetPercentage.show '  Not Log File

        Case 3

            If checkApility("FrmSalesRepCommValues") = False Then
                Exit Sub
            End If

            FrmSalesRepCommValues.show ' Not Log File

        Case 4

            If checkApility("FrmAcceleratePayment") = False Then
                Exit Sub
            End If

            FrmAcceleratePayment.show
    End Select

End Sub

Private Sub SalesTransactionssubss0_Click(Index As Integer)

    Select Case Index

        Case 0
            'If checkApility("FrmTemplate") = False Then
            '    Exit Sub
            'End If

            'FrmTemplate
            'OpenScreen TemplateScreen

    End Select

End Sub

Private Sub SalesTransactionssubss00_Click(Index As Integer)
Select Case Index
Case 0
GeneralPriceType = 0
 If checkApility("FrmPO") = False Then
                 Exit Sub
             End If
FrmPO.show

Case 1
  'FrmApprovalTransactions.screenName = "FrmPO"
 'FrmApprovalTransactions.show

 'FrmApprovalTransactions.loadFlexGrid
 
 GeneralPriceType = 0
 If checkApility("FrmPO1") = False Then
                 Exit Sub
             End If
FrmPO1.show

Case 2
      If checkApility("FrmQotation") = False Then
                 Exit Sub
             End If
FrmQotation.show
'GeneralPriceType = 1

'FrmPOApp.show
'Case 2
End Select
End Sub

Private Sub SalesTransactionssubss000_Click(Index As Integer)
 
Select Case Index
Case 0
GeneralPriceType = 0

 If checkApility("FrmPO2") = False Then
                 Exit Sub
             End If
             FrmPO2.show

Case 1
 
 If checkApility("FrmPO3") = False Then
                 Exit Sub
             End If
GeneralPriceType = 0
FrmPO3.show
 
 
            ' GeneralPriceType = 0

            'If checkApility("FrmShowPrice") = False Then
            '    Exit Sub
            'End If
'
'            OpenScreen ScreensName.ShowPriceScreen
End Select

 

End Sub

Private Sub SearchInHelp_Click()

End Sub

Private Sub Stores_Click(Index As Integer)

End Sub

Private Sub ShpmentBasicdata_Click(Index As Integer)
Select Case Index
Case 0

Case 1
frmShipmentFollow.show
Case 2
frmSipmentAllocation.show
Case 3

Case 4

          If checkApility("FrmProductionPlan") = False Then
                Exit Sub
            End If
            
            FrmProductionPlan.show
            FrmProductionPlan.CplanType.ListIndex = 3
            If SystemOptions.UserInterface = ArabicInterface Then
                FrmProductionPlan.Caption = "    Œÿ… «·‘Õ‰   "
                FrmProductionPlan.Ele(5).Caption = FrmProductionPlan.Caption
            Else
            FrmProductionPlan.Caption = "   Shipment Plan  "
                FrmProductionPlan.Ele(5).Caption = FrmProductionPlan.Caption

            End If
            
Case 5
      If checkApility("FrmShipmentOrder") = False Then
                Exit Sub
            End If
FrmShipmentOrder.show

Case 6
      If checkApility("FrmShipmentRegestration") = False Then
                Exit Sub
            End If
FrmShipmentRegestration.show
Case 7
     If checkApility("FrmShipmentRegestration1") = False Then
                Exit Sub
            End If
FrmShipmentRegestration1.show

Case 8
 
     If checkApility("FrmShippingReport") = False Then
                Exit Sub
            End If
 


FrmShippingReport.show
End Select
End Sub

Private Sub ShpmentBasicdatasub_Click(Index As Integer)
Select Case Index
        Case 0

            If checkApility("FrmCountriesData") = False Then
                Exit Sub
            End If

            FrmCountriesData.show

        Case 1

            If checkApility("FrmGovernmentData") = False Then
                Exit Sub
            End If

            FrmGovernmentData.show


        Case 2

            If checkApility("FrmCitiesDistance") = False Then
                Exit Sub
            End If

            FrmCitiesDistance.show


        Case 3

            If checkApility("FrmGovernCitiesData") = False Then
                Exit Sub
            End If

            FrmGovernCitiesData.show

        Case 4

            If checkApility("streets") = False Then
                Exit Sub
            End If

            streets.show
 
 Case 5
             If checkApility("CarTypes") = False Then
                Exit Sub
            End If

            CarTypes.show



        Case 6

            If checkApility("FrmCars") = False Then
                Exit Sub
            End If

            FrmCars.show

    Case 7
            If checkApility("FrmDrivers") = False Then
                Exit Sub
            End If

            FrmDrivers.show



    Case 8
            If checkApility("FrmTypesofshipping") = False Then
                Exit Sub
            End If

            FrmTypesofshipping.show



    Case 9
       '     If checkApility("FRMMaintenanceTypes") = False Then
       '         Exit Sub
       '     End If
'
'            FRMMaintenanceTypes.show
            


End Select
End Sub

Private Sub StockControlBasicSub_Click(Index As Integer)

    Select Case Index

        Case 0

            'FrmItems
            If checkApility("FrmItems") = False Then
                Exit Sub
            End If

            OpenScreen ItemsDataScreen

        Case 1

            If checkApility("FrmStoreData") = False Then
                Exit Sub
            End If

            'FrmStoreData
            OpenScreen StoresDataScreen

        Case 2

            If checkApility("FrmGroups") = False Then
                Exit Sub
            End If

            'FrmGroups
            OpenScreen ItemsGroupsScreen

        Case 3

       '     If checkApility("FrmSystemUnites") = False Then
       '         Exit Sub
       '     End If

       '     FrmSystemUnites.show

Unload FrmPay_Garanty_Shipment3M
             If checkApility("FrmPay_Garanty_Shipment3M") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment3M.SendForm = 0
FrmPay_Garanty_Shipment3M.show

        Case 4

         '   If checkApility("FrmItemsColor") = False Then
         '       Exit Sub
         '   End If
'
'            FrmItemsColor.show
Unload FrmPay_Garanty_Shipment3M
             If checkApility("FrmPay_Garanty_Shipment3M") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment3M.SendForm = 1
FrmPay_Garanty_Shipment3M.show


        Case 5

     '       If checkApility("FrmItemsSize") = False Then
     '           Exit Sub
     '       End If
'
'            FrmItemsSize.show

Unload FrmPay_Garanty_Shipment3M
             If checkApility("FrmPay_Garanty_Shipment3M") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment3M.SendForm = 2
FrmPay_Garanty_Shipment3M.show

        Case 6

             If checkApility("FrmItemsClass") = False Then
                 Exit Sub
             End If
             Unload FrmItemsClass
 FrmItemsClass.mIndex = 2
             FrmItemsClass.show '
 


        Case 7

     '       If checkApility("FrmStoresLocation") = False Then
     '           Exit Sub
     '       End If
'
'            FrmStoresLocation.show
Unload FrmPay_Garanty_Shipment3M
             If checkApility("FrmPay_Garanty_Shipment3M") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment3M.SendForm = 3
FrmPay_Garanty_Shipment3M.show

        Case 8

         '   If checkApility("FrmSalePriceNames") = False Then
         '       Exit Sub
         '   End If
'
'            FrmSalePriceNames.show

      '    If checkApility("FrmSpecification") = False Then
      '          Exit Sub
      '      End If
'
'            FrmSpecification.show
'
Unload FrmPay_Garanty_Shipment3M
             If checkApility("FrmPay_Garanty_Shipment3M") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment3M.SendForm = 4
FrmPay_Garanty_Shipment3M.show

  

        Case 9

      '      If checkApility("FrmProductionElements") = False Then
      '          Exit Sub
      '      End If
'
'            FrmProductionElements.show

Unload FrmPay_Garanty_Shipment3M
             If checkApility("FrmPay_Garanty_Shipment3M") = False Then
                 Exit Sub
             End If
FrmPay_Garanty_Shipment3M.SendForm = 5
FrmPay_Garanty_Shipment3M.show

        Case 10

            If checkApility("UnitsIndustrialCost") = False Then
                Exit Sub
            End If

            UnitsIndustrialCost.show

        Case 11

            If checkApility("frmitemsalessPlan") = False Then
                Exit Sub
            End If

Case 12
    If checkApility("FrmLinkItemToStore") = False Then
                Exit Sub
            End If
FrmLinkItemToStore.show
Case 13
    If checkApility("FrmBeforeInventory") = False Then
                Exit Sub
            End If
FrmBeforeInventory.show


            'frmitemsalessPlan

    End Select

End Sub

Private Sub StrategyBasicdata_Click(Index As Integer)
Select Case Index
Case 1
     If checkApility("FrmMinistryContract") = False Then
                Exit Sub
            End If


FrmMinistryContract.show
Case 2



     If checkApility("FrmSuperVisorSchoolAllocation") = False Then
                Exit Sub
            End If

 
FrmSuperVisorSchoolAllocation.show

Case 3


     If checkApility("FrmDriverAllocation") = False Then
                Exit Sub
            End If

 
FrmDriverAllocation.show


Case 4


     If checkApility("FrmVehicleAllocation") = False Then
                Exit Sub
            End If

 
FrmVehicleAllocation.show

Case 5
     If checkApility("FrmAttributionContract") = False Then
                Exit Sub
            End If


FrmAttributionContract.show




Case 6
     If checkApility("FrmConfirmVaction") = False Then
                Exit Sub
            End If


FrmConfirmVaction.show
FrmConfirmVaction.WindowState = 0
Case 7
     If checkApility("FrmConfirmViolation") = False Then
                Exit Sub
            End If
FrmConfirmViolation.show


Case 8

    If checkApility("FrmRequest_MinistryContract") = False Then
                Exit Sub
            End If
 

FrmRequest_MinistryContract.show


Case 9
     If checkApility("FrmRequest1") = False Then
                Exit Sub
            End If


FrmRequest1.show

Case 10
     If checkApility("FrmExchangeRequest") = False Then
                Exit Sub
            End If


FrmExchangeRequest.show






Case 11
      If checkApility("FrmPayments") = False Then
                Exit Sub
            End If

 
 FrmPayments.show
Case 12

      If checkApility("FrmStopDealing") = False Then
                Exit Sub
            End If

 
 FrmStopDealing.show
 
 
 Case 13
       If checkApility("FrmAddExceptionDays") = False Then
                Exit Sub
            End If

 
 FrmAddExceptionDays.show
 Case 14
 

      If checkApility("FrmReport_Scenes") = False Then
                Exit Sub
            End If

 
 frmReport_Scenes.show
 
 
End Select

End Sub

Private Sub StrategyBasicdatasub_Click(Index As Integer)
Select Case Index

Case 0
       If checkApility("FrmGovernmentData") = False Then
                Exit Sub
            End If

            FrmGovernmentData.show
            
  Case 1
  
        If checkApility("FrmManagerialArea") = False Then
                Exit Sub
            End If
            
  FrmManagerialArea.show
  FrmManagerialArea.WindowState = 0

  
Case 2

            If checkApility("FrmCompany") = False Then
                Exit Sub
            End If
FrmCompany.mIndex = 1
            FrmCompany.show
     
         
FrmCompany.EleHeader.Caption = "»Ì«‰«  «·„ ⁄ÂœÌ‰"
FrmCompany.Caption = FrmCompany.EleHeader.Caption
FrmCompany.chkCustomerandVendor.Visible = False
FrmCompany.Fra(4).Visible = False
FrmCompany.Fra(6).Visible = False
FrmCompany.CmdPriceList.Visible = False


 
            
  Case 3
  
       If checkApility("CarTypes") = False Then
                Exit Sub
            End If

            CarTypes.show
            CarTypes.Caption = "«‰Ê«⁄ «·Õ«›·« "
            CarTypes.Label1(2).Caption = CarTypes.Caption
             
     Case 5
            If checkApility("FrmCars") = False Then
              Exit Sub
           End If
            FrmCars.show
            FrmCars.Caption = "»Ì«‰«  «·Õ«›·« "
     '       FrmCars.Ele.Caption = FrmCars.Caption
                 FrmCars.Image2.Visible = True
                 FrmCars.lbl(7).Visible = False
                 FrmCars.TxtEquQty.Visible = False
                 FrmCars.Label4.Visible = False
                 
               FrmCars.WindowState = 0



     
  Case 6
              If checkApility("FrmSchooleFile") = False Then
              Exit Sub
           End If
            FrmSchooleFile.show
  FrmSchooleFile.WindowState = 0
  Case 7
              If checkApility("FrmYearDurations") = False Then
              Exit Sub
           End If
            FrmYearDurations.show
            FrmYearDurations.WindowState = 0
            
            
            
Case 8


       If checkApility("FrmViolationGroups") = False Then
              Exit Sub
           End If
            FrmViolationGroups.show
            
          '    FrmViolationGroups.WindowState = 0
            
            


          Case 9
              If checkApility("FrmViolationTypes") = False Then
              Exit Sub
           End If
            FrmViolationTypes.show
            
              FrmViolationTypes.WindowState = 0
            
              Case 10
              If checkApility("FrmVactionTypes") = False Then
              Exit Sub
           End If
            FrmVactionTypes.show
            
            FrmVactionTypes.WindowState = 0
            
                   Case 11
           Unload FrmEmployee

            'FrmEmployee
            If checkApility("FrmEmployee") = False Then
                Exit Sub
            End If
            
FrmEmployee.DriverOnly = 1
FrmEmployee.show
FrmEmployee.Caption = "«·”«∆ﬁÌ‰"
            FrmEmployee.EleHeader = FrmEmployee.Caption
            

 
End Select
End Sub

Private Sub SupBackColor_Click()
 
End Sub

Private Sub SupFont_Click()
     End Sub

Private Sub SupForeColor_Click()
 
 
End Sub

Private Sub StudentMenueSub_Click(Index As Integer)
Select Case Index
Case 0


            If checkApility("FrmStudentBasicData") = False Then
                Exit Sub
            End If

            FrmStudentBasicData.show
   
           
            
Case 1

  If checkApility("FrmInstructors") = False Then
                Exit Sub
            End If

            FrmInstructors.show

Case 2

            If checkApility("FrmCompanies") = False Then
                Exit Sub
            End If

            FrmCompanies.show


Case 3

             If checkApility("FrmTrainingRequest") = False Then
                Exit Sub
            End If

            FrmTrainingRequest.show
Case 4
            If checkApility("FrmStudents") = False Then
                Exit Sub
            End If

            FrmStudents.show

            
            

            

Case 5

            If checkApility("FrmContStudent") = False Then
                Exit Sub
            End If

            FrmContStudent.show
 Case 6
             If checkApility("FrmStudentsCandidacy") = False Then
                Exit Sub
            End If

            FrmStudentsCandidacy.show
            
 Case 7
             If checkApility("FrmStudCandidAccept") = False Then
                Exit Sub
            End If

            FrmStudCandidAccept.show
 

 
 Case 8
 
    If checkApility("FrmGroupStudents") = False Then
                Exit Sub
            End If

            FrmGroupStudents.show
            

 Case 9
 
    If checkApility("FrmAttendance") = False Then
                Exit Sub
            End If

            FrmAttendance.show
                       
 Case 10
 
    If checkApility("FrmStudentCalling") = False Then
                Exit Sub
            End If

            FrmStudentCalling.show
            
 Case 11
 
    If checkApility("FrmStudTermination") = False Then
                Exit Sub
            End If

            FrmStudTermination.show
           
           
           
           Case 12
                                If checkApility("FrmEndExtenGroups") = False Then
                Exit Sub
            End If
           FrmEndExtenGroups.show
  
  
 Case 13
                      If checkApility("FrmStudTermiCompany") = False Then
                Exit Sub
            End If
            
 FrmStudTermiCompany.show
 Case 14
         If checkApility("FrmIssuBillStudent") = False Then
                Exit Sub
            End If
            
 FrmIssuBillStudent.show
  
   Case 15
         If checkApility("FrmGroupStudentsAdd") = False Then
                Exit Sub
            End If
            
 FrmGroupStudentsAdd.show
  
  'FrmGroupStudentsAdd
  Case 16
                       
                         If checkApility("FrmReportsStudent") = False Then
                Exit Sub
            End If
               FrmReportsStudent.show
            
End Select
End Sub

Private Sub Tailorsub_Click(Index As Integer)
Select Case Index '
Case 0
            If checkApility("dean") = False Then 'jobs
                Exit Sub
            End If
            
dean.mIndex = 0
dean.show

Case 1 'size
            If checkApility("dean") = False Then
                Exit Sub
            End If
dean.mIndex = 1
dean.show
Case 2 'items
            If checkApility("FrmItems") = False Then
                Exit Sub
            End If
FrmItems.show
Case 3 'employee
Unload FrmEmployee

            'FrmEmployee
            If checkApility("FrmEmployee") = False Then
                Exit Sub
            End If

            OpenScreen EmployeesScreen
FrmEmployee.WorkShop_Job = 0

Case 4 'customer
       If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If

            OpenScreen CustomersScreen '




Case 5 'orders
            If checkApility("dean") = False Then
                Exit Sub
            End If
dean.mIndex = 3
dean.show
Case 6
'sales
            If checkApility("FrmSaleBill") = False Then
                Exit Sub
            End If
 
            OpenScreen InvoiceScreen
            
Case 7
'cashing
           'FrmCashing
            If checkApility("FrmCashing") = False Then
                Exit Sub
            End If

            OpenScreen CashingDataScreen


Case 8 'empprod
            If checkApility("dean") = False Then
                Exit Sub
            End If
dean.mIndex = 4
dean.show

 

End Select

End Sub

Private Sub TaxexSub_Click(Index As Integer)
Select Case Index
Case 0
          If checkApility("FrmBeforeInventoryK") = False Then
                Exit Sub
            End If
FrmBeforeInventoryK.show

Case 1
           If checkApility("FrmAddedValueVAT") = False Then
                Exit Sub
            End If
 Unload FrmAddedValueVAT
 FrmAddedValueVAT.CtranIndex = 22
  FrmAddedValueVAT.show
 
Case 2
           If checkApility("FrmAddedValueVAT") = False Then
                Exit Sub
            End If
 Unload FrmAddedValueVAT
 FrmAddedValueVAT.CtranIndex = 21
  FrmAddedValueVAT.show
 Case 3
           If checkApility("FrmAddedValueVAT") = False Then
                Exit Sub
            End If
 Unload FrmAddedValueVAT
 FrmAddedValueVAT.CtranIndex = 5
  FrmAddedValueVAT.show

Case 4
           If checkApility("FrmAddedValueVAT") = False Then
                Exit Sub
            End If
 Unload FrmAddedValueVAT
  FrmAddedValueVAT.show
FrmAddedValueVAT.CtranIndex = 9
Case 5
           If checkApility("FrmAddedValueVAT") = False Then
                Exit Sub
            End If
 Unload FrmAddedValueVAT
 FrmAddedValueVAT.CtranIndex = 11
  FrmAddedValueVAT.show
Case 6
           If checkApility("FrmAddedValueVAT") = False Then
                Exit Sub
            End If
 Unload FrmAddedValueVAT
 FrmAddedValueVAT.CtranIndex = 12
  FrmAddedValueVAT.show
             
             
             
             Case 7
                         If checkApility("FrmDiscounts") = False Then
                Exit Sub
            End If
             FrmDiscounts.show
             
             Case 8
                        If checkApility("FrmVATAvowal") = False Then
                Exit Sub
            End If
             FrmVATAvowal.show
             
             Case 9
           If checkApility("Ageng_all") = False Then
                Exit Sub
            End If
            Unload Ageng_all
            Ageng_all.Indx = 1
Ageng_all.show


Case 10
           If checkApility("FrmReCalVATPO") = False Then
                Exit Sub
            End If
            Unload FrmReCalVATPO
            FrmReCalVATPO.show
FrmReCalVATPO.show
End Select
End Sub

Private Sub Texh_Click(Index As Integer)
Dim Msg As String
    Select Case Index

        Case 0
    If user_id <> 1 Then
        '   MsgBox ""
        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
        '    Msg = Msg & Chr(13) & "Õ ‰Â“— Ê·««ÌÂ "
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    
            If checkApility("SMSSeTTings") = False Then
                Exit Sub
            End If

            SMSSeTTings.show

        Case 1

            If checkApility("FrmPlainMessage") = False Then
                Exit Sub
            End If

            FrmPlainMessage.show

        Case 2

            If checkApility("FrmDEfineMessage") = False Then
                Exit Sub
            End If

            FrmDEfineMessage.show

        Case 3

            If checkApility("FrmCustomerBalances1") = False Then
                Exit Sub
            End If

            FrmCustomerBalances1.show
            
         Case 4
             If user_id <> 1 Then
        '   MsgBox ""
        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
        '    Msg = Msg & Chr(13) & "Õ ‰Â“— Ê·««ÌÂ "
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    
'         EmailSettings.show
            
    End Select

End Sub

Private Sub TimerAlret_Timer()
Exit Sub
 '  If Messnger = False Then Exit Sub
If AlarmAutoTime < 5 Then
AlarmAutoTime = AlarmAutoTime + 1
'Exit Sub
Else
AlarmAutoTime = 0
End If
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
      
        Dim StrSQL As String

        StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.ExpectedtimeTime, dbo.ApprovalData.SendTime, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, "
 StrSQL = StrSQL + "        dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, dbo.ApprovalData.currorder, dbo.ApprovalData.FromUser, dbo.ApprovalData.Transaction_ID,"
 StrSQL = StrSQL + "      dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks, dbo.TbLLevels.Name, dbo.TbLLevels.Namee, dbo.Screens.ScreenCaption,"
 StrSQL = StrSQL + "     dbo.Screens.ScreenTitleEng, dbo.ApprovalData.Currcursor, dbo.ApprovalData.id AS searchid, dbo.ApprovalData.NoteSerial, dbo.ApprovalData.Transaction_Date"
 StrSQL = StrSQL + "   FROM         dbo.ApprovalData left JOIN"
 StrSQL = StrSQL + "    dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
 StrSQL = StrSQL + "    dbo.Screens ON dbo.ApprovalData.ScreenName = dbo.Screens.ScreenName"
 
        StrSQL = StrSQL + "   Where (dbo.ApprovalData.Currcursor = 1) And (dbo.ApprovalData.EmpID = " & user_id & ")"
       
          

     
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        sndPlay App.path & "\sound\NewSms.wav", SND_ASYNC Or SND_NODEFAULT
    Unload FrmApprovalTransactions
    FrmApprovalTransactions.show
    
        
     
    End If

    rs.Close
 
End Sub

Private Sub Timer3_Timer()
'MDIForm_Click
'Timer3.Enabled = False
End Sub

 

 

Private Sub Timer1_Timer()
On Error Resume Next

    If Messnger = False Then Exit Sub
If messengerTime < 5 Then
messengerTime = messengerTime + 1
Exit Sub
Else
messengerTime = 0
End If
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "SELECT  *  FROM  Messages  where recived=0 and  [to]='" & user_name & "' order by id desc"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        sndPlay App.path & "\sound\NewSms.wav", SND_ASYNC Or SND_NODEFAULT
        FrmMessnger.show
        FrmMessnger.Adodc4.Refresh

        FrmMessnger.DataGrid2.Refresh
        FrmMessnger.DataGrid4.Refresh

        FrmMessnger.Adodc3.Refresh

        FrmMessnger.DataGrid1.Refresh
        FrmMessnger.DataGrid3.Refresh
        FrmMessnger.SSTab1.Tab = 1
    Else
    End If

    rs.Close
 

End Sub

Private Sub TradingTransaction_Click(Index As Integer)
    Dim rsOut As New ADODB.Recordset
    Dim Msg As String

    Select Case Index

        Case 0

            'FrmOpeningBalance
            If checkApility("FrmOpeningBalance") = False Then
                Exit Sub
            End If

            OpenScreen OpenStockBalance

        Case 1

        Case 2
            Set rsOut = New ADODB.Recordset
            rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

            If Not (rsOut.EOF Or rsOut.BOF) Then
                If rsOut!checkinpo = True Then
                    If checkApility("FrmInpout") = False Then
                        Exit Sub
                    End If

                    FrmInpout.show
  'FrmInpout2.show
                ElseIf rsOut!checkbey = True Then
                    Msg = "⁄›Ê«  „ «Œ Ì«— ›« Ê—… «·‘—«¡ ··«÷«›…  ... ·«Ì„ﬂ‰ «·«÷«›…  „‰ «–‰ «·«÷«›… "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                Else
                End If
            End If

        Case 3

        Case 4

            'FrmMoving
            If checkApility("FrmMoving") = False Then
                Exit Sub
            End If

            OpenScreen StockTransfereScreen

        Case 5

            ' OpenScreen StockCountScreen
        Case 6

            'FrmStockSettlement
            If checkApility("FrmStockSettlement") = False Then
                Exit Sub
            End If

            OpenScreen StockSettlementScreen

        Case 7

       If checkApility("FrmDefinCompItem") = False Then
                Exit Sub
            End If

            FrmDefinCompItem.show


        Case 8
            On Error GoTo ErrTrap

            If checkApility("FrmSearchSerial") = False Then
                Exit Sub
            End If

            FrmSearchSerial.show vbModal
            Exit Sub
ErrTrap:

        Case 9
            'FrmSerialData
            OpenScreen CheckItemSerial

        Case 10

            If checkApility("FrmRequest") = False Then
                Exit Sub
            End If

'            If ShowRequest(True) = True Then
                FrmRequest.show
                FrmRequest.ZOrder 0
'            End If

        Case 11
            ShowItemsStatusReport WindowTarget

            'FrmInventoryStatus.Show
        Case 12

            If checkApility("ReportItems") = False Then
                Exit Sub
            End If

            FrmReports.show
            FrmReports.C1TabMain.CurrTab = 6

        Case 13
            GeneralPriceType = 5

            If checkApility("FrmShowPrice3") = False Then
                Exit Sub
            End If

            FrmShowPrice.show
    End Select

End Sub

Private Sub TradingTransactionSub_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("FrmStartGard") = False Then
                Exit Sub
            End If

            FrmStartGard.show

        Case 1

            If checkApility("FrmGardReport") = False Then
                Exit Sub
            End If

            FrmGardReport.show

        Case 2

            If checkApility("FrmNewGard") = False Then
                Exit Sub
            End If

            FrmNewGard.show

        Case 3

            If checkApility("FrmNewGard1") = False Then
                Exit Sub
            End If

            FrmNewGard1.show
            'OpenScreen StockCountScreen

    End Select

End Sub

Private Sub TradingTransactionSub1_Click(Index As Integer)
    Dim rsOut As New ADODB.Recordset
    Dim Msg As String

    Select Case Index
Case 0
               If checkApility("FrmPO11") = False Then
                        Exit Sub
                    End If

                    FrmPO11.show
                    

        Case 1
           
            Set rsOut = New ADODB.Recordset
            rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

            If Not (rsOut.EOF Or rsOut.BOF) Then
                If rsOut!checkout = True Then
                    If checkApility("FrmOut") = False Then
                        Exit Sub
                    End If

                    FrmOut.show
                ElseIf rsOut!checksal = True Then
                    Msg = "⁄›Ê«  „ «Œ Ì«— ›« Ê—… «·»Ì⁄ ··Œ’„  ... ·«Ì„ﬂ‰ «·Œ’„ „‰ «–‰ «·’—› "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                Else
                End If
            End If
            
        Case 2

            Set rsOut = New ADODB.Recordset
            rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

            If Not (rsOut.EOF Or rsOut.BOF) Then
                If rsOut!checkout = True Then
                    If checkApility("FrmOut1") = False Then
                        Exit Sub
                    End If

                    FrmOut1.show
                ElseIf rsOut!checksal = True Then
                    Msg = "⁄›Ê«  „ «Œ Ì«— ›« Ê—… «·»Ì⁄ ··Œ’„  ... ·«Ì„ﬂ‰ «·Œ’„ „‰ «–‰ «·’—› "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                Else
                End If
            End If
            
    End Select

End Sub

Private Sub TransporterSub_Click(Index As Integer)
 
    Select Case Index

        Case 0

            If checkApility("FrmGovernmentData") = False Then
                Exit Sub
            End If

            FrmGovernmentData.show

        Case 1

            If checkApility("FrmCitiesDistance") = False Then
                Exit Sub
            End If
FrmCitiesDistance.Indx = 0
            FrmCitiesDistance.show


        Case 2

            If checkApility("FrmCitiesDistance") = False Then
                Exit Sub
            End If
FrmCitiesDistance.Indx = 1
            FrmCitiesDistance.show
            
                    Case 3

            If checkApility("FrmCitiesDistance") = False Then
                Exit Sub
            End If
FrmCitiesDistance.Indx = 2
            FrmCitiesDistance.show
            
     Case 4
            If checkApility("FrmCitiesDistance") = False Then
                Exit Sub
            End If
FrmCitiesDistance.Indx = 3
            FrmCitiesDistance.show
             
             
            
     Case 5
            If checkApility("FrmCitiesDistance") = False Then
                Exit Sub
            End If
FrmCitiesDistance.Indx = 4
            FrmCitiesDistance.show
              
              
            'xxxxxxxxxxxxxx
        Case 6

            If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If

            OpenScreen CustomersScreen '

        Case 7

            If checkApility("FrmCompany") = False Then
                Exit Sub
            End If
FrmCompany.mIndex = 0
            FrmCompany.show

        Case 8

            If checkApility("FrmDrivers") = False Then
                Exit Sub
            End If

            FrmDrivers.show

        Case 9

            If checkApility("CarTypes") = False Then
                Exit Sub
            End If

            CarTypes.show
            
            
      Case 10 '«·ÿ—«“« 
  If checkApility("FrmCarModels") = False Then
                Exit Sub
            End If
            FrmCarModels.show
            
 
        Case 11

            If checkApility("insurancecompanies1") = False Then
                Exit Sub
            End If

            insurancecompanies.show

      Case 12 '«‰Ê«⁄ «·’Ì«‰Â
                   If checkApility("FrmMaintenTypes") = False Then
                Exit Sub
            End If
            
FrmMaintenTypes.show


        Case 13

            If checkApility("FrmCars") = False Then
                Exit Sub
            End If

            FrmCars.show

        Case 14

            If checkApility("FrmCarsPlan") = False Then
                Exit Sub
            End If

          FrmCarsPlan.show
          
        Case 15

            If checkApility("FrmClientTransContr") = False Then
                Exit Sub
            End If
'
          FrmClientTransContr.show
          
          Case 16
               If checkApility("FrmOrderUpload") = False Then
                Exit Sub
            End If
'
          FrmOrderUpload.show
        Case 17

            If checkApility("FrmTravelTransactions") = False Then
                Exit Sub
            End If

            FrmTravelTransactions.show

     Case 18

          If checkApility("FrmPaymenTransTrip") = False Then
                Exit Sub
           End If

            FrmPaymenTransTrip.show

      Case 19

                If checkApility("Nationality") = False Then
                Exit Sub
            End If
Nationality.mIndex = 2
            Nationality.show

            
        Case 20

            If checkApility("frmTravelRports") = False Then
                Exit Sub
            End If

            frmTravelRports.show

    End Select

End Sub

Private Sub UserAbility_Click()
    Dim Msg As String
    
    'If SystemOptions.usertype = UserNormal Then
    If user_id <> 1 And SystemOptions.PlaywithAuthorityMatrix = False Then
        '   MsgBox ""
        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
        '    Msg = Msg & Chr(13) & "Õ ‰Â“— Ê·««ÌÂ "
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

 
     
        FrmPermission.show
        FrmPermission.ZOrder 0
 

End Sub

Private Sub UserRpt_Click()
    Dim Msg As String
    'If user_id <> 1 Then
 
    '    Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
    '
    '    MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    Exit Sub
    'End If
 
    If checkApility("FrmUsersLogReports") = False Then
        Exit Sub
    End If



    FrmUsersLogReports.show
End Sub

Private Sub UsersData_Click()
    'MDIFrmMain.Arrange vbCascade
End Sub

Private Sub Load_MenusIcons()
    'On Error GoTo ErrTrap
    Dim MyFont As New StdFont

    With Me.PopMenu1

        If SystemOptions.UserInterface = ArabicInterface Then
            .RightToLeft = True
        Else
            .RightToLeft = False
        End If

        .OfficeXpStyle = True
        MyFont.Name = "MS Sans Serif"
        MyFont.Bold = False
        MyFont.Charset = 178
        MyFont.size = 8
        Set .Font = MyFont
        '.SubClassMenu Me
        .ImageList = Me.ilsIcons
        '.ItemIcon("BasicDataM(0)") = Me.ilsIcons.ItemIndex("Employess") - 1
    
        '   .ItemIcon("Employee(1)") = Me.ilsIcons.ItemIndex("Employess") - 1   'Me.ImgLstMenuIcons.ListImages("New").Index - 1
        '   .ItemIcon("Employee(3)") = Me.ilsIcons.ItemIndex("patients") - 1
        '   .ItemIcon("Employee(4)") = Me.ilsIcons.ItemIndex("User") - 1    '
        '   .ItemIcon("Groups") = Me.ilsIcons.ItemIndex("Groups") - 1
        '  .ItemIcon("Items(0)") = Me.ilsIcons.ItemIndex("Items") - 1
        '   .ItemIcon("Stores(0)") = Me.ilsIcons.ItemIndex("ClosedBox") - 1
        '   .ItemIcon("Stores(1)") = Me.ilsIcons.ItemIndex("Dollar") - 1
        '   .ItemIcon("Stores(2)") = Me.ilsIcons.ItemIndex("Dollar") - 1
        '   .ItemIcon("Exit") = Me.ilsIcons.ItemIndex("Exit") - 1
        '   .ItemIcon("PriceList") = Me.ilsIcons.ItemIndex("PriceList") - 1
        '   .ItemIcon("Trading(5)") = Me.ilsIcons.ItemIndex("invoice") - 1
        '   .ItemIcon("Trading(6)") = Me.ilsIcons.ItemIndex("Purchase") - 1
        '   .ItemIcon("Trading(7)") = Me.ilsIcons.ItemIndex("Return") - 1
        '   .ItemIcon("Trading(7)") = Me.ilsIcons.ItemIndex("Return") - 1
        '   .ItemIcon("Trading(9)") = Me.ilsIcons.ItemIndex("Maintenence") - 1
        '   .ItemIcon("Trading(12)") = Me.ilsIcons.ItemIndex("Cal") - 1
        '   .ItemIcon("Trading(13)") = Me.ilsIcons.ItemIndex("Store") - 1
        '   .ItemIcon("Trading(17)") = Me.ilsIcons.ItemIndex("task") - 1
        '   .ItemIcon("Trading(18)") = Me.ilsIcons.ItemIndex("Search") - 1
        '   .ItemIcon("Report") = Me.ilsIcons.ItemIndex("Report") - 1
        '   .ItemIcon("DailyReport") = Me.ilsIcons.ItemIndex("Reportd") - 1
        '
        '    If Me.ilsIcons.KeyExists("Connect") = True Then
        '        .ItemIcon("MnuToolsDataBase") = Me.ilsIcons.ItemIndex("Connect") - 1
        '    End If
        '    .ItemIcon("MnuDataBaseTools") = Me.ilsIcons.ItemIndex("DataBaseTools") - 1
        '        .ItemIcon("MnuDataBaseTools_Items(0)") = Me.ilsIcons.ItemIndex("DataBaseBackup") - 1
        '        .ItemIcon("MnuDataBaseTools_Items(1)") = Me.ilsIcons.ItemIndex("DataBaseRestore") - 1
        '        .ItemIcon("MnuDataBaseTools_Items(3)") = Me.ilsIcons.ItemIndex("DataBaseFilter") - 1
        '
        '    .ItemIcon("Barcode") = Me.ilsIcons.ItemIndex("BarCode") - 1
        '    .ItemIcon("Trading(14)") = Me.ilsIcons.ItemIndex("Execute") - 1
        '    .ItemIcon("UsersData") = Me.ilsIcons.ItemIndex("partners") - 1
        '        .ItemIcon("AddUser") = Me.ilsIcons.ItemIndex("AddUser") - 1
        '        .ItemIcon("DelUser") = Me.ilsIcons.ItemIndex("DelUser") - 1
        '        .ItemIcon("EditPw") = Me.ilsIcons.ItemIndex("UserPass") - 1
        '        .ItemIcon("UserAbility") = Me.ilsIcons.ItemIndex("Pass") - 1
        '    .ItemIcon("Options") = Me.ilsIcons.ItemIndex("Maintenence") - 1
        '    .ItemIcon("HelpFile") = Me.ilsIcons.ItemIndex("Help") - 1
        '    .ItemIcon("About") = Me.ilsIcons.ItemIndex("About") - 1
        '    .ItemIcon("ConnectUs") = Me.ilsIcons.ItemIndex("Phone") - 1
     
        '    .ItemIcon("ExpensesType(0)") = Me.ilsIcons.ItemIndex("copy") - 1
        '    .ItemIcon("ExpensesType(1)") = Me.ilsIcons.ItemIndex("copy") - 1
    
        ' .ItemIcon("Expenses") = Me.ilsIcons.ItemIndex("Bank") - 1
        '   .ItemIcon("Cashing") = Me.ilsIcons.ItemIndex("Currency") - 1
    
        '    .ItemIcon("MnuBackColor") = Me.ilsIcons.ItemIndex("Back") - 1
        '    .ItemIcon("MnuForeColor") = Me.ilsIcons.ItemIndex("Fore") - 1
        '    .ItemIcon("FormatFONT") = Me.ilsIcons.ItemIndex("Font") - 1
        '    .ItemIcon("ShowCol") = Me.ilsIcons.ItemIndex("Col") - 1
        '    .ItemIcon("ShowItems") = Me.ilsIcons.ItemIndex("clock") - 1
        ''    .ItemIcon("ItemsPrice") = Me.ilsIcons.ItemIndex("Bank") - 1
    
        '   .ItemIcon("AddItem") = Me.ilsIcons.ItemIndex("ADD") - 1
        '   .ItemIcon("DelItem") = Me.ilsIcons.ItemIndex("Del") - 1
        '  .ItemIcon("PriceChips") = Me.ilsIcons.ItemIndex("Bank") - 1
        ''   .ItemIcon("PriceOffer") = Me.ilsIcons.ItemIndex("Currency") - 1
        '  .ItemIcon("SupBackColor") = Me.ilsIcons.ItemIndex("Back") - 1
        '  .ItemIcon("SupForeColor") = Me.ilsIcons.ItemIndex("Fore") - 1
        '  .ItemIcon("SupFont") = Me.ilsIcons.ItemIndex("Font") - 1
        '
        '  .ItemIcon("PopEmployee") = Me.ilsIcons.ItemIndex("Employess") - 1 'Me.ImgLstMenuIcons.ListImages("New").Index - 1
        '  .ItemIcon("PopClients") = Me.ilsIcons.ItemIndex("patients") - 1
        '  .ItemIcon("PopSupliers") = Me.ilsIcons.ItemIndex("User") - 1    '
        '  .ItemIcon("PopGroups") = Me.ilsIcons.ItemIndex("Groups") - 1
        '  .ItemIcon("PopItems") = Me.ilsIcons.ItemIndex("Items") - 1
        '  .ItemIcon("PopStore") = Me.ilsIcons.ItemIndex("ClosedBox") - 1
        '  .ItemIcon("PopBanks") = Me.ilsIcons.ItemIndex("Dollar") - 1
        '  .ItemIcon("PopPriceList") = Me.ilsIcons.ItemIndex("PriceList") - 1
        '  .ItemIcon("PopSallBill") = Me.ilsIcons.ItemIndex("invoice") - 1
        '  .ItemIcon("PopPurchaseBill") = Me.ilsIcons.ItemIndex("Purchase") - 1
        '  .ItemIcon("PopReturn") = Me.ilsIcons.ItemIndex("Return") - 1
        '  .ItemIcon("PopMaintanence") = Me.ilsIcons.ItemIndex("Maintenence") - 1
        ''  .ItemIcon("PopBalance") = Me.ilsIcons.ItemIndex("Cal") - 1
        ' .ItemIcon("PopGard") = Me.ilsIcons.ItemIndex("Store") - 1
        ' .ItemIcon("PopAvailable") = Me.ilsIcons.ItemIndex("task") - 1
        ' .ItemIcon("PopSerialData") = Me.ilsIcons.ItemIndex("Search") - 1
        ' .ItemIcon("PpBarcode") = Me.ilsIcons.ItemIndex("BarCode") - 1
        ' .ItemIcon("Trading(19)") = Me.ilsIcons.ItemIndex("Less") - 1
        ' .ItemIcon("HelpIndex") = Me.ilsIcons.ItemIndex("PriceList") - 1
        ' .ItemIcon("SearchInHelp") = Me.ilsIcons.ItemIndex("Search") - 1
        '  .ItemIcon("Trading(0)") = Me.ilsIcons.ItemIndex("ShowPrice") - 1
        '  .ItemIcon("DelayVal") = Me.ilsIcons.ItemIndex("clock") - 1
        ' .ItemIcon("Trading(4)") = Me.ilsIcons.ItemIndex("Option") - 1
        '.ItemIcon("Payments") = Me.ilsIcons.ItemIndex("Edit") - 1
        '    .ItemIcon("ReportDesigner") = Me.ilsIcons.ItemIndex("Report") - 1
        ' .ItemIcon("ReceiptPart") = Me.ilsIcons.ItemIndex("Cascade") - 1
        ' If Me.ilsIcons.KeyExists("Recycle") = True Then
        '  '   .ItemIcon("Destruction") = Me.ilsIcons.ItemIndex("Recycle") - 1
        ' End If
        ' .ItemIcon("Trading(7)") = Me.ilsIcons.ItemIndex("Required") - 1

        ' .ItemIcon("VacancyType(2)") = Me.ilsIcons.ItemIndex("VacancyType") - 1
        ' .ItemIcon("VacancyType(3)") = Me.ilsIcons.ItemIndex("Planner") - 1
        '.ItemIcon("EmployeSalary") = Me.ilsIcons.ItemIndex("Report") - 1
        ' .ItemIcon("Premium") = Me.ilsIcons.ItemIndex("premium") - 1
        ' .ItemIcon("Discounts") = Me.ilsIcons.ItemIndex("discount") - 1
        ' .ItemIcon("ComingRecord") = Me.ilsIcons.ItemIndex("clock") - 1
        '    .ItemIcon("LeavingRecord") = Me.ilsIcons.ItemIndex("ComeTime") - 1
        ' .ItemIcon("AbsenceRecord") = Me.ilsIcons.ItemIndex("CardEdit") - 1
        ' .ItemIcon("EmployeSalary") = Me.ilsIcons.ItemIndex("Currency") - 1
        '--------------------------------------------------------------------
        ' If Me.ilsIcons.KeyExists("Refresh") = True Then
        '     .ItemIcon("MnuPopItemsTreePane_Array(0)") = Me.ilsIcons.ItemIndex("Refresh") - 1
        ' End If
        ' If Me.ilsIcons.KeyExists("Dock") = True Then
        '     .ItemIcon("MnuPopItemsTreePane_Array(2)") = Me.ilsIcons.ItemIndex("Dock") - 1
        ' End If
    End With

    Exit Sub
ErrTrap:

    If SystemOptions.SysRegisterState = DevelopVersion Then
        Stop
    End If

End Sub

Public Sub LoadInterface(IntInterface As SystemInterface)
    Dim XPanel As MSComctlLib.Panel
    Dim i As Integer
    Dim xPane As XtremeDockingPane.Pane
    Dim XFont As IFontDisp

    'XFont.name = "Tahoma"
    'XFont.Charset = 178
    'Set Me.PopMenu1.Font = XFont
    'Me.PopMenu1.Font.name = "Tahoma"
    'Me.PopMenu1.Font.Charset = 178
    Screen.MousePointer = vbArrowHourglass

    If IntInterface = ArabicInterface Then
        SystemOptions.UserInterface = ArabicInterface
        App.title = GetAppTitle
        Me.RightToLeft = True
        Me.PopMenu1.RightToLeft = True
    
        With Me.XPStusBar
            .Panels.Clear
            Set XPanel = .Panels.Add(, "Pan_Date", "", , mdifrmmain.ImgLstMenuIcons.ListImages("Calender").ExtractIcon)
            XPanel.Style = sbrDate
            XPanel.Alignment = sbrRight
            XPanel.ToolTipText = "«· «—ÌŒ «·Õ«·Ï ›Ï «·ÃÂ«“"
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrContents
            .Refresh
            Set XPanel = .Panels.Add(, "Pan_Time", "", , mdifrmmain.ImgLstMenuIcons.ListImages("Clock").ExtractIcon)
            XPanel.Style = sbrTime
            XPanel.Alignment = sbrRight
            XPanel.ToolTipText = "«·Êﬁ  «·Õ«·Ï ›Ï «·ÃÂ«“"
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrContents
            .Refresh
            Set XPanel = .Panels.Add(, "Pan_Caps", "", , mdifrmmain.ImgLstMenuIcons.ListImages("Caps").ExtractIcon)
            XPanel.Style = sbrCaps
            XPanel.Alignment = sbrRight
            XPanel.ToolTipText = "CapsLock-ﬂ «»… «·Õ—Ê› ﬂ»Ì—… √„ ’€Ì—… "
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrContents
            .Refresh
            Set XPanel = .Panels.Add(, "Pan_Num", "", , mdifrmmain.ImgLstMenuIcons.ListImages("Num").ExtractIcon)
            XPanel.Style = sbrNum
            XPanel.Alignment = sbrRight
            XPanel.ToolTipText = "NumLock-„›« ÌÕ «·√—ﬁ«„ ›Ï «·Ì„Ì‰ „‰ ·ÊÕ… «·„›« ÌÕ"
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrContents
            .Refresh
            Set XPanel = .Panels.Add(, "lang", "", , mdifrmmain.ImgLstMenuIcons.ListImages("KeyBorad").ExtractIcon)
            XPanel.Style = sbrText
            XPanel.Alignment = sbrRight
            XPanel.ToolTipText = "«··€… «·‰‘ÿ… „‰ ·ÊÕ… «·„›« ÌÕ"
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrContents
            .Refresh
            Set XPanel = .Panels.Add(, "User", "«”„ «·„” Œœ„:" & user_name, , mdifrmmain.ImgLstMenuIcons.ListImages("User").ExtractIcon)
            XPanel.Style = sbrText
            XPanel.Alignment = sbrRight
            XPanel.ToolTipText = "«·„” Œœ„ «·Õ«·Ï ··»—‰«„Ã"
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrContents
        
            If SystemOptions.SysDataBaseType = AccessDataBase Then
                Set XPanel = .Panels.Add(, "DataBase", "ﬁ«⁄œ… «·»Ì«‰« :„Ìﬂ—Ê”Ê›  «ﬂ””", , mdifrmmain.ImgLstMenuIcons.ListImages("DataBase").ExtractIcon)
            Else
                Set XPanel = .Panels.Add(, "DataBase", "ﬁ«⁄œ… «·»Ì«‰« :SQL Server 2000 ", , mdifrmmain.ImgLstMenuIcons.ListImages("DataBase").ExtractIcon)
            End If

            XPanel.Style = sbrText
            XPanel.Alignment = sbrRight
            XPanel.ToolTipText = "‰Ê⁄ ﬁ«⁄œ… «·»Ì«‰«  «· Ï Ì⁄„· ⁄·ÌÂ« «·»—‰«„Ã"
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrContents
        
            .Refresh

            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
                Set XPanel = .Panels.Add(, "AccountIntervalID", "«·› —… «·„Õ«”»Ì… «·Õ«·Ì… : " & SystemOptions.SysCurrentAccountIntervalID, , mdifrmmain.ImgLstMenuIcons.ListImages("DataBase").ExtractIcon)
                XPanel.Style = sbrText
                XPanel.Alignment = sbrRight
                XPanel.ToolTipText = "—ﬁ„ «·› —… «·„Õ«”»Ì… «·Õ«·Ì…"
                XPanel.Bevel = sbrInset
                XPanel.MinWidth = 1
                XPanel.AutoSize = sbrContents
            End If

            Set XPanel = .Panels.Add(, "Pan_Comment", App.title, , mdifrmmain.Icon)
            XPanel.Style = sbrText
            XPanel.Alignment = sbrRight
            XPanel.ToolTipText = "«·–Â«» ≈·Ï „Êﬁ⁄ BYTE"
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrSpring
            .Refresh
            .Panels("Pan_Comment").Width = .Width - (.Panels("Pan_Date").Width + .Panels("lang").Width + .Panels("Pan_Time").Width + .Panels("Pan_Caps").Width + .Panels("Pan_Num").Width + .Panels("User").Width)
            'MsgBox "End Me.XPStusBar"
        End With

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.OutBarPaneID)

        If Not xPane Is Nothing Then
            xPane.title = "‘—Ìÿ «·≈Œ ’«—« "
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.NewsBarPaneID)

        If Not xPane Is Nothing Then
            xPane.title = "„⁄·Ê„«  «·»—‰«„Ã"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID)

        If Not xPane Is Nothing Then
            xPane.title = "‘Ã—… «·√’‰«›"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.MantainceID)

        If Not xPane Is Nothing Then
            xPane.title = "«·’Ì«‰…"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.InternetNews)

        If Not xPane Is Nothing Then
            xPane.title = "„⁄·Ê„«  «·≈‰ —‰ "
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.DynamicHelp)

        If Not xPane Is Nothing Then
            xPane.title = ""
             DockingPane1.FindPane(DockingPanesIDs.DynamicHelp).MinTrackSize.setSize 10, 100
 
           xPane.Closed = False 'panelpanel
           xPane.Enabled = PaneEnableClient
       '   xPane.Enabled = PaneEnabled
            xPane.MaxTrackSize.setSize 150, 50
            xPane.MinTrackSize.setSize 100, 50
           'xPane.Type
          '  xPane.Enabled = PaneEnableClient ' PaneEnableActions ' PaneDisabled 'ABUSAUD
     xPane.Enabled = PaneEnabled
       xPane.Closed = True
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.CalendarPaneID)

        If Not xPane Is Nothing Then
            xPane.title = "«·”«⁄… "
        End If

        Me.XPStusBar.Refresh
        
    ElseIf IntInterface = EnglishInterface Then
        SystemOptions.UserInterface = EnglishInterface
        App.title = GetAppTitle
        Me.RightToLeft = False
        Me.PopMenu1.RightToLeft = False

        With Me.XPStusBar
            .Panels.Clear
            Set XPanel = .Panels.Add(, "Pan_Comment", App.title, , mdifrmmain.Icon)
            XPanel.Style = sbrText
            XPanel.Alignment = sbrLeft
         '   XPanel.ToolTipText = "Goto  BYTE"
            XPanel.Bevel = sbrInset
            XPanel.AutoSize = sbrSpring
        
            If SystemOptions.SysDataBaseType = AccessDataBase Then
                Set XPanel = .Panels.Add(, "DataBase", "DataBase:Microsoft Access", , mdifrmmain.ImgLstMenuIcons.ListImages("DataBase").ExtractIcon)
            Else
                Set XPanel = .Panels.Add(, "DataBase", "DataBase:SQL Server 2000", , mdifrmmain.ImgLstMenuIcons.ListImages("DataBase").ExtractIcon)
            End If

            XPanel.Style = sbrText
            XPanel.Alignment = sbrRight
            XPanel.ToolTipText = "The DataBase Type Which the Programe Used."
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrContents

            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
                Set XPanel = .Panels.Add(, "AccountIntervalID", "Current Accounting Interval : " & SystemOptions.SysCurrentAccountIntervalID, , mdifrmmain.ImgLstMenuIcons.ListImages("DataBase").ExtractIcon)
                XPanel.Style = sbrText
                XPanel.Alignment = sbrRight
         '       XPanel.ToolTipText = "Current Open Accounting Interval Number"
                XPanel.Bevel = sbrInset
                XPanel.MinWidth = 1
                XPanel.AutoSize = sbrContents
            End If
        
            Set XPanel = .Panels.Add(, "User", "Current User:" & user_name, , mdifrmmain.ImgLstMenuIcons.ListImages("User").ExtractIcon)
            XPanel.Style = sbrText
            XPanel.Alignment = sbrLeft
            XPanel.ToolTipText = "The Current System User"
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrContents
        
            Set XPanel = .Panels.Add(, "lang", "", , mdifrmmain.ImgLstMenuIcons.ListImages("KeyBorad").ExtractIcon)
            XPanel.Style = sbrText
            XPanel.Alignment = sbrLeft
            XPanel.ToolTipText = "The Active KeyBorad Language"
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrContents
        
            Set XPanel = .Panels.Add(, "Pan_Num", "", , mdifrmmain.ImgLstMenuIcons.ListImages("Num").ExtractIcon)
            XPanel.Style = sbrNum
            XPanel.Alignment = sbrLeft
            XPanel.ToolTipText = "The State Of The Num Lock Key"
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrContents
        
            Set XPanel = .Panels.Add(, "Pan_Caps", "", , mdifrmmain.ImgLstMenuIcons.ListImages("Caps").ExtractIcon)
            XPanel.Style = sbrCaps
            XPanel.Alignment = sbrLeft
            XPanel.ToolTipText = "The State Of The Caps Lock Key"
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrContents
        
            Set XPanel = .Panels.Add(, "Pan_Time", "", , mdifrmmain.ImgLstMenuIcons.ListImages("Clock").ExtractIcon)
            XPanel.Style = sbrTime
            XPanel.Alignment = sbrLeft
            XPanel.ToolTipText = "The Current System Time"
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrContents
        
            Set XPanel = .Panels.Add(, "Pan_Date", "", , mdifrmmain.ImgLstMenuIcons.ListImages("Calender").ExtractIcon)
            XPanel.Style = sbrDate
            XPanel.Alignment = sbrLeft
            XPanel.ToolTipText = "The Current System Date"
            XPanel.Bevel = sbrInset
            XPanel.MinWidth = 1
            XPanel.AutoSize = sbrContents
        
            .Panels("Pan_Comment").Width = .Width - (.Panels("Pan_Date").Width + .Panels("lang").Width + .Panels("Pan_Time").Width + .Panels("Pan_Caps").Width + .Panels("Pan_Num").Width + .Panels("User").Width)
        End With

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.OutBarPaneID)
'XPanel.Alignment =
        If Not xPane Is Nothing Then
            xPane.title = "Shortcut OutBar"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.NewsBarPaneID)

        If Not xPane Is Nothing Then
            xPane.title = "Programe Information"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID)

        If Not xPane Is Nothing Then
            xPane.title = "Items Tree"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.MantainceID)

        If Not xPane Is Nothing Then
            xPane.title = "Maintenance"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.InternetNews)

        If Not xPane Is Nothing Then
            xPane.title = "Internet Information"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.DynamicHelp)

        If Not xPane Is Nothing Then
            xPane.title = ""
                      xPane.Closed = True 'panelpanel abosaud
            xPane.Enabled = PaneEnableClient
      '         xPane.Enabled = PaneDisabled
      '      xPane.MinTrackSize.setSize 50, 100
         '   XPanel.Alignment = sbrRight
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.CalendarPaneID)

        If Not xPane Is Nothing Then
            xPane.title = "Calendar"
        End If

        Me.XPStusBar.Refresh
    End If

    Me.Caption = App.title

    With Me.PopMenu1

        If Me.PopMenu1.Tag = "" Then
            SetMenus
            .SubClassMenu Me
            .Tag = "1"
        Else
            .UnsubclassMenu
            SetMenus
            MenuItemShow True
            .SubClassMenu Me
        End If

    End With

    SetMenusHelp
    Load_MenusIcons
    MenuItemShow False

  '  If Not FrmOutBarPane Is Nothing Then
  '      FrmOutBarPane.LoadInterface SystemOptions.UserInterface '
  '  End If

   ' If Not FrmNewsBarPane Is Nothing Then
   '     FrmNewsBarPane.CreateTaskPanel
   ' End If

    'Public Enum DockingPanesIDs

    'End Enum
    Screen.MousePointer = vbDefault

End Sub

Private Sub MenuItemShow(BolShow As Boolean)

    'Me.MnuView_Item(3).Visible = BolShow

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        MnuReports_Assblied.Visible = BolShow
    End If

    If SystemOptions.SysAppAccoutingType = SimpleAccoutning Then
        Me.MnuAccounts.Visible = BolShow
    End If

    'Me.MnuCurrencyData.Visible = BolShow
End Sub

Private Sub VacancyType_Click(Index As Integer)

End Sub

Private Sub UsersGroup_Click()

         If checkApility("FrmGroupUsers") = False Then
                Exit Sub
            End If

            FrmGroupUsers.show
            
            
End Sub

Private Sub Vscstionsssub_Click(Index As Integer)

    Select Case Index


        Case 0

            If checkApility("FrmInstalVacation") = False Then
                Exit Sub
            End If

            FrmInstalVacation.show
            
            
            Case 1
                        If checkApility("FrmLastVacation") = False Then
                Exit Sub
            End If

            FrmLastVacation.show
            
            
        Case 2

            If checkApility("FrmHolidayPlan") = False Then
                Exit Sub
            End If

            FrmHolidayPlan.show

        Case 3

            If checkApility("formvocatinl") = False Then
                Exit Sub
            End If

           ' FrmHolidayorder.show

formvocatinl.show

 

      Case 4
 
'      If checkApility("FrmHolidayData") = False Then
'        Exit Sub
'    End If
    
'FrmHolidayData.show

        Case 5
 
            If checkApility("frmdriveassestMove") = False Then
                Exit Sub
            End If

           ' FrmFixedAssetMoving.show
frmdriveassestMove.show
        Case 6
'FrmHolidayorder2
            If checkApility("FrmVocationEntitlements") = False Then
                Exit Sub
            End If

            FrmVocationEntitlements.show
            
Case 7

          If checkApility("FrmExitvisasReturn") = False Then
                Exit Sub
            End If

FrmExitvisasReturn.show
        Case 8

            If checkApility("FrmEmbarkation") = False Then
                Exit Sub
            End If

           ' FrmHolidayorder3.show

FrmEmbarkation.show
    
  Case 9
            If checkApility("FrmRegsterSickleave") = False Then
                Exit Sub
            End If

           ' FrmHolidayorder3.show

FrmRegsterSickleave.show

End Select

End Sub

Private Sub XC_Click(Index As Integer)

    Select Case Index

        Case 0
        '    GeneralPriceType = 3

            If checkApility("FrmPO6") = False Then
                Exit Sub
            End If

            FrmPO6.show

        Case 1
       '     GeneralPriceType = 4

            If checkApility("FrmPO7") = False Then
                Exit Sub
            End If

            FrmPO7.show
            
    End Select

End Sub

Private Sub XPStusBar_PanelClick(ByVal Panel As MSComctlLib.Panel)

    Select Case Panel.Key

        Case "WebSite"
            OpenWebSite
    End Select

End Sub

Private Sub SetMenus()

    'On Error GoTo ErrTrap
    If SystemOptions.UserInterface = ArabicInterface Then
    
    
    
   SalesInsSub(0).Caption = "ÿ·» ‘—«¡ »«· ﬁ”Ìÿ"
    SalesInsSub(1).Caption = "ÿ·» › Õ Õ”«» ⁄„Ì·"
    SalesInsSub(2).Caption = "«·⁄„·«¡"
    SalesInsSub(3).Caption = "›« Ê—… „»Ì⁄«   ﬁ”Ìÿ"
    SalesInsSub(4).Caption = " Õ’Ì· «·«ﬁ”«ÿ"
    SalesInsSub(5).Caption = "«· ‰»ÌÂ« "
    
    
   SalesInsSub(6).Caption = "«· ﬁ«—Ì—"
   
   AgeingSub(0).Caption = "«⁄œ«œ «⁄„«— «·œÌÊ‰ ··„‘ —Ì« "
   AgeingSub(1).Caption = "«⁄œ«œ «⁄„«— «·œÌÊ‰ ··„»Ì⁄« "
   AgeingSub(2).Caption = " ”ÃÌ· ›Ê« Ì— «·„‘ —Ì«  «·”«»ﬁ…"
   AgeingSub(3).Caption = " ”ÃÌ· ›Ê« Ì— «·„»Ì⁄«  «·”«»ﬁ…"
   AgeingSub(4).Caption = "—»ÿ ›Ê« Ì— «·„»Ì⁄«  «·Õ«·Ì…"
   AgeingSub(5).Caption = "«· ﬁ«—Ì—"
    
   
    
    Strategy.Caption = "«·‰ﬁ· «·„œ—”Ì"
'    GoldMenu.Caption = "Ê—‘  Ê„⁄«—÷ «·œÂ» Ê «·«·„«” "
    dev.Caption = "«·„Â«„ Ê «·«œ«¡"
   CarMaintenance.Caption = "Ê—‘ ’Ì«‰… «·„⁄œ« /«·”Ì«—« "
    CarMaintenancesub(0).Caption = "«·»Ì«‰«  «·«”«”Ì…"
     CarMaintenancesub(1).Caption = "«·Õ—ﬂ« "
     
CarMaintenancesub1(0).Caption = "«‰Ê«⁄ «·„—ﬂ»« "
CarMaintenancesub1(1).Caption = "ÿ—«“«  «·„—ﬂ»« "
CarMaintenancesub1(2).Caption = "»Ì«‰«  «·„—ﬂ»« "
CarMaintenancesub1(3).Caption = "«‰Ê«⁄ «·«’·«Õ« "
CarMaintenancesub1(4).Caption = "«‰Ê«⁄ «·„‘ —Ì«  Ê «·«⁄„«· «·Œ«—ÃÌ…"
CarMaintenancesub1(5).Caption = " «·„‘ —Ì«  Ê «·«⁄„«· «·Œ«—ÃÌ… "

MnuElevatorssUB(0).Caption = " ⁄—Ì› „Õœœ«  «·⁄—Ê÷"
MnuElevatorssUB(1).Caption = "—»ÿ „Õœœ«  «·⁄—Ê÷"
MnuElevatorssUB(2).Caption = "⁄—Ê÷ «·«”⁄«— «·„ Œ’’…"
MnuElevatorssUB(3).Caption = "«·⁄—÷ «·›‰Ì"
MnuElevatorssUB(4).Caption = "«·’Ì«‰… Ê «·÷„«‰"

Elevatorsmaintenance(0).Caption = "«·÷„«‰ Ê⁄ﬁÊœ «·’Ì«‰…"
Elevatorsmaintenance(1).Caption = "’—› ﬁÿ⁄ «·€Ì«—"
Elevatorsmaintenance(2).Caption = "«·’Ì«‰… «·Êﬁ«∆Ì…"
Elevatorsmaintenance(3).Caption = " ‰»ÌÂ«  ⁄ﬁÊœ «·’Ì«‰… «·„‰ ÂÌ…"
Elevatorsmaintenance(4).Caption = " ‰»ÌÂ«  «·÷„«‰«  «·„‰ÂÌ…"

Elevatorsmaintenance(5).Caption = "«· ﬁ«—Ì—"





MnuElevatorssUB(6).Caption = "«· ﬁ«—Ì—"

CarMaintenancesub1(6).Caption = "«‰Ê«⁄  «⁄ÿ«· ›Õ’ «·ﬂ„»ÌÊ —"
CarMaintenancesub1(7).Caption = "«·Ê«‰ «·„—ﬂ»« "
CarMaintenancesub1(8).Caption = "»Ì«‰«  «·„Œ«“‰"
CarMaintenancesub1(9).Caption = "„Ã„Ê⁄«  «·«’‰«›"
CarMaintenancesub1(10).Caption = "ÊÕœ«  «·«’‰«›"
CarMaintenancesub1(11).Caption = "»Ì«‰«  «·«’‰«›"
CarMaintenancesub1(12).Caption = "»Ì«‰«  «·⁄„·«¡"
CarMaintenancesub1(13).Caption = "»Ì«‰«  «·„ÊŸ›Ì‰"
CarMaintenancesub1(17).Caption = "»Ì«‰«  «ﬁ”«„ «·Ê—‘…"
CarMaintenancesub1(18).Caption = "«·„‘—›Ì‰ Ê «·›‰ÌÌ‰"
 
CarMaintenancesub2(0).Caption = "«–‰ œŒÊ· «·’Ì«‰…"
CarMaintenancesub2(1).Caption = "›« Ê—… ›Õ’ ﬂ„»ÌÊ —"
CarMaintenancesub2(2).Caption = "”‰œ ’—› ﬁÿ⁄ €Ì«—"
CarMaintenancesub2(3).Caption = "«Ê«„— «·‘—«¡"
CarMaintenancesub2(4).Caption = "›« Ê—… ··’Ì«‰…"
CarMaintenancesub2(5).Caption = "«·⁄„Ê·«  «·„” Õﬁ…"
CarMaintenancesub2(6).Caption = "«ÃÊ— «·Ìœ"

Texh(0).Caption = "«⁄œ«œ«  ›‰Ì… ··—”«∆· «·‰’Ì… Ê «·«Ì„Ì·« "
Texh(1).Caption = "‰„«–Ã «·—”«∆·"
Texh(2).Caption = " ⁄—Ì› «·—”«∆· ··‘«‘« "
Texh(3).Caption = "—”«∆· «·⁄„·«¡"

CarMaintenancesub(2).Caption = "«· ﬁ«—Ì—"
 
 '*******************************
 HRProcedures(0).Caption = "ÿ·» ”·›…"
 HRProcedures(1).Caption = " ’—ÌÕ Œ—ÊÃ „ƒﬁ  -≈Ã«“… ⁄«—÷…"
 HRProcedures(2).Caption = "  ﬂ·Ì› »„Â„… ⁄„·"
 HRProcedures(3).Caption = "ÿ·» ’—› »œ· ”ﬂ‰"
 HRProcedures(4).Caption = "‰ﬁ· „ÊŸ›"
 HRProcedures(5).Caption = "„»«‘—… „ÊŸ›"
 HRProcedures(7).Caption = "«” »Ì«‰ ⁄‰ „ÊŸ›"
 HRProcedures(8).Caption = "ÿ·» «Ã«“…"
 HRProcedures(9).Caption = "»Ì«‰«  «·«Ã«“…"
 HRProcedures(10).Caption = " ”·Ì„ «·⁄Âœ «·⁄Ì‰Ì…"
 HRProcedures(11).Caption = " ”·Ì„ ÃÊ«“ «·”›—"
 HRProcedures(12).Caption = "  «‰–«— ·„ÊŸ›"
 HRProcedures(13).Caption = "Œÿ«» ·„‰ ÌÂ„Â «·«„—"
 HRProcedures(14).Caption = " ﬁ—Ì— «’«»Â ⁄„·"
 HRProcedures(15).Caption = "«” ·«„ „⁄«„·« "
 HRProcedures(16).Caption = "„Œ«·’… ‰Â«∆Ì…"
 HRProcedures(25).Caption = " ÕœÌÀ »Ì«‰«  «·„ÊŸ›Ì‰"
 HRProcedures(26).Caption = "ÿ·» «Œ·«¡ ÿ—›"
 HRProcedures(27).Caption = " ⁄ﬁÌ» »‘√‰ «Ã—«¡ «œ«—Ì"
 HRProcedures(28).Caption = "„–ﬂ—… Œ’„"
 
 HRProcedures(30).Caption = "Œÿ«»  ⁄—Ì› "
 HRProcedures(31).Caption = " ›ÊÌ÷ «·ﬁÌ«œ… "
 

 '*******************************
  
 POSTRansactiosG.Caption = "‰ﬁ«ÿ «·»Ì⁄"

POSTRansactios(0).Caption = "»Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄"
POSTRansactios(1).Caption = "»Ì«‰«  «·ﬂ«‘Ì—"

POSTRansactios(2).Caption = "»Ì«‰«  «·‘Ì› "
POSTRansactios(3).Caption = "»Ì«‰«  «·„Ê«ﬁ⁄"
POSTRansactios(4).Caption = "«⁄œ«œ«  ‰ﬁ«ÿ «·⁄„·«¡"

'FrmPoints
POSTRansactios(5).Caption = " ”ÃÌ· «·œŒÊ·"
POSTRansactios(6).Caption = "’—› «·„ﬂÊ‰« "

POSTRansactios(7).Caption = "ﬁ»÷ ⁄«„ ‰ﬁ«ÿ «·»Ì⁄"
POSTRansactios(8).Caption = "«· ﬁ«—Ì— "
POSTRansactios(9).Caption = "ÿ»«⁄Â ﬂ—Ê  «·⁄„·«¡ "
POSTRansactios(10).Caption = "«·ﬁ”«∆„ «·„Ã«‰Ì…"

mangDep.Caption = "«·„Ê«—œ «·»‘—Ì…"
mangDepSub(0).Caption = "⁄‰«’— «· ﬁÌÌ„"
mangDepSub(1).Caption = "‰„Ê–Ã «· ÊŸÌ›"
mangDepSub(2).Caption = "«Œÿ«— «Õ Ì«Ã«  ÊŸÌﬁÌ…"



 MarketingMnu.Caption = " ≈œ«—… «· ”ÊÌﬁ"
MarketingMnusub(0).Caption = "«·«⁄œ«œ«  «·⁄«„…"
MarketingMnusub(1).Caption = "⁄—Ê÷ «·«’‰«›"
MarketingMnusub(2).Caption = "„ «»⁄Â  «·⁄„·«¡"
MarketingMnusub(3).Caption = "«· ﬁ«—Ì—"
MarketingMnusub(4).Caption = " ﬁ«—Ì— «·« ’«·« "

MarketingMnusubsub(0).Caption = " ”ÃÌ· „Ê«⁄Ìœ «·⁄„·«¡"
MarketingMnusubsub(1).Caption = " ”ÃÌ· “Ì«—«  «·⁄„·«¡"
MarketingMnusubsub(2).Caption = "„ «»⁄Â “Ì«—«  «·⁄„·«¡"
MarketingMnusubsub(3).Caption = "«” ÿ·«⁄ —√Ì «·⁄„·«¡"
MarketingMnusubsub(4).Caption = " ”ÃÌ· ‘ﬂ«ÊÌ «·⁄„·«¡"
MarketingMnusubsub(5).Caption = "„ «»⁄Â ‘ﬂ«ÊÌ «·⁄„·«¡"
'MarketingMnusubsub(58).Caption = "œ·Ì· «·Â« ›"
MdiContextMenu.Caption = "ﬁ«∆„… «·»—«„Ã"
        Me.Basicdata.Caption = " «·»Ì«‰«   «·«”«”Ì…"
        Me.BasicDataM(0).Caption = "  «‰Ê«⁄ «·„’—Ê›« "
        Me.BasicDataM(1).Caption = "  «‰Ê«⁄ «·«Ì—«œ« "
        Me.BasicDataM(2).Caption = " »Ì«‰«  «·»‰Êﬂ"
        Me.BasicDataM(3).Caption = "»Ì«‰«  «·Œ“‰ Ê «·⁄Âœ"
        Me.BasicDataM(4).Caption = "ÿ—ﬁ «·œ›⁄ "
        Me.BasicDataM(5).Caption = "»Ì«‰«  «·„Ê—œÌ‰"
        
        Me.BasicDataM(6).Caption = "»Ì«‰«  «·⁄„·«¡"

If SystemOptions.AllowScInterface = True Then
Me.BasicDataM(6).Caption = "√Ê·Ì«¡ «·«„Ê—"
SalesBasicSub(1).Caption = Me.BasicDataM(6).Caption
SalesBasicSubsub(1).Caption = Me.BasicDataM(6).Caption
End If




Me.BasicDataM(7).Caption = "»Ì«‰«  «·„ÊŸ›Ì‰"
Me.BasicDataM(8).Caption = "»Ì«‰«  «·«’‰«›"

        Me.BasicDataM(9).Caption = "»Ì«‰«  «·⁄„·« "
        Me.BasicDataM(10).Caption = "»Ì«‰«  «·œÊ· «·Ã‰”Ì« "
        Me.BasicDataM(11).Caption = "»Ì«‰«  «·œÌ«‰« "
        Me.BasicDataM(12).Caption = "»Ì«‰«   «·œÊ·"
        Me.BasicDataM(13).Caption = "»Ì«‰«  «·„œ‰"
        Me.BasicDataM(14).Caption = "»Ì«‰«  «·«ÕÌ«¡"
        Me.BasicDataM(15).Caption = "»Ì«‰«  «·‘Ê«—⁄"
        Me.BasicDataM(17).Caption = "«‰Ê«⁄ «·„” ‰œ«   "
        'Me.BasicDataM(15).Caption = "»Ì«‰«  «·«’‰«›  "
Me.BasicDataM(16).Caption = "«·„‘«—Ì⁄"
        Me.BasicDataM(17).Caption = " ﬁ«—Ì—"
        
        Me.BasicDataM(20).Caption = "  Œ—ÊÃ"
        AssetsMngBase.Caption = "«œ«—… «·«„·«ﬂ"
        
        MnuToolsSetPrinters0sub(0).Caption = "ÿ·» œ⁄„ ›‰Ì"
        MnuToolsSetPrinters0sub(1).Caption = "„ «»⁄Â «·ﬂ«„Ì—« "
       MnuToolsSetPrinters0sub(2).Caption = "œ⁄„ ›‰Ì „ Œ’’"
       MnuToolsSetPrinters0sub(3).Caption = "«·«ﬁ›«·"
       MnuToolsSetPrinters0sub(4).Caption = "„“«„‰Â «·„«ﬂÌ‰« "
       MnuToolsSetPrinters0sub(5).Caption = "«·«”‰«œ"
       MnuToolsSetPrinters0sub(6).Caption = "„Êﬁ› «·“Ì«—« "
       MnuToolsSetPrinters0sub(7).Caption = "„Êﬁ› «· ÃÂÌ“"
       MnuToolsSetPrinters0sub(8).Caption = "≈⁄«œ… «Õ‰”«» «· ﬂ·›…"
       MnuToolsSetPrinters0sub(9).Caption = "÷»ÿ «· ﬂ·›…."
              MnuToolsSetPrinters0sub(10).Caption = "«·œ⁄„ ⁄‰ »⁄œ"
              
              
       UsersGroup.Caption = "„Ã„Ê⁄«  «·„” Œœ„Ì‰"
       
       
        
        mnuEmployee.Caption = "‘∆Ê‰ «·„ÊŸ›Ì‰"
        MnuAccDEV(0).Caption = "«·«ÿ·«⁄ ⁄·Ì «·ﬁÌÊœ «·„Õ«”»Ì…"
        MnuAccDEV(1).Caption = "  ﬁÌÊœ «· ”ÊÌ… «·ÌœÊÌ…"
        
        MnuAccDEV_Post.Caption = "  „—«Ã⁄Â ﬁÌÊœ «·ÌÊ„Ì…"
        xxx(0).Caption = "  «‰Ê«⁄ „—«ﬂ“ «· ﬂ·›…"
        xxx(1).Caption = "  »Ì«‰«  „—«ﬂ“ «· ﬂ·›…"

        xxy(0).Caption = "  «·„Ê«“‰… «·⁄«„…"
        xxy(1).Caption = "  «· œ›ﬁ «·‰ﬁœÌ  "
        xxy(2).Caption = "   »ÊÌ» «·ﬁÊ«∆„"
        xxy(3).Caption = "  Œÿ…  Ê“Ì⁄ «·Õ”«»« "
        xxy(4).Caption = "  «⁄œ«œ „⁄«œ·«  «· Õ·Ì· «·„«·Ì"
        xxy(5).Caption = "  «ŸÂ«— ‰ «∆Ã «· Õ·Ì· «·„«·Ì"
        xxy(6).Caption = "  «·Õ”«»«  «·„Ã„⁄Â  "
        xxy(7).Caption = " «·„’«œﬁ« "
        xxy(8).Caption = "√Ã‰œ… «·⁄„·«¡"
        
        taxes.Caption = "«·ﬁÌ„… «·„÷«›…"
        TaxexSub(0).Caption = "≈⁄œ«œ««  «·ﬁÌ„… «·„÷«›…"
TaxexSub(1).Caption = " ”ÃÌ· «·„‘ —Ì«  ÌœÊÌ«"
TaxexSub(2).Caption = " ”ÃÌ· «·„»Ì⁄«  ÌœÊÌ«"
TaxexSub(3).Caption = " ”ÃÌ· „—œÊœ«  «·„‘ —Ì«  ÌœÊÌ«"
TaxexSub(4).Caption = " ”ÃÌ· „—œÊœ«   «·„»Ì⁄«  ÌœÊÌ«"
TaxexSub(5).Caption = " ”ÃÌ· „‘ —Ì«  «·«·«  Ê«·„⁄œ« "
TaxexSub(6).Caption = " ”ÃÌ·    „»Ì⁄«  «·«·«  Ê«·„⁄œ« "
TaxexSub(7).Caption = "«·«‘⁄«—« "
TaxexSub(8).Caption = "«·«ﬁ—«— «·÷—Ì»Ì"
TaxexSub(9).Caption = "«· ﬁ«—Ì—"
TaxexSub(10).Caption = "«‰‘«¡ ﬁÌÊœ ﬁ .„ ·‰ﬁ«ÿ «·»Ì⁄"


xxy(9).Caption = "«” œ⁄«¡ „Ì“«‰ „—«Ã⁄Â"

xxy(10).Caption = "«·„œ›Ê⁄«  «·„ﬁœ„…"
advancedPayment(0).Caption = " ⁄—Ì›   «·„ﬁœ„« "
advancedPayment(1).Caption = "«À»«    «·„ﬁœ„« "
advancedPayment(2).Caption = "«ÿ›«¡   «·„ﬁœ„« "
advancedPayment(3).Caption = "«À»«    «·»œ·«  «·„ﬁœ„…"

        ProductionPlan.Caption = "„—«ﬁ»… «·ÃÊœ…"
        'xxx(4).Caption = "  «· Õ·Ì· «·„«·Ì"
        ProductionPlansub(0).Caption = "ŒÿÂ «·«‰ «Õ"
        ProductionPlansub(1).Caption = " ⁄—Ì› ⁄‰«’— „—«ﬁ»… «·ÃÊœ…"
        ProductionPlansub(2).Caption = " ’‰Ì› «·„‰ Ã« "
        ProductionPlansub(3).Caption = " ⁄—Ì› «·«Ã—«¡«  «· ’ÕÌÕÌ…"
        ProductionPlansub(4).Caption = "‰„Ê–Ã «· ‘€Ì·"
        ProductionPlansub(5).Caption = "‰„Ê–Ã «·›Õ’"
        ProductionPlansub(6).Caption = "«·ÃÊœ…"
        ProductionPlansub(7).Caption = "„·«ÕŸ… «·«·« "
        
        xxx(12).Caption = "   ﬁ«—Ì— «·Õ”«»« "
        Me.MnuProjects.Caption = " ≈œ«—… «·„‘«—Ì⁄"
        Me.MnuProjectsBasic.Caption = "«·»Ì«‰«  «·«”«”Ì…"
        Me.MnuProjectsBasicSub(0).Caption = "Õ«·«  «·„‘«—Ì⁄"
        Me.MnuProjectsBasicSub(1).Caption = " «‰Ê«⁄ «·⁄ﬁÊœ"
        Me.MnuProjectsBasicSub(2).Caption = "»Ì«‰«  „ﬁ«Ê·Ì «·»«ÿ‰"
Me.MnuProjectsBasicSub(3).Caption = " ⁄—Ì› «·»‰Êœ"

        Me.MnuProjectsBasicSub(4).Caption = "ÊÕœ«  «·⁄„·Ì« "
        Me.MnuProjectsBasicSub(5).Caption = " ⁄—Ì› «·⁄„·Ì« "
        Me.MnuProjectsBasicSub(6).Caption = "»Ì«‰«  «·„⁄œ«  Ê «·«·« "
              
      Me.MnuProjectsTransactions(0).Caption = "»Ì«‰«  «·„‘«—Ì⁄"
        Me.MnuProjectsTransactions(1).Caption = " ”‰œ ’—› „Ê«œ ··„‘«—Ì⁄"
        Me.MnuProjectsTransactions(2).Caption = " ”‰œ „—œÊœ«   „Ê«œ ··„‘«—Ì⁄"
        
        Me.MnuProjectsTransactions(3).Caption = "   Œ’Ì’ «·⁄„«·…"
        Me.MnuProjectsTransactions(4).Caption = "  ‰ﬁ· «·⁄„«·Â"
        
        Me.MnuProjectsTransactions(5).Caption = "   Œ’Ì’ «·„⁄œ«  Ê «··«·«  ··„‘«—Ì⁄"
        Me.MnuProjectsTransactions(6).Caption = "  ‰ﬁ·  «·„⁄œ«  Ê «··«·«  ··„‘«—Ì⁄"
        
        
        Me.MnuProjectsTransactions(7).Caption = "  „ «»⁄Â «·⁄„·Ì«  "
        Me.MnuProjectsTransactions(8).Caption = "  „” Œ·’«  «·„‘«—Ì⁄"
       Me.MnuProjectsTransactions(9).Caption = "  ≈’œ«— „” Œ·’«  «·„‘«—Ì⁄"
        Me.MnuProjectsTransactions(10).Caption = "   ﬁ«—Ì— «·„‘«—Ì⁄"
        mnuEmployeeBasic(0).Caption = "  «·»Ì«‰«  «·«”«”ÌÂ"
        mnuEmployeeBasicSub(0).Caption = "«⁄œ«œ «Êﬁ«  ⁄„· «·‘—ﬂ…"
        mnuEmployeeBasict(0).Caption = "«⁄œ«œ«  «· ﬁÌÌ„"
        mnuEmployeeBasict(1).Caption = "«· ﬁÌÌ„"
        mnuEmployeeBasicSub(1).Caption = "«·‘Ì› « "
        mnuEmployeeBasicSub(2).Caption = "«·«Ã«“« "
        mnuEmployeeBasicSub(3).Caption = "«‰Ê«⁄ «·⁄ﬁÊœ"
        mnuEmployeeBasicSub(4).Caption = "Õ«·«  «·⁄„·"
        mnuEmployeeBasicSub(5).Caption = "»Ì«‰«  «·«œ«—« / «·«ﬁ”«„"
        mnuEmployeeBasicSub(6).Caption = " »Ì«‰«  «·ÊŸ«∆›"
        mnuEmployeeBasicSub(7).Caption = "›—ﬁ «·⁄„·"
mnuEmployeeBasicSub(8).Caption = "«·œ—Ã«  «·ÊŸÌ›Ì…"
mnuEmployeInsuranceSub(0).Caption = "√⁄œ«œ«  «· √„Ì‰«  «·«Ã „«⁄Ì…"
        mnuEmployeInsuranceSub(1).Caption = "»Ì«‰«  ‘—ﬂ«  «· √„Ì‰"
        mnuEmployeInsuranceSub(2).Caption = "»Ì«‰«  «‰Ê«⁄ «· √„Ì‰"
        mnuEmployeInsuranceSub(3).Caption = "»Ì«‰«  ›∆«  «· √„Ì‰"
        mnuEmployeInsuranceSub(4).Caption = "≈Õ ”«» «· √„Ì‰«  «·«Ã „«⁄Ì…"
     '   mnuEmployeeBasicSub(11).Caption = "⁄‰«’— «· ﬁÌÌ„"
     
     mnuEmployeeBasicSub(13).Caption = "«‰Ê«⁄ «–Ê‰«  «·Œ—ÊÃ"
     mnuEmployeeBasicSub(14).Caption = "„Ê«ﬁ⁄ «·⁄„·"
     mnuEmployeeBasicSub(15).Caption = "«·Ã‰”Ì« "
     mnuEmployeeBasicSub(16).Caption = "«·œÌ«‰« "
     mnuEmployeeBasicSub(17).Caption = " ⁄—Ì› «·„ÊÃÊÊœ«  «·⁄Ì‰Ì… - «·⁄Âœ"
     mnuEmployeeBasicSub(18).Caption = "’·Â «· «»⁄Ì‰"
     mnuEmployeeBasicSub(19).Caption = "»Ì«‰«  «·„‰«ÿﬁ / «·ﬁÿ«⁄«  "
    mnuEmployeeBasicSub(20).Caption = "»«‰«  «· √‘Ì—« "
    mnuEmployeeBasicSub(21).Caption = "«‰Ê«⁄ «·Ã“«¡«  «·«œ«—Ì…"
    mnuEmployeeBasicSub(22).Caption = "«⁄œ«œ«  «·«Ã«“… «·„—÷Ì…"
    mnuEmployeeBasicSub(23).Caption = "”Ì«”… «·«Ã«“« "
    
       mnuEmployeeBasic(2).Caption = "«· √„Ì‰«  «·«Ã „«⁄Ì… Ê «·ÿ»Ì…"
       mnuEmployeeBasic(3).Caption = "„ƒ‘—«  «·√œ«¡ «·—∆Ì”Ì…"
        
        mnuEmployeeBasict(0).Caption = "⁄‰«’— «· ﬁÌÌ„"
         mnuEmployeeBasict(1).Caption = "   «· ﬁÌÌ„"
         mnuEmployeeBasict(2).Caption = "«” Õﬁ«ﬁ «· ﬁÌÌ„"
         
        mnuEmployeeBasic(4).Caption = "«·Õ÷Ê— Ê «·«‰’—«›"
        EmployeeAttendanceSub(0).Caption = "«‰Ê«⁄ «·⁄ÿ·« "
         EmployeeAttendanceSub(1).Caption = "«⁄œ«œ«  «·‘Ìﬁ « "
        EmployeeAttendanceSub(2).Caption = "«⁄œ«œ«  «·‰ ÌÃ…"
        EmployeeAttendanceSub(3).Caption = " ”ÃÌ· «·Õ÷Ê— Ê «·«‰’—«› «·Ì"
        EmployeeAttendanceSub(4).Caption = " ”ÃÌ· «·Õ÷Ê— Ê «·«‰’—«› ÌœÊÌ"
        EmployeeAttendanceSub(5).Caption = "«·«⁄ „«œ "
    '    EmployeeAttendanceSub(4).Caption = "«·⁄—÷ «·⁄«„ ·„Ê«⁄Ìœ «·Õ÷Ê— Ê «·«‰’—«›"
       
       mnuEmployeeBasic(5).Caption = "‰„«–Ã «·«Ã—«¡« "
        mnuEmployeeBasic(6).Caption = "«·—Ê« »"
        EmployeeSalarySub(0).Caption = "«‰Ê«⁄ „›—œ«  «·—« »"
        EmployeeSalarySub(1).Caption = "„⁄«œ·«   „›—œ«  «·—« »"
        EmployeeSalarySub(2).Caption = "«·„ﬂ«›√ "
        EmployeeSalarySub(3).Caption = "«·Œ’Ê„« "
        EmployeeSalarySub(4).Caption = " ”ÃÌ· ”·› «·„ÊŸ›Ì‰"
        EmployeeSalarySub(5).Caption = "—œ ”·› «·„ÊŸ›Ì‰"
        EmployeeSalarySub(6).Caption = "„”Ì— «·—« »"
        EmployeeSalarySub(7).Caption = "Õ”«» „ﬂ«›√… ‰Â«Ì… «·Œœ„Â"
        EmployeeSalarySub(8).Caption = " ”ÃÌ· «·„›—œ«  «·„ €Ì—…"
        EmployeeSalarySub(9).Caption = " ”ÃÌ·   «·„Œ’’«  ··«Ã«“«  Ê ‰Â«Ì… «·Œœ„…"
        EmployeeSalarySub(10).Caption = " ”ÃÌ· «” Õﬁ«ﬁ «·„›—œ«  «·”‰ÊÌ…"
        EmployeeSalarySub(11).Caption = "«·“Ì«œ«   "
        EmployeeSalarySub(12).Caption = " €ÌÌ— «—ÌŒ «Ê «Ìﬁ«› ”·›…"

        mnuEmployeeBasic(7).Caption = "«Ã«“«  «·„ÊŸ›Ì‰"
Vscstionsssub(0).Caption = " ”ÃÌ· «·»Ì«‰«  «·«›  «ÕÌ…"
Vscstionsssub(1).Caption = " ”ÃÌ·  «·«Ã«“«  «·”«»ﬁ… "

        Vscstionsssub(2).Caption = "ŒÿÂ «·«Ã«“« "
        Vscstionsssub(3).Caption = "ÿ·» «Ã«“…"
        Vscstionsssub(4).Caption = "»Ì«‰«  «·«Ã«“…"
        Vscstionsssub(5).Caption = " ”·Ì„ Ê ”·„ ⁄Âœ ⁄Ì‰Ì…"
        Vscstionsssub(6).Caption = "„” Õﬁ«  «·«Ã«“…"
Vscstionsssub(7).Caption = "«· √‘Ì—« "
        Vscstionsssub(8).Caption = "„»«‘—… «·⁄„·"
        Vscstionsssub(9).Caption = "«œŒ«· «·«Ã«“«  «·„—÷Ì…"
        
mnuEmployeeBasic(8).Caption = "”·› «·„ÊŸ›Ì‰"
        mnuEmployeeBasic(9).Caption = "«‰Â«¡ «·Œœ„Â"
mnuEmployeeBasic(10).Caption = "Œÿ… «·»œ·«  «·„ﬁœ„…"
        mnuEmployeeBasic(11).Caption = "«· ﬁ«—Ì—   "
        
        FinishSevicersub(0).Caption = "ÿ·» ‰Â«Ì… «·Œœ„Â"
        FinishSevicersub(1).Caption = "Õ”«» „ﬂ«›√… ‰Â«Ì… «·Œœ„Â"
  
        mnuEmployeeBasic(1).Caption = "  »Ì«‰«  «·„ÊŸ›Ì‰"
        EmployeeDataicSub(0).Caption = "  „·› «·„ÊŸ›Ì‰"
        EmployeeDataicSub(1).Caption = "  ⁄ﬁÊœ «·„ÊŸ›Ì‰"
        TransporterMain.Caption = " ≈œ«—… «·‰ﬁ·Ì« "
        TransporterSub(0).Caption = "»Ì«‰«  «·„œ‰"
        TransporterSub(1).Caption = "«·„”«›«  »Ì‰ «·„œ‰"
        TransporterSub(2).Caption = "«·„Ê«‰Ì¡"
        TransporterSub(3).Caption = "«·”›‰"
        TransporterSub(4).Caption = "«‰Ê«⁄ «·‰ﬁ·"
        TransporterSub(5).Caption = " ⁄—Ì› «·—œÊœ"
        
        
        TransporterSub(6).Caption = "»Ì«‰«  «·⁄„·«¡"
        TransporterSub(7).Caption = "»Ì«‰«  «·„Ê—œÌ‰"
        TransporterSub(8).Caption = "»Ì«‰«  «·”«∆ﬁÌ‰"
        TransporterSub(9).Caption = "«‰Ê«⁄ «·„—ﬂ»« "
        TransporterSub(10).Caption = "ÿ—«“«  «·„—ﬂ»« "
        TransporterSub(11).Caption = "‘—ﬂ«  «· √„Ì‰"
        TransporterSub(12).Caption = "«‰Ê«⁄ «·’Ì«‰… «·œÊ—Ì…"
        TransporterSub(13).Caption = "»Ì«‰«  «·„—ﬂ»« "
       TransporterSub(14).Caption = "Œÿ… «·’Ì«‰…"
        TransporterSub(15).Caption = "« ›«ﬁÌ«  «·⁄„·«¡"
        TransporterSub(16).Caption = "√Ê«„— «· Õ„Ì·"
        TransporterSub(17).Caption = "»Ì«‰«  «·—Õ·« "
        TransporterSub(18).Caption = "›Ê« Ì «·⁄„·«¡"
        TransporterSub(19).Caption = "  ’›ÌÂ  «·⁄Âœ… ··”«∆ﬁÌ‰"
TransporterSub(20).Caption = " «· ﬁ«—Ì—"
        Me.StockControl.Caption = "„—«ﬁ»… «·„Œ“Ê‰"
        Me.StockControlBasic.Caption = "«·»Ì«‰«  «·«”«”Ì…"
        StockControlBasicSub(0).Caption = "»Ì«‰«  «·«’‰«›"
        StockControlBasicSub(1).Caption = "»Ì«‰«  «·„Œ«“‰  "
        StockControlBasicSub(2).Caption = "„Ã„Ê⁄«  «·«’‰«›"
        StockControlBasicSub(3).Caption = "«·ÊÕœ« "
        StockControlBasicSub(4).Caption = "«·Ê«‰ «·«’‰«›"
        StockControlBasicSub(5).Caption = "„ﬁ«”«  «·«’‰«›"
        StockControlBasicSub(6).Caption = "›—“ «·«’‰«›"
        StockControlBasicSub(7).Caption = "«⁄œ«œ «„«ﬂ‰ «· Œ“Ì‰"
        StockControlBasicSub(8).Caption = "„Ê«’›«  «·«’‰«›"

        'StockControlBasicSub(9).Caption = "⁄‰«’—  ﬂ«·Ì› «·«‰ «Ã  "
        'StockControlBasicSub(10).Caption = " «· ﬂ«·Ì› «·’‰«⁄Ì… ÿ»ﬁ« ··ÊÕœ…"
        StockControlBasicSub(11).Caption = "Œÿ… „»Ì⁄«  «·«’‰«›"
         StockControlBasicSub(12).Caption = "—»ÿ «·«’‰«› »«·„Œ«“‰"
         StockControlBasicSub(13).Caption = "«⁄œ«œ«  Õœ «·ÿ·»"
         
        Me.TradingTransaction(0).Caption = " «·—’Ìœ «·«›  «ÕÌ"
        Me.TradingTransaction(1).Caption = "«·ÿ·»«  «·œ«Œ·Ì…"
        XC(0).Caption = "ÿ·»«  œ«Œ·Ì…"
        XC(1).Caption = "”‰œ ÕÃ“"
        Me.TradingTransaction(2).Caption = "”‰œ«  «·«” ·«„"
        Me.TradingTransaction(3).Caption = "”‰œ«  «·’—›"
        Me.TradingTransaction(4).Caption = "«· ÕÊÌ· »Ì‰ «·„Œ«“‰"
        Me.TradingTransaction(5).Caption = "Ã—œ «·„Œ«“‰"
        TradingTransactionSub(0).Caption = "»œ√  Ã—œ «·„Œ«“‰"
        TradingTransactionSub(1).Caption = "ÿ»«⁄Â ﬂ‘Ê›«  «·Ã—œ"
        TradingTransactionSub(2).Caption = "«œŒ«· «·ﬂ„Ì«  «·›⁄·Ì…"
        TradingTransactionSub(3).Caption = " ‰›Ì– «·Ã—œ"

        Me.TradingTransaction(6).Caption = " ”ÊÌ… «·„Œ“Ê‰"
        Me.TradingTransaction(7).Caption = "”‰œ«  «· Ã„Ì⁄"
        Me.TradingTransaction(8).Caption = " «·«” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰›"
        Me.TradingTransaction(9).Caption = "»ÕÀ ⁄‰ ”Ì—Ì«·"
        Me.TradingTransaction(10).Caption = "«·«’‰«› «· Ì »·€  Õœ «·ÿ·»"
        Me.TradingTransaction(11).Caption = "„Êﬁ› «·«’‰«› «·Õ«·Ì"
        Me.TradingTransaction(12).Caption = "«· ﬁ«—Ì—"
TradingTransactionSub1(0).Caption = " ÿ·» «·’— › «·„»œ∆Ì"
        TradingTransactionSub1(1).Caption = "”‰œ«  «·’—›/ ”·Ì„"
        TradingTransactionSub1(2).Caption = "”‰œ«  ’—› «·Â«·ﬂ Ê«·⁄Ì‰« "

        Me.Purchase.Caption = "«·„‘ —Ì«  Ê «·„Ê—œÌ‰"
        Me.PurchaseBasicRoot.Caption = "«·»Ì«‰«  «·«”«”Ì…"
        Me.PurchaseBasic(0).Caption = "»Ì«‰«  «·„Ê—œÌ‰"
        Me.PurchaseBasic(1).Caption = "⁄ﬁÊœ «·„Ê—œÌ‰"
        Me.PurchaseBasic(2).Caption = "«⁄œ«œ «⁄„«— «·œÌÊ‰"
        Me.PurchaseBasic(3).Caption = "«‰Ê«⁄ «·‘Õ‰"
        Me.PurchaseBasic(4).Caption = "«‰Ê«⁄ «·÷„«‰« "
        Me.PurchaseBasic(5).Caption = "ÿ—ﬁ «·œ›⁄"

Me.PurchaseBasic(6).Caption = "„Ã„Ê⁄«  «·„‰«œÌ»"
Me.PurchaseBasic(7).Caption = "»Ì«‰«  «·„‰«œÌ» "
Me.PurchaseBasic(8).Caption = "  ÿ—ﬁ «·‘Õ‰"

        Me.PurchaseTransactions(0).Caption = "⁄—Ê÷ «·«”⁄«— Ê √Ê«„—  «·‘—«¡ "
 
        PurchaseTransactionssubd(0).Caption = "⁄—Ê÷ «·«”⁄«—"
        PurchaseTransactionssubs(0).Caption = "ÿ·» ⁄—Ê÷ «”⁄«—"
        PurchaseTransactionssubs(1).Caption = "⁄—Ê÷ «·«”⁄«—"
        PurchaseTransactionssubs(2).Caption = "„ﬁ«—‰Â ⁄—Ê÷ «·«”⁄«—"

        PurchaseTransactionssubd(1).Caption = "ÿ·»«  / √Ê«„— «·‘—«¡"
        PurchaseTransactionssubs1(0).Caption = "ÿ·»«  «·‘—«¡"
        PurchaseTransactionssubs1(1).Caption = "≈⁄ „«œ √„— ‘—«¡"
        PurchaseTransactionssubs1(2).Caption = "√Ê«„— «·‘—«¡"

        FinAnalysis.Caption = "«· Õ·Ì· «·„«·Ì"
  
        Me.PurchaseTransactions(1).Caption = "»Ì«‰«  «·‘Õ‰"
        Me.PurchaseTransactions(2).Caption = "«·«⁄ „«œ«  Ê «·÷„«‰«  «·»‰ﬂÌ…"

        LCTransactions(0).Caption = " «‰Ê«⁄ «·«⁄ „«œ«  Ê«·÷„«‰«  «·»‰ﬂÌ…"
        LCTransactions(1).Caption = "«·›Ê« Ì— «·„»œ∆Ì…"
        LCTransactions(2).Caption = "› Õ «⁄ „«œ „” ‰œÌ/»‰ﬂÌ"
        LCTransactions(3).Caption = " ⁄œÌ·  «⁄ „«œ „” ‰œÌ/»‰ﬂÌ"
        LCTransactions(4).Caption = "„ «»⁄Â «·‘Õ‰« "
        LCTransactions(5).Caption = "”‰œ «” ·«„ ‘Õ‰« "
        LCTransactions(6).Caption = " ›« Ê—… ‰Â«∆Ì…"
        LCTransactions(7).Caption = "€·ﬁ «⁄ „«œ „” ‰œÌ "
        LCTransactions(8).Caption = "ÿ·» ÷„«‰ »‰ﬂÌ"
        LCTransactions(9).Caption = "ÿ·»   „œÌœ ÷„«‰ »‰ﬂÌ"
        LCTransactions(10).Caption = " ÷„«‰ »‰ﬂÌ ‰Â«∆Ì"
        LCTransactions(11).Caption = "‘—«¡ «·„‰«›”Â"

        Me.PurchaseTransactions(3).Caption = "›« Ê—… „‘ —Ì« "
 Me.PurchaseTransactions(4).Caption = "›« Ê—… „‘ —Ì«  „Ã„⁄Â"
 
        Me.PurchaseTransactions(5).Caption = "„—œÊœ«  «·„‘ —Ì« "
        Me.PurchaseTransactions(6).Caption = " ﬁ—Ì— «⁄„«— «·œÌÊ‰"
        Me.PurchaseTransactions(7).Caption = " ﬁ«—Ì— «·„‘ —Ì« "
 
        Me.Sales.Caption = "«·„»Ì⁄«  Ê «·⁄„·«¡"
   
        Me.SalesBasic.Caption = "«·»Ì«‰«  «·«”«”Ì…"
        Me.SalesBasicSub(0).Caption = "«‰Ê«⁄ «·⁄„·«¡"
        Me.SalesBasicSub(1).Caption = "»Ì«‰«  «·⁄„·«¡"
        Me.SalesBasicSub(2).Caption = "⁄ﬁÊœ «·⁄„·«¡"
        Me.SalesBasicSub(3).Caption = "«⁄œ«œ «⁄„«— «·œÌÊ‰ "
        Me.SalesBasicSub(4).Caption = "   ⁄—Ì› «”⁄«— «·»Ì⁄"
        Me.SalesBasicSub(5).Caption = "«⁄œ«œ «·«’‰«› «·—«ﬂœ… "
        Me.SalesBasicSub(6).Caption = "«⁄œ«œ Âœ› «·„»Ì⁄« "
        Me.SalesBasicSub(7).Caption = "„Ã„Ê⁄«  «·„‰«œÌ»"
        Me.SalesBasicSub(8).Caption = "»Ì«‰«  «·„‰«œÌ»"
   Me.SalesBasicSub(9).Caption = "«‰Ê«⁄ ÷„«‰«  «· ﬁ”Ìÿ "
   Me.SalesBasicSub(10).Caption = "«‰Ê«⁄ «·„—œÊœ«   "
   SalesBasicSubsub(0).Caption = "„Ã„Ê⁄«  «·⁄„·«¡"
    SalesBasicSubsub(1).Caption = " ’‰Ì›«  «·⁄„·«¡"
    
   SalesBasicSubsub(2).Caption = "ÿ·» › Õ Õ”«» ⁄„Ì·"
      SalesBasicSubsub(3).Caption = "»Ì«‰«  «·⁄„·«¡"
SalesBasicSubsub(4).Caption = "»Ì«‰«  «·⁄„·«¡ «·‰ﬁœÌ"

        Me.SalesTransactions(0).Caption = "⁄—Ê÷ «·«”⁄«— Ê √Ê«„— «·»Ì⁄ "
 
        SalesTransactionssubss0(0).Caption = "⁄—Ê÷ «·«”⁄«—"
        SalesTransactionssubss00(0).Caption = "ÿ·»«  ⁄—Ê÷ «·«”⁄«— „‰ «·⁄„·«¡"
   '     SalesTransactionssubss00(1).Caption = "«⁄ „«œ ⁄—Ê÷ «·«”⁄«—"
        SalesTransactionssubss00(1).Caption = "⁄—Ê÷ «·«”⁄«— "
   
        SalesTransactionssubss0(1).Caption = "√Ê«„— «·»Ì⁄"
        SalesTransactionssubss000(0).Caption = " √Ê«„— «·»Ì⁄ «·„»œ∆Ì…"
       ' SalesTransactionssubss000(1).Caption = "≈⁄ „«œ √„— »Ì⁄"
        SalesTransactionssubss000(1).Caption = " √Ê«„— «·»Ì⁄"
  
        Me.SalesTransactions(1).Caption = "ÿ·»«  «·»Ì⁄"
        Me.SalesTransactions(2).Caption = "›« Ê—… «·„»Ì⁄« "
        Me.SalesTransactions(3).Caption = "„—œÊœ«  «·„»Ì⁄« "
        Me.SalesTransactions(4).Caption = "›« Ê—… „Ã„⁄Â"
        Me.SalesTransactions(5).Caption = "⁄—Ê÷ «·«’‰«›"
        Me.SalesTransactions(6).Caption = "ŒÿÂ  ”⁄Ì—  «·«’‰«› "
        Me.SalesTransactions(7).Caption = "ﬁ«∆„… «·«”⁄«—"
        Me.SalesTransactions(8).Caption = "„ «»⁄Â «·„‰«œÌ»"
        Me.SalesTransactions(9).Caption = " ﬁ—Ì— «⁄„«— «·œÌÊ‰"
        Me.SalesTransactions(10).Caption = " ﬁ«—Ì— «·„»Ì⁄« "
        Me.SalesTransactions(11).Caption = " ﬁ«—Ì— «·⁄„·«¡ «·‰ﬁœÌÌ‰"
        SalesTransactionsEmp(0).Caption = "«⁄œ«œ ⁄„Ê·«  «·„»Ì⁄«  Ê «· Õ’Ì·« "
        SalesTransactionsEmp(1).Caption = "ŒÿÂ   «·„»Ì⁄«  Ê «· Õ’Ì·« "
        SalesTransactionsEmp(2).Caption = "‰”»Â  Õﬁﬁ   ŒÿÂ ⁄„Ê·«  «·„»Ì⁄«  Ê «· Õ’Ì·« "
        SalesTransactionsEmp(3).Caption = "⁄„Ê·«  «·„‰«œÌ» «·„” Õ›…"
        SalesTransactionsEmp(4).Caption = "”Ì«”Â/⁄—Ê÷  ⁄ÃÌ· «·œ›⁄"
        Archiving.Caption = "«·«—‘Ì› ‰Ÿ«„ «·’«œ— Ê«·Ê—«œ "
        ArchivingSub(0).Caption = "«·«œ«—«  Ê «·«ﬁ”«„"
        ArchivingSub(1).Caption = "»Ì«‰«  «·«—‘Ì›"
        ArchivingSub(2).Caption = "«·€—› ›Ì ﬂ· «—‘Ì›"
        ArchivingSub(3).Caption = "’‰«œÌﬁ «·Õ›Ÿ ›Ì ﬂ· «—‘Ì›"
        ArchivingSub(4).Caption = "«—›› «·Õ›Ÿ ›Ì ﬂ· ’‰œÊﬁ"
        ArchivingSub(5).Caption = "«‰Ê«⁄ «·„⁄«„·« "
        ArchivingSub(6).Caption = "«÷«›… «·‰„«–Ã"
        ArchivingSub(7).Caption = " ”ÃÌ· «·„⁄«„·« "
        ArchivingSub(8).Caption = "„ «»⁄Â «·„⁄«„·« "
        ArchivingSub(9).Caption = " ‰»ÌÂ«  «·„⁄«„·« "
        ArchivingSub(10).Caption = "«· ﬁ«—Ì—"
 taxes.Caption = "«·ﬁÌ„… «·„÷«›…"
 TaxexSub(0).Caption = "«·«⁄œ«œ«« "
 LIFEINDICATORMNU.Caption = "«·„ƒ‘—«  «·ÕÌ…"
 AgeingMAster.Caption = "«⁄„«— «·œÌÊ‰"
 SalesIns.Caption = "«·»Ì⁄ »«· ﬁ”Ìÿ"
ProductionPlan.Caption = "«· ŒÿÌÿ Ê „—«ﬁ»Â «·ÃÊœ…"
MnuElevators.Caption = "«œ«—… «·„’«⁄œ"
rsInvestment.Caption = "«·«” À„«— «·⁄ﬁ«—Ì"
hajMnu.Caption = "«·ÕÃ Ê «·⁄„—…"
StudentMenue.Caption = "«·„⁄«Âœ «· ⁄·Ì„Ì…"
         Me.Currency.Caption = "«·„⁄«„·«  «·„«·ÌÂ"
        Me.ExpensesType(0).Caption = "«‰Ê«⁄ «·„’—Ê›« "
        Me.ExpensesType(1).Caption = "  «‰Ê«⁄ «·«Ì—«œ« "
      Me.ExpensesType(2).Caption = "œ›« — «·‘Ìﬂ« "
      
        Me.Expenses(0).Caption = "«·›Ê« Ì— «·„«·Ì…"
Me.Expenses(1).Caption = "›« Ê—… Œœ„Ì…"
        Me.Expenses(2).Caption = "”‰œ«  «·’—›"
ExpensesSub(0).Caption = "«‰Ê«⁄ «·’—› "
ExpensesSub(1).Caption = "ÿ·» ’—› "
        ExpensesSub(2).Caption = "”‰œ«  «·’—›- Õ·Ì·Ì „’—Ê›«  "
        ExpensesSub(3).Caption = "”‰œ«  «·’—›- «·„œ›Ê⁄«  "
        ExpensesSub(4).Caption = "”‰œ ’—› „ ⁄œœ "
        
        '  Me.Payments(0).Caption = "«·„œ›Ê⁄« "

        Me.Cashing(0).Caption = "«·„ﬁ»Ê÷« "
        Me.Cashing(1).Caption = "”‰œ «·ﬁ»÷ «·’‰œÊﬁ «·⁄«„"
        
       BankOp.Caption = "«·„⁄«„·«  «·»‰ﬂÌ…"
        Me.BankOpsub(0).Caption = "«·«Ìœ«⁄«  «·»‰ﬂÌ…"
        Me.BankOpsub(1).Caption = " Õ’Ì·  Ê”œ«œ «·‘Ìﬂ« "
          Me.BankOpsub(2).Caption = " «· ”ÊÌ«  «·»‰ﬂÌ…  "
          Me.BankOpsub(3).Caption = "„–ﬂ—… »‰ﬂ  "
        Me.BankOpsub(4).Caption = "ÿ»«⁄Â «·‘Ìﬂ« "
        Me.BankOpsub(5).Caption = "«· ﬁ«—Ì—"
        
        
        CeramicEstimation.Caption = "«·„ﬁ«Ì”« "
        CeramicEstimationsub(0).Caption = "ÊÕœ«  «·⁄„·Ì« "
        CeramicEstimationsub(1).Caption = " ⁄—Ì› «·⁄„·Ì« "
        CeramicEstimationsub(2).Caption = "ÿ·» —›⁄ „ﬁ«”"
        CeramicEstimationsub(3).Caption = "Õ—ﬂ… «·ÿ·»« "
        CeramicEstimationsub(4).Caption = "«·« ›«ﬁÌ« "
        CeramicEstimationsub(5).Caption = "«·„‘«—Ì⁄"
        CeramicEstimationsub(6).Caption = " ”ÃÌ· «·«⁄„«· «·ÌÊ„Ì…"
        CeramicEstimationsub(7).Caption = " «·›Ê« Ì—"
        
        CeramicEstimationsub(8).Caption = "«· ﬁ«—Ì—"
        
        
        
        '*********************************************
StudentMenueSub(0).Caption = "«·»Ì«‰«  «·«”«”Ì…"
StudentMenueSub(1).Caption = "«·„œ—»Ì‰"
StudentMenueSub(2).Caption = "«·‘—ﬂ« "
StudentMenueSub(3).Caption = "ÿ·»  œ—Ì»"
StudentMenueSub(4).Caption = "«·ÿ·«»"
StudentMenueSub(5).Caption = "«·⁄ﬁÊœ"
StudentMenueSub(6).Caption = "«· —‘ÌÕ"
StudentMenueSub(7).Caption = "«·„Ê«›ﬁÂ ⁄·Ì «· —‘ÌÕ"
StudentMenueSub(8).Caption = "«·„Ã„Ê⁄« "
StudentMenueSub(9).Caption = "«·Õ÷Ê—"
StudentMenueSub(10).Caption = "«·« ’«·« "
StudentMenueSub(11).Caption = "«·›’·"
StudentMenueSub(12).Caption = " „œÌœ Ê«‰Â«¡ «·„Ã„Ê⁄« "
StudentMenueSub(13).Caption = "«‰Â«¡ ⁄ﬁÊœ «·‘—ﬂ« "
StudentMenueSub(14).Caption = "«’œ«— «·›Ê« Ì—"
StudentMenueSub(15).Caption = "«÷«›… ÊÕ–› Ê‰ﬁ· «·ÿ·«» »Ì‰ «·„Ã„Ê⁄« "
StudentMenueSub(16).Caption = "«· ﬁ«—Ì—"

'****************************************

        '   Me.Cashing(6).Caption = " ’›Ì… «·⁄Âœ "
        
        Me.MnuFinDiscounts.Caption = "«·Œ’Ê„«  «·„”„ÊÕ… Ê «·„ﬂ ”»…"
        Me.DelayVal(0).Caption = "«·«Ê—«ﬁ «·„«·ÌÂ «·„” Õﬁ…"
        
        Me.ReceiptPart.Caption = " Õ’Ì· Ê”œ«œ «·«ﬁ”«ÿ"
        Me.RequiredInstallment.Caption = "«·«ﬁ”«ÿ «·„ÿ·Ê»…"
        Me.MnuCheckBriefcase.Caption = "cheque Briefcase"
        '   Me.MnuCheckOperations.Caption = "‰Õ’Ì·  Ê”œ«œ «·‘Ìﬂ« "
        Me.MnuBoxDeposit(0).Caption = "«·«—’œ… «·«›  «ÕÌ…"
        Me.MnuBoxDeposit(1).Caption = " „ÊÌ· «·Œ“‰ Ê «” ⁄«÷… «·⁄Âœ"
        Me.MnuBoxDeposit(2).Caption = " ’›Ì… «·⁄Âœ…"
        
        Me.MnuBoxDrawing.Caption = " ÕÊÌ·«  „«·Ì…"
        Me.MnuBoxAccouns.Caption = "—’Ìœ «·Œ“‰ «·«‰"
        Me.MnuBoxIncapacity_Increase(0).Caption = "“Ì«œ… Ê⁄Ã“ ›Ì ‰ﬁœÌ… «·Œ“Ì‰…"
'        Me.MnuBoxIncapacity_Increase(1).Caption = "›« Ê—… Œœ„Ì…"
        
        'Me.MnuBoxStock.Caption = "Ã—œ «·Œ“Ì‰…"
        dev.Caption = "«·«œ«¡ Ê«·„Â«„"
        devsub(0).Caption = " ﬁ—Ì— ”Ì— «·⁄„· «·ÌÊ„Ì"
        devsub(1).Caption = "„—«Ã⁄Â Ê ﬁÌÌ„ ”Ì— «·⁄„· «·ÌÊ„Ì"
        devsub(2).Caption = " ⁄—Ì› «·„Â«„ Ê«·⁄„·Ì« "
        devsub(3).Caption = "„ «»⁄Â «·„Â«„ Ê«·⁄„·Ì« "
        devsub(4).Caption = " »ÌÂ«  «·„Â«„ Ê«·⁄„·Ì« "
        devsub(5).Caption = " ﬁ«—Ì— «·„Â«„ Ê«·⁄„·Ì« "
        
        Me.MnuAccounts.Caption = "«·Õ”«»«  «·⁄«„Â"
        Me.MnuAccCharts(0).Caption = "  œ·Ì· «·Õ”«»« "
        Me.MnuAccCharts(1).Caption = " «·ﬁÌœ «·«›  «ÕÌ  "

        Me.Reports.Caption = " «· ﬁ«—Ì—                                     "
        Me.Report.Caption = "«· ﬁ«—Ì— «·⁄«„…"
        Me.DailyReport.Caption = "«· ﬁ—Ì— «·ÌÊ„Ì"
        Me.MnuReports_Assblied.Caption = "«· ﬁ—Ì— «·„Ã„⁄ ⁄‰ › —…"
        Me.Tools.Caption = "„œÌ— «·‰Ÿ«„"
         
        Me.Barcode.Caption = " ’„Ì„ «·»«—ﬂÊœ..."
        Me.MnuPrintItemsCodes.Caption = "ÿ»«⁄Â «·»«—ﬂÊœ ..."
        'Me.MnuCorrectSerial.Caption = " ⁄œÌ· ”Ì—Ì·«  «·«’‰«›"
        'Me.MnuBoxDetectErrors.Caption = " ’ÕÌÕ «—’œ… «·Œ“‰"
        Me.MnuToolCustomers.Caption = " ⁄œÌ· ›Ê« Ì— «·⁄„·«¡"

        'Me.MnuToolRepaireItemsCost.Caption = " ⁄œÌ· «· ﬂ·›… ›Ì ›Ê« Ì— «·»Ì⁄"
        Me.MnuToolsDataBase(0).Caption = " ÕœÌÀ «·« ’«· »ﬁ«⁄œ… «·»Ì«‰« "
        Me.MnuToolsDataBase(1).Caption = " ÕœÌÀ «·‰Ÿ«„ "
        '        Me.MnuToolsDataBase(2).Caption = " €ÌÌ— ﬁ«⁄œ… «·»Ì«‰«  "
        Me.MnuDataBaseTools.Caption = "«œÊ«  ﬁ«⁄œ… «·»Ì«‰« "
        Me.UsersData.Caption = "«·„” Œœ„Ì‰"
        Me.AddUser.Caption = "»Ì«‰«  «·„” Œœ„Ì‰  ..."
'        Me.DelUser.Caption = "Õ–›  „” Œœ„  ..."
        Me.EditPw.Caption = " ⁄œÌ· «·—ﬁ„ ·”—Ì..."
        UserRpt.Caption = " ﬁ«—Ì— «·„” Œœ„Ì‰ "
            
            advanceMenu(0).Caption = "ÿ·» ”·›…"
             advanceMenu(1).Caption = " »Ì«‰«  «·”·› «·«›  «ÕÌ…"
              advanceMenu(2).Caption = " ⁄œÌ· /«Ìﬁ«› / —œ  ”·›…"
              
              
         
              
        Me.UserAbility.Caption = "’·«ÕÌ«  «·„” Œœ„Ì‰..."
        'Me.MnuUsersScreensPremission.Caption = "’·«ÕÌ«  «·„” Œœ„Ì‰ ⁄·Ï «·‘«‘« "
        Me.Options.Caption = "«⁄œ«œ«  «·‰Ÿ«„"
        Me.ShortCuts.Caption = "«·«Œ ’«—« "
         
        Me.MnuToolsSetPrinters0(0).Caption = "  «·œ⁄„ ›‰Ì"
         
         Me.MnuToolsSetPrinters0(1).Caption = "«⁄œ«œ «·ÿ«»⁄Â «·Õ«·Ì… ›Ì «·ÃÂ«“ «·Õ«·Ì..."
         
         
        Me.MnuToolsSetPrinters(1).Caption = " «⁄œ«œ«  œ·Ì· «·Õ”«»« "
        Me.MnuToolsSetPrinters(2).Caption = "«‰Ê«⁄ «·”‰œ« "
        Me.MnuToolsSetPrinters(3).Caption = "«·«ÿ·«⁄  ⁄·Ï  «· ‰»ÌÂ« "
         
        Me.MnuToolsSetPrinters(4).Caption = " ﬂÊÌœ «·”‰œ« "
        Me.MnuToolsSetPrinters(5).Caption = "  ﬂÊÌœ «·ÕﬁÊ·"
        Me.MnuToolsSetPrinters(6).Caption = "«·—”«∆· «·œ«Œ·Ì…"
        Me.MnuToolsSetPrinters7.Caption = "≈⁄œ«œ«  —”«∆· «·ÃÊ«· Ê «·«Ì„Ì·« "
         Me.MnuToolsSetPrinters(7).Caption = " «·ﬁ«„Ê”"
 
 
       
       
       
        Me.MnuInterface.Caption = "«·Ê«ÃÂ…   "
        Me.MnuInterfaceSub(0).Caption = "Ê«ÃÂÂ ⁄—»Ì…"
        Me.MnuInterfaceSub(1).Caption = "English Interface"
        'Me.MnuWindowsList.Caption = "«·‘«‘«  «·„› ÊÕÂ"
        'Me.MnuWindowsListOpen.Caption = "«·‘«‘«  «·„› ÊÕÂ"
        Me.Help.Caption = "„”«⁄œÂ"
        help_list(0).Caption = "  ⁄œÌ· «·ﬁ«∆„…"
        Me.HelpFileSub(0).Caption = "«·„Õ ÊÌ« ..."
       Me.HelpFileSub(1).Caption = "«·œ·Ì·..."
        Me.HelpFileSub(2).Caption = "«·»ÕÀ..."
        Me.HelpFileSub(3).Caption = "‰’«∆Õ..."
        Me.FavoritesMenue.Caption = "«·ﬁ«∆„… «·„›÷·…"
        Me.HelpFileSub(4).Caption = "⁄‰«..."
       Me.HelpFileSub(5).Caption = " ”ÃÌ·..."
 Me.HelpFileSub(6).Caption = "ﬁ«∆„… «·„Â«„"
 Me.HelpFileSub(7).Caption = "„‰ œÌ«  «·œ⁄„ «·› Ì..."
        prdo.Caption = "«·«‰ «Ã"


        prdo1(0).Caption = " «·»Ì«‰«  «·«”«”Ì…"
        prdo1sub(0).Caption = "»Ì«‰«  «·«·«  Ê «·„⁄œ« "
        prdo1sub(1).Caption = "⁄‰«’— «· ﬂ·Ì› «·’‰«⁄Ì… "
        prdo1sub(2).Caption = "«· ﬂ·›… «· ﬁœÌ—Ì… ÿ»ﬁ« ··ÊÕœ…"
        prdo1sub(3).Caption = "»Ì«‰«  «·ﬁÊ«·»"
        
         prdo1sub(4).Caption = "«‰Ê«⁄ «·«‰ «Ã "
         prdo1sub(5).Caption = "«· ﬂ«·Ì› «·’‰«⁄Ì… ÿ»ﬁ« ··«’‰«›"
          
        
        prdo1(4).Caption = " ŒÿÊÿ «·«‰ «Ã"
        prosub1(0).Caption = " ⁄—Ì› ŒÿÊÿ «·«‰ «Ã"
        prosub1(1).Caption = " Œ’Ì’  Ê‰ﬁ· «·⁄„«· »Ì‰ ŒÿÊÿ «·«‰ «Ã"

        prdo1(5).Caption = "„—«Õ· «·«‰ «Ã"

        prdo1(6).Caption = "”‰œ ÕÃ“ «‰ «Ã"
        prdo1(7).Caption = "«„— «·«‰ «Ã / «·‘€·"
        prdo1(8).Caption = "”‰œ ’—› „Ê«œ Œ«„ ··«‰ «Ã"
        prdo1(9).Caption = "”‰œ «” ·«„  «‰ «Ã  «„"

        prdo1(10).Caption = " ﬂ«·Ì› «·«‰ «Ã  «·‰„ÿÌ"
        prdo1(11).Caption = " Ê“Ì⁄ «· ﬂ«·Ì› €Ì— «·„»«‘—…"
       prdo1(12).Caption = " Œ’Ì’ ŒÿÊÿ «·«‰ «Ã ·√Ê«„— «·‘€·"
prdo1(13).Caption = "«÷«›«  «·—œÊœ Ê«„ «— «·„‘€·Ì‰"
        prdo1(14).Caption = "”‰œ«  «· Ã„Ì⁄"
         prdo1(15).Caption = " ﬁ«—Ì— «·«‰ «Ã"
 PrbH(0).Caption = " «„— «‰ «Ã ‰’› „’‰⁄"
        PrbH(1).Caption = " ”‰œ ’—› „—«Õ· «‰ «Ã"
        
        PrbH(2).Caption = " ”‰œ «” ·«„ «‰ «Ã ‰’› „’‰⁄"
 ScreenSetting.Caption = "«⁄œ«œ«  «·‘«‘« "
        MnuLevels(0).Caption = "«⁄ „«œ «·œÊ—… «·„” ‰œÌ…"
        MnuLevelsSub(0).Caption = " ⁄—Ì› „” ÊÌ«  «·«⁄ „«œ ··‘«‘« "
        MnuLevelsSub(1).Caption = "≈⁄œ«œ«  «⁄ „«œ «·‘«‘« "
        
          MnuLevels(1).Caption = "„Õœœ«  «·‘«‘« "
        MnuLevelsSub2(0).Caption = " ⁄—Ì›  „Õœœ«  «·‘«‘« "
        MnuLevelsSub2(1).Caption = "«⁄œ«œ „Õœœ«  «·‘«‘« "
        
        
        MNUFixedAssets.Caption = "«·«’Ê· «·À«» …"
        xxxxx(0).Caption = "„Ã„Ê⁄«  «·«’Ê· «·À«» …"
        xxxxx(1).Caption = "»Ì«‰«  «·«’Ê· «·À«» …"
        xxxxx(2).Caption = "›Ê« Ì— ‘—«¡ «·«’Ê· «·À«» …"
        xxxxx(3).Caption = "«ﬁ”«ÿ «·«Â·«ﬂ «·«’Ê· «·À«» …"
        xxxxx(4).Caption = "«· Œ·’ «Ê «” »⁄«œ«  «·«’Ê· "
        xxxxx(5).Caption = "«÷«›«  «·«’Ê· "
        xxxxx(6).Caption = "‰ﬁ· «·«’Ê· "
xxxxx(7).Caption = "Ã—œ «·«’Ê· "
xxxxx(8).Caption = " ﬁ«—Ì— "
        
        'ArrowsBase.Caption = " «·«”Â„"
        'ArrowsFollow(0).Caption = "»Ì«‰«  «·»Ê—’« "
        'ArrowsFollow(1).Caption = "»Ì«‰«  „Ã„Ê⁄«  «·«”Â„"
        'ArrowsFollow(2).Caption = "»Ì«‰«  «·‘—ﬂ« "
        'ArrowsFollow(3).Caption = " Õ„Ì· «·«”⁄«—"
        'ArrowsFollow(4).Caption = "  «·«”⁄«— «· «—ÌŒÌ…"
        'ArrowsFollow(5).Caption = "«·„Õ«›Ÿ"
'
'        ArrowsFollowBocket(0).Caption = " »Ì«‰«  «·„Õ«›Ÿ"
'        ArrowsFollowBocket(1).Caption = "‘—«¡ «·«”Â„"
'        ArrowsFollowBocket(2).Caption = "»Ì⁄ «·«”Â„"
'        ArrowsFollowBocket(3).Caption = "«·ﬁÌ„… «·«”„Ì… ··«”Â„"

'        ArrowsFollow(6).Caption = "„Ê«ﬁ⁄ Â«„…"
'        ArrowsFollow(7).Caption = " ﬁ«—Ì—"
'
        MnuMaintnance.Caption = "«·’Ì«‰…  "
        MnuMaintnanceBasic.Caption = "»Ì«‰«  «”«”Ì…"
        MnuMaintnanceBasicSub(0).Caption = "√‰Ê«⁄ «·’Ì«‰…"
        MnuMaintnanceBasicSub(1).Caption = "√‰Ê«⁄ «·„—ﬂ»« "
        MnuMaintnanceBasicSub(2).Caption = "ÿ—«“«  «·„—ﬂ»« "
        MnuMaintnanceBasicSub(3).Caption = "«·Ê«‰ «·„—ﬂ»« "
        MnuMaintnanceBasicSub(4).Caption = "»Ì«‰«  «·„—ﬂ»« "
        MnuMaintnanceBasicSub(5).Caption = "«ﬁ”«„ «·‘—ﬂ…"
        MnuMaintnanceBasicSub(6).Caption = "«·„‘—›Ì‰ Ê«·›‰ÌÌ‰"
        
        MnuMaintnanceBasicSub(7).Caption = "‘—ﬂ«  «·’Ì«‰…"
        MnuMaintnanceBasicSub(8).Caption = " ⁄—Ì› «·„’—Ê›« "
        
       ' MnuMaintnanceBasicSub1.Caption = "‘—ﬂ«  «·’Ì«‰…"
MnuMaintnanceTransactions(0).Caption = "Œÿ… «·’Ì«‰…"
MnuMaintnanceTransactions(1).Caption = "ÿ·»«  «·’Ì«‰…"

'MnuMaintnanceTransactionssub(0).Caption = "ÿ·» ’Ì«‰…"
'MnuMaintnanceTransactionssub(1).Caption = "√„— ‘€·"

        MnuMaintnanceTransactions(2).Caption = "√„— ‘€·"
        MnuMaintnanceTransactions(3).Caption = "ÿ·»«  ﬁÿ⁄ «·€Ì«— "
MnuMaintnanceTransactions(4).Caption = "”‰œ «” ·«„ ﬁÿ⁄ €Ì«— ··’Ì«‰…"
        MnuMaintnanceTransactions(5).Caption = "”‰œ ’—› ﬁÿ⁄ €Ì«— ··’Ì«‰…"

        'MnuMaintnanceTransactions(5).Caption = " ”·Ì„ «·’Ì«‰…"
        'MnuMaintnanceTransactions(6).Caption = "«· ÕÊÌ· „‰ Ê—‘… «·Ï Ê—‘… "
        'MnuMaintnanceTransactions(6).Caption = "—’Ìœ «›  «ÕÌ „Œ“‰ «·’Ì«‰…"
         MnuMaintnanceTransactions(8).Caption = " ”·Ì„ Ê ”·„ «·⁄Âœ «·⁄Ì‰Ì…"
         MnuMaintnanceTransactions(9).Caption = " ›ÊÌ÷ «·ﬁÌ«œ…"
         MnuMaintnanceTransactions(10).Caption = "«·÷„«‰"
         MnuMaintnanceTransactions(11).Caption = " ”ÃÌ· «·ÕÊ«œÀ"
         MnuMaintnanceTransactions(12).Caption = " ﬁ«—Ì— «·’Ì«‰…"
 
        tech.Caption = "√œÊ«  ›‰Ì…"
        MnuManToolsSub5.Caption = "„ «»⁄Â «·’Ì«‰…"
 
 shipmentMnu.Caption = " «·‘Õ‰ Ê «· Ê“Ì⁄"

ShpmentBasicdata(0).Caption = "«·»Ì«‰«  «·«”«”ÌÂ"
ShpmentBasicdata(1).Caption = "«·»÷«∆⁄ ﬁÌœ «· ”·Ì„"
ShpmentBasicdata(2).Caption = " Œ’Ì’  «·‘«Õ‰« "
ShpmentBasicdata(3).Caption = " ”ÃÌ·  ÊﬁÌ «  «· ”·Ì„ "
ShpmentBasicdata(4).Caption = "ŒÿÂ «·‘Õ‰"
ShpmentBasicdata(5).Caption = "ÿ·» ‘Õ‰"
ShpmentBasicdata(6).Caption = "«–‰ «·‘Õ‰ / «· ”·Ì„"
ShpmentBasicdata(7).Caption = "”‰œ «” ·«„  ‘Õ‰"
ShpmentBasicdata(8).Caption = "«· ﬁ«—Ì—"

ShpmentBasicdatasub(0).Caption = "»Ì«‰«  «·œÊ·"
ShpmentBasicdatasub(1).Caption = "»Ì«‰«  «·„Õ«›Ÿ«  Ê «·„‰«ÿﬁ"
ShpmentBasicdatasub(2).Caption = "«·„”«›«  »Ì‰ «·„œ‰"
ShpmentBasicdatasub(3).Caption = "»Ì«‰«  «·«ÕÌ«¡"
ShpmentBasicdatasub(4).Caption = "»Ì«‰«  «·‘Ê«—⁄"
ShpmentBasicdatasub(5).Caption = "«‰Ê«⁄ «·„—ﬂ»« "
ShpmentBasicdatasub(6).Caption = "»Ì«‰«  «·„—ﬂ»« "
ShpmentBasicdatasub(7).Caption = "»Ì«‰«  «·”«∆ﬁÌ‰"
 ShpmentBasicdatasub(8).Caption = "«‰Ê«⁄ «·‘Õ‰"
 ShpmentBasicdatasub(9).Caption = "«‰Ê«⁄ «·’Ì«‰…"
 



    ElseIf SystemOptions.UserInterface = EnglishInterface Then
                    MnuToolsSetPrinters0sub(0).Caption = "Technical Request"
        MnuToolsSetPrinters0sub(1).Caption = "Camera Follow"
       MnuToolsSetPrinters0sub(2).Caption = "Technical Reques 2 "
       MnuToolsSetPrinters0sub(3).Caption = "Close/open System"
       
             MnuToolsSetPrinters0sub(4).Caption = "Vending Machine"
       MnuToolsSetPrinters0sub(5).Caption = "Contracting"
       MnuToolsSetPrinters0sub(6).Caption = "Visit"
       MnuToolsSetPrinters0sub(7).Caption = "Implementaions"
     MnuToolsSetPrinters0sub(8).Caption = "Cost Manipulation"
MnuToolsSetPrinters0sub(9).Caption = "Re calcualte Issue Vchr"
MnuToolsSetPrinters0sub(10).Caption = "Team Viewer"


       UsersGroup.Caption = "Users Group"
       
       
    CarMaintenance.Caption = "Car Maintenance"
    CarMaintenancesub(0).Caption = "Basic Data"
     CarMaintenancesub(1).Caption = "Transactions"
     
      CarMaintenancesub1(0).Caption = "Vehicle Type"
CarMaintenancesub1(1).Caption = "Vehicle Style"
CarMaintenancesub1(2).Caption = "Vehicle Data"
CarMaintenancesub1(3).Caption = "Types of reforms"
CarMaintenancesub1(4).Caption = "Purchase Types"
CarMaintenancesub1(5).Caption = "Define Purchase"

        '*********************************************




   AgeingSub(0).Caption = "Purchase Aging Settings"
   AgeingSub(1).Caption = "Sales Aging Settings"
   AgeingSub(2).Caption = "Register old purchase invoice"
   AgeingSub(3).Caption = "Register old Sales invoice"
   AgeingSub(4).Caption = "Link Current Sales Invoice"
   AgeingSub(5).Caption = "Reports"







 
StudentMenueSub(0).Caption = "Basic Data"
StudentMenueSub(1).Caption = "Instructor"
StudentMenueSub(2).Caption = "Companies"
StudentMenueSub(3).Caption = "Training Request"
StudentMenueSub(4).Caption = "Students"
StudentMenueSub(5).Caption = "Contract"
StudentMenueSub(6).Caption = "nomination"
StudentMenueSub(7).Caption = "«nomination Approval"
StudentMenueSub(8).Caption = "Groups"
StudentMenueSub(9).Caption = "Attendance"
StudentMenueSub(10).Caption = "Calling"
StudentMenueSub(11).Caption = "Termination"
StudentMenueSub(12).Caption = "End Groups"
StudentMenueSub(13).Caption = "End Contract"
StudentMenueSub(14).Caption = "Bill Vouchers"
StudentMenueSub(15).Caption = "Groups Add/Delete Students"
StudentMenueSub(16).Caption = "Reports"



        '****************************************
        
MnuElevatorssUB(0).Caption = "Define Criteria"
MnuElevatorssUB(1).Caption = "Link Criteria"
MnuElevatorssUB(2).Caption = "Special Quotations"
MnuElevatorssUB(3).Caption = "Technical Quotations"
MnuElevatorssUB(4).Caption = "Maintenance Contracts And Warranty"
Elevatorsmaintenance(0).Caption = "Warranty and  Maintenance Contracts "
Elevatorsmaintenance(1).Caption = "Spare Parts Issue Voucher"
Elevatorsmaintenance(2).Caption = "Preventive maintenance"

Elevatorsmaintenance(3).Caption = "Maintenance Contracts Alarms"
Elevatorsmaintenance(4).Caption = "Warranty and  gurantee Alarms"
Elevatorsmaintenance(5).Caption = "Reports"

MnuElevatorssUB(6).Caption = "Reports"

'***********************************************

CarMaintenancesub1(6).Caption = "Computer Checks Codes"
CarMaintenancesub1(7).Caption = "Vehicle Colors"
CarMaintenancesub1(8).Caption = "Stores"
CarMaintenancesub1(9).Caption = "Items Groups"
CarMaintenancesub1(10).Caption = "Items Units"
CarMaintenancesub1(11).Caption = "Items Data"
CarMaintenancesub1(12).Caption = "Customers"
CarMaintenancesub1(13).Caption = "Employees Data"
CarMaintenancesub1(17).Caption = "Departements Data"
CarMaintenancesub1(18).Caption = "Supervisors And Technicals"
 
CarMaintenancesub2(0).Caption = "Maintenance Entry"
CarMaintenancesub2(1).Caption = "Computer Check"
CarMaintenancesub2(2).Caption = "Issue Voucher"
CarMaintenancesub2(3).Caption = "Purchase Orders"
CarMaintenancesub2(4).Caption = "Maintenance Invoice"
CarMaintenancesub2(5).Caption = "Commissions"
CarMaintenancesub2(6).Caption = "Hand Cost"

CarMaintenancesub(2).Caption = "Reports"
 
          Strategy.Caption = "School transport"
'    GoldMenu.Caption = "Gold and Diamond "
    dev.Caption = "Development"
    
    
      POSTRansactiosG.Caption = "POS"
 
 HRProcedures(0).Caption = "Advance Request"
 HRProcedures(1).Caption = "Temparary vacation-authorized exit"
 HRProcedures(2).Caption = "Task Request"
 HRProcedures(3).Caption = "Housing Allowance Request "
 HRProcedures(4).Caption = "Employee Transfer "
 HRProcedures(5).Caption = "Direct Employee Request"
 HRProcedures(7).Caption = "Employee Questionnaire "
 HRProcedures(8).Caption = "Vacation Request"
 HRProcedures(9).Caption = "Vacation Date"
 HRProcedures(10).Caption = "Assets Delivery"
 HRProcedures(11).Caption = "Passport submit "
 HRProcedures(12).Caption = " Employee Warning of Penality "
 HRProcedures(13).Caption = "To Whom It May Concern Letter"
 HRProcedures(14).Caption = "Work injury Report"
 HRProcedures(15).Caption = "Receiving Transactions"
 HRProcedures(16).Caption = "Final Clearance "
 HRProcedures(25).Caption = "Employee Update info"
 HRProcedures(26).Caption = "Employee Disclaimer Request "
 HRProcedures(27).Caption = "Administrative Feedback"
 HRProcedures(28).Caption = "Deduction Note"
 HRProcedures(30).Caption = "HR Letter"
HRProcedures(31).Caption = "Driving Letter"

 

POSTRansactios(0).Caption = "POS Data"
POSTRansactios(1).Caption = "Cashier Data"

POSTRansactios(2).Caption = "Shift Data"
POSTRansactios(3).Caption = "Locations Data"
POSTRansactios(4).Caption = "Points"

POSTRansactios(5).Caption = "Login"
POSTRansactios(6).Caption = "General Issue Voucher"
POSTRansactios(7).Caption = "Pos Geneal Collect"
POSTRansactios(8).Caption = "Reports"
POSTRansactios(9).Caption = "Customer Card"
POSTRansactios(10).Caption = "Coupons"

Texh(0).Caption = "Technical Settings"
Texh(1).Caption = "Messages Form"
Texh(2).Caption = "Define SMS For Screens"
Texh(3).Caption = "Customers SMS"

     
 shipmentMnu.Caption = "Shipping and Distribution"

ShpmentBasicdata(0).Caption = "Basic Data"
ShpmentBasicdata(1).Caption = "Non-delivered goods"
ShpmentBasicdata(2).Caption = "Allocation of vehicles"
ShpmentBasicdata(3).Caption = "Recording  delivery timing    "
ShpmentBasicdata(4).Caption = "Shipping Plan"


ShpmentBasicdata(5).Caption = "Shipping Order"
ShpmentBasicdata(6).Caption = "Shipping Voucher"
ShpmentBasicdata(7).Caption = "Shipping Recived Voucher"
ShpmentBasicdata(8).Caption = "Shipping Recived Voucher"


SalesInsSub(0).Caption = "Request for purchase by installments"
    SalesInsSub(1).Caption = "Request to open a client account"
    SalesInsSub(2).Caption = "Customers"
    SalesInsSub(3).Caption = "Installment bill"
    SalesInsSub(4).Caption = "Collection of premiums"
    SalesInsSub(5).Caption = "Alerts"
    
    
   SalesInsSub(6).Caption = "Reports"

ShpmentBasicdatasub(0).Caption = "Country data"
ShpmentBasicdatasub(1).Caption = "Cities Data"
ShpmentBasicdatasub(2).Caption = "Distance between Cities"
ShpmentBasicdatasub(3).Caption = "Neighborhoods Data "
ShpmentBasicdatasub(4).Caption = "Streets Data"
ShpmentBasicdatasub(5).Caption = "Vehicles Types"
ShpmentBasicdatasub(6).Caption = "Vehicles Data"
ShpmentBasicdatasub(7).Caption = "Drivers"
ShpmentBasicdatasub(8).Caption = "Shipment Types"
ShpmentBasicdatasub(9).Caption = "Maintenance  Types"

     MarketingMnu.Caption = "Marketing"
     mangDep.Caption = "HR Mangement"
mangDepSub(0).Caption = "Indicators"
mangDepSub(1).Caption = "Application Form"
mangDepSub(2).Caption = "Jobs Requirments"


 MarketingMnusub(0).Caption = "Basic Data"
MarketingMnusub(1).Caption = "Items Overs"
MarketingMnusub(2).Caption = "Customers Follow"
MarketingMnusub(3).Caption = "Reports"
 MarketingMnusub(4).Caption = "Call Reports"
 

MdiContextMenu.Caption = "Programs List "

MarketingMnusubsub(0).Caption = "Visits Allocations"
MarketingMnusubsub(1).Caption = "Register customer visits"
MarketingMnusubsub(2).Caption = "Follow customer visits"
MarketingMnusubsub(3).Caption = "Poll customers"
MarketingMnusubsub(4).Caption = "Customer complaint registration"
MarketingMnusubsub(5).Caption = "Customer complaint Follow"
 


        MnuManToolsSub5.Caption = "Maintenance Follow"

        MnuMaintnance.Caption = "Maintenence"
        MnuMaintnanceBasic.Caption = "Basic Data"
        MnuMaintnanceBasicSub(0).Caption = "Maintenence Types"
  
     
        MnuMaintnanceBasicSub(1).Caption = "Vechile Type"
        MnuMaintnanceBasicSub(2).Caption = "Vechile Model"
        MnuMaintnanceBasicSub(3).Caption = "Vechile Colors"
        MnuMaintnanceBasicSub(4).Caption = "Vechile Data"
      MnuMaintnanceBasicSub(5).Caption = "Maintenence Departement"
      MnuMaintnanceBasicSub(6).Caption = "Maintenence Supervisors and Technicals"
      
        MnuMaintnanceBasicSub(7).Caption = "Maintenence Companies"
               MnuMaintnanceBasicSub(8).Caption = "Define Expenses"
 
      
       
    '
    '    MnuMaintnanceBasicSub1.Caption = "Maintenence Companies"
'

MnuMaintnanceTransactions(0).Caption = "Maintenance Plan"
       MnuMaintnanceTransactions(1).Caption = "Internal Maintenance Req."
'MnuMaintnanceTransactionssub(0).Caption = " Maintenance Request"
'MnuMaintnanceTransactionssub(1).Caption = "Work Order"


        MnuMaintnanceTransactions(2).Caption = "Work Order"
        MnuMaintnanceTransactions(3).Caption = "Internal Request"
MnuMaintnanceTransactions(4).Caption = "Spare part Recieve Voucher"
        MnuMaintnanceTransactions(5).Caption = "Spare part Issue Voucher"

       ' MnuMaintnanceTransactions(4).Caption = "Maintenance Delivery"
       ' MnuMaintnanceTransactions(5).Caption = "Back Guarantee From The Supplier"
       ' MnuMaintnanceTransactions(6).Caption = "Opening Balance For Maintenance Store"
         MnuMaintnanceTransactions(8).Caption = "F.A moving"
        MnuMaintnanceTransactions(9).Caption = "Drive License "
        MnuMaintnanceTransactions(10).Caption = "Gurantee "
       
       MnuMaintnanceTransactions(11).Caption = "Accedient Reports"
        MnuMaintnanceTransactions(12).Caption = "Maintenance Reports"
        
        tech.Caption = "Technical Tools"

        Me.Basicdata.Caption = "Basic Data"
        Me.BasicDataM(0).Caption = " Expenses Types"
        Me.BasicDataM(1).Caption = " Revenues Types"
        Me.BasicDataM(2).Caption = " Banks Data"
        Me.BasicDataM(3).Caption = " Cash On Hand Data"
        Me.BasicDataM(4).Caption = " Payment  Type"
        Me.BasicDataM(5).Caption = " Supplier Data"
        Me.BasicDataM(6).Caption = " Customer Data"

        Me.BasicDataM(7).Caption = " Employee Data"
        Me.BasicDataM(8).Caption = " Items Data"
        
       Me.BasicDataM(9).Caption = " Currency Data"
        Me.BasicDataM(10).Caption = "Countries\ Nationality Data"
         
        Me.BasicDataM(11).Caption = " Religons Data"
        Me.BasicDataM(12).Caption = " Countries Data"
        Me.BasicDataM(13).Caption = " Government Data"
        Me.BasicDataM(14).Caption = " Neighborhoods Data"
        Me.BasicDataM(15).Caption = " Street Data"
        Me.BasicDataM(16).Caption = "Projects"
        Me.BasicDataM(17).Caption = "Reports"
        
        'Me.BasicDataM(15).Caption = " Items Data"
        'Me.BasicDataM(16).Caption = "Employee Date"
        Me.BasicDataM(20).Caption = "  Exit"
        FinAnalysis.Caption = "Fin. Analysis"
        AssetsMngBase.Caption = "RealState Mangement"
         mnuEmployee.Caption = "Personal And Payroll"
 
        MnuItemTools_ItemCart.Caption = "Item Card"
        MnuItemTools_ItemCostTrans.Caption = "Item Cost Price"
        MnuItemTools_ItemData.Caption = "Items Data"
        MnuItemTools_ItemQty.Caption = "Items Search Qty"
        MnuItemTools_Sep.Caption = "Alternatives"
        
        MnuItemTools_ItemSerial.Caption = "Items Serials"

        MnuAccDEV(0).Caption = " J L Viewer"
              MnuAccDEV(1).Caption = " J L Manual Entry"
        MnuAccDEV_Post.Caption = "Auditing   J LEntry"
        xxx(0).Caption = "Cost Centers Type"
        xxx(1).Caption = "Cost Centers"
        ProductionPlansub(0).Caption = "Production Plan"
        ProductionPlansub(1).Caption = "Defining QC Items"
        ProductionPlansub(2).Caption = "Production Classification "

        ProductionPlansub(3).Caption = "Corrective Action"
        ProductionPlansub(4).Caption = "Batch Sheet"
        ProductionPlansub(5).Caption = "Test Certificate"
        ProductionPlansub(6).Caption = "Quality"
        ProductionPlansub(7).Caption = "Process Report"

        xxy(0).Caption = "Budget"
        ProductionPlan.Caption = "Planning and Quality Control"
        'xxx(4).Caption = "Financial Analysis"
        xxy(1).Caption = "Cash Flow"
        xxy(3).Caption = "Accounts Distribution"
        'xxx(7).Caption = "Prepare BalanceSheet"
        xxy(2).Caption = "Tab list Sheet"
        xxy(4).Caption = "perpare  Fin Equations"
        xxy(5).Caption = "View Fin Equations"

        xxy(6).Caption = "Composite Accounts"
        xxy(7).Caption = "Acc. Asurance"
        xxy(8).Caption = "Agenda customers"
      xxy(9).Caption = "Load Trial Balance"
      xxy(10).Caption = "Adnanced Expenses"
      xxy(11).Caption = "Plans "
advancedPayment(0).Caption = "Adnanced Expenses Types"
advancedPayment(1).Caption = "Adnanced Expenses Vchr"
advancedPayment(2).Caption = "Adnanced Expenses Allocations"
advancedPayment(3).Caption = "Adnanced Allowance  Vchr"


        taxes.Caption = "«VAT"
        TaxexSub(0).Caption = "VAT Settings"
TaxexSub(1).Caption = "Purchase Transactions Entery"
TaxexSub(2).Caption = "Sales Transactions Entery"
TaxexSub(3).Caption = "Return Purchase Transactions Entery"
TaxexSub(4).Caption = "Return Sales Transactions Entery"
TaxexSub(5).Caption = "F.A Purchase Transactions Entery"
TaxexSub(6).Caption = "F.A   Sales Transactions Entery"
TaxexSub(7).Caption = "Notes"
TaxexSub(8).Caption = "VAT Forms"
TaxexSub(9).Caption = "Reports"

TaxexSub(10).Caption = "Create Voucher For POS VAT"
  '*******************************
  Tailorsub(0).Caption = "Tasks"
  Tailorsub(1).Caption = "Sizes"
  Tailorsub(3).Caption = "Employee"
  Tailorsub(4).Caption = "Customer"
  Tailorsub(5).Caption = "Job Order"
  Tailorsub(6).Caption = "Sales Invoice"
  Tailorsub(7).Caption = "Cash Voucher"
  Tailorsub(8).Caption = "Employee Production"
 Tailorsub(9).Caption = "reports"
Tailor.Caption = "Tailor"
 Beauty.Caption = " Beauty"
     Beautysub(0).Caption = "Services Group"
     Beautysub(1).Caption = "Services"
     Beautysub(2).Caption = "Labors"
     Beautysub(3).Caption = "Shifts"
     Beautysub(4).Caption = "Rest Types"
     Beautysub(5).Caption = "Link Employee To Services"
     Beautysub(6).Caption = "Customers"
     Beautysub(7).Caption = "Booking  Type"
     Beautysub(8).Caption = "Rest Plan"
     Beautysub(9).Caption = "Booking appointments "
     Beautysub(10).Caption = "View reservations"

     Beautysub(11).Caption = "POS"
     Beautysub(12).Caption = "Cashier"
     Beautysub(13).Caption = "POS Login"
     Beautysub(14).Caption = "Catch a deposit"
     Beautysub(15).Caption = "Public receipt Voucher"
     Beautysub(16).Caption = "Reports"
     Beautysub(17).Caption = "CustomersReports"
        xxx(12).Caption = "Accounts Reports"

        Me.MnuProjects.Caption = "Projects Mangment"
        Me.MnuProjectsBasic.Caption = "Basic Data"
        Me.MnuProjectsBasicSub(0).Caption = "Projects Status"
        Me.MnuProjectsBasicSub(1).Caption = "Contract Type"

        Me.MnuProjectsBasicSub(2).Caption = "Sub-contractor  Data"
        
         Me.MnuProjectsBasicSub(3).Caption = "Projects Items"
         
        Me.MnuProjectsBasicSub(4).Caption = "Processes Unit"
        Me.MnuProjectsBasicSub(5).Caption = "Processes"
        Me.MnuProjectsBasicSub(6).Caption = "Equipmemts Data  "
        Me.MnuProjectsTransactions(0).Caption = "Projects Data"
        Me.MnuProjectsTransactions(1).Caption = "Projects Issue Vouchers"
         Me.MnuProjectsTransactions(2).Caption = "Projects Return Vouchers"
         
        Me.MnuProjectsTransactions(3).Caption = "Projects Labors Allocate"
        Me.MnuProjectsTransactions(4).Caption = "Projects Labors Transfer"
        
       Me.MnuProjectsTransactions(5).Caption = "Projects Equipments Allocation"
        Me.MnuProjectsTransactions(6).Caption = "Projects Equipments Transfer"
        
        
        
        Me.MnuProjectsTransactions(7).Caption = "Follow Up Processes "
        Me.MnuProjectsTransactions(8).Caption = "Projects Invoice"
        Me.MnuProjectsTransactions(9).Caption = "Monthly Projects Invoice"
 Me.MnuProjectsTransactions(10).Caption = "Projects Reports"
        mnuEmployeeBasic(0).Caption = "Basic Data"
        mnuEmployeeBasicSub(0).Caption = "Prepare Company Attendance Times"
        mnuEmployeeBasicSub(1).Caption = "Shifts"
        mnuEmployeeBasicSub(2).Caption = "Vacations Types"
        mnuEmployeeBasict(0).Caption = "Evaluation Settings"
        mnuEmployeeBasict(1).Caption = "Evaluation"
        mnuEmployeeBasicSub(3).Caption = "Contract Type"
        mnuEmployeeBasicSub(4).Caption = "Job Status"
        mnuEmployeeBasicSub(5).Caption = "Departrment\Sections Data"
        mnuEmployeeBasicSub(6).Caption = "Job Types Data"
        mnuEmployeeBasicSub(7).Caption = "Team Data"
       mnuEmployeeBasicSub(8).Caption = "Employees Grades"
       mnuEmployeInsuranceSub(0).Caption = "Insurance Settings"
        mnuEmployeInsuranceSub(1).Caption = "Insurance Companies"
        mnuEmployeInsuranceSub(2).Caption = "Insurance  Types"
        mnuEmployeInsuranceSub(3).Caption = "Insurance  Classe"
        mnuEmployeInsuranceSub(4).Caption = "GOSI Calc"
        
        mnuEmployeeBasicSub(12).Caption = "Elements of assessment"
                
                mnuEmployeeBasicSub(13).Caption = "Types of requests out"
                mnuEmployeeBasicSub(14).Caption = "Job Locations"
                mnuEmployeeBasicSub(15).Caption = "Nationality"
                mnuEmployeeBasicSub(16).Caption = "Religions"
                mnuEmployeeBasicSub(17).Caption = "Definition of Assets"
                mnuEmployeeBasicSub(18).Caption = "Definition of Relations"
                 mnuEmployeeBasicSub(19).Caption = "Regions and sectors"
                 mnuEmployeeBasicSub(20).Caption = "Visas Data"
                 mnuEmployeeBasicSub(21).Caption = "Punch Basic Data"
                 mnuEmployeeBasicSub(22).Caption = "Sick Settings"
                 mnuEmployeeBasicSub(23).Caption = "Vacations Policy"
                 
                mnuEmployeeBasic(2).Caption = "GOSI and GOMI "
           mnuEmployeeBasic(3).Caption = "KPI"
                
           mnuEmployeeBasict(0).Caption = "KPI Def" 'key performance Indicator
         mnuEmployeeBasict(1).Caption = "KPI Manual Evaluation"
      mnuEmployeeBasict(2).Caption = "KPI Results"
         
        mnuEmployeeBasic(4).Caption = "Atendance"
        EmployeeAttendanceSub(0).Caption = "Vacation Type"
     '   EmployeeAttendanceSub(0).Caption = "Prepare Employee Attendance Times"
        EmployeeAttendanceSub(1).Caption = " Shifts  Settings"
         EmployeeAttendanceSub(2).Caption = " Calender  Settings"
          EmployeeAttendanceSub(3).Caption = " Manual Attendance "
           EmployeeAttendanceSub(4).Caption = " Import  Shifts  Attendance"
        EmployeeAttendanceSub(5).Caption = "Attendance  Approve"
         
      '  EmployeeAttendanceSub(4).Caption = "View Attendance Times"
        mnuEmployeeBasic(5).Caption = "Procedures Form"
        mnuEmployeeBasic(6).Caption = "Salaries"
        EmployeeSalarySub(0).Caption = "Salary Components Types"
        EmployeeSalarySub(1).Caption = "Salary Components Equations"
        EmployeeSalarySub(2).Caption = "Bonus"
 
        EmployeeSalarySub(3).Caption = "Punishments"
        EmployeeSalarySub(4).Caption = "Record Advances to staff"
        EmployeeSalarySub(5).Caption = "Return Advances to staff"""
        EmployeeSalarySub(6).Caption = "Payroll"
        EmployeeSalarySub(7).Caption = "Calcualte End of service"
        EmployeeSalarySub(8).Caption = "Register Changed Components"
        EmployeeSalarySub(9).Caption = "Register  Employee Allocations  "
        EmployeeSalarySub(10).Caption = "Register  Employee  Annual Components "
        EmployeeSalarySub(11).Caption = "Salary Increases "
        EmployeeSalarySub(12).Caption = "Change Advance Due Date "

        mnuEmployeeBasic(1).Caption = "Employees Data"
        EmployeeDataicSub(0).Caption = "Employees Files"
        EmployeeDataicSub(1).Caption = "Employees Contracts"

        mnuEmployeeBasic(7).Caption = "Employees vacations"
Vscstionsssub(0).Caption = "Opening Data "
Vscstionsssub(1).Caption = "Opening Vacations "

        Vscstionsssub(2).Caption = "Vacations Plan"
        Vscstionsssub(3).Caption = "Vacations Request"
        Vscstionsssub(4).Caption = "Vacations Data"
        Vscstionsssub(5).Caption = "Assets Transfer"
        Vscstionsssub(6).Caption = "Vacations Dues"
        Vscstionsssub(7).Caption = "Exit And ReturnVisa"
        Vscstionsssub(8).Caption = "Start Work"
       Vscstionsssub(9).Caption = "Sick vacation"
mnuEmployeeBasic(8).Caption = "Advanced"
        mnuEmployeeBasic(9).Caption = "Termination"
mnuEmployeeBasic(10).Caption = "ˆAdvanced Allowance Plan"
        mnuEmployeeBasic(11).Caption = "Reports"
        
        FinishSevicersub(0).Caption = "Service Termination Request"
        FinishSevicersub(1).Caption = "Service Termination"

        TransporterMain.Caption = "Trasportation"
        TransporterSub(0).Caption = "Cities Data"
        TransporterSub(1).Caption = "Distance Cities Cities"
        
         TransporterSub(2).Caption = "Port Data"
          TransporterSub(3).Caption = "Ship Date"
          TransporterSub(4).Caption = "Transport Type"
          TransporterSub(5).Caption = "Trip Type "
          
        TransporterSub(6).Caption = "Customer Data "
        TransporterSub(7).Caption = "Supplier Data"
        TransporterSub(8).Caption = "Driver Data"
        TransporterSub(9).Caption = "Vehicles Types"
        TransporterSub(10).Caption = "Vehicles Model"
        
        TransporterSub(11).Caption = "Insurance Company"
        TransporterSub(12).Caption = "Regular Maintenance Type"
        TransporterSub(13).Caption = "Vehicles Data"
        TransporterSub(14).Caption = "Maintenance Plan"
        TransporterSub(15).Caption = "Customer Contract"
        TransporterSub(16).Caption = "Carry Orders"
        TransporterSub(17).Caption = "Trip Data"
        TransporterSub(18).Caption = "Customers Invoices"
       TransporterSub(19).Caption = "Dribver era"
        TransporterSub(20).Caption = "Reports"

        Me.StockControl.Caption = "StockControl"
        Me.StockControlBasic.Caption = "Basic Data"
        StockControlBasicSub(0).Caption = "Items Data"

        StockControlBasicSub(1).Caption = "Store Data"
        StockControlBasicSub(2).Caption = "Items Groups"
        StockControlBasicSub(3).Caption = "Units"
        StockControlBasicSub(4).Caption = "Items Colors"
        StockControlBasicSub(5).Caption = "Items Sizes"
        StockControlBasicSub(6).Caption = "Items Classes"
        StockControlBasicSub(7).Caption = "Define Stores Locations"
        StockControlBasicSub(8).Caption = "Items Specifications"

        'StockControlBasicSub(9).Caption = "Production Cost component   "
        'StockControlBasicSub(10).Caption = "Unit  Cost Of Production"
        StockControlBasicSub(11).Caption = "Plan For Items Sales "
        StockControlBasicSub(12).Caption = "Linking items With stores "
        StockControlBasicSub(13).Caption = "Re-Order Limit Settings"

        Me.TradingTransaction(0).Caption = "Stock Opening Balances"
        Me.TradingTransaction(1).Caption = "Internal Orders"
        XC(0).Caption = "Internal Order"
        XC(1).Caption = "reservation Voucher "
        Me.TradingTransaction(2).Caption = "Recieve  Vouchers"
        Me.TradingTransaction(3).Caption = "Issue  Vouchers"
        Me.TradingTransaction(4).Caption = "Transfer Items Between Stores"
        Me.TradingTransaction(5).Caption = "Stock Count"
        TradingTransactionSub(0).Caption = "Start Inventory"
        TradingTransactionSub(1).Caption = "Print Inventory Report"
        TradingTransactionSub(2).Caption = "˝Actual Inventory"
        TradingTransactionSub(3).Caption = "Stock Settlement Auto "

        Me.TradingTransaction(6).Caption = "Stock Settlement"
        Me.TradingTransaction(7).Caption = "MiX Voucher"
        Me.TradingTransaction(8).Caption = "tems Qty Query"
        Me.TradingTransaction(9).Caption = "Items Serial Search"
        Me.TradingTransaction(10).Caption = "On Demand Items"
        Me.TradingTransaction(11).Caption = "Items Current Status"
        Me.TradingTransaction(12).Caption = "Reports"

TradingTransactionSub1(0).Caption = "Issue  Request  "
        TradingTransactionSub1(1).Caption = "Issue  Vouchers  "
        TradingTransactionSub1(2).Caption = "Damage and Sample Issue  Vouchers"

        Me.Purchase.Caption = "Purchase "
        Me.PurchaseBasicRoot.Caption = "Basic Data"
        Me.PurchaseBasic(0).Caption = "Supplier Data"
        Me.PurchaseBasic(1).Caption = "Supplier Contract"
        Me.PurchaseBasic(2).Caption = "Prepare Ageing Data"
        Me.PurchaseBasic(3).Caption = "Shipment type"
        Me.PurchaseBasic(4).Caption = "Gurantee Type"
        Me.PurchaseBasic(5).Caption = "Payment Method"
 
 
Me.PurchaseBasic(6).Caption = "Purchae Pesron Groups"
Me.PurchaseBasic(7).Caption = "Purchae Pesron Data "
Me.PurchaseBasic(8).Caption = "Shipment Method"

        Me.PurchaseTransactions(0).Caption = "Quotations and Purchase Orders"
 
        PurchaseTransactionssubd(0).Caption = "Quotations"
        PurchaseTransactionssubs(0).Caption = "'Quotations Request"
        PurchaseTransactionssubs(1).Caption = "Quotations"
        PurchaseTransactionssubs(2).Caption = "Quotations Comparison Sheet"

        PurchaseTransactionssubd(1).Caption = "Purchase Orders"
        PurchaseTransactionssubs1(0).Caption = "Purchase Request  "
        PurchaseTransactionssubs1(1).Caption = "Purchase Order Approval"
        PurchaseTransactionssubs1(2).Caption = " Purchase Order"

        Me.PurchaseTransactions(1).Caption = "Shipment Data"
        Me.PurchaseTransactions(2).Caption = "LC"

        LCTransactions(0).Caption = "Types of LC"
        LCTransactions(1).Caption = "Performa Invoices"
        LCTransactions(2).Caption = "Open LC"
        LCTransactions(3).Caption = "Edit LC"
        LCTransactions(4).Caption = "Shipments Follow"
        LCTransactions(5).Caption = "Shipment Recieve Voucher"
        LCTransactions(6).Caption = "Final Invoice"
        LCTransactions(7).Caption = "Close LC"

LCTransactions(8).Caption = "Request Bank Guarantee"
LCTransactions(9).Caption = "Request for extension of Bank Guarantee"
LCTransactions(10).Caption = "Final Bank Guarantee"

LCTransactions(11).Caption = "Purchase Form"
        Me.PurchaseTransactions(3).Caption = "Purchase Invoices"
 Me.PurchaseTransactions(4).Caption = "Composite Purchase Invoices"
 
        Me.PurchaseTransactions(5).Caption = "Return Purchase"
        Me.PurchaseTransactions(6).Caption = "Ageing Report"
        Me.PurchaseTransactions(7).Caption = "Purchase Reports"
 
        Me.Sales.Caption = "Sales "
 
        Me.SalesBasic.Caption = "Basic Data"
        Me.SalesBasicSub(0).Caption = "Customers Type"
        Me.SalesBasicSub(1).Caption = "Customers Data"
        Me.SalesBasicSub(2).Caption = "Cusettomers Contract"
        Me.SalesBasicSub(3).Caption = "Perpare Ageing "
        Me.SalesBasicSub(4).Caption = "Define Sales Price "
        Me.SalesBasicSub(5).Caption = "Items stagnant"
        Me.SalesBasicSub(6).Caption = "Prepare Sales Target"
        Me.SalesBasicSub(7).Caption = "Sales Rep Groups"
        Me.SalesBasicSub(8).Caption = "Sales Rep Data"
   Me.SalesBasicSub(9).Caption = "Installments Gurantee Type "
   Me.SalesBasicSub(10).Caption = "Returns Types "
   
   SalesBasicSubsub(0).Caption = "Customer Groups  "
   SalesBasicSubsub(1).Caption = "Customer Calassifications  "
   
      SalesBasicSubsub(2).Caption = "Customer account Request"
      SalesBasicSubsub(3).Caption = "Customer Data"
SalesBasicSubsub(4).Caption = "Cash Customer Data"

        Me.SalesTransactions(0).Caption = "Quotations and Sales Orders"
 
        SalesTransactionssubss0(0).Caption = "Quotations"
        SalesTransactionssubss00(0).Caption = "Customers Quotations  Requests"
      '  SalesTransactionssubss00(1).Caption = "Quotations Approval  "
        SalesTransactionssubss00(1).Caption = "Quotations"
   
        SalesTransactionssubss0(1).Caption = "Sales Orders"
        SalesTransactionssubss000(0).Caption = "Primary Sales Orders"
     '   SalesTransactionssubss000(1).Caption = "Sales Orders Approval"
        SalesTransactionssubss000(1).Caption = "Sales Orders"
  
        Me.SalesTransactions(1).Caption = "Sales Order"
        Me.SalesTransactions(2).Caption = "Sales Invoices"
        Me.SalesTransactions(3).Caption = "Sales Return"
        Me.SalesTransactions(4).Caption = "Bill compound"
        Me.SalesTransactions(5).Caption = "Items Offers"
        Me.SalesTransactions(6).Caption = "Pricing plan"
 
        Me.SalesTransactions(7).Caption = "Price List"
        Me.SalesTransactions(8).Caption = "CRM"
        Me.SalesTransactions(9).Caption = "Ageing Report"
        Me.SalesTransactions(10).Caption = "Sales Reports"
        Me.SalesTransactions(11).Caption = "Cash Customer Reports"
        
        SalesTransactionsEmp(0).Caption = "Preparation of sales commissions and collections"

        SalesTransactionsEmp(1).Caption = "sales commissions and collections Plan"
        SalesTransactionsEmp(2).Caption = "Ratios achieve the objectives of sales and collections"

        SalesTransactionsEmp(3).Caption = "Commissions receivable For SalesPersons"
        SalesTransactionsEmp(4).Caption = "Quick Pay Offers"
        Archiving.Caption = "Electronic Archiving"
   
   
     
        ArchivingSub(0).Caption = "Departements"
        ArchivingSub(1).Caption = "Archive Data"
        ArchivingSub(2).Caption = "Rooms in Archive"
        ArchivingSub(3).Caption = "Boxes in Each Rooms "
        ArchivingSub(4).Caption = "Shelves in Each Boxes "
        ArchivingSub(5).Caption = "Documents Types "
        ArchivingSub(6).Caption = "Tempkates"
        ArchivingSub(7).Caption = "New Document"
        ArchivingSub(8).Caption = "Follow Document"
        ArchivingSub(9).Caption = "Alarms"
        ArchivingSub(10).Caption = "Reports"
        
   
        Me.Currency.Caption = "Fi&nancial Transactions"
        Me.ExpensesType(0).Caption = "Expenses Types"
        Me.ExpensesType(1).Caption = "Revenues Types"
        Me.ExpensesType(2).Caption = "Cheques Notes"
        
        Me.Expenses(0).Caption = "Financial Invoice"
        Me.Expenses(1).Caption = "Service Invoices"
            Me.Expenses(2).Caption = "Expenses Voucher"
            ExpensesSub(0).Caption = "Expenses Type"
       ExpensesSub(1).Caption = "Expenses Request"
        ExpensesSub(2).Caption = "Payments Voucher"
        ExpensesSub(3).Caption = "Payable Voucher "
        ExpensesSub(4).Caption = "Multiple Payments Voucher  "
        
        Me.Payments(0).Caption = "Notes Payable"
taxes.Caption = "VAT"
TaxexSub(0).Caption = "Settings"
LIFEINDICATORMNU.Caption = "Dash Board"
        Me.Cashing(0).Caption = "Cash Receipt Voucher"
        Me.Cashing(1).Caption = "General Cashing Voucher"
         Me.Cashing(1).Visible = False
        BankOp.Caption = "Banks Operations"
        
        Me.BankOpsub(0).Caption = "Bank Deposite"
        Me.BankOpsub(1).Caption = "cheque Release"
       Me.BankOpsub(2).Caption = "Bank Setellments"
        Me.BankOpsub(3).Caption = "Bank Report"
         Me.BankOpsub(4).Caption = "Print Cheque"
        Me.BankOpsub(5).Caption = "ÒReports "
        
        
                CeramicEstimation.Caption = "The assays"
        CeramicEstimationsub(0).Caption = "Units"
        CeramicEstimationsub(1).Caption = "Operations"
        CeramicEstimationsub(2).Caption = "Request Measurement"
        CeramicEstimationsub(3).Caption = "Orders Distribution"
        CeramicEstimationsub(4).Caption = "Conventions"
        CeramicEstimationsub(5).Caption = "Projects"
        CeramicEstimationsub(6).Caption = " Daily works"
       CeramicEstimationsub(7).Caption = "Ivvoicing"
        CeramicEstimationsub(8).Caption = "Reports"
        
        
        Me.MnuFinDiscounts.Caption = "Allowed and acquired Discounts"
        Me.DelayVal(0).Caption = "Debits Notes"
        '        Me.DelayVal(1).Caption = "Ageing Setting"
        '        Me.DelayVal(2).Caption = "Payable Ageing Report"
        
        Me.ReceiptPart.Caption = "Getting Installment"
        Me.RequiredInstallment.Caption = "Required Installment"
        Me.MnuCheckBriefcase.Caption = "cheque Briefcase"
        '  Me.MnuCheckOperations.Caption = "cheque Release"
        Me.MnuBoxDeposit(0).Caption = "Box Opening Balance"
        Me.MnuBoxDeposit(1).Caption = "Box Recharge and BT-cash"
        Me.MnuBoxDeposit(2).Caption = "Petty Cash Settlement"

        Me.MnuBoxDrawing.Caption = "Transfer Money "
        Me.MnuBoxAccouns.Caption = "Current Box Balance"
        Me.MnuBoxIncapacity_Increase(0).Caption = "Box Incapacity && Increase"
'        Me.MnuBoxIncapacity_Increase(1).Caption = "Service Invoice"
        
        'Me.MnuBoxStock.Caption = "Box Stock"
        
        Me.MnuAccounts.Caption = "Accounting"
        Me.MnuAccCharts(0).Caption = "Chart Of Accounts"
        Me.MnuAccCharts(1).Caption = "Accounts Opening Balance"
        '---------------------------------------------
        Me.AgeingMAster.Caption = "Ageing"
        Me.rsInvestment.Caption = "RS Investment"
        Me.MnuElevators.Caption = "Elevators"
        Me.SalesIns.Caption = "Insatllments Sales"
        Me.hajMnu.Caption = "Hajj"
         Me.StudentMenue.Caption = "Institutes"
         
        
        
        
        
        '--------------------------------------------
        ReportDesign.Caption = "Report Designer"
        Me.Reports.Caption = "Reports"
        Me.Report.Caption = "General Reports"
        Me.DailyReport.Caption = "Daily Reports"
        Me.MnuReports_Assblied.Caption = "Assblied Interval Report"
        Me.Tools.Caption = "System Manger"
         
        Me.Barcode.Caption = "Barcode Design..."
        Me.MnuPrintItemsCodes.Caption = "Items Codes Barcode Print..."
        'Me.MnuCorrectSerial.Caption = "Repaire Items Serial Number Errors"
        'Me.MnuBoxDetectErrors.Caption = "Repaire Box Balance Errors"
        Me.MnuToolCustomers.Caption = "Edit Customers Invoices"

        'Me.MnuToolRepaireItemsCost.Caption = "Adjust Items Cost in Bill Invoices"
        Me.MnuToolsDataBase(0).Caption = "Refresh DataBase Connectoion"
        Me.MnuToolsDataBase(1).Caption = "Update DataBase "
        '         Me.MnuToolsDataBase(2).Caption = "Change DataBase "
        Me.MnuDataBaseTools.Caption = "Data Base Tools"
        Me.UsersData.Caption = "Users"
        Me.AddUser.Caption = "Users Data..."
        'Me.DelUser.Caption = "Delete User..."
        Me.EditPw.Caption = "Change Password..."
        UserRpt.Caption = "Users Log File   "
             
            advanceMenu(0).Caption = "Advance request"
             advanceMenu(1).Caption = "Opening  Advance Registeration"
              advanceMenu(2).Caption = "Advance Modifications"
              
              
        Me.UserAbility.Caption = "Authority Matrix..."
        'Me.MnuUsersScreensPremission.Caption = "Users Screens Premission"
        Me.Options.Caption = "Options"
        Me.ShortCuts.Caption = "Shortcuts"
         
         Me.MnuToolsSetPrinters0(0).Caption = "IT Service Ticket"
        Me.MnuToolsSetPrinters0(1).Caption = "Set Local Printer..."
        Me.MnuToolsSetPrinters(1).Caption = "Accounts Coding"
        Me.MnuToolsSetPrinters(2).Caption = "Doc Type  "
        Me.MnuToolsSetPrinters(3).Caption = "Show Alarms "
         
        Me.MnuToolsSetPrinters(4).Caption = "Voucher Coding"
        Me.MnuToolsSetPrinters(5).Caption = "Fields Coding"
        Me.MnuToolsSetPrinters(6).Caption = " Local Messenger "
        
        Me.MnuToolsSetPrinters(7).Caption = " Dictionary "
         
        Me.MnuToolsSetPrinters7.Caption = " SMS Settings "
        
        Me.MnuInterface.Caption = "User Interface"
        Me.MnuInterfaceSub(0).Caption = "Arabic Interface"
        Me.MnuInterfaceSub(1).Caption = "English Interface"
        'Me.MnuWindowsList.Caption = "Programe Windows"
        'Me.MnuWindowsListOpen.Caption = "Opened Windows"
        Me.Help.Caption = "Help"
        'Me.HelpFile.Caption = "Contents..."
        'Me.HelpIndex.Caption = "Index..."
        'Me.SearchInHelp.Caption = "Search..."
        'Me.DailyToolTip.Caption = "Daily Tool Tip..."
   '     Me.FavoritesMenue.Caption = "Favorites Menue"
        'Me.About.Caption = "About..."
        'Me.ConnectUs.Caption = "Register..."
 
 help_list(0).Caption = "Modify Menue"
 
 '***************************************************************************
         dev.Caption = "Tasks And Performance"
        devsub(0).Caption = "Daily Tasks"
        devsub(1).Caption = "Follow Daily Workflow"
        devsub(2).Caption = "Define Tasks"
        devsub(3).Caption = "Follow Tasks"
        devsub(4).Caption = "Tasks Alarms"
        devsub(5).Caption = "Tasks Reports"
       
 '*******************************************************
        Me.HelpFileSub(0).Caption = "Contents..."
       Me.HelpFileSub(1).Caption = "Index..."
        Me.HelpFileSub(2).Caption = "Search..."
        Me.HelpFileSub(3).Caption = "Daily Tool Tip..."
        Me.FavoritesMenue.Caption = "Favorites Menu"
        Me.HelpFileSub(4).Caption = "About..."
       Me.HelpFileSub(5).Caption = "Register..."
Me.HelpFileSub(6).Caption = "Check List"

 Me.HelpFileSub(7).Caption = "Technical Support Forum ..."
 
 
        prdo.Caption = "Production"


  prdo1(0).Caption = "Basic Data"
  
        prdo1sub(0).Caption = "Equipments Data"
        prdo1sub(1).Caption = "Production Cost component   "
        prdo1sub(2).Caption = "Unit  Cost Of Production"
        prdo1sub(3).Caption = "Templates Data"
        
         prdo1sub(4).Caption = "Production Types"
         prdo1sub(5).Caption = "Items Indirect Cost"
          
        
       ' prdo1(0).Caption = "Templates Data"
       ' prdo1(1).Caption = "Equipments Data"
        
       ' prdo1(2).Caption = "Production Cost component   "
       ' prdo1(3).Caption = "Unit  Cost Of Production"
        
        
        prdo1(4).Caption = "Production Lines "
        prosub1(0).Caption = "Define Production Lines"
        prosub1(1).Caption = "Allocate and Trannsfer Employee "

        prdo1(5).Caption = "Production Cycle"

        prdo1(6).Caption = "Production Reservation Vchr"
        prdo1(7).Caption = "Production/Work Order"
        prdo1(8).Caption = "Issue Voucher-Row Material Items"
        prdo1(9).Caption = "Receive Voucher- Production Items"

        prdo1(10).Caption = "Typical production costs"
        prdo1(11).Caption = "Indirect Costs Distributions"
        prdo1(12).Caption = "Allocation Of Production order"
        LIFEINDICATORMNU.Caption = "Dash Board"
         
        
         prdo1(13).Visible = False
       prdo1(13).Caption = "Add Meter"
        prdo1(14).Caption = "Assembly Voucher"
        prdo1(15).Caption = "Production Reports"
 
      PrbH(0).Caption = "Production work order"
        PrbH(1).Caption = "Production Issue Voucher"
        
        PrbH(2).Caption = "Production Recieve Voucher "
 
        MNUFixedAssets.Caption = "FixedAssets"
        xxxxx(0).Caption = "Fixed Assets Groups"
        xxxxx(1).Caption = "Fixed Assets Data"
        xxxxx(2).Caption = "Fixed Assets Invoice"
        xxxxx(3).Caption = "Depreciation Installments Issueing"
        xxxxx(4).Caption = "Disposal  OF Fixed Assets"
        xxxxx(5).Caption = "FA Additions"
        xxxxx(6).Caption = "Assets Movements"
        
        xxxxx(7).Caption = "FA Adjustements"
        xxxxx(8).Caption = "Reports"
 
 
 
  ScreenSetting.Caption = "Screens Settings"
           MnuLevels(0).Caption = "Documents Approvals"
        MnuLevelsSub(0).Caption = "Approval Levels"
        MnuLevelsSub(1).Caption = "Approval for Documents"
 
          MnuLevels(1).Caption = "Screen criteria"
        MnuLevelsSub2(0).Caption = "Define Screen criteria"
        MnuLevelsSub2(1).Caption = "Screen  criteria Settings"
        
        
        
        
       ' ArrowsBase.Caption = "Arrows Mangements"
       ' ArrowsFollow(0).Caption = "Capital Market Data"
       ' ArrowsFollow(1).Caption = "Groups of Arrows"
       ' ArrowsFollow(2).Caption = "Companies Data"
       ' ArrowsFollow(3).Caption = "Loading Prices"
       ' ArrowsFollow(4).Caption = "Historical prices"
       ' ArrowsFollow(5).Caption = "Bockets"
'
'        ArrowsFollowBocket(0).Caption = "Bockets Data"
'        ArrowsFollowBocket(1).Caption = "Arrows Purchases"
'        ArrowsFollowBocket(2).Caption = "Arrows Salling"
'        ArrowsFollowBocket(3).Caption = "Arrows Current Value"

'        ArrowsFollow(6).Caption = "Links"
'        ArrowsFollow(7).Caption = "Reports"

        '
        Me.MnuPopItemsTreePane_Array(0).Caption = "Refresh"
        Me.MnuPopItemsTreePane_Array(2).Caption = "Dock"
        Me.MnuPopItemsTreePane_Array(3).Caption = "Close"
'        Me.MnuPopItemsTreePane_Array(5).Caption = "Groups Sort"
         
    End If

    Exit Sub
ErrTrap:

    Stop
End Sub

Private Sub SetMenusHelp()

End Sub

Public Function GetDayTransSQL(IntTransType) As String

End Function

Public Function AskForExit() As Boolean
    Dim Msg As String
    Dim IntRes As Integer

    'Stop
    If SystemOptions.UserInterface = EnglishInterface Then
        Msg = "Confirm Exit"
    Else
        Msg = "Â·  —Ìœ «·Œ—ÊÃ „‰ «·»—‰«„Ã .øø"
    End If

    IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

    If IntRes = vbYes Then
        'End
        '    Exit Function
        AddToLogFile CInt(user_id), 0, Date, Time, "  ”ÃÌ·   «·Œ—ÊÃ „‰ «·‰Ÿ«„ ", " System LogOut", Me.Name, "L", "", ""
        AskForExit = True
        'Me.TimerData.Enabled = False
        ClosePanes
        SystemOptions.BolStopUpdateTask = True

        If Forms.count = 1 Then
            SystemOptions.BolUpdateTaskInProgress = False
        End If

        Do While SystemOptions.BolUpdateTaskInProgress = True
            DoEvents

            If Forms.count = 1 Then
                SystemOptions.BolUpdateTaskInProgress = False
            End If

            'SystemOptions.BolUpdateTaskInProgress = False
        Loop

        'ClearTaskPanel Me.TaskPanel1
        CloseApplication
    Else
        AskForExit = False
        Exit Function
    End If

    Unload Me
End Function

Private Sub CreateDocks(Optional BolWithCheck As Boolean = True)
    Dim i As Integer

    Dim x As XtremeDockingPane.Pane
    Dim Y As XtremeDockingPane.Pane
    Dim xItemsTreePane As XtremeDockingPane.Pane
    Dim xMantaincePane As XtremeDockingPane.Pane
    Dim xInternetPane As XtremeDockingPane.Pane
    Dim xHelpPane As XtremeDockingPane.Pane
    Dim xCalendarPane As XtremeDockingPane.Pane
    Dim XTip As XtremeDockingPane.ToolTipContext

    '------------------------------------------------
    For i = 1 To Me.ImgLstMenuIcons.ListImages.count

        If Me.ImgLstMenuIcons.ListImages.Item(i).Tag <> "" Then
            'Stop
        End If

        Me.ImgLstMenuIcons.ListImages.Item(i).Tag = Me.ImgLstMenuIcons.ListImages.Item(i).Index
    Next i

    '------------------------------------------------
    Set DockingPane1.ImageList = Me.ImgLstTree

    Set x = Me.DockingPane1.CreatePane(DockingPanesIDs.NewsBarPaneID, 250, 200, DockLeftOf, Nothing)
    x.IconId = 2

    Set Y = Me.DockingPane1.CreatePane(DockingPanesIDs.OutBarPaneID, 150, 200, DockRightOf, Nothing)
    Y.IconId = 1 'Me.ImgLstMenuIcons.ListImages("").Index

    Set xItemsTreePane = Me.DockingPane1.CreatePane(DockingPanesIDs.ItemsTreeID, 250, 200, DockLeftOf, Nothing)
    'xItemsTreePane.IconId = Me.ImgLstMenuIcons.ListImages("TreeItems").Tag
    xItemsTreePane.Options = PaneHasMenuButton

    Set xInternetPane = Me.DockingPane1.CreatePane(DockingPanesIDs.InternetNews, 250, 250, DockLeftOf, Nothing)
    '    xInternetPane.IconId = Me.ImgLstMenuIcons.ListImages("Options").Index
    xInternetPane.Options = PaneHasMenuButton
    Set xHelpPane = Me.DockingPane1.CreatePane(DockingPanesIDs.DynamicHelp, 250, 250, DockLeftOf, Nothing)
    '    xHelpPane.IconId = 6 'Me.ImgLstMenuIcons.ListImages("Help2").Index
    xHelpPane.Options = PaneNoCloseable
    
    ' DockingPane1.FindPane(DockingPanesIDs.DynamicHelp).MinTrackSize.setSize 100, 100
 
    If SystemOptions.SysMantainceAllow = True Then
        Set xMantaincePane = Me.DockingPane1.CreatePane(DockingPanesIDs.MantainceID, 250, 200, DockLeftOf, Nothing)

        If SystemOptions.UserInterface = ArabicInterface Then
            xMantaincePane.title = "«·’Ì«‰…"
        Else
            xMantaincePane.title = "Mantaince"
        End If

        xMantaincePane.Options = PaneHasMenuButton
        '    xMantaincePane.IconId = Me.ImgLstMenuIcons.ListImages("Tools").Index
    End If

    Set xCalendarPane = Me.DockingPane1.CreatePane(DockingPanesIDs.CalendarPaneID, 250, 250, DockLeftOf, Nothing)
      xCalendarPane.IconId = Me.ImgLstMenuIcons.ListImages("OpenAcc").Index
    xCalendarPane.Options = PaneHasMenuButton
    
    If SystemOptions.UserInterface = ArabicInterface Then
        x.title = "„⁄·Ê„«  «·»—‰«„Ã"
        Y.title = "‘—Ìÿ «·√Œ ’«—« "
        xItemsTreePane.title = "‘Ã—… «·√’‰«›"
        xInternetPane.title = "√Œ»«— «·√‰ —‰ "
        xHelpPane.title = "«·œ⁄«Ì…"
        xCalendarPane.title = "«·”«⁄…"
    Else
        x.title = "Information OutBar"
        Y.title = "Shortcut OutBar"
        xItemsTreePane.title = "Items Tree"
        xInternetPane.title = "Internet News"
        xHelpPane.title = "Dynamic Help"
        xCalendarPane.title = "Calendar"
    End If

    DockingPane1.VisualTheme = ThemeVisio
    DockingPane1.HidePane x
    DockingPane1.HidePane xItemsTreePane
    DockingPane1.HidePane xInternetPane
'    DockingPane1.HidePane xCalendarPane

    DockingPane1.ToolTipContext.ShowShadow = True
    DockingPane1.ToolTipContext.Style = xtpToolTipOffice2007

    If Not xMantaincePane Is Nothing Then

        DockingPane1.HidePane xMantaincePane
    End If

    Me.DockingPane1.LoadState "bisegypt", "SmallAccount", "DockingPanes"
    'If BolWithCheck = True Then
    '    Me.DockingPane1.LoadState "bisegypt", "SmallAccount", "DockingPanes"
    '    If Me.DockingPane1.PanesCount = 0 Then
    '        CreateDocks False
    '    End If
    'End If

    '-----------------------

End Sub

Private Sub ClosePanes()
    Dim i As Integer
    SaveDockingPanes

    For i = 1 To Me.DockingPane1.PanesCount
        Me.DockingPane1(i).Hide
        Me.DockingPane1(i).Close

        DoEvents
    Next i

    'If Not FrmOutBarPane Is Nothing Then
        'Unload FrmOutBarPane
    'End If

    'If Not FrmNewsBarPane Is Nothing Then
        'Unload FrmNewsBarPane
    'End If

    'If Not ItemsTreePane Is Nothing Then
    '    Unload ItemsTreePane
    'End If

    If Not FrmDynamicHelpPane Is Nothing Then
    '    Unload FrmDynamicHelpPane
    End If

    'If Not FrmCalendarPane Is Nothing Then
    '    Unload FrmCalendarPane
    'End If

End Sub

Private Sub LoadDockingPanes()

End Sub

Private Sub SaveDockingPanes()

    Dim xPaneRec As PaneRecorde
    Dim IntFreeFile As Integer
    Dim StrFile As String
    Dim i As Integer
    Dim xx As XtremeDockingPane.PaneContainer
    Me.DockingPane1.SaveState "bisegypt", "SmallAccount", "DockingPanes"
    IntFreeFile = FreeFile
    StrFile = App.path & "\Temp.dat"

    If Dir(StrFile) <> "" Then
        Kill StrFile
    End If

    Open StrFile For Random As #IntFreeFile Len = Len(xPaneRec)

    For i = 1 To Me.DockingPane1.PanesCount
        xPaneRec.PaneID = Me.DockingPane1.Panes(i).ID
        xPaneRec.PanePositon = Me.DockingPane1(i).Position
        xPaneRec.PaneTitle = Me.DockingPane1(i).title
        xPaneRec.PaneClosed = Me.DockingPane1(i).Closed
        xPaneRec.PaneEnabled = Me.DockingPane1(i).Enabled
        xPaneRec.PaneFloated = Me.DockingPane1(i).Floating
        xPaneRec.PaneHidden = Me.DockingPane1(i).Hidden
        Put #IntFreeFile, , xPaneRec
    Next i

    Close #IntFreeFile
End Sub

Private Sub CreateWindowList()
    On Error Resume Next
    Dim i As Integer, j As Integer
    Dim Lparent As Long
    Dim BolTemp As Boolean
    Dim IntCount As Integer
    Dim StrOldFrmName As String

    If mdifrmmain.ActiveForm Is Nothing Then
        Me.PopMenu1.ClearSubMenusOfItem ("MnuWindowsListOpen")
  '      MnuWindowsListOpen.Enabled = False
        Exit Sub
    Else
        'MnuWindowsListOpen.Enabled = True
    End If

    Me.PopMenu1.ClearSubMenusOfItem ("MnuWindowsListOpen")

    For i = 0 To Forms.count - 1

        If Forms(i).Name <> "MDIFrmMain" Then
            If Forms(i).MDIChild = True Then

                With Me.PopMenu1
                    Lparent = .MenuIndex("MnuWindowsListOpen")

                    If ImgInImgList(Forms(i).Name) = -1 Then
                        Dim CCCC As Long
                        'Me.ImgLstMenuIcons.ListImages.Add , Forms(I).name, Forms(I).Icon
                        'me.ImgLstMenuIcons.ListImages.Add
                        'cccc=me.ImgLstMenuIcons.ListImages(forms(i).name).
                        Dim xx As IPictureDisp
                        Set xx = Forms(i).Icon
                        Me.ilsIcons.AddFromHandle xx.Handle, IMAGE_ICON, Forms(i).Name
                    End If

                    BolTemp = False

                    For j = 1 To .count

                        If StrOldFrmName <> Forms(i).Name Then
                            IntCount = 0
                            StrOldFrmName = Forms(i).Name
                        End If

                        If .MenuKey(j) = Forms(i).Name Then
                            IntCount = IntCount + 1
                            StrOldFrmName = Forms(i).Name
                            BolTemp = True
                        End If

                    Next j

                    If BolTemp = False Then
                        .AddItem Forms(i).Caption, Forms(i).Name, , 2000 + .count, Lparent, Me.ilsIcons.ItemIndex(Forms(i).Name) - 1, True, True
                    ElseIf BolTemp = True Then
                        .AddItem Forms(i).Caption & " " & IntCount, , , 2000 + .count, Lparent, Me.ilsIcons.ItemIndex(Forms(i).Name) - 1, True, True
                    End If

                    If mdifrmmain.ActiveForm.Name = Forms(i).Name Then
                        .MenuDefault(Forms(i).Name) = True
                    Else
                        .MenuDefault(Forms(i).Name) = False
                    End If

                End With

            End If
        End If

    Next i

End Sub

Private Sub xxx_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("FrmCostCenterType1") = False Then
                Exit Sub
            End If

            FrmCostCenterType1.show

        Case 1

            If checkApility("CostCenter") = False Then
                Exit Sub
            End If

            CostCenter.show

            'Frmcostcenter.Show' Ì „  ›⁄Ì·Â« ﬁ—Ì»«
            ' frm_marakez_taklefa.Show
        Case 2

        Case 3

            If checkApility("mowazna") = False Then
                Exit Sub
            End If

      '      mowazna.show

        Case 4
          '  tahlil_maly.show

        Case 5

            If checkApility("Cash_flow") = False Then
                Exit Sub
            End If

            Cash_flow.show

        Case 6

            If checkApility("FrmAccountDestribution") = False Then
                Exit Sub
            End If

            FrmAccountDestribution.show

        Case 7

 

        Case 8

            If checkApility("FrmBalanceSheet") = False Then
                Exit Sub
            End If

            'BaklanceSheetvIEW.Show
            FrmBalanceSheet.show

        Case 9

            If checkApility("FinancialAnalysis") = False Then
                Exit Sub
            End If

            FinancialAnalysis.show

        Case 10

            If checkApility("FinancialAnalysisView") = False Then
                Exit Sub
            End If

            FinancialAnalysisView.show

        Case 11
            FrmCompositeAccounts.show

        Case 12

            If checkApility("FrmAccountingReport") = False Then
                Exit Sub
            End If

            FrmAccountingReport.show

    End Select

End Sub

Private Sub xxxxx_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("FixedAssetsGroup") = False Then
                Exit Sub
            End If

            FixedAssetsGroup.show

        Case 1

            If checkApility("FixedAssets") = False Then
                Exit Sub
            End If

            FixedAssets.show

        Case 2

            If checkApility("FrmExpenses4") = False Then
                Exit Sub
            End If

            FrmExpenses4.show
 
        Case 3

            If checkApility("FrmCase1") = False Then
                Exit Sub
            End If

            FrmCase1.show

        Case 4

            If checkApility("FrmExpenses40E") = False Then
                Exit Sub
            End If
    
            'FrmExpenses40.show
           FrmExpenses40E.show

        Case 5
        
            If checkApility("FrmExpenses40A") = False Then
                Exit Sub
            End If
            FrmExpenses40A.show

        Case 6
         If checkApility("FrmTransferAssets") = False Then
                Exit Sub
            End If
            
            FrmTransferAssets.show
            
            

        Case 7

      FrmNewGard10.show
      
      
Case 8


            If checkApility("ShowFixedAssets") = False Then
                Exit Sub
            End If
    
            frmFixedAsseteports.show
            
            
      

    End Select

End Sub

 
Private Sub xxy_Click(Index As Integer)

    Select Case Index
            
                  Case 11
           If checkApility("FrmProductionPlan") = False Then
                Exit Sub
            End If
            
             FrmProductionPlan.show
             
             
        Case 0

'            If checkApility("mowazna") = False Then
'                Exit Sub
'            End If
'
'            mowazna.show
 If checkApility("FrmEstimations") = False Then
                Exit Sub
            End If
FrmEstimations.show

 


        Case 1

            If checkApility("Cash_flow") = False Then
                Exit Sub
            End If

            Cash_flow.show

        Case 2

            If checkApility("FrmBalanceSheet") = False Then
                Exit Sub
            End If
 
            FrmBalanceSheet.show
            'FrmBalanceSheet.Show

        Case 3

            If checkApility("FrmAccountDestribution") = False Then
                Exit Sub
            End If

            FrmAccountDestribution.show

        Case 4

            If checkApility("FinancialAnalysis") = False Then
                Exit Sub
            End If

            FinancialAnalysis.show

        Case 5

            If checkApility("FinancialAnalysisView") = False Then
                Exit Sub
            End If

            FinancialAnalysisView.show

        Case 6

            If checkApility("FrmCompositeAccounts") = False Then
                Exit Sub
            End If

            FrmCompositeAccounts.show

        Case 7
 If checkApility("FrmCorBalaCusDet") = False Then
                Exit Sub
            End If

    FrmCorBalaCusDet.show
            
           ' If checkApility("FrmStatistics") = False Then
           '     Exit Sub
           ' End If
'
'            OpenScreen StatisticsShow
'
        Case 8

            If checkApility("FrmCustomersAgenda") = False Then
                Exit Sub
            End If

 

        Case 9
        
             If checkApility("FrmLoadTrialBalance") = False Then
                Exit Sub
            End If
         '   FrmBalanceSheet1.show
FrmLoadTrialBalance.show

    End Select

End Sub

 Private Sub xyzSub_Click(Index As Integer)
Select Case Index
Case 0

            If checkApility("FrmCompanies") = False Then
                Exit Sub
            End If

            FrmCompanies.show


Case 1

            If checkApility("FrmContStudent") = False Then
                Exit Sub
            End If

            FrmContStudent.show
            
            
 Case 2
 
                    If checkApility("FrmVisa") = False Then
                Exit Sub
            End If
           FrmVisa.show
           
Case 3
             If checkApility("FrmStudentsCandidacy") = False Then
                Exit Sub
            End If

            FrmStudentsCandidacy.show
    Case 4


            If checkApility("Projects") = False Then
                Exit Sub
            End If

       '     Projects1.show
         Projects.show


Case 5
            If checkApility("FrmEmpSalary3") = False Then
                 Exit Sub
            End If

           
            FrmEmpSalary3.show
 Case 6
            If checkApility("projectsbill") = False Then
                Exit Sub
            End If
 
            projectsbill.show
Case 7
            If checkApility("FrmProjectMonthBill") = False Then
                Exit Sub
            End If
FrmProjectMonthBill.show
 
 End Select
End Sub



