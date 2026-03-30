VERSION 5.00
Object = "{798A85D3-625A-4512-A9E4-BA96E09CA6A6}#1.0#0"; "ciaXPIML30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3D800911-77E3-43DE-82EA-7FC87C713180}#1.2#0"; "cPopMenu6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "DOCKIN~1.OCX"
Begin VB.MDIForm mdifrmmain 
   BackColor       =   &H00E2E9E9&
   Caption         =   " "
   ClientHeight    =   5670
   ClientLeft      =   5730
   ClientTop       =   4275
   ClientWidth     =   9645
   Icon            =   "MDIFrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2520
      Top             =   1680
   End
   Begin MSComctlLib.StatusBar XPStusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   5325
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   360
      Top             =   1200
   End
   Begin cPopMenu6.PopMenu PopMenu1 
      Left            =   6420
      Top             =   2370
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgLstTree 
      Left            =   5310
      Top             =   2100
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
            Picture         =   "MDIFrmMain.frx":324A
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":35E4
            Key             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":397E
            Key             =   "Refresh"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3D18
            Key             =   "receipt"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":40B2
            Key             =   "Required"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":464C
            Key             =   "Balance"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":49E6
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4D80
            Key             =   "Dollar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":63DA
            Key             =   "Item2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":6774
            Key             =   "Design"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":6B0E
            Key             =   "Request"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":70A8
            Key             =   "Favorites"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":7442
            Key             =   "Wizared"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":77DC
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":7B76
            Key             =   "Excute"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":7F10
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":84AA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":8844
            Key             =   "save"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":8BDE
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":8F78
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":9312
            Key             =   "Sall"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":96AC
            Key             =   "Clients"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":9A46
            Key             =   "Groups"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":9DE0
            Key             =   "Maintenance"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":A17A
            Key             =   "Items"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":A514
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":A8AE
            Key             =   "Supplier"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":AC48
            Key             =   "barcode"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":AFE2
            Key             =   "ReturnBack"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":B57C
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":B916
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":BCB0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":C04A
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":C3E4
            Key             =   "Purchase"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":C77E
            Key             =   "store"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":CB18
            Key             =   "LIST"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":CEB2
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":D24C
            Key             =   "DReport"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":D5E6
            Key             =   "From"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":D980
            Key             =   "To"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":DD1A
            Key             =   "User"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":E0B4
            Key             =   "Tax"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":E44E
            Key             =   "Currency"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":E7E8
            Key             =   "Discount"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":EB82
            Key             =   "DiscountType"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":EF1C
            Key             =   "Tick"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":F2B6
            Key             =   "Date"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":F650
            Key             =   "Ask"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":FBEA
            Key             =   "number"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":FF84
            Key             =   "qty"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1031E
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":106B8
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":10A52
            Key             =   "Closed_Node"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":10DEC
            Key             =   "Open_Node"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":11186
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":11720
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":11ABA
            Key             =   "Serial"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":11E54
            Key             =   "code"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":121EE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":12588
            Key             =   "Plus"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":12922
            Key             =   "Minus"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":12CBC
            Key             =   "FillData"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":13056
            Key             =   "GridOptions"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":133F0
            Key             =   "Tree"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1378A
            Key             =   "Assblied"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":13B24
            Key             =   "LinkItem"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":13EBE
            Key             =   "ItemPart"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":14258
            Key             =   "Filter"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Cmdlg 
      Left            =   6600
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgLstMenuIcons 
      Left            =   4680
      Top             =   2040
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
            Picture         =   "MDIFrmMain.frx":145F2
            Key             =   "Salles"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1498C
            Key             =   "Warn"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":14D26
            Key             =   "Screen"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":150C0
            Key             =   "Execute"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1545A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":157F4
            Key             =   "Purashes"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":15B8E
            Key             =   "DEV_Preview"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":15F28
            Key             =   "OpenAcc"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":164C2
            Key             =   "AccReports"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1685C
            Key             =   "Options"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":16BF6
            Key             =   "Emp"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":17190
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1752A
            Key             =   "Items"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1837C
            Key             =   "store"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":18C56
            Key             =   "Invoice"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":18FF0
            Key             =   "NewAccout"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1938A
            Key             =   "NewGroupAccount"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":19724
            Key             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":19ABE
            Key             =   "ToGroup"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1A058
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1A3F2
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1A78C
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1AB26
            Key             =   "Screens"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1AEC0
            Key             =   "HotKey"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1B1DA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1B574
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1B90E
            Key             =   "Tools"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1BCA8
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1C042
            Key             =   "PrintSetup"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1C3DC
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1C776
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1CB10
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1CEAA
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1D244
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1D5DE
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1D978
            Key             =   "MoveFirst"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1DD12
            Key             =   "MovePrevious"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1E0AC
            Key             =   "MoveNext"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1E446
            Key             =   "MoveLast"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1E7E0
            Key             =   "Money1"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1F0BA
            Key             =   "ToolTip"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1F454
            Key             =   "DEV_Edit"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1F76E
            Key             =   "Reports"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":1FB08
            Key             =   "Suppliers"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":200A2
            Key             =   "Customers"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":20EF4
            Key             =   "Help1"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":22BFE
            Key             =   "Cal"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":22F98
            Key             =   "OpenStore"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":233EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":23784
            Key             =   "EditTree"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":23B1E
            Key             =   "NewItem"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":23EB8
            Key             =   "Users"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":24252
            Key             =   "AddUser"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":245EC
            Key             =   "DeleteUser"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":24986
            Key             =   "UserPass"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":24D20
            Key             =   "UserPremis"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":250BA
            Key             =   "DataBaseBackup"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":25454
            Key             =   "DataBaseRestore"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":257EE
            Key             =   "DataBaseRepaire"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":25B88
            Key             =   "NewDataBase"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":25F22
            Key             =   "DataBaseReg"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":262BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2670E
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":273E8
            Key             =   "Tick"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":27782
            Key             =   "Home"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":27B1C
            Key             =   "TreeItems"
            Object.Tag             =   "65"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":27EB6
            Key             =   "NewGroup"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":28250
            Key             =   "DataBase"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":285EA
            Key             =   "About"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":28984
            Key             =   "WindowMin"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":28D1E
            Key             =   "WindowMax"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":290B8
            Key             =   "City"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":29992
            Key             =   "GridDelRow"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":29D2C
            Key             =   "Bank"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2A046
            Key             =   "Pur"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2A1A0
            Key             =   "OutOrder"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2A53A
            Key             =   "InOrder"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2A8D4
            Key             =   "Dev_Screen"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2AC6E
            Key             =   "Prop"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2B008
            Key             =   "Money2"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2B3A2
            Key             =   "Money3"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2B73C
            Key             =   "DefColor"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2BAD6
            Key             =   "CusColor"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2BE70
            Key             =   "Caps"
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2C20A
            Key             =   "Clock"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2C5A4
            Key             =   "Num"
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2C93E
            Key             =   "Calender"
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2CCD8
            Key             =   "User"
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2D072
            Key             =   "KeyBorad"
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2D94C
            Key             =   "LogOFF"
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2DEE6
            Key             =   "Interface"
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2E280
            Key             =   "BarCode"
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2E61A
            Key             =   "UserOptions"
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2E9B4
            Key             =   "InvoiceDesign"
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2ED4E
            Key             =   "Unit"
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2F0E8
            Key             =   "grd"
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2F482
            Key             =   "StoreCon"
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2FA1C
            Key             =   "StoreEx"
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":2FDB6
            Key             =   "StoreIm"
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":30150
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":306EA
            Key             =   "Web"
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":30C84
            Key             =   "wazrid"
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3101E
            Key             =   "Vertical"
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":313B8
            Key             =   "Horizental"
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":31752
            Key             =   "TabDown"
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":31AEC
            Key             =   "TabRight"
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":31E86
            Key             =   "TabUp"
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":32220
            Key             =   "TabLeft"
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":325BA
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":32954
            Key             =   "ItemsPrice"
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":32CEE
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":33088
            Key             =   "Unlock"
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":33422
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":339BC
            Key             =   "Help2"
            Object.Tag             =   "113"
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":33D56
            Key             =   "SearchHelp"
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":340F0
            Key             =   "Hide"
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3468A
            Key             =   "SortASC"
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":34A24
            Key             =   "SortDESC"
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":34DBE
            Key             =   "BrowseFile"
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":35358
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":356F2
            Key             =   "ExportExcel"
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":35A8C
            Key             =   "ExportPDF"
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":35E26
            Key             =   "ExportWord"
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":361C0
            Key             =   "ExportHTML"
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3655A
            Key             =   "ExportMail"
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":368F4
            Key             =   "Plus"
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":36C8E
            Key             =   "Mins"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5340
      Top             =   3600
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
            Picture         =   "MDIFrmMain.frx":37028
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":37704
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":37DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":384DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":38BB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":39291
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":39981
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3A074
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3A751
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3AE38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3B518
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3BC04
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3C2F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3C9E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3D0C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3D7A7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4740
      Top             =   3600
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
            Picture         =   "MDIFrmMain.frx":3DE92
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3E579
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3EC4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3F309
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":3F9C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":40084
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":40738
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":40E0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":414CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":41B9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":42277
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4294F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4301D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":436DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":43D9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmMain.frx":4446C
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
      Images          =   "MDIFrmMain.frx":44B3A
      KeyCount        =   11
      Keys            =   "ˇˇˇˇˇˇˇˇˇˇ"
   End
   Begin vbalIml6.vbalImageList ilsIcons 
      Left            =   5340
      Top             =   2670
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   32
      Size            =   81508
      Images          =   "MDIFrmMain.frx":473BE
      Version         =   131072
      KeyCount        =   71
      Keys            =   $"MDIFrmMain.frx":5B242
   End
   Begin XtremeDockingPane.DockingPane DockingPane1 
      Left            =   480
      Top             =   2040
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin VB.Menu BasicData 
      Caption         =   "«·»Ì«‰«  «·«”«”Ì…"
      Begin VB.Menu BasicDataM 
         Caption         =   "«⁄œ«œ«  «·—»ÿ „⁄  «·Õ”«»« "
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·«‰‘ÿÂ  Ê «·›—Ê⁄"
         Index           =   1
         Visible         =   0   'False
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
         Caption         =   "»Ì«‰«  «·⁄„·« "
         Index           =   7
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "«·Ã‰”Ì« "
         Index           =   8
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "«·œÌ«‰« "
         Index           =   9
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·œÊ·"
         Index           =   10
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·„Õ«›Ÿ«  Ê«·„‰«ÿﬁ"
         Index           =   11
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·«ÕÌ«¡"
         Index           =   12
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·‘Ê«—⁄"
         Index           =   13
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "«‰Ê«⁄ «·„” ‰œ« "
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "»Ì«‰«  «·«’‰«›"
         Index           =   15
      End
      Begin VB.Menu BasicDataM 
         Caption         =   "Œ—ÊÃ"
         Index           =   17
         Shortcut        =   ^K
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
         Caption         =   "»Ì«‰«  «·⁄„·«¡"
         Index           =   2
      End
      Begin VB.Menu TransporterSub 
         Caption         =   "»Ì«‰«  «·„Ê—œÌ‰"
         Index           =   3
      End
      Begin VB.Menu TransporterSub 
         Caption         =   "»Ì«‰«  «·”«∆ﬁÌ‰"
         Index           =   4
      End
      Begin VB.Menu TransporterSub 
         Caption         =   "«‰Ê«⁄ «·„—ﬂ»« "
         Index           =   5
      End
      Begin VB.Menu TransporterSub 
         Caption         =   "‘—ﬂ«  «· √„Ì‰"
         Index           =   6
      End
      Begin VB.Menu TransporterSub 
         Caption         =   "«‰Ê«⁄ «·’Ì«‰… «·œÊ—Ì…"
         Index           =   7
      End
      Begin VB.Menu TransporterSub 
         Caption         =   "»Ì«‰«  «·„—ﬂ»« "
         Index           =   8
      End
      Begin VB.Menu TransporterSub 
         Caption         =   "»Ì«‰«  «·—Õ·« "
         Index           =   9
      End
      Begin VB.Menu TransporterSub 
         Caption         =   "«· ﬁ«—Ì—"
         Index           =   10
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
            Caption         =   "ÊÕœ«  «·⁄„·Ì« "
            Index           =   3
         End
         Begin VB.Menu MnuProjectsBasicSub 
            Caption         =   "  ⁄—Ì› «·⁄„·Ì«  "
            Index           =   4
         End
         Begin VB.Menu MnuProjectsBasicSub 
            Caption         =   "»Ì«‰«  «·„‘«—Ì⁄"
            Index           =   5
         End
      End
      Begin VB.Menu MnuProjectsTransactions 
         Caption         =   "’—› „Ê«œ ⁄·Ï „‘—Ê⁄"
         Index           =   0
      End
      Begin VB.Menu MnuProjectsTransactions 
         Caption         =   " Œ’Ì’ ⁄„«·Â ·„‘—Ê⁄"
         Index           =   1
      End
      Begin VB.Menu MnuProjectsTransactions 
         Caption         =   "«‰Â«¡  Œ’Ì’ Ê‰ﬁ· ⁄„«·Â »Ì‰ «·„‘«—Ì⁄"
         Index           =   2
      End
      Begin VB.Menu MnuProjectsTransactions 
         Caption         =   "„ «»⁄Â «·⁄„·Ì« "
         Index           =   3
      End
      Begin VB.Menu MnuProjectsTransactions 
         Caption         =   "›« Ê—… „‘—Ê⁄"
         Index           =   4
      End
      Begin VB.Menu MnuProjectsTransactions 
         Caption         =   " ﬁ«—Ì— «·„‘«—Ì€"
         Index           =   5
      End
   End
   Begin VB.Menu prdo 
      Caption         =   "«·«‰ «Ã Ê√Ê«„— «·‘€·"
      Index           =   0
      Begin VB.Menu prdo1 
         Caption         =   "«‰Ê«⁄ «·œÊ«„ / «·Ê—œÌ« "
         Index           =   0
      End
      Begin VB.Menu prdo1 
         Caption         =   "»Ì«‰«  «·„⁄œ«  / «·„«ﬂÌ‰« "
         Index           =   1
      End
      Begin VB.Menu prdo1 
         Caption         =   "ŒÿÊÿ «·«‰ «Ã"
         Index           =   2
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
         Index           =   3
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
         Caption         =   "ÿ·»Ì… ‘—«¡"
         Index           =   4
      End
      Begin VB.Menu prdo1 
         Caption         =   "«„— «·«‰ «Ã/«·‘€·"
         Index           =   5
      End
      Begin VB.Menu prdo1 
         Caption         =   "”‰œ ’—› „Ê«œ Œ«„"
         Index           =   6
      End
      Begin VB.Menu prdo1 
         Caption         =   "”‰œ «” ·«„ «‰ «Ã  «„"
         Index           =   7
      End
      Begin VB.Menu prdo1 
         Caption         =   "Õ”«»  ﬂ«·Ì› «·«‰ «Ã «·‰„ÿÌ"
         Index           =   8
      End
      Begin VB.Menu prdo1 
         Caption         =   " Ê“Ì⁄ «· ﬂ«·Ì› €Ì— «·„Ì«‘—…"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu prdo1 
         Caption         =   " ﬁ«—Ì— «·«‰ «Ã"
         Index           =   10
      End
   End
   Begin VB.Menu ProductionPlan 
      Caption         =   " «· ŒÿÌÿ Ê„—«ﬁ»Â «·ÃÊœ…"
      Visible         =   0   'False
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
   End
   Begin VB.Menu MnuMaintnance 
      Caption         =   " «·’Ì«‰…"
      Begin VB.Menu MnuMaintnanceBasic 
         Caption         =   "»Ì«‰«  «”«”ÌÂ       "
         Begin VB.Menu MnuMaintnanceBasicSub 
            Caption         =   "«‰Ê«⁄ «·’Ì«‰…"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu MnuMaintnanceBasicSub1 
            Caption         =   "‘—ﬂ«  «·’Ì«‰Â"
         End
      End
      Begin VB.Menu MnuMaintnanceTransactions 
         Caption         =   "œ ŒÊ· «·’Ì«‰Â"
         Index           =   0
      End
      Begin VB.Menu MnuMaintnanceTransactions 
         Caption         =   "„Œ“‰ «·’Ì«‰Â"
         Index           =   1
      End
      Begin VB.Menu MnuMaintnanceTransactions 
         Caption         =   "’—› ﬁÿ⁄ €Ì«— ··’Ì«‰…"
         Index           =   2
      End
      Begin VB.Menu MnuMaintnanceTransactions 
         Caption         =   " ”·Ì„ «·’Ì«‰…"
         Index           =   3
      End
      Begin VB.Menu MnuMaintnanceTransactions 
         Caption         =   "—ÃÊ⁄ ÷„«‰ „‰ „Ê—œ"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu MnuMaintnanceTransactions 
         Caption         =   "—’Ìœ «›  «ÕÌ ·„Œ“‰ «·’Ì«‰…"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu MnuMaintnanceTransactions 
         Caption         =   "Ã—œ „Œ“‰ «·’Ì«‰…"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu MnuMaintnanceTransactions 
         Caption         =   "«—”«·  ‰»ÌÂ  Ã„Ì⁄ «ÃÂ“…"
         Index           =   7
      End
      Begin VB.Menu MnuMaintnanceTransactions 
         Caption         =   " ﬁ«—Ì— «·’Ì«‰Â"
         Index           =   8
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
            Caption         =   "  ⁄—Ì› «”⁄«—  «·»Ì⁄"
            Index           =   8
         End
         Begin VB.Menu StockControlBasicSub 
            Caption         =   "⁄‰«’— «· ﬂ«·Ì› «·’‰«⁄ÌÂ"
            Index           =   9
         End
         Begin VB.Menu StockControlBasicSub 
            Caption         =   "«· ﬂ·›… «· ﬁœÌ—Ì… ÿ»ﬁ« ·„Ã„Ê⁄«  «·«’‰«›"
            Index           =   10
         End
         Begin VB.Menu StockControlBasicSub 
            Caption         =   "Œÿ… „»Ì⁄«  «·«’‰«›"
            Index           =   11
            Visible         =   0   'False
         End
      End
      Begin VB.Menu TradingTransaction 
         Caption         =   "«·—’Ìœ «·«›  «ÕÌ"
         Index           =   0
         Shortcut        =   ^Q
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
            Caption         =   "”‰œ ’—› »÷«⁄Â"
            Index           =   0
         End
         Begin VB.Menu TradingTransactionSub1 
            Caption         =   "”‰œ ’—› Â«·ﬂ «Ê ⁄Ì‰« "
            Index           =   1
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
         Caption         =   "≈–‰ ’—› »÷«⁄…"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu TradingTransaction 
         Caption         =   "≈” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰›"
         Index           =   8
         Shortcut        =   ^S
      End
      Begin VB.Menu TradingTransaction 
         Caption         =   "»ÕÀ ⁄‰ »Ì«‰«  ”Ì—Ì«·"
         Index           =   9
         Shortcut        =   ^T
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
            Caption         =   "ÿ—ﬁ «·‘Õ‰"
            Index           =   3
         End
         Begin VB.Menu PurchaseBasic 
            Caption         =   "«‰Ê«⁄ «·÷„«‰« "
            Index           =   4
         End
         Begin VB.Menu PurchaseBasic 
            Caption         =   "«⁄œ««  «·«’‰«› «·—«ﬂœ…"
            Index           =   5
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
               Caption         =   "ÿ·» «„— ‘—«¡"
               Index           =   0
            End
            Begin VB.Menu PurchaseTransactionssubs1 
               Caption         =   "«⁄ „«œ «„— ‘—«¡"
               Index           =   1
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
      End
      Begin VB.Menu PurchaseTransactions 
         Caption         =   "›« Ê—… „‘ —Ì« "
         Index           =   3
         Shortcut        =   ^N
      End
      Begin VB.Menu PurchaseTransactions 
         Caption         =   "„—œÊœ«  «·„‘ —Ì« "
         Index           =   4
         Shortcut        =   ^O
      End
      Begin VB.Menu PurchaseTransactions 
         Caption         =   "  ﬁ«—Ì— «⁄„«— œÌÊ‰ «·„Ê—œÌ‰"
         Index           =   5
      End
      Begin VB.Menu PurchaseTransactions 
         Caption         =   " ﬁ«—Ì— «·„‘ —Ì«  Ê «·„Ê—œÌ‰"
         Index           =   6
      End
   End
   Begin VB.Menu MarketingMnu 
      Caption         =   "«· ”ÊÌﬁ"
      Begin VB.Menu MarketingMnusub 
         Caption         =   "ŒÿÂ „»Ì⁄«  «·«’‰«›"
         Index           =   0
      End
      Begin VB.Menu MarketingMnusub 
         Caption         =   "⁄—Ê÷ «·«’‰«›"
         Index           =   1
      End
      Begin VB.Menu MarketingMnusub 
         Caption         =   "„ «»⁄Â «·⁄„·«¡"
         Index           =   2
         Begin VB.Menu MarketingMnusubsub 
            Caption         =   " ”ÃÌ· “Ì«—«  «·⁄„·«¡"
            Index           =   0
         End
         Begin VB.Menu MarketingMnusubsub 
            Caption         =   "„ «»⁄Â “Ì«—«  «·⁄„·«¡"
            Index           =   1
         End
         Begin VB.Menu MarketingMnusubsub 
            Caption         =   "«” ÿ·«⁄ —√Ì «·⁄„·«¡"
            Index           =   2
         End
         Begin VB.Menu MarketingMnusubsub 
            Caption         =   " ”ÃÌ· ‘ﬂÊÏ «·⁄„·«¡"
            Index           =   3
         End
         Begin VB.Menu MarketingMnusubsub 
            Caption         =   "„ «»⁄Â ‘ﬂÊÏ «·⁄„·«¡"
            Index           =   4
         End
         Begin VB.Menu MarketingMnusubsub 
            Caption         =   "œ·Ì· «·Â« ›"
            Index           =   5
         End
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
            Caption         =   "»Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu SalesBasicSub 
            Caption         =   "»Ì«‰«  ﬂ«‘Ì—"
            Index           =   5
            Visible         =   0   'False
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
               Caption         =   "«⁄ „«œ ⁄—Ê÷ «·«”⁄«—"
               Index           =   1
            End
            Begin VB.Menu SalesTransactionssubss00 
               Caption         =   "⁄—Ê÷ «”⁄«— ‰Â«∆Ì… "
               Index           =   2
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
               Caption         =   "≈⁄ „«œ √„— »Ì⁄"
               Index           =   1
            End
            Begin VB.Menu SalesTransactionssubss000 
               Caption         =   "√„— »Ì⁄"
               Index           =   2
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
         Shortcut        =   ^M
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
         Visible         =   0   'False
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
            Caption         =   "“Ì«—«  «·⁄„·«¡"
            Index           =   4
            Visible         =   0   'False
         End
      End
      Begin VB.Menu SalesTransactions 
         Caption         =   " ﬁ—Ì— «⁄„«— œÌÊ‰ «·⁄„·«¡"
         Index           =   9
         Shortcut        =   ^P
      End
      Begin VB.Menu SalesTransactions 
         Caption         =   " ﬁ«—Ì— «·„»Ì⁄«  Ê«·⁄„·«¡"
         Index           =   10
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
         Caption         =   "„—œÊœ«  «·‘Õ‰"
         Index           =   4
      End
   End
   Begin VB.Menu POSTRansactiosG 
      Caption         =   "‰ﬁ«ÿ «·»Ì⁄"
      Begin VB.Menu POSTRansactios 
         Caption         =   "»Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄"
         Index           =   0
      End
      Begin VB.Menu POSTRansactios 
         Caption         =   "»Ì«‰«  «·‘Ì› "
         Index           =   1
      End
      Begin VB.Menu POSTRansactios 
         Caption         =   "»Ì«‰«  «·„Ê«ﬁ⁄"
         Index           =   2
      End
      Begin VB.Menu POSTRansactios 
         Caption         =   "»Ì«‰«  ﬂ«‘Ì—"
         Index           =   3
      End
      Begin VB.Menu POSTRansactios 
         Caption         =   " ”ÃÌ· «·œŒÊ·"
         Index           =   4
      End
      Begin VB.Menu POSTRansactios 
         Caption         =   " ﬁ«—Ì— ‰ﬁ«ÿ «·»Ì⁄"
         Index           =   5
      End
   End
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
         Caption         =   "«’œ«— «·ﬁÌÊœ «· ﬂ—«—Ì…"
         Index           =   1
         Visible         =   0   'False
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
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu ExpensesType 
         Caption         =   "√‰Ê«⁄ «·≈Ì—«œ« "
         Index           =   1
      End
      Begin VB.Menu MnuFinSep1 
         Caption         =   "-"
      End
      Begin VB.Menu Expenses 
         Caption         =   "›« Ê—… „«·Ì…"
         Index           =   0
      End
      Begin VB.Menu Expenses 
         Caption         =   "”‰œ«  «·’—›"
         Index           =   1
         Begin VB.Menu ExpensesSub 
            Caption         =   "”‰œ«  «·’—› -  Õ·Ì·Ì „’—Ê›« "
            Index           =   0
         End
         Begin VB.Menu ExpensesSub 
            Caption         =   "”‰œ«  «·’—› - «·„œ›Ê⁄« "
            Index           =   1
         End
      End
      Begin VB.Menu Payments 
         Caption         =   "«·„œ›Ê⁄« "
         Index           =   0
         Shortcut        =   ^{F3}
         Visible         =   0   'False
      End
      Begin VB.Menu Cashing 
         Caption         =   "«·„ﬁ»Ê÷« "
         Index           =   0
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu Cashing 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Cashing 
         Caption         =   "ÿ»«⁄… «·‘Ìﬂ« "
         Index           =   2
      End
      Begin VB.Menu Cashing 
         Caption         =   "«Ìœ«⁄«  »‰ﬂÌÂ"
         Index           =   3
      End
      Begin VB.Menu Cashing 
         Caption         =   " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« "
         Index           =   4
      End
      Begin VB.Menu Cashing 
         Caption         =   "„–ﬂ—… »‰ﬂ"
         Index           =   5
      End
      Begin VB.Menu DelayVal 
         Caption         =   "«·√Ê—«ﬁ «·„«·Ì… «·„” Õﬁ…"
         Index           =   0
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu MnuFinSep6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFinDiscounts 
         Caption         =   "«·Œ’Ê„«  «·„”„ÊÕ… Ê«·„ﬂ ”»…"
      End
      Begin VB.Menu MnuFinSep3 
         Caption         =   "-"
      End
      Begin VB.Menu ReceiptPart 
         Caption         =   " Õ’Ì· Ê”œ«œ √ﬁ”«ÿ"
         Shortcut        =   ^{F6}
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
      Begin VB.Menu MnuFinSep4 
         Caption         =   "-"
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
      Begin VB.Menu MnuFinSep7 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuBoxAccouns 
         Caption         =   "—’Ìœ «·Œ“‰… «·√‰..."
      End
      Begin VB.Menu MnuBoxStock 
         Caption         =   "Ã—œ «·Œ“‰…"
      End
      Begin VB.Menu MnuBoxIncapacity_Increase 
         Caption         =   "“Ì«œ… Ê⁄Ã“ ›Ï ‰ﬁœÌ… «·Œ“‰…"
      End
      Begin VB.Menu MnuFinSep5 
         Caption         =   "-"
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
         Caption         =   "≈Õ’«∆Ì« "
         Index           =   7
      End
      Begin VB.Menu xxy 
         Caption         =   "√Ã‰œ… «·⁄„·«¡"
         Index           =   8
      End
      Begin VB.Menu xxy 
         Caption         =   " ﬁ—Ì—"
         Index           =   9
         Visible         =   0   'False
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
         Caption         =   " ”·Ì„ Ê ”·„ «·«’Ê·"
         Index           =   6
      End
      Begin VB.Menu xxxxx 
         Caption         =   "«· ﬁ«—Ì—"
         Index           =   7
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
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuEmployeeBasicSub 
            Caption         =   "»Ì«‰«   Œ’’«  «·⁄„· ›Ï «·‘—ﬂ…"
            Index           =   7
         End
         Begin VB.Menu mnuEmployeeBasicSub 
            Caption         =   "‘—ﬂ«  «· √„Ì‰"
            Index           =   8
         End
         Begin VB.Menu mnuEmployeeBasicSub 
            Caption         =   "«‰Ê«⁄ «· √„Ì‰"
            Index           =   9
         End
         Begin VB.Menu mnuEmployeeBasicSub 
            Caption         =   "›∆«  «· √„Ì‰"
            Index           =   10
         End
         Begin VB.Menu mnuEmployeeBasicSub 
            Caption         =   "⁄‰«’— «· ﬁÌÌ„"
            Index           =   11
         End
      End
      Begin VB.Menu mnuEmployeeBasic 
         Caption         =   "»Ì«‰«  «·„ÊŸ›Ì‰"
         Index           =   1
         Begin VB.Menu EmployeeDataicSub 
            Caption         =   "„·› «·„ÊŸ›Ì‰"
            Index           =   0
            Shortcut        =   ^B
         End
         Begin VB.Menu EmployeeDataicSub 
            Caption         =   "⁄ﬁÊœ «·„ÊŸ›Ì‰"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEmployeeBasic 
         Caption         =   "«·Õ÷Ê— Ê«·«‰’—«›"
         Index           =   2
         Begin VB.Menu EmployeeAttendanceSub 
            Caption         =   "  ≈⁄œ«œ „Ê«⁄Ìœ «·Õ÷Ê— Ê«·√‰’—«› ·„ÊŸ›"
            Index           =   0
         End
         Begin VB.Menu EmployeeAttendanceSub 
            Caption         =   " ”ÃÌ·  „Ê«⁄Ìœ «·Õ÷Ê— Ê «·«‰’—«› ÌœÊÌ«"
            Index           =   1
         End
         Begin VB.Menu EmployeeAttendanceSub 
            Caption         =   " ”ÃÌ· „Ê«⁄Ìœ «·Õ÷Ê— Ê «·«‰’—«›  «·Ì«"
            Index           =   2
         End
         Begin VB.Menu EmployeeAttendanceSub 
            Caption         =   " ”ÃÌ· «·€Ì«»"
            Index           =   3
         End
         Begin VB.Menu EmployeeAttendanceSub 
            Caption         =   "«·⁄—÷ «·⁄«„ ·„Ê«⁄Ìœ «·Õ÷Ê— Ê«·√‰’—«›"
            Index           =   4
         End
      End
      Begin VB.Menu mnuEmployeeBasic 
         Caption         =   "«·—Ê« »"
         Index           =   3
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
         End
         Begin VB.Menu EmployeeSalarySub 
            Caption         =   "—œ ”·›… „ÊŸ›"
            Index           =   5
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
         End
         Begin VB.Menu EmployeeSalarySub 
            Caption         =   " ”ÃÌ·  —ﬂ «·Œœ„…"
            Index           =   11
            Visible         =   0   'False
         End
         Begin VB.Menu EmployeeSalarySub 
            Caption         =   " €ÌÌ— „Ì⁄«œ ”·›…"
            Index           =   12
         End
      End
      Begin VB.Menu mnuEmployeeBasic 
         Caption         =   "«Ã«“«  «·„ÊŸ›Ì‰"
         Index           =   4
         Begin VB.Menu Vscstionsssub 
            Caption         =   "Œÿ… «·«Ã«“« "
            Index           =   0
         End
         Begin VB.Menu Vscstionsssub 
            Caption         =   "ÿ·» «Ã«“…"
            Index           =   1
         End
         Begin VB.Menu Vscstionsssub 
            Caption         =   " ”·Ì„ Ê≈” ·«„ ⁄Âœ ⁄Ì‰Ì…"
            Index           =   2
         End
         Begin VB.Menu Vscstionsssub 
            Caption         =   "„” Õﬁ«  «·«Ã«“…"
            Index           =   3
         End
         Begin VB.Menu Vscstionsssub 
            Caption         =   "‰”ÃÌ· «·Õ÷Ê— „‰ «Ã«“…"
            Index           =   4
         End
      End
      Begin VB.Menu mnuEmployeeBasic 
         Caption         =   "«‰Â«¡ «·Œœ„…"
         Index           =   5
         Begin VB.Menu FinishSevicersub 
            Caption         =   " ”ÃÌ·  —ﬂ «·Œœ„…"
            Index           =   0
         End
         Begin VB.Menu FinishSevicersub 
            Caption         =   "Õ”«» „ﬂ«›√… ‰Â«Ì… «·Œœ„…"
            Index           =   1
         End
      End
   End
   Begin VB.Menu Archiving 
      Caption         =   "«·«—‘Ì› «·«·ﬂ —Ê‰Ì"
      Visible         =   0   'False
      Begin VB.Menu ArchivingSub 
         Caption         =   "«÷«›… ‰„«–Ã ÃœÌœ…"
         Index           =   0
      End
      Begin VB.Menu m2 
         Caption         =   "„ «»⁄Â «·„œ«—”"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu ArrowsBase 
      Caption         =   "„ «»⁄Â «·«”Â„"
      Visible         =   0   'False
      Begin VB.Menu ArrowsFollow 
         Caption         =   "»Ì«‰«  «·»Ê—’« "
         Index           =   0
      End
      Begin VB.Menu ArrowsFollow 
         Caption         =   "»Ì«‰«  „Ã„Ê⁄«  «·«”Â„"
         Index           =   1
      End
      Begin VB.Menu ArrowsFollow 
         Caption         =   "»Ì«‰«  «·‘—ﬂ« "
         Index           =   2
      End
      Begin VB.Menu ArrowsFollow 
         Caption         =   " Õ„Ì· «·«”⁄«—              "
         Index           =   3
      End
      Begin VB.Menu ArrowsFollow 
         Caption         =   "«·«”⁄«— «· «—ÌŒÌ…"
         Index           =   4
      End
      Begin VB.Menu ArrowsFollow 
         Caption         =   "«·„Õ€Ÿ…"
         Index           =   5
         Begin VB.Menu ArrowsFollowBocket 
            Caption         =   "»Ì«‰«  «·„Õ«›Ÿ «·„„·ÊﬂÂ"
            Index           =   0
         End
         Begin VB.Menu ArrowsFollowBocket 
            Caption         =   "‘—«¡ «”Â„"
            Index           =   1
         End
         Begin VB.Menu ArrowsFollowBocket 
            Caption         =   "»Ì⁄ «”Â„"
            Index           =   2
         End
         Begin VB.Menu ArrowsFollowBocket 
            Caption         =   "«·ﬁÌ„… «·”ÊﬁÌ… ·Ã„Ì⁄ «·«”Â„ «·„„·ÊﬂÂ"
            Index           =   3
         End
      End
      Begin VB.Menu ArrowsFollow 
         Caption         =   "„Ê«ﬁ⁄ „Â„Â…"
         Index           =   6
      End
      Begin VB.Menu ArrowsFollow 
         Caption         =   "«· ﬁ«—Ì—"
         Index           =   7
      End
   End
   Begin VB.Menu AssetsMngBase 
      Caption         =   "«œ«—… «·«„·«ﬂ"
      Visible         =   0   'False
      Begin VB.Menu AssetsMng 
         Caption         =   "„·›«  «”«”Ì…       "
         Index           =   0
         Begin VB.Menu AssetsMngBasicFiles 
            Caption         =   " ⁄—Ì› «·⁄ﬁ«—« "
            Index           =   0
            Begin VB.Menu AssetsMngBasicFilesR 
               Caption         =   "⁄ﬁ«—« "
               Index           =   0
               Begin VB.Menu AssetsMngBasicFiles2 
                  Caption         =   "⁄„«∆—"
                  Index           =   0
               End
               Begin VB.Menu AssetsMngBasicFiles2 
                  Caption         =   "‘ﬁﬁ"
                  Index           =   1
               End
               Begin VB.Menu AssetsMngBasicFiles2 
                  Caption         =   "€—›"
                  Index           =   2
               End
               Begin VB.Menu AssetsMngBasicFiles2 
                  Caption         =   "„Õ·« "
                  Index           =   3
               End
            End
            Begin VB.Menu AssetsMngBasicFilesR 
               Caption         =   "›··"
               Index           =   1
            End
            Begin VB.Menu AssetsMngBasicFilesR 
               Caption         =   "«—«÷Ì"
               Index           =   2
            End
            Begin VB.Menu AssetsMngBasicFilesR 
               Caption         =   "«·„” Êœ⁄« "
               Index           =   3
            End
            Begin VB.Menu AssetsMngBasicFilesR 
               Caption         =   "«·Ê—‘"
               Index           =   4
            End
            Begin VB.Menu AssetsMngBasicFilesR 
               Caption         =   "«·„—«ﬂ“ «· Ã«—ÌÂ"
               Index           =   5
            End
         End
         Begin VB.Menu AssetsMngBasicFiles 
            Caption         =   " ⁄—Ì›  «·„Œÿÿ« "
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu AssetsMngBasicFiles 
            Caption         =   " ⁄—Ì› «·„·«ﬂ"
            Index           =   3
         End
         Begin VB.Menu AssetsMngBasicFiles 
            Caption         =   "  ⁄—Ì›  «·„” √Ã—Ì‰ Ê«·„‘ —Ì‰"
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
            Caption         =   " ⁄—Ì›  «·‘Ê«—⁄"
            Index           =   8
         End
         Begin VB.Menu AssetsMngBasicFiles 
            Caption         =   "œ·Ì· «·Â« ›"
            Index           =   9
         End
      End
      Begin VB.Menu AssetsMng 
         Caption         =   "«·Õ—ﬂ« "
         Index           =   1
         Begin VB.Menu AssetsMngTrans 
            Caption         =   " ”ÃÌ· ÿ·»«  «·»Ì⁄ Ê «·‘—«¡ Ê «·«ÌÃ«—"
            Index           =   0
         End
         Begin VB.Menu AssetsMngTrans 
            Caption         =   " ”ÃÌ·  ⁄—Ê÷   «·»Ì⁄ Ê «·‘—«¡  Ê «·«ÌÃ«—"
            Index           =   1
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
            Caption         =   "„ «»⁄Â «·ÌÌ⁄ »«· ﬁ”Ìÿ"
            Index           =   7
         End
         Begin VB.Menu AssetsMngTrans 
            Caption         =   " Õ’Ì· «ÌÃ«—« "
            Index           =   8
         End
         Begin VB.Menu AssetsMngTrans 
            Caption         =   "«·ﬁ«∆„Â «·”Êœ«¡"
            Index           =   9
         End
         Begin VB.Menu AssetsMngTrans 
            Caption         =   "«’œ«— «‘⁄«—  ”œÌœ - «‰–«—"
            Index           =   10
         End
         Begin VB.Menu AssetsMngTrans 
            Caption         =   "«·’Ì«‰Â"
            Index           =   11
            Visible         =   0   'False
            Begin VB.Menu estateMain 
               Caption         =   "’Ì«‰Â ⁄ﬁ«—"
               Index           =   0
            End
            Begin VB.Menu estateMain 
               Caption         =   "’Ì«‰Â ÊÕœÂ"
               Index           =   1
            End
            Begin VB.Menu estateMain 
               Caption         =   "√Ê«„— «·‘€·"
               Index           =   2
            End
         End
      End
      Begin VB.Menu AssetsMng 
         Caption         =   "«· ﬁ«—Ì—"
         Index           =   2
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â ⁄—Ê÷ «·«ÌÃ«—"
            Index           =   0
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â ⁄—Ê÷ «·‘—«¡ Ê«·»Ì⁄"
            Index           =   1
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â ÿ·»«  «·«ÌÃ«—"
            Index           =   2
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â ÿ·»«  «·‘—«¡ Ê«·»Ì⁄"
            Index           =   3
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â «·⁄„·Ì«  «· Ì  „  ⁄·Ï ÊÕœÂ «Ê ⁄ﬁ«—"
            Index           =   4
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â ⁄ﬁœ «ÌÃ«— ÊÕœÂ «Ê ⁄ﬁ«—"
            Index           =   5
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â ⁄ﬁœ »Ì⁄ ÊÕœÂ «Ê ⁄ﬁ«—"
            Index           =   6
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â »Ì«‰«  «· ﬁ”Ìÿ ··«ÌÃ«— Ê«·»Ì⁄"
            Index           =   7
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   " ﬁ—Ì— «·’Ì«‰… ·ÊÕœÂ «Ê ⁄ﬁ«—"
            Index           =   8
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â «Ê«„— «·‘€·"
            Index           =   9
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â ﬂ‘› «·«ÌÃ«—«  «·„ √Œ—Â"
            Index           =   10
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â  ﬁ—Ì— ⁄«„ ·„”‹«Ã—Ï «·›··"
            Index           =   11
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â  ﬁ—Ì— ⁄«„ ·„”‹«Ã—Ï «·«—«÷Ì"
            Index           =   12
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â  ﬁ—Ì— ⁄«„ ·„”‹«Ã—Ï «·„—«ﬂ“ «· Ã«—Ì…"
            Index           =   13
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â  ﬁ—Ì— ⁄«„ ·„”‹«Ã—Ï «·Ê—‘"
            Index           =   14
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â  ﬁ—Ì— ⁄«„ ·„”‹«Ã—Ï «·„” Êœ⁄« "
            Index           =   15
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â  ﬁ—Ì— ⁄«„ ·„”‹«Ã—Ï «·⁄„«∆—"
            Index           =   16
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â  ﬁ—Ì— ⁄«„ ·„”‹«Ã—Ï «·‘ﬁﬁ"
            Index           =   17
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â  ﬁ—Ì— ⁄«„ ·„”‹«Ã—Ï «·€—›"
            Index           =   18
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â  ﬁ—Ì— ⁄«„ ·„”‹«Ã—Ï «·„Õ·« "
            Index           =   19
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â ”‰œ«  «·’—›"
            Index           =   20
         End
         Begin VB.Menu AssetsMngReport 
            Caption         =   "ÿ»«⁄Â ”‰œ«  «·ﬁ»÷"
            Index           =   21
         End
      End
      Begin VB.Menu AssetsMng 
         Caption         =   "—”«∆· ··⁄„·«¡"
         Index           =   3
      End
   End
   Begin VB.Menu Reports 
      Caption         =   "«· ﬁ«—Ì—"
      Begin VB.Menu Report 
         Caption         =   "«· ﬁ«—Ì— «·⁄«„…"
         Shortcut        =   ^U
      End
      Begin VB.Menu sss 
         Caption         =   "-"
      End
      Begin VB.Menu DailyReport 
         Caption         =   "«· ﬁ—Ì— «·ÌÊ„Ì"
         Shortcut        =   ^Y
      End
      Begin VB.Menu MnuReports_Assblied 
         Caption         =   "«· ﬁ—Ì— «·„Ã„⁄ ⁄‰ › —…"
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "„œÌ— «·‰Ÿ«„"
      Begin VB.Menu Options 
         Caption         =   "«⁄œ«œ«  «·‰Ÿ«„"
      End
      Begin VB.Menu MNUCloseYear 
         Caption         =   "⁄„·ÌÂ «·«ﬁ›«·"
         Visible         =   0   'False
      End
      Begin VB.Menu UsersData 
         Caption         =   "„” Œœ„Ì «·‰Ÿ«„"
         Begin VB.Menu AddUser 
            Caption         =   "≈÷«›… „” Œœ„..."
         End
         Begin VB.Menu DelUser 
            Caption         =   "Õ–› „” Œœ„..."
         End
         Begin VB.Menu EditPw 
            Caption         =   " ⁄œÌ· ﬂ·„… «·„—Ê—..."
         End
         Begin VB.Menu Sep7 
            Caption         =   "-"
         End
         Begin VB.Menu MnuLevels 
            Caption         =   "«⁄ „«œ «·„” ‰œ« "
            Begin VB.Menu MnuLevelsSub 
               Caption         =   " ⁄—Ì› „” ÊÌ«  «·«⁄ „«œ"
               Index           =   0
            End
            Begin VB.Menu MnuLevelsSub 
               Caption         =   " ⁄—Ì› «⁄ „«œ«  «·„” œ« "
               Index           =   1
            End
         End
         Begin VB.Menu UserAbility 
            Caption         =   "’·«ÕÌ«  «·„” Œœ„Ì‰"
         End
         Begin VB.Menu MnuUsersScreensPremission 
            Caption         =   "’·«ÕÌ… «·„” Œœ„Ì‰ ⁄·Ï «·‘«‘« "
         End
         Begin VB.Menu UserRpt 
            Caption         =   " ﬁ«—Ì— «·„” Œœ„Ì‰"
         End
      End
      Begin VB.Menu ShortCuts 
         Caption         =   "„›« ÌÕ «·«Œ ’«—"
      End
      Begin VB.Menu Sep30 
         Caption         =   "-"
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
   End
   Begin VB.Menu MnuWindowsList 
      Caption         =   "‘«‘«  «·»—‰«„Ã"
      Visible         =   0   'False
   End
   Begin VB.Menu MnuWindowsListOpen 
      Caption         =   "«·‰Ê«›– «·„› ÊÕ…"
      Visible         =   0   'False
   End
   Begin VB.Menu Tech 
      Caption         =   "«·«œÊ«  «·›‰Ì…"
      Begin VB.Menu MnuToolsSetPrinters0 
         Caption         =   "≈⁄œ«œ «·ÿ«»⁄… ›Ï «·ÃÂ«“ «·Õ«·Ì"
      End
      Begin VB.Menu Barcode 
         Caption         =   " ’„Ì„ «·»«—ﬂÊœ"
         Shortcut        =   ^W
      End
      Begin VB.Menu MnuPrintItemsCodes 
         Caption         =   "ÿ»«⁄… »«—ﬂÊœ  ·√ﬂÊ«œ «·√’‰«›"
      End
      Begin VB.Menu MnuToolsSetPrinters7 
         Caption         =   " ≈⁄œ«œ«  —”«∆· «·ÃÊ«·"
         Begin VB.Menu Texh 
            Caption         =   "≈⁄œ«œ«  ›‰Ì…"
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
      End
      Begin VB.Menu MnuCorrectSerial 
         Caption         =   "«·ﬂ‘› ⁄‰ √Œÿ«¡ «·”Ì—Ì«· ··√’‰«›"
      End
      Begin VB.Menu MnuBoxDetectErrors 
         Caption         =   "«·ﬂ‘› ⁄‰ √Œÿ«¡ ﬂ‘› Õ”«» «·Œ“‰…"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuToolCustomers 
         Caption         =   "Ÿ»ÿ ›Ê« Ì— «·⁄„·«¡"
      End
      Begin VB.Menu MnuToolRepaireItemsCost 
         Caption         =   "⁄—÷ „ Ê”ÿ «· ﬂ·›… ··√’‰«› ›Ï ›Ê« Ì— «·»Ì⁄"
      End
      Begin VB.Menu MnuToolsDataBase 
         Caption         =   " ‰‘Ìÿ «·√ ’«· »ﬁ«⁄œ… «·»Ì«‰« "
         Index           =   0
      End
      Begin VB.Menu MnuToolsDataBase 
         Caption         =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Index           =   1
      End
      Begin VB.Menu MnuDataBaseTools 
         Caption         =   "√œÊ«  ﬁ«⁄œ… «·»Ì«‰« "
      End
   End
   Begin VB.Menu Help 
      Caption         =   "„”«⁄œ…"
      Begin VB.Menu HelpFile 
         Caption         =   "„·›«  «·„”«⁄œ…"
      End
      Begin VB.Menu HelpIndex 
         Caption         =   "›Â—” „·›«  «·„”«⁄œ…"
      End
      Begin VB.Menu SearchInHelp 
         Caption         =   "«·»ÕÀ ›Ì „·›«  «·„”«⁄œ…"
      End
      Begin VB.Menu DailyToolTip 
         Caption         =   "«· ·„ÌÕ «·ÌÊ„Ì"
      End
      Begin VB.Menu Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuHelpForums 
         Caption         =   "„‰ œÌ«  «·œ⁄„ «·›‰Ì"
      End
      Begin VB.Menu About 
         Caption         =   "⁄‰ «·»—‰«„Ã..."
      End
      Begin VB.Menu ConnectUs 
         Caption         =   " ”ÃÌ· «·»—‰«„Ã..."
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
    frmabout.show vbModal
End Sub

Private Sub AddItem_Click()
    FrmMainPriceList.XPBtnAdd_Click
End Sub

Private Sub AddUser_Click()
    Dim Msg As String

    If user_id <> 1 Then
        '   MsgBox ""
        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
        '    Msg = Msg & Chr(13) & "Õ ‰Â“— Ê·««ÌÂ "
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    'If user_id <> 1 Then
    '    Msg = "€Ì— „”„ÊÕ ·ﬂ »«· ⁄«„· „⁄ Â–Â «·‰«›–…"
    '    MsgBox Msg, vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "’·«ÕÌ«  «·„” Œœ„Ì‰"
    '    Exit Sub
    'End If

    If checkApility("FrmAddUser") = False Then
        Exit Sub
    End If

    FrmAddUser.show vbModal
End Sub

Private Sub Asset_Click(Index As Integer)
End Sub

Private Sub ArchivingSub_Click(Index As Integer)

    Select Case Index

        Case 0
            loading_temolates.show

    End Select

End Sub

Private Sub ArrowsFollow_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("ArrowsFinancialMarkets") = False Then
                Exit Sub
            End If

            ArrowsFinancialMarkets.show

        Case 1

            If checkApility("ArrowsGroup") = False Then
                Exit Sub
            End If

            ArrowsGroup.show

        Case 2

            If checkApility("ArrowsAllCompanyilstDetails1") = False Then
                Exit Sub
            End If

            ArrowsAllCompanyilstDetails1.show

        Case 3

            If checkApility("Arrows") = False Then
                Exit Sub
            End If

            Arrows.show

        Case 4

            If checkApility("ArrowsHistory") = False Then
                Exit Sub
            End If

            ArrowsHistory.show
            'ArrowsAllCompanyilstDetails.Show

    End Select

End Sub

Private Sub ArrowsFollowa_Click(Index As Integer)

    Select Case Index

        Case 0
            ArrowsAccount.show
    End Select

End Sub

Private Sub ArrowsFollowBocket_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("ArrowsAccount") = False Then
                Exit Sub
            End If

            ArrowsAccount.show

        Case 1

            If checkApility("ArrowsPurchase") = False Then
                Exit Sub
            End If

            ArrowsPurchase.show

        Case 2

            'ArrowsSale.Show
            If checkApility("ArrowsSale1") = False Then
                Exit Sub
            End If

            ArrowsSale1.show

        Case 3

            If checkApility("ArrowsCurrentValue") = False Then
                Exit Sub
            End If

            ArrowsCurrentValue.show
    End Select

End Sub

Private Sub AssetsMng_Click(Index As Integer)

    Select Case Index

        Case 3

            If checkApility("messages_frm") = False Then
                Exit Sub
            End If

            messages_frm.show
    End Select

End Sub

Private Sub AssetsMngBasicFiles_Click(Index As Integer)

    Select Case Index

        Case 3

            If checkApility("RSOwner") = False Then
                Exit Sub
            End If

            RSOwner.show

        Case 4

            If checkApility("RsCustomers") = False Then
                Exit Sub
            End If

            RsCustomers.show

        Case 5

            If checkApility("FrmCountriesData1") = False Then
                Exit Sub
            End If

            FrmCountriesData.show

        Case 6

            If checkApility("FrmGovernmentData1") = False Then
                Exit Sub
            End If

            FrmGovernmentData.show

        Case 7

            If checkApility("FrmGovernCitiesData1") = False Then
                Exit Sub
            End If

            FrmGovernCitiesData.show
 
        Case 8

            If checkApility("streets1") = False Then
                Exit Sub
            End If

            streets.show

        Case 9

            If checkApility("RSPhoneBook") = False Then
                Exit Sub
            End If

            RSPhoneBook.show
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

            RsApartement.show

        Case 2

            If checkApility("RsRoom") = False Then
                Exit Sub
            End If

            RsRoom.show

        Case 3

            If checkApility("RsStore") = False Then
                Exit Sub
            End If

            RsStore.show

    End Select

End Sub

Private Sub AssetsMngBasicFilesR_Click(Index As Integer)

    Select Case Index

        Case 1
            RsVila.show

        Case 2
            RSland.show

        Case 3
            RsStores.show

        Case 4
            RSWorkShop.show

        Case 5
            RSTradingCenter.show

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

            If checkApility("RSContract1") = False Then
                Exit Sub
            End If

            RSContract.show
    End Select

End Sub

Private Sub AssetsMngTrans_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("RsOrders") = False Then
                Exit Sub
            End If

            RsOrders.show

        Case 1

            If checkApility("RsOrders1") = False Then
                Exit Sub
            End If

            RsOrders.show

        Case 5

            If checkApility("RsCashing") = False Then
                Exit Sub
            End If

            RsCashing.show

        Case 6

            If checkApility("RsExpenses") = False Then
                Exit Sub
            End If

            RsExpenses.show

        Case 7

            If checkApility("RSContractInstallments") = False Then
                Exit Sub
            End If

            RSContractInstallments.show

        Case 8

            If checkApility("RsPayemntReport") = False Then
                Exit Sub
            End If

            RsPayemntReport.show

        Case 9

            If checkApility("black_list") = False Then
                Exit Sub
            End If

            black_list.show

        Case 10

            If checkApility("RsCustomerAlarm") = False Then
                Exit Sub
            End If

            RsCustomerAlarm.show
    End Select

End Sub

Private Sub balancsheet_Click(Index As Integer)

    Select Case Index

        Case 0
            BaklanceSheet.show

        Case 1
            BaklanceSheetvIEW.show
    End Select

    'FrmAccountingReport1.Show

End Sub

Private Sub BankAdM_Click()

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

    Select Case Index

        Case 0

            If checkApility("baranches") = False Then
                Exit Sub
            Else
                baranches.show
            End If
 
        Case 1

            If checkApility("FrmBranchesData") = False Then
                Exit Sub
            End If
             
            FrmBranchesData.show

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

            FrmPaymentType.show

        Case 5

            If checkApility("FrmCompany") = False Then
                Exit Sub
            End If

            FrmCompany.show

        Case 6

            If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If

            OpenScreen CustomersScreen '

        Case 7

            If checkApility("FRMcurrency") = False Then
                Exit Sub
            End If

            FRMcurrency.show

        Case 8

            If checkApility("nationality") = False Then
                Exit Sub
            End If

            nationality.show

        Case 9

            If checkApility("dean") = False Then
                Exit Sub
            End If

            dean.show
 
        Case 10

            If checkApility("FrmCountriesData") = False Then
                Exit Sub
            End If

            FrmCountriesData.show

        Case 11

            If checkApility("FrmGovernmentData") = False Then
                Exit Sub
            End If

            FrmGovernmentData.show

        Case 12

            If checkApility("FrmGovernCitiesData") = False Then
                Exit Sub
            End If

            FrmGovernCitiesData.show

        Case 13

            If checkApility("streets") = False Then
                Exit Sub
            End If

            streets.show
 
        Case 14
            ' FrmDocType.Show

        Case 15

            If checkApility("FrmItems") = False Then
                Exit Sub
            End If

            OpenScreen ItemsDataScreen

        Case 17
            AskForExit

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

        Case 1

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

            If checkApility("FrmChiqueRelease") = False Then
                Exit Sub
            End If

            'FrmChiqueRelease.Show

            FrmBankDeposite1.show

        Case 5

            If checkApility("FrmBankAdj") = False Then
                Exit Sub
            End If

            FrmBankAdj.show

    End Select

End Sub

Private Sub ComingTimes_Click()
    Dim Frm As FrmTimeSetting

    If checkApility("FrmTimeSetting") = False Then
        Exit Sub
    End If

    Set Frm = New FrmTimeSetting
    Frm.WorkType = 0
    Frm.show
    Frm.ZOrder 0

End Sub

Private Sub ConnectUs_Click()
    'FrmConect_US.Show
    'FrmConect_US.ZOrder 0
    Dim Msg As String

    If SystemOptions.SysRegisterState = DemoRun Or SystemOptions.SysRegisterState = DemoStop Then
        FrmRegisteration.show vbModal
    Else
        Msg = "‰”Œ… „”Ã·… "
        Msg = Msg & Chr(13) & "‘ﬂ—« .. .·≈” Œœ«„ﬂ„ »—‰«„Ã ‰Ÿ«„ œÌ‰«„Ìﬂ »«Ì "
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If

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
    FrmDailyToolTip.show
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
    FrmMainPriceList.XPBtnRemove_Click
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
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If checkApility("FrmDelUser") = False Then
        Exit Sub
    End If

    FrmDelUser.show vbModal
End Sub

Private Sub Destruction_Click()
    OpenScreen DestructionScreen
End Sub

Private Sub DockingPane1_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, _
                                ByVal Pane As XtremeDockingPane.IPane, _
                                ByVal Container As XtremeDockingPane.IPaneActionContainer, _
                                Cancel As Boolean)
  
    Dim Frm As Form
    Dim i  As Integer
    Dim Msg As String

    On Error GoTo hErr

    If Pane.id = DockingPanesIDs.NewsBarPaneID Then
        If Not FrmNewsBarPane Is Nothing Then
            If Action = PaneActionClosed Then
                FrmNewsBarPane.TimerData.Enabled = False
            ElseIf Action = PaneActionCollapsed Then
                FrmNewsBarPane.TimerData.Enabled = False
            ElseIf Action = PaneActionCollapsing Then
                FrmNewsBarPane.TimerData.Enabled = False
            ElseIf Action = PaneActionExpanding Then
                FrmNewsBarPane.TimerData.Enabled = True
            ElseIf Action = PaneActionExpanded Then
                FrmNewsBarPane.TimerData.Enabled = True
            End If
        End If

    ElseIf Pane.id = DockingPanesIDs.MantainceID Then

        If Not FrmMantaincePane Is Nothing Then
            If Action = PaneActionExpanded Or Action = PaneActionExpanding Then
                FrmMantaincePane.SetDcboSearch
            End If
        End If
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
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub DockingPane1_AttachPane(ByVal Item As XtremeDockingPane.IPane)

    If Not Item Is Nothing Then
        If Item.id = DockingPanesIDs.NewsBarPaneID Then
            Set FrmNewsBarPane = New FrmPane
            FrmNewsBarPane.PanelType = 1
            Item.Handle = FrmNewsBarPane.hWnd
            FrmNewsBarPane.backcolor = &HE2E9E9
        ElseIf Item.id = DockingPanesIDs.OutBarPaneID Then
            Set FrmOutBarPane = New FrmOurBarPane
            Item.Handle = FrmOutBarPane.hWnd
            FrmOutBarPane.backcolor = &HE2E9E9
        ElseIf Item.id = DockingPanesIDs.ItemsTreeID Then
            Set ItemsTreePane = New FrmPaneTree
            Item.Handle = ItemsTreePane.hWnd
            ItemsTreePane.backcolor = &HE2E9E9
        ElseIf Item.id = DockingPanesIDs.MantainceID Then
            Set FrmMantaincePane = New FrmPane
            FrmMantaincePane.PanelType = 3
            Item.Handle = FrmMantaincePane.hWnd
            FrmMantaincePane.backcolor = &HE2E9E9
        ElseIf Item.id = DockingPanesIDs.InternetNews Then
            Set FrmInternetNews = New FrmPane
            FrmInternetNews.PanelType = 2
            Item.Handle = FrmInternetNews.hWnd
            FrmInternetNews.backcolor = &HE2E9E9
        ElseIf Item.id = DockingPanesIDs.DynamicHelp Then
            Set FrmDynamicHelpPane = New FrmPaneHelp
            Item.Handle = FrmDynamicHelpPane.hWnd
            FrmDynamicHelpPane.backcolor = &HE2E9E9
        ElseIf Item.id = DockingPanesIDs.CalendarPaneID Then
            Set FrmCalendarPane = New FrmPaneCalendar
            Item.Handle = FrmCalendarPane.hWnd 'salim found
            FrmCalendarPane.backcolor = &HE2E9E9
        End If
    End If

End Sub

Private Sub DockingPane1_PanePopupMenu(ByVal Pane As XtremeDockingPane.IPane, _
                                       ByVal x As Long, _
                                       ByVal Y As Long, _
                                       Handled As Boolean)

    Select Case Pane.id

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

Private Sub EmployeeAttendanceSub_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("FrmTimeSetting1") = False Then
                Exit Sub
            End If

            Dim Frm As New FrmTimeSetting
            Frm.WorkType = 1
            Frm.show
            Frm.ZOrder 0

        Case 1

            If checkApility("FrmPresentTime") = False Then
                Exit Sub
            End If

            FrmPresentTime.show
            FrmPresentTime.ZOrder 0
 
        Case 2

            If checkApility("FrmEmpSalary2") = False Then
                Exit Sub
            End If

            FrmEmpSalary2.show

        Case 3

            If checkApility("FrmAbsent") = False Then
                Exit Sub
            End If

            FrmAbsent.show
            FrmAbsent.ZOrder 0

        Case 4

            If checkApility("FrmEmpMonthShow") = False Then
                Exit Sub
            End If

            FrmEmpMonthShow.show
    End Select

End Sub

Private Sub EmployeeDataicSub_Click(Index As Integer)

    Select Case Index

        Case 0

            'FrmEmployee
            If checkApility("FrmEmployee") = False Then
                Exit Sub
            End If

            OpenScreen EmployeesScreen

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

            If checkApility("FrmEmpsAdvancePayed") = False Then
                Exit Sub
            End If

            FrmEmpsAdvancePayed.show

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

            If checkApility("FrmChangedComponentData3") = False Then
                Exit Sub
            End If

            FrmChangedComponentData3.show

        Case 11

        Case 12

            If checkApility("FrmEmpsAdvancePayed1") = False Then
                Exit Sub
            End If

            FrmEmpsAdvancePayed1.show

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

    End Select

End Sub

Private Sub ExpensesSub_Click(Index As Integer)

    Select Case Index

        Case 0

            '           OpenScreen ExpensesDataScreen
            If checkApility("FrmExpenses5") = False Then
                Exit Sub
            End If

            FrmExpenses5.show

        Case 1

            'FrmPayments.Show
            If checkApility("FrmPayments") = False Then
                Exit Sub
            End If

            OpenScreen PaymentsDataScreen

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

Private Sub FormatFONT_Click()
    On Error GoTo ErrTrap

    With FrmMainPriceList.FgMain
        Cmdlg.FontBold = .FontBold
        Cmdlg.FontItalic = .FontItalic
        Cmdlg.FontName = .FontName
        Cmdlg.fontsize = .fontsize
        Cmdlg.Flags = cdlCFBoth
        Cmdlg.ShowFont
        .FontBold = Cmdlg.FontBold
        .FontItalic = Cmdlg.FontItalic
        .FontName = Cmdlg.FontName
        .fontsize = Cmdlg.fontsize
        .Cell(flexcpFontBold, .FixedRows, .ColIndex("Tree"), .Rows - 1, .ColIndex("Tree")) = Cmdlg.FontBold
        .Cell(flexcpFontItalic, .FixedRows, .ColIndex("Tree"), .Rows - 1, .ColIndex("Tree")) = Cmdlg.FontItalic
        .Cell(flexcpFontSize, .FixedRows, .ColIndex("Tree"), .Rows - 1, .ColIndex("Tree")) = Cmdlg.fontsize
        .Cell(flexcpFontName, .FixedRows, .ColIndex("Tree"), .Rows - 1, .ColIndex("Tree")) = Cmdlg.FontName
        .AutoSize 0, .Cols - 1, False
        .Refresh
    End With

    FrmMainPriceList.SaveFontSetting
    Exit Sub
ErrTrap:
End Sub

Private Sub Groups_Click()

End Sub

Private Sub HelpFile_Click()
    SystemOptions.SysHelp.HHDisplayContents Me.hWnd
End Sub

Private Sub HelpIndex_Click()
    SystemOptions.SysHelp.HHDisplayIndex Me.hWnd
End Sub

Private Sub insurance_type_Click()

End Sub

Private Sub Items_Click(Index As Integer)

End Sub

Private Sub ItemsPrice_Click()
    On Error GoTo ErrTrap

    With FrmMainPriceList

        If .XPOptViewType(0).value = True Then
            If .FgMain.Rowdata(.FgMain.Row) <> "" Then
                If right(.FgMain.Rowdata(.FgMain.Row), 1) = "I" Then
                    FrmItemsPrice.XPLblItemName.Caption = .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("tree"))
                    FrmItemsPrice.txtqty.text = .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("Qty"))
                    FrmItemsPrice.XPLblItemCode.Caption = .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("ItemCode"))
                    FrmItemsPrice.XPTxtPrice.text = .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("DefalutPrice"))
                    FrmItemsPrice.TxtCompareValue.text = .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("DefalutPrice"))
                    FrmItemsPrice.XPLblItemID.Caption = left(.FgMain.Rowdata(.FgMain.Row), (Len(.FgMain.Rowdata(.FgMain.Row)) - 1))
                    FrmItemsPrice.show vbModal
                End If
            End If

        ElseIf .XPOptViewType(1).value = True Then
            FrmItemsPrice.XPLblItemName.Caption = .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("tree"))
            FrmItemsPrice.txtqty.text = .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("Qty"))
            FrmItemsPrice.XPLblItemCode.Caption = .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("ItemCode"))
            FrmItemsPrice.XPTxtPrice.text = .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("DefalutPrice"))
            FrmItemsPrice.TxtCompareValue.text = .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("DefalutPrice"))
            FrmItemsPrice.XPLblItemID.Caption = .FgMain.Rowdata(.FgMain.Row)
            FrmItemsPrice.show vbModal
        End If

    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub LeavingRecord_Click()

    If checkApility("FrmGoTime") = False Then
        Exit Sub
    End If

    FrmGoTime.show
    FrmGoTime.ZOrder 0
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
                shipment.show
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
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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

    End Select

End Sub

Private Sub m2_Click()
    xx.show
    xx.top = 0
    xx.left = 11500
    ' xx.SmartMenuXP1_Click (0)
 
End Sub

Private Sub MarketingMnusub_Click(Index As Integer)
Select Case Index
Case 0
Case 1

            If checkApility("overs") = False Then
                Exit Sub
            End If

            overs.show
 

Case 2


End Select
End Sub

Private Sub MarketingMnusubsub_Click(Index As Integer)
Select Case Index
Case 0

            If checkApility("FrmCustomerssFollow") = False Then
                Exit Sub
            End If

            FrmCustomerssFollow.show
            
            
            

End Select
End Sub

Private Sub MDIForm_DblClick()

    With Cmdlg
        '*.jpg,*.jpeg,*.jpe,*.jfif
        .CancelError = False
        .DialogTitle = " ≈Œ Ì«— ’Ê—…"
        'Set The Filter to show pictures only
        .Filter = "Bitmap (*.bmp)|*.bmp|JPEG(*.JPG,*.JPEG,*.JPE,*.JFIF)|*.jpg;*.jpeg;*.jpe;*.jfif|" & "GIF (*.gif)|*.gif|All Files|*.*" ' choose formats to include
        .ShowOpen
    
        If .FileName <> "" Then
            'Set Me.ImgPic.Picture = LoadPicture(.FileName)
            Me.Picture = LoadPicture(.FileName)
            WebForm.Picture = LoadPicture(.FileName)
            SaveSetting StrAppRegPath, "View_Type", "BackGroundImag", .FileName
        Else

            If Dir(App.path & "\Garphics\wallpaper_Main.jpg") <> "" Then
                Me.Picture = LoadPicture(App.path & "\Garphics\wallpaper_Main.jpg")
                WebForm.Picture = LoadPicture(.FileName)
                SaveSetting StrAppRegPath, "View_Type", "BackGroundImag", App.path & "\Garphics\wallpaper_Main.jpg"
                                
            End If

        End If

    End With

    ' €ÌÌ— «·Œ·›Ì…

End Sub

Private Sub MDIForm_Load()
    Dim BGround As ClsBackGroundPic
    Dim BolShowRequest As Boolean
    'On Local Error GoTo ErrTrap
    Me.backcolor = vbWhite
    Me.Caption = GetAppTitle  'App.Title
    CreateDocks
    LoadInterface SystemOptions.UserInterface
 
    If Messnger = False Then Timer1.Enabled = True

    BackGroundImag = GetSetting(StrAppRegPath, "View_Type", "BackGroundImag", App.path & "\Garphics\wallpaper_Main.jpg")

    If Dir(BackGroundImag) <> "" Then
        '   Me.Picture = LoadPicture(App.path & "\Garphics\wallpaper_Main.jpg")
        Me.Picture = LoadPicture(BackGroundImag)
        'AskOption
        'Set Me.PopMenu1.BackgroundPicture = Me.Picture
    End If

    'Grid_WallPaper.jpg
    If Dir(App.path & "\Garphics\Grid_WallPaper.jpg") <> "" Then
        '   Set Me.PopMenu1.BackgroundPicture = LoadPicture(App.Path & "\Garphics\Grid_WallPaper.jpg")
    End If

    'If Dir(App.Path & "\ReportDesign.exe") = "" Then
    '    ReportDesigner.Visible = False
    '    Sep30.Visible = False
    'End If
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

Private Sub MnuAccAnalysis_Click()
    FrmAccountingAnalysis.show
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
            keddawrym.show

    End Select

End Sub

Private Sub MnuAccDEV_Post_Click()
    Frm_General_Journal.show
End Sub

Private Sub MnuAccIntervals_Click()
    FrmAccountIntervals.show
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

            FrmBoxDeposit.show
            FrmBoxDeposit.ZOrder 0

        Case 1

            If checkApility("FrmBoxDeposit") = False Then
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

    FrmBoxDetetErrors.show
End Sub

Private Sub MnuBoxStock_Click()

    If checkApility("FrmBoxStock") = False Then
        Exit Sub
    End If

    OpenScreen BoxesStockScreen
End Sub

Private Sub MnuCheckBriefcase_Click()
    FrmChecksBriefcase.show
End Sub

Private Sub MNUCloseYear_Click()
    FrmClose.show
End Sub

Private Sub MnuCorrectSerial_Click()

    If checkApility("FrmToolsSerials") = False Then
        Exit Sub
    End If

    FrmToolsSerials.show
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

    If checkApility("FrmDataBaseTools") = False Then
        Exit Sub
    End If

    If Me.ActiveForm Is Nothing Then
        FrmDataBaseTools.show vbModal
    Else
        Msg = "ÌÃ» €·ﬁ «Ï ‘«‘… „‰ ‘«‘«  «·»—‰«„Ã ﬁ»·"
        Msg = Msg & Chr(13) & "«‰  ” Œœ„ Â–« «·‘«‘…....!!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If

End Sub

Private Sub MnuEmpDepartmentData_Click()

End Sub

Private Sub MnuEmpJobsData_Click()

End Sub

Private Sub MnuEmpsEmpTimeSeeting_Click()

End Sub

Private Sub mnuEmployeeBasicSub_Click(Index As Integer)

    Select Case Index

        Case 0
            Dim Frm As FrmTimeSetting

            If checkApility("FrmTimeSetting") = False Then
                Exit Sub
            End If

            Set Frm = New FrmTimeSetting

            Frm.WorkType = 0
            Frm.show
            Frm.ZOrder 0

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

            If checkApility("insurancecompanies") = False Then
                Exit Sub
            End If
            
            insurancecompanies.show

        Case 9

            If checkApility("insurancetype") = False Then
                Exit Sub
            End If
            
            insurancetype.show

        Case 10

            If checkApility("Insurance_class") = False Then
                Exit Sub
            End If
            
            Insurance_class.show

        Case 11

            If checkApility("frmtakeem") = False Then
                Exit Sub
            End If

            frmtakeem.show

    End Select

End Sub

Private Sub MnuHelpForums_Click()
    OpenWebSite "http://www.sattaryah.com/userGuide.pdf"
End Sub

Private Sub MnuInvPrintReceipt_Click()
    MnuInvPrintReceipt.Checked = Not MnuInvPrintReceipt.Checked
End Sub

Private Sub MnuInvPrintSave_Click()
    MnuInvPrintSave.Checked = Not MnuInvPrintSave.Checked
End Sub

Private Sub MnuInvSalesOptions_Click()
    On Error GoTo ErrTrap

    If SystemOptions.UserInvoiceShowProfit = 1 Then
        If Me.ActiveForm.name = "FrmSaleBill" Then
            Me.ActiveForm.Ele(8).Visible = Not Me.ActiveForm.Ele(8).Visible
            MnuInvSalesOptions.Checked = Me.ActiveForm.Ele(8).Visible
        End If
    End If

    Exit Sub
ErrTrap:
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If

End Sub

Private Sub MnuItemTools_ItemData_Click()
    Dim LngCurrentItemID As Long
    LngCurrentItemID = val(MnuItemTools_ItemData.Tag)

    If LngCurrentItemID <> 0 Then
        OpenScreen ItemsDataScreen, LngCurrentItemID
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

Private Sub MnuMaintnanceBasicSub_Click(Index As Integer)

    Select Case Index

        Case 0

        Case 1

    End Select

End Sub

Private Sub MnuMaintnanceBasicSub1_Click()

    If checkApility("FrmCompany") = False Then
        Exit Sub
    End If

    FrmCompany.show

End Sub

Private Sub MnuMaintnanceTransactions_Click(Index As Integer)

    Select Case Index

        Case 0
            Load FrmManAddNew
            FrmManAddNew.TxtModFlg.text = "N"
            FrmManAddNew.show
            
        Case 1

            If checkApility("FrmManStore") = False Then
                Exit Sub
            End If

            FrmManStore.show
            FrmManStore.ZOrder 0
 
        Case 2

            If checkApility("FrmOut") = False Then
                Exit Sub
            End If

            FrmOut.show
            FrmOut.TxtTicketNo.Visible = True
            FrmOut.lbl(32).Visible = True
              
        Case 3
            FrmManCusRecive.show

        Case 4
            FrmManGoBack.show

        Case 5
            FrmManOpenBalance.show

        Case 6
            FrmManStoreStock.show

        Case 7
            FrmManAlram.show

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
        Case 8

            If checkApility("FrmManStore") = False Then
                Exit Sub
            End If

            '    FrmManStore.Show
            '    FrmManStore.ZOrder 0
            FrmReports.show
            FrmReports.C1TabMain.CurrTab = 4

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

    Load FrmManEmpReport
    FrmManEmpReport.TxtOrgManID.text = val(VarTemp(0))
    FrmManEmpReport.TxtTicketNo.text = val(VarTemp(1))
    FrmManEmpReport.lblReciptNumber.Caption = val(VarTemp(2))
    FrmManEmpReport.show vbModal

End Sub

Private Sub MnuManToolsSub6_Click()
    Dim StrTemp As String
    Dim VarTemp As Variant
    Dim LngItemID As Long
    Dim StrItemSerial  As String

    If mdifrmmain.MnuManToolsSub6.Tag <> "" Then
        StrTemp = mdifrmmain.MnuManToolsSub6.Tag
        VarTemp = Split(StrTemp, ";", , vbTextCompare)
        LngItemID = val(VarTemp(0))
        StrItemSerial = Trim$(VarTemp(1))
        OpenScreen CheckItemSerial, LngItemID, StrItemSerial
    End If

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

    For i = 0 To FrmOutBarPane.OutBar.Groups.count - 1
        FrmOutBarPane.OutBar.Groups(i).ItemsStyle = x
    Next i

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

            If Not ItemsTreePane Is Nothing Then
                ItemsTreePane.LoadData ItemsTreePane.GroupsSort, ItemsTreePane.ItemsSort
            End If

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

    If checkApility("FrmPrintItemsBarcodes") = False Then
        Exit Sub
    End If

    FrmPrintItemsBarcodes.show
End Sub

Private Sub MnuProjectsBasicSub_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("project_status") = False Then
                Exit Sub
            End If

            project_status.show

        Case 1

            If checkApility("Contract_type") = False Then
                Exit Sub
            End If

            Contract_type.show

        Case 2

            If checkApility("FrmOtherCustomers") = False Then
                Exit Sub
            End If

            OpenScreen OtherCustomersScreen '

        Case 3
            FrmProcessUnit.show

        Case 4
            FrmProcessDef.show

        Case 5

            If checkApility("Projects") = False Then
                Exit Sub
            End If

            Projects.show

    End Select

End Sub

Private Sub MnuProjectsTransactions_Click(Index As Integer)

    Select Case Index

        Case 0

            'FrmDestruction
            If checkApility("FrmDestruction") = False Then
                Exit Sub
            End If

            OpenScreen DestructionScreen

        Case 1

            If checkApility("FrmEmpSalary3") = False Then
                Exit Sub
            End If

            FrmEmpSalary3.show

        Case 2

            If checkApility("FrmEmpSalary4") = False Then
                Exit Sub
            End If

            FrmEmpSalary4.show

        Case 3

            If checkApility("FrmOperationsFollow") = False Then
                Exit Sub
            End If

            FrmOperationsFollow.show
 
        Case 4

            If checkApility("projectsbill") = False Then
                Exit Sub
            End If
 
            projectsbill.show

        Case 5

            If checkApility("projectsReports") = False Then
                Exit Sub
            End If

            Projects.ShowReports
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

    FrmToolsRepireItemsCost.show
End Sub

Private Sub MnuToolsDataBase_Click(Index As Integer)
    Dim Msg As String

    Select Case Index

        Case 0

            If checkApility("open_my_connection") = False Then
                Exit Sub
            End If

            open_my_connection

        Case 1

            If checkApility("AdminLogin") = False Then
                Exit Sub
            End If

            AdminLogin.show

        Case 2
            Unload WebForm

            If Me.ActiveForm Is Nothing Then

                FrmNEWlOGIN.show
            Else
                Msg = "ÌÃ» €·ﬁ «Ï ‘«‘… „‰ ‘«‘«  «·»—‰«„Ã ﬁ»·"
                Msg = Msg & Chr(13) & "«‰  ” Œœ„ Â–« «·‘«‘…....!!!!"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

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

            If checkApility("System_alarms") = False Then
                Exit Sub
            End If

            System_alarms.show

        Case 4

            If checkApility("System_manger2") = False Then
                Exit Sub
            End If

            System_manger2.show

        Case 5

            If checkApility("coding") = False Then
                Exit Sub
            End If

            coding.show

        Case 6

            If checkApility("FrmMessnger") = False Then
                Exit Sub
            End If

            FrmMessnger.show

        Case 7

            If checkApility("SMSSeTTings") = False Then
                Exit Sub
            End If

            SMSSeTTings.show
            'WebForm.Show
    End Select

End Sub

Private Sub MnuToolsSetPrinters0_Click()
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
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

End Sub

Private Sub MnuUsersScreensPremission_Click()
    Dim Msg As String
    
    If SystemOptions.usertype = UserNormal Then
    
        '   MsgBox ""
        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
        '    Msg = Msg & Chr(13) & "Õ ‰Â“— Ê·««ÌÂ "
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If Not mdifrmmain.ActiveForm Is Nothing Then
        ModPremis.ShowScreenPermission Me.ActiveForm.name
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
   '     PopupMenu mdifrmmain.MdiContextMenu  ', vbPopupMenuRightAlign, X, Y + 200
    End If

ErrTrap:
End Sub

Private Sub MDIForm_Resize()

    Dim i As Integer
    On Error Resume Next

    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then

        For i = 0 To Forms.count - 1

            If Forms(i).name <> "MDIFrmMain" Then
                If Forms(i).MDIChild = True Then
                    Resize_Form Forms(i)
                End If
            End If

        Next i

    End If

End Sub

Private Sub MnuBackColor_Click()
    On Error GoTo ErrTrap
    Cmdlg.ShowColor

    With FrmMainPriceList
        .FgMain.Cell(flexcpBackColor, 1, .FgMain.Col, .FgMain.Rows - 1, .FgMain.Col) = Cmdlg.color
        .SaveMeSetting
    End With

    Exit Sub
ErrTrap:
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

Private Sub MnuBoxIncapacity_Increase_Click()

    If checkApility("FrmBoxIncapacity") = False Then
        Exit Sub
    End If

    FrmBoxIncapacity.show
End Sub

Private Sub MnuFinDiscounts_Click()

    'FrmDiscounts
    If checkApility("FrmDiscounts") = False Then
        Exit Sub
    End If

    OpenScreen AllowsDiscountsScreen
End Sub

Private Sub MnuForeColor_Click()
    On Error GoTo ErrTrap
    Cmdlg.ShowColor

    With FrmMainPriceList
        .FgMain.Cell(flexcpForeColor, 1, .FgMain.Col, .FgMain.Rows - 1, .FgMain.Col) = Cmdlg.color
        .SaveMeSetting
    End With

    Exit Sub
ErrTrap:
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

    If Not mdifrmmain.ActiveForm Is Nothing Then
        'GetMsgs 156, vbExclamation
        ' Exit Sub
    End If

    Unload System_alarms

    Select Case Index

        Case 0 'Load Arabic Interface
            LoadInterface ArabicInterface

        Case 1 'Load English Interface
            LoadInterface EnglishInterface
    End Select

    System_alarms.show

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

    For i = MPITP_GSort_Option.LBound To MPITP_GSort_Option.UBound
        MPITP_GSort_Option(i).Checked = False
    Next i

    MPITP_GSort_Option(Index).Checked = True

    If Not ItemsTreePane Is Nothing Then
        ItemsTreePane.GroupsSort = StrTemp
        ItemsTreePane.LoadData StrTemp, ItemsTreePane.ItemsSort
    End If

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

    For i = MPITP_ISort_Option.LBound To MPITP_ISort_Option.UBound
        MPITP_ISort_Option(i).Checked = False
    Next i

    MPITP_ISort_Option(Index).Checked = True

    If Not ItemsTreePane Is Nothing Then
        ItemsTreePane.ItemsSort = StrTemp
        ItemsTreePane.LoadData ItemsTreePane.GroupsSort, StrTemp
    End If

End Sub

Private Sub Options_Click()

    If checkApility("FrmOptions") = False Then
        Exit Sub
    End If

    OpenScreen OptionsScreen
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
    Dim Temp As String
    Dim TempArry As Variant
    Dim i As Integer

    With Me.PopMenu1
        Lparent = .MenuIndex("MnuWindowsList")
        Temp = .HierarchyPath(.MenuKey(ItemNumber), 1, "-")

        If Temp <> "" Then
            TempArry = Split(Temp, "-", , vbTextCompare)

            If CStr(TempArry(1)) Like .Caption("MnuWindowsList") Then

                For i = 0 To Forms.count - 1

                    If Forms(i).name = .MenuKey(ItemNumber) Then

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

            If checkApility("k") = False Then
                Exit Sub
            End If

            FrmPOSDATA.show

        Case 1
 
            If checkApility("frm_sheft") = False Then
                Exit Sub
            End If

            frm_sheft.show
 
        Case 2
 
            If checkApility("FrmTables") = False Then
                Exit Sub
            End If

            FrmTables.show

        Case 3

            If checkApility("cachierData") = False Then
                Exit Sub
            End If

            cachierData.show

        Case 4

            If checkApility("CashierLogin") = False Then
                Exit Sub
            End If
 
            CashierLogin.show
            'frmsalebill1.Show
 
        Case 5

            If checkApility("ReportSales") = False Then
                Exit Sub
            End If

            FrmReports.show
            FrmReports.C1TabMain.CurrTab = 0
 
    End Select

End Sub

Private Sub PpBarcode_Click()
    'Barcode_Click
End Sub

Private Sub PrbH_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("FrmOutProductionOrder1") = False Then
                Exit Sub
            End If

            FrmOutProductionOrder1.show

        Case 1

            If checkApility("FrmProductionOrder1") = False Then
                Exit Sub
            End If

            FrmProductionOrder1.show

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

            If checkApility("frm_sheft") = False Then
                Exit Sub
            End If

            frm_sheft.show

        Case 1

            If checkApility("FrmıEquipment") = False Then
                Exit Sub
            End If

            FrmıEquipment.show
            'Case 2
            'If checkApility("frmProductLine") = False Then
            '    Exit Sub
            'End If

            'frmProductLine.Show

        Case 4

            If checkApility("FrmShowPrice1") = False Then
                Exit Sub
            End If

            'FrmCustomerOrder.Show
            GeneralPriceType = 1
            FrmShowPrice.show

        Case 5

            If checkApility("FrmProductionOrder") = False Then
                Exit Sub
            End If

            FrmProductionOrder.show
 
        Case 6

            If checkApility("FrmOutProductionOrder") = False Then
                Exit Sub
            End If

            FrmOutProductionOrder.show

            'FrmOut.Show
            'FrmOutForOrder.Show
        Case 7

            If checkApility("FrmInpoutWorkOrder") = False Then
                Exit Sub
            End If
 
            FrmInpoutWorkOrder.show

        Case 8

            If checkApility("FrmCalcCostPrice") = False Then
                Exit Sub
            End If

            FrmCalcCostPrice.show

        Case 9

            If checkApility("FrmCalcCostPrice1") = False Then
                Exit Sub
            End If

            FrmCalcCostPrice2.show

        Case 10

            If checkApility("FrmProductionReport") = False Then
                '    Exit Sub
            End If

            frmProductionreport.show

    End Select

End Sub

Private Sub PriceChips_Click()
    FrmMainPriceList.FgMain_DblClick
End Sub

Private Sub PriceOffer_Click()
    On Error GoTo ErrTrap

    With FrmMainPriceList

        If .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("ItemID")) = "" Then Exit Sub
        FrmPurchasePrice.XPLblItemName.Caption = .FgMain.Cell(flexcpTextDisplay, .FgMain.Row, .FgMain.ColIndex("Tree"))
        FrmPurchasePrice.XPLblItemID.Caption = .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("ItemID"))
    End With

    FrmPurchasePrice.show vbModal
    Exit Sub
ErrTrap:
End Sub

Private Sub ProductionPlansub_Click(Index As Integer)

    Select Case Index

        Case 0
            FrmProductionPlan.show

        Case 1
            FrmQCitems.show

        Case 2
            FrmItemsClass.show
            FrmItemsClass.Caption = " ’‰Ì› «·„‰ Ã« "
            FrmItemsClass.EleHeader.Caption = FrmItemsClass.Caption

        Case 3
            frmcorrectaction.show

        Case 4
            FrmInpoutWorkOrder.show
            FrmInpoutWorkOrder.Caption = "›Õ’  ÃÊœ… «·„‰ Ã «· «„"
            FrmInpoutWorkOrder.Ele(6).Caption = FrmInpoutWorkOrder.Caption

        Case 5
            FrmProductionOrder.show
            FrmProductionOrder.Caption = "«„— ‘€· «’·«Õ «·„‰ Ã«  «·„⁄Ì»…"
            FrmProductionOrder.Ele(6).Caption = FrmProductionOrder.Caption
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

            If checkApility("FrmShipment_mode") = False Then
                Exit Sub
            End If

            FrmShipment_mode.show

        Case 4

            If checkApility("FrmGaranty_type") = False Then
                Exit Sub
            End If

            FrmGaranty_type.show

        Case 5
            AgengItem.show

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
                shipment.show
            End If

        Case 3
            Set RsOptions = New ADODB.Recordset
            RsOptions.Open "tbloptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

            If checkApility("FrmBillBuy") = False Then
                Exit Sub
            End If

            OpenScreen PurchaseScreen

            'FrmBillBuy
        Case 4

            If checkApility("FrmReturnpurchases") = False Then
                Exit Sub
            End If

            OpenScreen RetrunPurchse

            'FrmReturnpurchases
        Case 5

            If checkApility("Ageng_all") = False Then
                Exit Sub
            End If

            Ageng_all.show

        Case 6

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

Private Sub PurchaseTransactionssubs1_Click(Index As Integer)

    Select Case Index

        Case 0
            GeneralPriceType = 6

            If checkApility("FrmShowPrice1") = False Then
                Exit Sub
            End If

            OpenScreen ScreensName.ShowPriceScreen

        Case 1

        Case 2
            GeneralPriceType = 1

            If checkApility("FrmShowPrice1") = False Then
                Exit Sub
            End If

            OpenScreen ScreensName.ShowPriceScreen

    End Select

End Sub

Private Sub ReceiptPart_Click()

    'FrmReceiptPart
    If checkApility("FrmReceiptPart") = False Then
        Exit Sub
    End If

    OpenScreen ReceiptPartScreen
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

Private Sub RequiredInstallment_Click()

    'FrmInstallmentMustPay
    If checkApility("FrmInstallmentMustPay") = False Then
        Exit Sub
    End If

    OpenScreen PopUpShowInstallmentMustPay
End Sub

Private Sub SalesBasicSub_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("FrmCustomerType") = False Then
                Exit Sub
            End If

            FrmCustomerType.show

        Case 1

            If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If

            'FrmCustemers
            OpenScreen CustomersScreen '

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

            If checkApility("posdata") = False Then
                Exit Sub
            End If

            FrmPOSDATA.show

        Case 5

            If checkApility("cachierData") = False Then
                Exit Sub
            End If

            cachierData.show

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

            If checkApility("FrmSalesRepData") = False Then
                Exit Sub
            End If

            FrmSalesRepData.show
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

            Dim RsOptions As New ADODB.Recordset
            Set RsOptions = New ADODB.Recordset
            RsOptions.Open "tbloptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
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
            frmsalebillCompose.show

        Case 5

            If checkApility("overs") = False Then
                Exit Sub
            End If

            overs.show

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

            If checkApility("FrmCustomerssFollow") = False Then
                Exit Sub
            End If

            FrmCustomerssFollow.show
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

Private Sub SalesTransactionssubss000_Click(Index As Integer)

    Select Case Index

        Case 2
            GeneralPriceType = 0

            If checkApility("FrmShowPrice") = False Then
                Exit Sub
            End If

            OpenScreen ScreensName.ShowPriceScreen
    End Select

End Sub

Private Sub SearchInHelp_Click()
    SystemOptions.SysHelp.HHDisplaySearch Me.hWnd
End Sub

Private Sub ShortCuts_Click()
    FrmShortCut.show
    FrmShortCut.ZOrder 0
End Sub

Private Sub ShowCol_Click()
    On Error GoTo ErrTrap

    With FrmShowCol.FG
        .TextMatrix(0, .ColIndex("show")) = Not (FrmMainPriceList.FgMain.ColHidden(FrmMainPriceList.FgMain.ColIndex("ItemID")))
        .TextMatrix(1, .ColIndex("show")) = Not (FrmMainPriceList.FgMain.ColHidden(FrmMainPriceList.FgMain.ColIndex("ItemCode")))
        .TextMatrix(2, .ColIndex("show")) = Not (FrmMainPriceList.FgMain.ColHidden(FrmMainPriceList.FgMain.ColIndex("Qty")))
        .TextMatrix(3, .ColIndex("show")) = Not (FrmMainPriceList.FgMain.ColHidden(FrmMainPriceList.FgMain.ColIndex("DefalutPrice")))
        .TextMatrix(4, .ColIndex("show")) = Not (FrmMainPriceList.FgMain.ColHidden(FrmMainPriceList.FgMain.ColIndex("LastUpdate")))
    End With

    FrmShowCol.show vbModal
    Exit Sub
ErrTrap:
End Sub

Private Sub ShowItems_Click()
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    On Error GoTo ErrTrap

    With FrmMainPriceList

        If .FgMain.Row = -1 Then Exit Sub
        If .FgMain.Col = -1 Then Exit Sub
        If .XPOptViewType(0).value = True Then
            If right(.FgMain.Rowdata(.FgMain.Row), 1) = "I" Then
                If .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("ItemID")) <> "" Then
                    StrSQL = "select * From TblItems where ItemID=" & .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("ItemID"))
                    Set RsTemp = New ADODB.Recordset
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
                    FrmSearchSerial.Tag = RsTemp("ItemCode").value
                    FrmSearchSerial.Txt.text = "PriceList"
                    FrmSearchSerial.show vbModal
                    RsTemp.Close
                End If
            End If

        ElseIf .XPOptViewType(1).value = True Then

            If .FgMain.Row = 0 Then Exit Sub
            If .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("ItemID")) <> "" Then
                StrSQL = "select * From TblItems where ItemID=" & .FgMain.TextMatrix(.FgMain.Row, .FgMain.ColIndex("ItemID"))
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
                FrmSearchSerial.Tag = RsTemp("ItemCode").value
                FrmSearchSerial.Txt.text = "PriceList"
                FrmSearchSerial.show vbModal
                RsTemp.Close
            End If
        End If

    End With

    Exit Sub
ErrTrap:
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

            If checkApility("FrmSystemUnites") = False Then
                Exit Sub
            End If

            FrmSystemUnites.show

        Case 4

            If checkApility("FrmItemsColor") = False Then
                Exit Sub
            End If

            FrmItemsColor.show

        Case 5

            If checkApility("FrmItemsSize") = False Then
                Exit Sub
            End If

            FrmItemsSize.show

        Case 6

            If checkApility("FrmItemsClass") = False Then
                Exit Sub
            End If

            FrmItemsClass.show

        Case 7

            If checkApility("FrmStoresLocation") = False Then
                Exit Sub
            End If

            FrmStoresLocation.show

        Case 8

            If checkApility("FrmSalePriceNames") = False Then
                Exit Sub
            End If

            FrmSalePriceNames.show

        Case 9

            If checkApility("FrmProductionElements") = False Then
                Exit Sub
            End If

            FrmProductionElements.show

        Case 10

            If checkApility("UnitsIndustrialCost") = False Then
                Exit Sub
            End If

            UnitsIndustrialCost.show

        Case 11

            If checkApility("frmitemsalessPlan") = False Then
                Exit Sub
            End If

            'frmitemsalessPlan

    End Select

End Sub

Private Sub SupBackColor_Click()
    On Error GoTo ErrTrap
    Cmdlg.ShowColor

    With FrmMainPriceList
        .FgMain.Cell(flexcpBackColor, 1, .FgMain.Col, .FgMain.Rows - 1, .FgMain.Col) = Cmdlg.color
        .SaveSupPriceSetting
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub SupFont_Click()
    On Error GoTo ErrTrap

    With FrmMainPriceList.FgMain
        Cmdlg.FontBold = .FontBold
        Cmdlg.FontItalic = .FontItalic
        Cmdlg.FontName = .FontName
        Cmdlg.fontsize = .fontsize
        Cmdlg.Flags = cdlCFBoth
        Cmdlg.ShowFont
        .FontBold = Cmdlg.FontBold
        .FontItalic = Cmdlg.FontItalic
        .FontName = Cmdlg.FontName
        .fontsize = Cmdlg.fontsize
        .AutoSize 0, .Cols - 1, False
        .Refresh
        '    .Cell(flexcpFontBold, .FixedRows, .ColIndex("Tree"), .Rows - 1, .ColIndex("Tree")) = Cmdlg.FontBold
        '    .Cell(flexcpFontItalic, .FixedRows, .ColIndex("Tree"), .Rows - 1, .ColIndex("Tree")) = Cmdlg.FontItalic
        '    .Cell(flexcpFontSize, .FixedRows, .ColIndex("Tree"), .Rows - 1, .ColIndex("Tree")) = Cmdlg.FontSize
        '    .Cell(flexcpFontName, .FixedRows, .ColIndex("Tree"), .Rows - 1, .ColIndex("Tree")) = Cmdlg.FontName
    End With

    FrmMainPriceList.SaveFontSetting
    Exit Sub
ErrTrap:
End Sub

Private Sub SupForeColor_Click()
    On Error GoTo ErrTrap
    Cmdlg.ShowColor

    With FrmMainPriceList
        .FgMain.Cell(flexcpForeColor, 1, .FgMain.Col, .FgMain.Rows - 1, .FgMain.Col) = Cmdlg.color
        .SaveSupPriceSetting
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Texh_Click(Index As Integer)

    Select Case Index

        Case 0

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
    End Select

End Sub

Private Sub Timer1_Timer()

    If Messnger = False Then Exit Sub
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

                ElseIf rsOut!checkbey = True Then
                    Msg = "⁄›Ê«  „ «Œ Ì«— ›« Ê—… «·‘—«¡ ··«÷«›…  ... ·«Ì„ﬂ‰ «·«÷«›…  „‰ «–‰ «·«÷«›… "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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

            If ShowRequest(True) = True Then
                FrmRequest.show
                FrmRequest.ZOrder 0
            End If

        Case 11
            ShowItemsStatusReport WindowTarget

            'FrmInventoryStatus.Show
        Case 12

            If checkApility("ReportItems") = False Then
                Exit Sub
            End If

            FrmReports.show
            FrmReports.C1TabMain.CurrTab = 7

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
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                Else
                End If
            End If
            
        Case 1

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
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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

            FrmCitiesDistance.show

        Case 2

            If checkApility("FrmCustemers") = False Then
                Exit Sub
            End If

            OpenScreen CustomersScreen '

        Case 3

            If checkApility("FrmCompany") = False Then
                Exit Sub
            End If

            FrmCompany.show

        Case 4

            If checkApility("FrmDrivers") = False Then
                Exit Sub
            End If

            FrmDrivers.show

        Case 5

            If checkApility("CarTypes") = False Then
                Exit Sub
            End If

            CarTypes.show
 
        Case 6

            If checkApility("insurancecompanies1") = False Then
                Exit Sub
            End If

            insurancecompanies.show

        Case 7

            If checkApility("FRMMaintenanceTypes") = False Then
                Exit Sub
            End If

            FRMMaintenanceTypes.show

        Case 8

            If checkApility("FrmCars") = False Then
                Exit Sub
            End If

            FrmCars.show

        Case 9

            If checkApility("FrmTravelTransactions") = False Then
                Exit Sub
            End If

            FrmTravelTransactions.show

        Case 10

            If checkApility("frmTravelRports") = False Then
                Exit Sub
            End If

            frmTravelRports.show

    End Select

End Sub

Private Sub UserAbility_Click()
    Dim Msg As String
    
    'If SystemOptions.usertype = UserNormal Then
    If user_id <> 1 Then
        '   MsgBox ""
        Msg = "·Ì” ·œÌﬂ «·’·«ÕÌ… ··œŒÊ· ⁄·Ï Â–Â «·‘«‘…"
        '    Msg = Msg & Chr(13) & "Õ ‰Â“— Ê·««ÌÂ "
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        FrmUserAbility.show
        FrmUserAbility.ZOrder 0
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        FrmPermission.show
        FrmPermission.ZOrder 0
    End If

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
        MyFont.name = "MS Sans Serif"
        MyFont.Bold = False
        MyFont.Charset = 178
        MyFont.Size = 8
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
        App.Title = GetAppTitle
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

            Set XPanel = .Panels.Add(, "Pan_Comment", App.Title, , mdifrmmain.Icon)
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
            xPane.Title = "‘—Ìÿ «·≈Œ ’«—« "
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.NewsBarPaneID)

        If Not xPane Is Nothing Then
            xPane.Title = "„⁄·Ê„«  «·»—‰«„Ã"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID)

        If Not xPane Is Nothing Then
            xPane.Title = "‘Ã—… «·√’‰«›"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.MantainceID)

        If Not xPane Is Nothing Then
            xPane.Title = "«·’Ì«‰…"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.InternetNews)

        If Not xPane Is Nothing Then
            xPane.Title = "„⁄·Ê„«  «·≈‰ —‰ "
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.DynamicHelp)

        If Not xPane Is Nothing Then
            xPane.Title = "«·„”«⁄œ… «··ÕŸÌ…"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.CalendarPaneID)

        If Not xPane Is Nothing Then
            xPane.Title = "«·”«⁄… "
        End If

        Me.XPStusBar.Refresh
    ElseIf IntInterface = EnglishInterface Then
        SystemOptions.UserInterface = EnglishInterface
        App.Title = GetAppTitle
        Me.RightToLeft = False
        Me.PopMenu1.RightToLeft = False

        With Me.XPStusBar
            .Panels.Clear
            Set XPanel = .Panels.Add(, "Pan_Comment", App.Title, , mdifrmmain.Icon)
            XPanel.Style = sbrText
            XPanel.Alignment = sbrLeft
            XPanel.ToolTipText = "Goto  BYTE"
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
                XPanel.ToolTipText = "Current Open Accounting Interval Number"
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

        If Not xPane Is Nothing Then
            xPane.Title = "Shortcut OutBar"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.NewsBarPaneID)

        If Not xPane Is Nothing Then
            xPane.Title = "Programe Information"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID)

        If Not xPane Is Nothing Then
            xPane.Title = "Items Tree"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.MantainceID)

        If Not xPane Is Nothing Then
            xPane.Title = "Maintenance"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.InternetNews)

        If Not xPane Is Nothing Then
            xPane.Title = "Internet Information"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.DynamicHelp)

        If Not xPane Is Nothing Then
            xPane.Title = "Dynamic Help"
        End If

        Set xPane = Me.DockingPane1.FindPane(DockingPanesIDs.CalendarPaneID)

        If Not xPane Is Nothing Then
            xPane.Title = "Calendar"
        End If

        Me.XPStusBar.Refresh
    End If

    Me.Caption = App.Title

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

    If Not FrmOutBarPane Is Nothing Then
        FrmOutBarPane.LoadInterface SystemOptions.UserInterface '
    End If

    If Not FrmNewsBarPane Is Nothing Then
        FrmNewsBarPane.CreateTaskPanel
    End If

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

Private Sub Vscstionsssub_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("FrmHolidayPlan") = False Then
                Exit Sub
            End If

            FrmHolidayPlan.show

        Case 1

            If checkApility("FrmHolidayorder") = False Then
                Exit Sub
            End If

            FrmHolidayorder.show

        Case 2
 
            If checkApility("FrmFixedAssetMoving") = False Then
                Exit Sub
            End If

            FrmFixedAssetMoving.show

        Case 3

            If checkApility("FrmHolidayorder2") = False Then
                Exit Sub
            End If

            FrmHolidayorder2.show

        Case 4

            If checkApility("FrmHolidayorder3") = False Then
                Exit Sub
            End If

            FrmHolidayorder3.show

    End Select

End Sub

Private Sub XC_Click(Index As Integer)

    Select Case Index

        Case 0
            GeneralPriceType = 3

            If checkApility("FrmShowPrice3") = False Then
                Exit Sub
            End If

            FrmShowPrice.show

        Case 1
            GeneralPriceType = 4

            If checkApility("FrmShowPrice4") = False Then
                Exit Sub
            End If

            FrmShowPrice.show
            
    End Select

End Sub

Private Sub XPStusBar_PanelClick(ByVal Panel As MSComctlLib.Panel)

    Select Case Panel.key

        Case "WebSite"
            OpenWebSite
    End Select

End Sub

Private Sub SetMenus()

    'On Error GoTo ErrTrap
    If SystemOptions.UserInterface = ArabicInterface Then
 
 
 POSTRansactiosG.Caption = "‰ﬁ«ÿ «·»Ì⁄"

POSTRansactios(0).Caption = "»Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄"
POSTRansactios(1).Caption = "»Ì«‰«  «·‘Ì› "
POSTRansactios(2).Caption = "»Ì«‰«  «·„Ê«ﬁ⁄"
POSTRansactios(3).Caption = "»Ì«‰«  «·ﬂ«‘Ì—"
POSTRansactios(4).Caption = " ”ÃÌ· «·œŒÊ·"
POSTRansactios(5).Caption = "«· ﬁ«—Ì—"



 MarketingMnu.Caption = "«· ”ÊÌﬁ"
MarketingMnusub(0).Caption = "ŒÿÂ „»Ì⁄«  «·«’‰«›"
MarketingMnusub(1).Caption = "⁄—Ê÷ «·«’‰«›"
MarketingMnusub(2).Caption = "„ «»⁄Â  «·⁄„·«¡"


MarketingMnusubsub(0).Caption = " ”ÃÌ· “Ì«—«  «·⁄„·«¡"
MarketingMnusubsub(1).Caption = "„ «»⁄Â “Ì«—«  «·⁄„·«¡"
MarketingMnusubsub(2).Caption = "«” ÿ·«⁄ —√Ì «·⁄„·«¡"
MarketingMnusubsub(3).Caption = " ”ÃÌ· ‘ﬂ«ÊÌ «·⁄„·«¡"
MarketingMnusubsub(4).Caption = "„ «»⁄Â ‘ﬂ«ÊÌ «·⁄„·«¡"
MarketingMnusubsub(5).Caption = "œ·Ì· «·Â« ›"

        Me.BasicData.Caption = "«·»Ì«‰«   «·«”«”Ì…"
        Me.BasicDataM(0).Caption = "  —»ÿ «·Õ”«»« "
        Me.BasicDataM(1).Caption = "  «·«‰‘ÿ… Ê «·›—Ê⁄"
        Me.BasicDataM(2).Caption = "   »Ì«‰«  «·»‰Êﬂ"
        Me.BasicDataM(3).Caption = "  »Ì«‰«  «·Œ“‰ Ê «·⁄Âœ"
        Me.BasicDataM(4).Caption = "  ÿ—ﬁ «·œ›⁄ "
        Me.BasicDataM(5).Caption = "  »Ì«‰«  «·„Ê—œÌ‰"
        Me.BasicDataM(6).Caption = "  »Ì«‰«  «·⁄„·«¡"

        Me.BasicDataM(7).Caption = "  »Ì«‰«  «·⁄„·« "
        Me.BasicDataM(8).Caption = "  »Ì«‰«  «·Ã‰”Ì« "
        Me.BasicDataM(9).Caption = "  »Ì«‰«  «·œÌ«‰« "
        Me.BasicDataM(10).Caption = "  »Ì«‰«   «·œÊ·"
        Me.BasicDataM(11).Caption = "  »Ì«‰«  «·„œ‰"
        Me.BasicDataM(12).Caption = "  »Ì«‰«  «·«ÕÌ«¡"
        Me.BasicDataM(13).Caption = "  »Ì«‰«  «·‘Ê«—⁄"
        Me.BasicDataM(14).Caption = "  «‰Ê«⁄ «·„” ‰œ«   "
        Me.BasicDataM(15).Caption = "  »Ì«‰«  «·«’‰« ›  "

        Me.BasicDataM(17).Caption = "  Œ—ÊÃ"
        AssetsMngBase.Caption = "«œ«—… «·«„·«ﬂ"
        mnuEmployee.Caption = "‘∆Ê‰ «·„ÊŸ›Ì‰"
        MnuAccDEV(0).Caption = "  ﬁÌœ «·ÌÊ„Ì…"
        MnuAccDEV_Post.Caption = "  „—«Ã⁄Â ﬁÌÊœ «·ÌÊ„Ì…"
        xxx(0).Caption = "  «‰Ê«⁄ „—«ﬂ“ «· ﬂ·›…"
        xxx(1).Caption = "  »Ì«‰«  „—«ﬂ“ «· ﬂ·›…"

        xxy(0).Caption = "  «·„Ê«“‰… «·⁄«„…"
        xxy(1).Caption = "  «· œ›ﬁ «·‰ﬁœÌ  "
        xxy(2).Caption = "   »ÊÌ» «·„Ì“«‰Ì…"
        xxy(3).Caption = "   Ê“Ì⁄ «·Õ”«»« "
        xxy(4).Caption = "  «⁄œ«œ „⁄«œ·«  «· Õ·Ì· «·„«·Ì"
        xxy(5).Caption = "  «ŸÂ«— ‰ «∆Ã «· Õ·Ì· «·„«·Ì"
        xxy(6).Caption = "  «·Õ”«»«  «·„Ã„⁄Â  "
        xxy(7).Caption = " «Õ’«∆Ì« "
        xxy(8).Caption = "√Ã‰œ… «·⁄„·«¡"

        ProductionPlan.Caption = "«· ŒÿÌÿ  Ê «·ÃÊœ…"
        'xxx(4).Caption = "  «· Õ·Ì· «·„«·Ì"
        ProductionPlansub(0).Caption = "ŒÿÂ «·«‰ «Õ"
        ProductionPlansub(1).Caption = " ⁄—Ì› ⁄‰«’— „—«ﬁ»… «·ÃÊœ…"
        ProductionPlansub(2).Caption = " ’‰Ì› «·„‰ Ã« "
        ProductionPlansub(3).Caption = " ⁄—Ì› «·«Ã—«¡«  «· ’ÕÌÕÌ…"
        ProductionPlansub(4).Caption = "›Õ’  ÃÊœ… «·„‰ Ã «· «„"
        ProductionPlansub(5).Caption = "„ «»⁄Â Ê ”ÃÌ· «’·«Õ «·„‰ Ã«  «·„⁄Ì»Â"
        xxx(12).Caption = "   ﬁ«—Ì— «·Õ”«»« "
        Me.MnuProjects.Caption = "«·„‘«—Ì⁄"
        Me.MnuProjectsBasic.Caption = "«·»Ì«‰«  «·«”«”…"
        Me.MnuProjectsBasicSub(0).Caption = "  Õ«·«  «·„‘«—Ì⁄"
        Me.MnuProjectsBasicSub(1).Caption = " «‰Ê«⁄ «·⁄ﬁÊœ"
        Me.MnuProjectsBasicSub(2).Caption = "»Ì«‰«  „ﬁ«Ê·Ì «·»«ÿ‰"
        Me.MnuProjectsBasicSub(4).Caption = "ÊÕœ«  «·⁄„·Ì« "
        Me.MnuProjectsBasicSub(4).Caption = " ⁄—Ì› «·⁄„·Ì« "
        Me.MnuProjectsBasicSub(5).Caption = "»Ì«‰«  «·„‘«—Ì⁄"
              
        Me.MnuProjectsTransactions(0).Caption = " ”‰œ ’—› „Ê«œ ··„‘«—Ì⁄"
        Me.MnuProjectsTransactions(1).Caption = "   Œ’Ì’ «·⁄„«·…"
        Me.MnuProjectsTransactions(2).Caption = "  ‰ﬁ· «·⁄„«·Â"
        Me.MnuProjectsTransactions(3).Caption = "  „ «»⁄Â «·⁄„·Ì«  "
        Me.MnuProjectsTransactions(4).Caption = "  ›« Ê—… „‘—Ê⁄"
        Me.MnuProjectsTransactions(5).Caption = "   ﬁ«—Ì— «·„‘«—Ì⁄"
        mnuEmployeeBasic(0).Caption = "  «·»Ì«‰«  «·«”«”ÌÂ"
        mnuEmployeeBasicSub(0).Caption = "«⁄œ«œ «Êﬁ«  ⁄„· «·‘—ﬂ…"
        mnuEmployeeBasicSub(1).Caption = "«·‘Ì› « "
        mnuEmployeeBasicSub(2).Caption = "«·«Ã«“« "
        mnuEmployeeBasicSub(3).Caption = "«‰Ê«⁄ «·⁄ﬁÊœ"
        mnuEmployeeBasicSub(4).Caption = "  Õ«·«  «·⁄„·"
        mnuEmployeeBasicSub(5).Caption = "»Ì«‰«  «·«ﬁ”«„"
        mnuEmployeeBasicSub(6).Caption = " »Ì«‰«  «·ÊŸ«∆›"
        mnuEmployeeBasicSub(7).Caption = "»Ì«‰«  «· Œ’’« "
        mnuEmployeeBasicSub(8).Caption = "»Ì«‰«  ‘—ﬂ«  «· √„Ì‰"
        mnuEmployeeBasicSub(9).Caption = "»Ì«‰«  «‰Ê«⁄ «· √„Ì‰"
        mnuEmployeeBasicSub(10).Caption = "»Ì«‰«  ›∆«  «· √„Ì‰"
        mnuEmployeeBasicSub(11).Caption = "⁄‰«’— «· ﬁÌÌ„"
        mnuEmployeeBasic(2).Caption = "  «·Õ÷Ê— Ê «·«‰’—«›"
        EmployeeAttendanceSub(0).Caption = "«⁄œ«œ «·Õ÷Ê— Ê «·«‰’—«› ··‘—ﬂÂ"
        EmployeeAttendanceSub(0).Caption = "«⁄œ«œ «·Õ÷Ê— Ê «·«‰’—«› ·„ÊŸ›"
        EmployeeAttendanceSub(1).Caption = " ”ÃÌ· „Ê«⁄Ìœ «·Õ÷Ê— Ê «·«‰’—«› ÌœÊÌ«"
        EmployeeAttendanceSub(2).Caption = " ”ÃÌ·  „Ê«⁄Ìœ «·Õ÷Ê— Ê «·«‰’—«› «·Ì«"
        EmployeeAttendanceSub(3).Caption = " ”ÃÌ· «·€Ì«»"
        EmployeeAttendanceSub(4).Caption = "«·⁄—÷ «·⁄«„ ·„Ê«⁄Ìœ «·Õ÷Ê— Ê «·«‰’—«›"
        mnuEmployeeBasic(3).Caption = "«·—Ê« »"
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
        EmployeeSalarySub(11).Caption = " ”ÃÌ·  —ﬂ «·Œœ„… "
        EmployeeSalarySub(12).Caption = " €ÌÌ— «—ÌŒ «Ê «Ìﬁ«› ”·›…"

        mnuEmployeeBasic(4).Caption = "«Ã«“«  «·„ÊŸ›Ì‰"

        Vscstionsssub(0).Caption = "ŒÿÂ «·«Ã«“« "
        Vscstionsssub(1).Caption = "ÿ·» «Ã«“…"
        Vscstionsssub(2).Caption = " ”·Ì„ Ê ”·„ ⁄Âœ ⁄Ì‰Ì…"
        Vscstionsssub(3).Caption = "„” Õﬁ«  «·«Ã«“…"
        Vscstionsssub(4).Caption = " ”ÃÌ· «·Õ÷Ê— „‰ «Ã«“…"

        mnuEmployeeBasic(5).Caption = "«‰Â«¡ «·Œœ„Â"
        FinishSevicersub(0).Caption = " ”ÃÌ·  —ﬂ «·Œœ„Â"
        FinishSevicersub(1).Caption = "Õ”«» „ﬂ«›√… ‰Â«Ì… «·Œœ„Â"
  
        mnuEmployeeBasic(1).Caption = "  »Ì«‰«  «·„ÊŸ›»‰"
        EmployeeDataicSub(0).Caption = "  „·› «·„ÊŸ›Ì‰"
        EmployeeDataicSub(1).Caption = "  ⁄ﬁÊœ «·„ÊŸ›Ì‰"
        TransporterMain.Caption = "«·‰ﬁ·Ì« "
        TransporterSub(0).Caption = "»Ì«‰«  «·„œ‰"
        TransporterSub(1).Caption = "«·„”«›«  »Ì‰ «·„œ‰"
        TransporterSub(2).Caption = "»Ì«‰«  «·⁄„·«¡"
        TransporterSub(3).Caption = "»Ì«‰«  «·„Ê—œÌ‰"
        TransporterSub(4).Caption = "»Ì«‰«  «·”«∆ﬁÌ‰"
        TransporterSub(5).Caption = "«‰Ê«⁄ «·„—ﬂ»« "
        TransporterSub(6).Caption = "‘—ﬂ«  «· √„Ì‰"
        TransporterSub(7).Caption = "«‰Ê«⁄ «·’Ì«‰… «·œÊ—Ì…"
        TransporterSub(8).Caption = "»Ì«‰«  «·„—ﬂ»« "
        TransporterSub(9).Caption = "»Ì«‰«  «·—Õ·« "
        TransporterSub(10).Caption = "«· ﬁ«—Ì—"

        Me.StockControl.Caption = " «·„Œ“Ê‰"
        Me.StockControlBasic.Caption = "«·»Ì«‰«  «·«”«”Ì…"
        StockControlBasicSub(0).Caption = "»Ì«‰«  «·«’‰«›"
        StockControlBasicSub(1).Caption = "»Ì«‰«  «·„Œ«“‰  "
        StockControlBasicSub(2).Caption = "„Ã„Ê⁄«  «·«’‰«›"
        StockControlBasicSub(3).Caption = "«·ÊÕœ« "
        StockControlBasicSub(4).Caption = "«·Ê«‰ «·«’‰«›"
        StockControlBasicSub(5).Caption = "„ﬁ«”«  «·«’‰«›"
        StockControlBasicSub(6).Caption = "›—“ «·«’‰«›"
        StockControlBasicSub(7).Caption = "«⁄œ«œ «„«ﬂ‰ «· Œ“Ì‰"
        StockControlBasicSub(8).Caption = "«”„«¡ «”⁄«— »Ì⁄ «·«’‰«›"

        StockControlBasicSub(9).Caption = "⁄‰«’—  ﬂ«·Ì› «·«‰ «Ã  "
        StockControlBasicSub(10).Caption = " «· ﬂ«·Ì› «·’‰«⁄Ì… ÿ»ﬁ« ··ÊÕœ…"
        StockControlBasicSub(11).Caption = "Œÿ… „»Ì⁄«  «·«’‰«›"
        Me.TradingTransaction(0).Caption = " «·—’Ìœ «·«›  «ÕÌ"
        Me.TradingTransaction(1).Caption = "«·ÿ·»«  «·œ«Œ·Ì…"
        XC(0).Caption = "ÿ·»«  œ«Œ·Ì…"
        XC(1).Caption = "”‰œ ÕÃ“ »÷«⁄Â œ«Œ·Ì"
        Me.TradingTransaction(2).Caption = "”‰œ«  «·«” ·«„"
        Me.TradingTransaction(3).Caption = "”‰œ«  «·’—›"
        Me.TradingTransaction(4).Caption = "«· ÕÊÌ· »Ì‰ «·„Œ«“‰"
        Me.TradingTransaction(5).Caption = "Ã—œ «·„Œ«“‰"
        TradingTransactionSub(0).Caption = "»œ√  Ã—œ «·„Œ«“‰"
        TradingTransactionSub(1).Caption = "ÿ»«⁄Â ﬂ‘Ê›«  «·Ã—œ"
        TradingTransactionSub(2).Caption = "«œŒ«· «·ﬂ„Ì«  «·›⁄·Ì…"
        TradingTransactionSub(3).Caption = " ‰›Ì– «·Ã—œ"

        Me.TradingTransaction(6).Caption = " ”ÊÌ… «·„Œ“Ê‰"
        Me.TradingTransaction(7).Caption = "”‰œ«  «·’—›"
        Me.TradingTransaction(8).Caption = " «·«” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰›"
        Me.TradingTransaction(9).Caption = "»ÕÀ ⁄‰ ”Ì—Ì«·"
        Me.TradingTransaction(10).Caption = "«·«’‰«› «· Ì »·€  Õœ «·ÿ·»"
        Me.TradingTransaction(11).Caption = "„Êﬁ› «·«’‰«› «·Õ«·Ì"
        Me.TradingTransaction(12).Caption = "«· ﬁ«—Ì—"

        TradingTransactionSub1(0).Caption = "”‰œ«  «·’—›"
        TradingTransactionSub1(1).Caption = "”‰œ«  ’—› «·Â«·ﬂ Ê«·⁄Ì‰« "

        Me.Purchase.Caption = "«·„‘ —Ì« "
        Me.PurchaseBasicRoot.Caption = "«·»Ì«‰«  «·«”«”Ì…"
        Me.PurchaseBasic(0).Caption = "»Ì«‰«  «·„Ê—œÌ‰"
        Me.PurchaseBasic(1).Caption = "⁄ﬁÊœ «·„Ê—œÌ‰"
        Me.PurchaseBasic(2).Caption = "«⁄œ«œ «⁄„«— «·œÌÊ‰"
        Me.PurchaseBasic(3).Caption = "ÿ—ﬁ «·‘Õ‰"
        Me.PurchaseBasic(4).Caption = "«‰Ê«⁄ «·÷„«‰« "
        Me.PurchaseBasic(5).Caption = "«⁄œ«œ«  «·«’‰«› «·—«ﬂœ…"

        Me.PurchaseTransactions(0).Caption = "⁄—Ê÷ «·«”⁄«— Ê ÿ·»«  «·‘—«¡ "
 
        PurchaseTransactionssubd(0).Caption = "⁄—Ê÷ «·«”⁄«—"
        PurchaseTransactionssubs(0).Caption = "ÿ·» ⁄—Ê÷ «”⁄«—"
        PurchaseTransactionssubs(1).Caption = "⁄—Ê÷ «·«”⁄«—"
        PurchaseTransactionssubs(2).Caption = "„ﬁ«—‰Â ⁄—Ê÷ «·«”⁄«—"

        PurchaseTransactionssubd(1).Caption = "√Ê«„— «·‘—«¡"
        PurchaseTransactionssubs1(0).Caption = "ÿ·» √„— ‘—«¡"
        PurchaseTransactionssubs1(1).Caption = "≈⁄ „«œ √„— ‘—«¡"
        PurchaseTransactionssubs1(2).Caption = "√„— ‘—«¡"

        FinAnalysis.Caption = "«· Õ·Ì· «·„«·Ì"
  
        Me.PurchaseTransactions(1).Caption = "»Ì«‰«  «·‘Õ‰"
        Me.PurchaseTransactions(2).Caption = "«·«⁄ „«œ«  «·„” ‰œÌ…"

        LCTransactions(0).Caption = " «‰Ê«⁄ «·«⁄ „«œ«  «·„” ‰œÌ…"
        LCTransactions(1).Caption = "«·›Ê« Ì— «·„»œ∆Ì…"
        LCTransactions(2).Caption = "› Õ «⁄ „«œ „” ‰œÌ"
        LCTransactions(3).Caption = " ⁄œÌ·  «⁄ „«œ „” ‰œÌ"
        LCTransactions(4).Caption = "„ «»⁄Â «·‘Õ‰« "
        LCTransactions(5).Caption = "”‰œ «” ·«„ ‘Õ‰« "
        LCTransactions(6).Caption = " ›« Ê—… ‰Â«∆Ì…"
        LCTransactions(7).Caption = "€·ﬁ «⁄ „«œ „” ‰œÌ"

        Me.PurchaseTransactions(3).Caption = "›« Ê—… „‘ —Ì« "
 
        Me.PurchaseTransactions(4).Caption = "„—œÊœ«  «·„‘ —Ì« "
        Me.PurchaseTransactions(5).Caption = " ﬁ—Ì— «⁄„«— «·œÌÊ‰"
        Me.PurchaseTransactions(6).Caption = " ﬁ«—Ì— «·„‘ —Ì« "
 
        Me.Sales.Caption = "«·„»Ì⁄« "
        Me.SalesBasic.Caption = "«·»Ì«‰«  «·«”«”Ì…"
        Me.SalesBasicSub(0).Caption = "«‰Ê«⁄ «·⁄„·«¡"
        Me.SalesBasicSub(1).Caption = "»Ì«‰«  «·⁄„·«¡"
        Me.SalesBasicSub(2).Caption = "⁄ﬁÊœ «·⁄„·«¡"
        Me.SalesBasicSub(3).Caption = "«⁄œ«œ «⁄„«— «·œÌÊ‰ "
        Me.SalesBasicSub(4).Caption = "»Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄"
        Me.SalesBasicSub(5).Caption = "»Ì«‰«  «·ﬂ«‘Ì—"
        Me.SalesBasicSub(6).Caption = "«⁄œ«œ Âœ› «·„»Ì⁄« "
        Me.SalesBasicSub(7).Caption = "„Ã„Ê⁄«  «·„‰«œÌ»"
        Me.SalesBasicSub(8).Caption = "»Ì«‰«  «·„‰«œÌ»"
   
        Me.SalesTransactions(0).Caption = "⁄—Ê÷ «·«”⁄«— Ê √Ê«„— «·»Ì⁄ "
 
        SalesTransactionssubss0(0).Caption = "⁄—Ê÷ «·«”⁄«—"
        SalesTransactionssubss00(0).Caption = "ÿ·» ⁄—Ê÷ «”⁄«— „»œ∆Ì… „‰ «·⁄„·«¡"
        SalesTransactionssubss00(1).Caption = "«⁄ „«œ ⁄—Ê÷ «·«”⁄«—"
        SalesTransactionssubss00(2).Caption = "⁄—Ê÷ «·«”⁄«— «·‰Â«∆Ì…"
   
        SalesTransactionssubss0(1).Caption = "√Ê«„— «·»Ì⁄"
        SalesTransactionssubss000(0).Caption = "ÿ·» «„— »Ì⁄"
        SalesTransactionssubss000(1).Caption = "≈⁄ „«œ √„— »Ì⁄"
        SalesTransactionssubss000(2).Caption = "√„— »Ì⁄"
  
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
        SalesTransactionsEmp(0).Caption = "«⁄œ«œ ⁄„Ê·«  «·„»Ì⁄«  Ê «· Õ’Ì·« "
        SalesTransactionsEmp(1).Caption = "ŒÿÂ   «·„»Ì⁄«  Ê «· Õ’Ì·« "
        SalesTransactionsEmp(2).Caption = "‰”»Â  Õﬁﬁ   ŒÿÂ ⁄„Ê·«  «·„»Ì⁄«  Ê «· Õ’Ì·« "
        SalesTransactionsEmp(3).Caption = "⁄„Ê·«  «·„‰«œÌ» «·„” Õ›…"
        SalesTransactionsEmp(4).Caption = "„ «»⁄Â “Ì«—«  «·⁄„·«¡"
        Archiving.Caption = "«·«—‘Ì› "
        ArchivingSub(0).Caption = "«÷«›… ‰„Ê–Ã ÃœÌœ"
 
        Me.Currency.Caption = "«·„⁄«„·«  «·„«·ÌÂ"
        Me.ExpensesType(0).Caption = "«‰Ê«⁄ «·„’—Ê›« "
        Me.ExpensesType(1).Caption = "  «‰Ê«⁄ «·«Ì—«œ« "
        Me.Expenses(0).Caption = "«·›Ê« Ì— «·„«·Ì…"
        Me.Expenses(1).Caption = "”‰œ«  «·’—›"
        ExpensesSub(0).Caption = "”‰œ«  «·’—›- Õ·Ì·Ì „’—Ê›«  "
        ExpensesSub(1).Caption = "”‰œ«  «·’—›- «·„œ›Ê⁄«  "
        
        '  Me.Payments(0).Caption = "«·„œ›Ê⁄« "

        Me.Cashing(0).Caption = "«·„ﬁ»Ê÷« "
        ' Me.Cashing(1).Caption = "›« Ê—… „‘—Ê⁄"
        Me.Cashing(2).Caption = "ÿ»«⁄Â «·‘Ìﬂ« "
        Me.Cashing(3).Caption = "«·«Ìœ«⁄«  «·»‰ﬂÌ…"
        Me.Cashing(4).Caption = " Õ’Ì·  Ê”œ«œ «·‘Ìﬂ« "
        Me.Cashing(5).Caption = "„–ﬂ—… »‰ﬂ  "
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
        Me.MnuBoxIncapacity_Increase.Caption = "“Ì«œ… Ê⁄Ã“ ›Ì ‰ﬁœÌ… «·Œ“Ì‰…"
        Me.MnuBoxStock.Caption = "Ã—œ «·Œ“Ì‰…"
        
        Me.MnuAccounts.Caption = "«·Õ”«»«  «·⁄«„Â"
        Me.MnuAccCharts(0).Caption = "  œ·Ì· «·Õ”«»« "
        Me.MnuAccCharts(1).Caption = " «·ﬁÌœ «·«›  «ÕÌ  "

        Me.Reports.Caption = "«· ﬁ«—Ì—"
        Me.Report.Caption = "«· ﬁ«—Ì— «·⁄«„…"
        Me.DailyReport.Caption = "«· ﬁ—Ì— «·ÌÊ„Ì"
        Me.MnuReports_Assblied.Caption = "«· ﬁ—Ì— «·„Ã„⁄ ⁄‰ › —…"
        Me.Tools.Caption = "„œÌ— «·‰Ÿ«„"
         
        Me.Barcode.Caption = " ’„Ì„ «·»«—ﬂÊœ..."
        Me.MnuPrintItemsCodes.Caption = "ÿ»«⁄Â «·»«—ﬂÊœ ..."
        Me.MnuCorrectSerial.Caption = " ⁄œÌ· ”Ì—Ì·«  «·«’‰«›"
        Me.MnuBoxDetectErrors.Caption = " ’ÕÌÕ «—’œ… «·Œ“‰"
        Me.MnuToolCustomers.Caption = " ⁄œÌ· ›Ê« Ì— «·⁄„·«¡"

        Me.MnuToolRepaireItemsCost.Caption = " ⁄œÌ· «· ﬂ·›… ›Ì ›Ê« Ì— «·»Ì⁄"
        Me.MnuToolsDataBase(0).Caption = " ÕœÌÀ «·« ’«· »ﬁ«⁄œ… «·»Ì«‰« "
        Me.MnuToolsDataBase(1).Caption = " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰«  "
        '        Me.MnuToolsDataBase(2).Caption = " €ÌÌ— ﬁ«⁄œ… «·»Ì«‰«  "
        Me.MnuDataBaseTools.Caption = "«œÊ«  ﬁ«⁄œ… «·»Ì«‰« "
        Me.UsersData.Caption = "«·„” Œœ„Ì‰"
        Me.AddUser.Caption = "«÷«›… „” Œœ„ ÃœÌœ..."
        Me.DelUser.Caption = "Õ–›  „” Œœ„  ..."
        Me.EditPw.Caption = " ⁄œÌ· «·—ﬁ„ ·”—Ì..."
        UserRpt.Caption = " ﬁ«—Ì— «·„” Œœ„Ì‰ "
            
        Me.UserAbility.Caption = "’·«ÕÌ«  «·„” Œœ„Ì‰..."
        Me.MnuUsersScreensPremission.Caption = "’·«ÕÌ«  «·„” Œœ„Ì‰ ⁄·Ï «·‘«‘« "
        Me.Options.Caption = "«⁄œ«œ«  «·‰Ÿ«„"
        Me.ShortCuts.Caption = "«·«Œ ’«—« "
         
        Me.MnuToolsSetPrinters0.Caption = "«⁄œ«œ «·ÿ«»⁄Â «·Õ«·Ì… ›Ì «·ÃÂ«“ «·Õ«·Ì..."
        Me.MnuToolsSetPrinters(1).Caption = " «⁄œ«œ«  œ·Ì· «·Õ”«»« "
        Me.MnuToolsSetPrinters(2).Caption = "«‰Ê«⁄ «·”‰œ« "
        Me.MnuToolsSetPrinters(3).Caption = "«·«ÿ·«⁄  ⁄·Ï  «· ‰»ÌÂ« "
         
        Me.MnuToolsSetPrinters(4).Caption = " ﬂÊÌœ «·”‰œ« "
        Me.MnuToolsSetPrinters(5).Caption = "  ﬂÊÌœ «·ÕﬁÊ·"
        Me.MnuToolsSetPrinters(6).Caption = "  «·—”«∆· «·œ«Œ·Ì…"
        Me.MnuToolsSetPrinters7.Caption = "≈⁄œ«œ«  —”«∆· «·ÃÊ«·"
         
        Me.MnuInterface.Caption = "«·Ê«ÃÂ…"
        Me.MnuInterfaceSub(0).Caption = "Ê«ÃÂÂ ⁄—»Ì…"
        Me.MnuInterfaceSub(1).Caption = "English Interface"
        Me.MnuWindowsList.Caption = "«·‘«‘«  «·„› ÊÕÂ"
        Me.MnuWindowsListOpen.Caption = "«·‘«‘«  «·„› ÊÕÂ"
        Me.Help.Caption = "„”«⁄œÂ"
        Me.HelpFile.Caption = "«·„Õ ÊÌ« ..."
        Me.HelpIndex.Caption = "«·œ·Ì·..."
        Me.SearchInHelp.Caption = "«·»ÕÀ..."
        Me.DailyToolTip.Caption = "‰’«∆Õ..."
        Me.MnuHelpForums.Caption = "„‰ œÏ «·œ⁄„ «·›‰Ì"
        Me.About.Caption = "⁄‰«..."
        Me.ConnectUs.Caption = " ”ÃÌ·..."
 
        prdo(0).Caption = "«·«‰ «Ã"

        prdo1(0).Caption = "»Ì«‰«  «·‘Ì› "
        prdo1(1).Caption = "»Ì«‰«  «·«·«  Ê «·„⁄œ« "
        prdo1(2).Caption = " ŒÿÊÿ «·«‰ «Ã"
        prosub1(0).Caption = " ⁄—Ì› ŒÿÊÿ «·«‰ «Ã"
        prosub1(1).Caption = " Œ’Ì’  Ê‰ﬁ· «·⁄„«· »Ì‰ ŒÿÊÿ «·«‰ «Ã"

        prdo1(3).Caption = "„—«Õ· «·«‰ «Ã"

        prdo1(4).Caption = "ÿ·»«  ‘—«¡ «·⁄„·«¡"
        prdo1(5).Caption = "«„— «·«‰ «Ã / «·‘€·"
        prdo1(6).Caption = "”‰œ ’—› „Ê«œ Œ«„ ··«‰ «Ã"
        prdo1(7).Caption = "”‰œ «” ·«„  «‰ «Ã  «„"

        prdo1(8).Caption = " ﬂ«·Ì› «·«‰ «Ã  «·‰„ÿÌ"
        prdo1(9).Caption = " Ê“Ì⁄ «· ﬂ«·Ì› €Ì— «·„»«‘—…"
        prdo1(10).Caption = " ﬁ«—Ì— «·«‰ «Ã"
 
        PrbH(0).Caption = " ”‰œ ’—› „—«Õ· «‰ «Ã"
        PrbH(1).Caption = " «„— «‰ «Ã ‰’› „’‰⁄"
        PrbH(2).Caption = " ”‰œ «” ·«„ «‰ «Ã ‰’› „’‰⁄"
 
        MnuLevels.Caption = "«⁄ „«œ «·„” ‰œ« "
        MnuLevelsSub(0).Caption = " ⁄—Ì› „” ÊÌ«  «·„” ‰œ« "
        MnuLevelsSub(1).Caption = " ⁄—Ì› «⁄ „«œ «·„” ‰œ« "
        MNUFixedAssets.Caption = "«·«’Ê· «·À«» …"
        xxxxx(0).Caption = "„Ã„Ê⁄«  «·«’Ê· «·À«» …"
        xxxxx(1).Caption = "»Ì«‰«  «·«’Ê· «·À«» …"
        xxxxx(2).Caption = "›Ê« Ì— ‘—«¡ «·«’Ê· «·À«» …"
        xxxxx(3).Caption = "«ﬁ”«ÿ «·«Â·«ﬂ «·«’Ê· «·À«» …"
        xxxxx(4).Caption = "«· Œ·’ «Ê «” »⁄«œ«  «·«’Ê· "
        xxxxx(5).Caption = "«÷«›«  «·«’Ê· "
        xxxxx(6).Caption = "‰ﬁ· «” ·«„ «·«’Ê· "
        xxxxx(7).Caption = " ﬁ«—Ì— "
        ArrowsBase.Caption = "«·«”Â„"
        ArrowsFollow(0).Caption = "»Ì«‰«  «·»Ê—’« "
        ArrowsFollow(1).Caption = "»Ì«‰«  „Ã„Ê⁄«  «·«”Â„"
        ArrowsFollow(2).Caption = "»Ì«‰«  «·‘—ﬂ« "
        ArrowsFollow(3).Caption = " Õ„Ì· «·«”⁄«—"
        ArrowsFollow(4).Caption = "  «·«”⁄«— «· «—ÌŒÌ…"
        ArrowsFollow(5).Caption = "«·„Õ«›Ÿ"

        ArrowsFollowBocket(0).Caption = " »Ì«‰«  «·„Õ«›Ÿ"
        ArrowsFollowBocket(1).Caption = "‘—«¡ «·«”Â„"
        ArrowsFollowBocket(2).Caption = "»Ì⁄ «·«”Â„"
        ArrowsFollowBocket(3).Caption = "«·ﬁÌ„… «·«”„Ì… ··«”Â„"

        ArrowsFollow(6).Caption = "„Ê«ﬁ⁄ Â«„…"
        ArrowsFollow(7).Caption = " ﬁ«—Ì—"

        MnuMaintnance.Caption = "«·’Ì«‰…"
        MnuMaintnanceBasic.Caption = "»Ì«‰«  «”«”Ì…"
        MnuMaintnanceBasicSub(0).Caption = "√‰Ê«⁄ «·’Ì«‰…"
        MnuMaintnanceBasicSub1.Caption = "‘—ﬂ«  «·’Ì«‰…"

        MnuMaintnanceTransactions(0).Caption = "œŒÊ· «·’Ì«‰…"
        MnuMaintnanceTransactions(1).Caption = "„Œ“‰ «·’Ì«‰…"
        MnuMaintnanceTransactions(2).Caption = "”‰œ ’—› ﬁÿ⁄ €Ì«— ··’Ì«‰…"

        MnuMaintnanceTransactions(3).Caption = " ”·Ì„ «·’Ì«‰…"
        MnuMaintnanceTransactions(4).Caption = "—ÃÊ⁄ ÷„«‰ „‰ „Ê—œ "
        MnuMaintnanceTransactions(5).Caption = "—’Ìœ «›  «ÕÌ „Œ“‰ «·’Ì«‰…"
        MnuMaintnanceTransactions(6).Caption = "Ã—œ „Œ“‰ «·’Ì«‰…"
        MnuMaintnanceTransactions(7).Caption = "«—”«·  ‰»Ì…  Ã„Ì⁄ «ÃÂ“…"
        MnuMaintnanceTransactions(8).Caption = " ﬁ«—Ì— «·’Ì«‰…"
 
        Tech.Caption = "√œÊ«  ›‰Ì…"
'        MnuManToolsSub5.Caption = "„ «»⁄Â «·’Ì«‰…"
 
 shipmentMnu.Caption = "«·‘Õ‰"

ShpmentBasicdata(0).Caption = "«·»Ì«‰«  «·«”«”ÌÂ"
ShpmentBasicdata(1).Caption = "«·»÷«∆⁄ ﬁÌœ «· ”·Ì„"
ShpmentBasicdata(2).Caption = " Œ’Ì’  «·‘«Õ‰« "
ShpmentBasicdata(3).Caption = " ”ÃÌ·  ÊﬁÌ «  «· ”·Ì„ "
ShpmentBasicdata(4).Caption = "„—œÊœ«  «·‘Õ‰"


ShpmentBasicdatasub(0).Caption = "»Ì«‰«  «·œÊ·"
ShpmentBasicdatasub(1).Caption = "»Ì«‰«  «·„Õ«›Ÿ«  Ê «·„‰«ÿﬁ"
ShpmentBasicdatasub(2).Caption = "«·„”«›«  »Ì‰ «·„œ‰"
ShpmentBasicdatasub(3).Caption = "»Ì«‰«  «·«ÕÌ«¡"
ShpmentBasicdatasub(4).Caption = "»Ì«‰«  «·‘Ê«—⁄"
ShpmentBasicdatasub(5).Caption = "«‰Ê«⁄ «·„—ﬂ»« "
ShpmentBasicdatasub(6).Caption = "»Ì«‰«  «·„—ﬂ»« "
ShpmentBasicdatasub(7).Caption = "»Ì«‰«  «·”«∆ﬁÌ‰"
 



    ElseIf SystemOptions.UserInterface = EnglishInterface Then
      POSTRansactiosG.Caption = "POS"

POSTRansactios(0).Caption = "Boxes Data"
POSTRansactios(1).Caption = "POS Data"
POSTRansactios(2).Caption = "Locations Data"
POSTRansactios(3).Caption = "Cashier Data"
POSTRansactios(4).Caption = "Login"
POSTRansactios(5).Caption = "Reports"

     
 shipmentMnu.Caption = "Shipping and Distribution"

ShpmentBasicdata(0).Caption = "Basic Data"
ShpmentBasicdata(1).Caption = "Non-delivered goods"
ShpmentBasicdata(2).Caption = "Allocation of vehicles"
ShpmentBasicdata(3).Caption = "Recording  delivery timing    "
ShpmentBasicdata(4).Caption = "Shipping returns"


ShpmentBasicdatasub(0).Caption = "Country data"
ShpmentBasicdatasub(1).Caption = "Cities Data"
ShpmentBasicdatasub(2).Caption = "Distance between Cities"
ShpmentBasicdatasub(3).Caption = "Neighborhoods Data "
ShpmentBasicdatasub(4).Caption = "Streets Data"
ShpmentBasicdatasub(5).Caption = "Vehicles Types"
ShpmentBasicdatasub(6).Caption = "Vehicles Data"
ShpmentBasicdatasub(7).Caption = "Drivers"

     MarketingMnu.Caption = "Marketing"
MarketingMnusub(0).Caption = "Sales items Plan"
MarketingMnusub(1).Caption = "Items Overs"
MarketingMnusub(2).Caption = "Customers Follow"


MarketingMnusubsub(0).Caption = "Register customer visits"
MarketingMnusubsub(1).Caption = "Follow customer visits"
MarketingMnusubsub(2).Caption = "Poll customers"
MarketingMnusubsub(3).Caption = "Customer complaint registration"
MarketingMnusubsub(4).Caption = "Customer complaint Follow"
MarketingMnusubsub(5).Caption = "Phone Directory"


'        MnuManToolsSub5.Caption = "Maintenance Follow"

        MnuMaintnance.Caption = "Maintenence"
        MnuMaintnanceBasic.Caption = "Basic Data"
        MnuMaintnanceBasicSub(0).Caption = "Maintenence Types"
        MnuMaintnanceBasicSub1.Caption = "Maintenence Companies"

        MnuMaintnanceTransactions(0).Caption = "Maintenance Order"
        MnuMaintnanceTransactions(1).Caption = "Maintenance Store"
        MnuMaintnanceTransactions(2).Caption = "Spare part Issue Voucher"

        MnuMaintnanceTransactions(3).Caption = "Maintenance Delivery"
        MnuMaintnanceTransactions(4).Caption = "Back Guarantee From The Supplier"
        MnuMaintnanceTransactions(5).Caption = "Opening Balance For Maintenance Store"
        MnuMaintnanceTransactions(6).Caption = "Maintenance Store Stock"
        MnuMaintnanceTransactions(7).Caption = "Send an alert collection devices"
        MnuMaintnanceTransactions(8).Caption = "Maintenance Reports"
        Tech.Caption = "Technical Tools"

        Me.BasicData.Caption = "Basic Data"
        Me.BasicDataM(0).Caption = " System Accounts Link"
        Me.BasicDataM(1).Caption = " Activity  And Branches"
        Me.BasicDataM(2).Caption = " Banks Data"
        Me.BasicDataM(3).Caption = " Boxes Data"
        Me.BasicDataM(4).Caption = " Payment  Type"
        Me.BasicDataM(5).Caption = " Vendors Data"
        Me.BasicDataM(6).Caption = " Customer Data"

        Me.BasicDataM(7).Caption = " Currency Data"
        Me.BasicDataM(8).Caption = " Nationality Data"
        Me.BasicDataM(9).Caption = " Religons Data"
        Me.BasicDataM(10).Caption = " Countries Data"
        Me.BasicDataM(11).Caption = " Government Data"
        Me.BasicDataM(12).Caption = " Neighborhoods Data"
        Me.BasicDataM(13).Caption = " Street Data"
        Me.BasicDataM(14).Caption = " Documents Type"
        Me.BasicDataM(15).Caption = " Items Data"
        Me.BasicDataM(17).Caption = "  Exit"
        FinAnalysis.Caption = "Fin. Analysis"
        AssetsMngBase.Caption = "RealState Mangement"
        mnuEmployee.Caption = "HR Mangement"
 
'        MnuItemTools_ItemCart.Caption = "Item Card"
        'MnuItemTools_ItemCostTrans.Caption = "Item Cost Price"
        'MnuItemTools_ItemData.Caption = "Items Data"
        'MnuItemTools_ItemQty.Caption = "Items Qty"
        'MnuItemTools_ItemSerial.Caption = "Items Serials"

        MnuAccDEV(0).Caption = " J L Entry"
        MnuAccDEV_Post.Caption = "Auditing   J LEntry"
        xxx(0).Caption = "Cost Centers Type"
        xxx(1).Caption = "Cost Centers"
        ProductionPlansub(0).Caption = "Production Plan"
        ProductionPlansub(1).Caption = "Defining QC Items"
        ProductionPlansub(2).Caption = "Production Classification "

        ProductionPlansub(3).Caption = "Register corrective action"
        ProductionPlansub(4).Caption = "Fully examine the quality of the product"
        ProductionPlansub(5).Caption = "Follow-up and repair of defective product registration"

        xxy(0).Caption = "Budget"
        ProductionPlan.Caption = " Planning and Quality Control"
        'xxx(4).Caption = "Financial Analysis"
        xxy(1).Caption = "Cash Flow"
        xxy(3).Caption = "Accounts Distribution"
        'xxx(7).Caption = "Prepare BalanceSheet"
        xxy(2).Caption = "View BalanceSheet"
        xxy(4).Caption = "perpare  Fin Equations"
        xxy(5).Caption = "View Fin Equations"

        xxy(6).Caption = "Composite Accounts"
        xxy(7).Caption = "Statistics"
        xxy(8).Caption = "Agenda customers"

        xxx(12).Caption = "Accounts Reports"

        Me.MnuProjects.Caption = "Projects Mangment"
        Me.MnuProjectsBasic.Caption = "Basic Data"
        Me.MnuProjectsBasicSub(0).Caption = "Projects Status"
        Me.MnuProjectsBasicSub(1).Caption = "Contract Type"

        Me.MnuProjectsBasicSub(2).Caption = "Sub-contractor  Data"
        Me.MnuProjectsBasicSub(3).Caption = "Projects Data"
        Me.MnuProjectsBasicSub(4).Caption = "Define Processes"
        Me.MnuProjectsBasicSub(5).Caption = "Projects Data"
              
        Me.MnuProjectsTransactions(0).Caption = "Project Row Of Matrial Issue Voucher"
        Me.MnuProjectsTransactions(1).Caption = "Projects Labors Allocate"
        Me.MnuProjectsTransactions(2).Caption = "Projects Labors Transfer"
        Me.MnuProjectsTransactions(3).Caption = "Follow Up Processes "
        Me.MnuProjectsTransactions(4).Caption = "Projects Invoice"
        Me.MnuProjectsTransactions(5).Caption = "Projects Reports"
 
        mnuEmployeeBasic(0).Caption = "Basic Data"
        mnuEmployeeBasicSub(0).Caption = "Prepare Company Attendance Times"
        mnuEmployeeBasicSub(1).Caption = "Shifts"
        mnuEmployeeBasicSub(2).Caption = "Vacations"
        mnuEmployeeBasicSub(3).Caption = "Contract Type"
        mnuEmployeeBasicSub(4).Caption = "Job Status"
        mnuEmployeeBasicSub(5).Caption = "Departrment Data"
        mnuEmployeeBasicSub(6).Caption = "Job Types Data"
        mnuEmployeeBasicSub(7).Caption = "Specifications Data"
        mnuEmployeeBasicSub(8).Caption = "Insurance Companies"
        mnuEmployeeBasicSub(9).Caption = "Insurance  Types"
        mnuEmployeeBasicSub(10).Caption = "Insurance  Classe"
        mnuEmployeeBasicSub(11).Caption = "Elements of assessment"
        mnuEmployeeBasic(2).Caption = "Atendance"
        EmployeeAttendanceSub(0).Caption = "Prepare Company Attendance Times"
        EmployeeAttendanceSub(0).Caption = "Prepare Employee Attendance Times"
        EmployeeAttendanceSub(1).Caption = " Attendance  Manual Record"
        EmployeeAttendanceSub(2).Caption = "Attendance  Auto Record"
        EmployeeAttendanceSub(3).Caption = "Absence Record"
        EmployeeAttendanceSub(4).Caption = "View Attendance Times"
        mnuEmployeeBasic(3).Caption = "Salaries"
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
        EmployeeSalarySub(11).Caption = "Register  End of service "
        EmployeeSalarySub(12).Caption = "Change Advance Due Date "

        mnuEmployeeBasic(1).Caption = "Employees Data"
        EmployeeDataicSub(0).Caption = "Employees Files"
        EmployeeDataicSub(1).Caption = "Employees Contracts"

        mnuEmployeeBasic(4).Caption = "Employees vacations"

        Vscstionsssub(0).Caption = "Vacations Plan"
        Vscstionsssub(1).Caption = "Vacations Request"
        Vscstionsssub(2).Caption = "Delivery and receipt of the era of in-kind"
        Vscstionsssub(3).Caption = "Vacations Dues"
        Vscstionsssub(4).Caption = "Record attendance of vacation"

        mnuEmployeeBasic(5).Caption = "Termination"
        FinishSevicersub(0).Caption = "Record Service Termination "
        FinishSevicersub(1).Caption = "Service Indemnity "

        TransporterMain.Caption = "Trasportation"
        TransporterSub(0).Caption = "Cities Data"
        TransporterSub(1).Caption = "Distance Cities Cities"
        TransporterSub(2).Caption = "Customer Data "
        TransporterSub(3).Caption = "Supplier Data"
        TransporterSub(4).Caption = "Driver Data"
        TransporterSub(5).Caption = "Vehicles Types"
        TransporterSub(6).Caption = "Insurance Company"
        TransporterSub(7).Caption = "Regular Maintenance Type"
        TransporterSub(8).Caption = "Vehicles Data"
        TransporterSub(9).Caption = "Trip Data"
        TransporterSub(10).Caption = "Reports"

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
        StockControlBasicSub(8).Caption = "Items Sales Price Names"

        StockControlBasicSub(9).Caption = "Production Cost component   "
        StockControlBasicSub(10).Caption = "Unit  Cost Of Production"
        StockControlBasicSub(11).Caption = "Plan For Items Sales "

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
        Me.TradingTransaction(7).Caption = "Issue Voucher"
        Me.TradingTransaction(8).Caption = "tems Qty Query"
        Me.TradingTransaction(9).Caption = "Items Serial Search"
        Me.TradingTransaction(10).Caption = "On Demand Items"
        Me.TradingTransaction(11).Caption = "Items Current Status"
        Me.TradingTransaction(12).Caption = "Reports"

        TradingTransactionSub1(0).Caption = "Issue  Vouchers  "
        TradingTransactionSub1(1).Caption = "Damage and Sample Issue  Vouchers"

        Me.Purchase.Caption = "Purchase "
        Me.PurchaseBasicRoot.Caption = "Basic Data"
        Me.PurchaseBasic(0).Caption = "Supplier Data"
        Me.PurchaseBasic(1).Caption = "Supplier Contract"
        Me.PurchaseBasic(2).Caption = "Prepare Ageing Data"
        Me.PurchaseBasic(3).Caption = "Shipment Method"
        Me.PurchaseBasic(4).Caption = "Gurantee Type"
        Me.PurchaseBasic(5).Caption = "Settings Items  stagnant"
 
        Me.PurchaseTransactions(0).Caption = "Quotations and Purchase Orders"
 
        PurchaseTransactionssubd(0).Caption = "Quotations"
        PurchaseTransactionssubs(0).Caption = "'Quotations Request"
        PurchaseTransactionssubs(1).Caption = "Quotations"
        PurchaseTransactionssubs(2).Caption = "Quotations Comparison Sheet"

        PurchaseTransactionssubd(1).Caption = "Purchase Orders"
        PurchaseTransactionssubs1(0).Caption = "Purchase Order Request"
        PurchaseTransactionssubs1(1).Caption = "Purchase Order Approval"
        PurchaseTransactionssubs1(2).Caption = "Purchase Order"

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

        Me.PurchaseTransactions(3).Caption = "Purchase Invoices"
 
        Me.PurchaseTransactions(4).Caption = "Return Purchase"
        Me.PurchaseTransactions(5).Caption = "Ageing Report"
        Me.PurchaseTransactions(6).Caption = "Purchase Reports"
 
        Me.Sales.Caption = "Sales "
 
        Me.SalesBasic.Caption = "Basic Data"
        Me.SalesBasicSub(0).Caption = "Customers Type"
        Me.SalesBasicSub(1).Caption = "Customers Data"
        Me.SalesBasicSub(2).Caption = "Cusettomers Contract"
        Me.SalesBasicSub(3).Caption = "Perpare Ageing "
        Me.SalesBasicSub(4).Caption = "POS Data"
        Me.SalesBasicSub(5).Caption = "Cashier Data"
        Me.SalesBasicSub(6).Caption = "Prepare Sales Target"
        Me.SalesBasicSub(7).Caption = "Sales Rep Groups"
        Me.SalesBasicSub(8).Caption = "Sales Rep Data"
   
        Me.SalesTransactions(0).Caption = "Quotations and Sales Orders"
 
        SalesTransactionssubss0(0).Caption = "Quotations"
        SalesTransactionssubss00(0).Caption = "Customes Quotations"
        SalesTransactionssubss00(1).Caption = "Quotations Approval  "
        SalesTransactionssubss00(2).Caption = "Final Quotations"
   
        SalesTransactionssubss0(1).Caption = "Sales Orders"
        SalesTransactionssubss000(0).Caption = "Sales Orders Request"
        SalesTransactionssubss000(1).Caption = "Sales Orders Approval"
        SalesTransactionssubss000(2).Caption = "Sales Orders"
  
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
        SalesTransactionsEmp(0).Caption = "Preparation of sales commissions and collections"

        SalesTransactionsEmp(1).Caption = "sales commissions and collections Plan"
        SalesTransactionsEmp(2).Caption = "Ratios achieve the objectives of sales and collections"

        SalesTransactionsEmp(3).Caption = "Commissions receivable For SalesPersons"
        SalesTransactionsEmp(4).Caption = "Customers Visits Follow"
        Archiving.Caption = "Electronic Archiving"
        ArchivingSub(0).Caption = "Add new Form"
   
        Me.Currency.Caption = "Fi&nancial Transactions"
        Me.ExpensesType(0).Caption = "Expenses Types"
        Me.ExpensesType(1).Caption = "Revenues Types"
        Me.Expenses(0).Caption = "Financial Invoice"
        Me.Expenses(1).Caption = "Expenses Voucher"
            
        ExpensesSub(0).Caption = "Expenses Voucher - Detailed "
        ExpensesSub(1).Caption = "Expenses Voucher-Payments "
        
        Me.Payments(0).Caption = "Notes Payable"

        Me.Cashing(0).Caption = "Notes Receivable"
        Me.Cashing(1).Caption = "-"
        Me.Cashing(2).Caption = "Print Cheque"
        Me.Cashing(3).Caption = "Bank Deposite"
        Me.Cashing(4).Caption = "cheque Release"
        Me.Cashing(5).Caption = "Bank Report"
        
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
        Me.MnuBoxDeposit(2).Caption = "Era Close"

        Me.MnuBoxDrawing.Caption = "Transfer Money "
        Me.MnuBoxAccouns.Caption = "Current Box Balance"
        Me.MnuBoxIncapacity_Increase.Caption = "Box Incapacity && Increase"
        Me.MnuBoxStock.Caption = "Box Stock"
        
        Me.MnuAccounts.Caption = "Accounting"
        Me.MnuAccCharts(0).Caption = "Chart Of Accounts"
        Me.MnuAccCharts(1).Caption = "Accounts Opening Balance"
        '
        '
        
        Me.Reports.Caption = "Reports"
        Me.Report.Caption = "General Reports"
        Me.DailyReport.Caption = "Daily Reports"
        Me.MnuReports_Assblied.Caption = "Assblied Interval Report"
        Me.Tools.Caption = "System Manger"
         
        Me.Barcode.Caption = "Barcode Design..."
        Me.MnuPrintItemsCodes.Caption = "Items Codes Barcode Print..."
        Me.MnuCorrectSerial.Caption = "Repaire Items Serial Number Errors"
        Me.MnuBoxDetectErrors.Caption = "Repaire Box Balance Errors"
        Me.MnuToolCustomers.Caption = "Edit Customers Invoices"

        Me.MnuToolRepaireItemsCost.Caption = "Adjust Items Cost in Bill Invoices"
        Me.MnuToolsDataBase(0).Caption = "Refresh DataBase Connectoion"
        Me.MnuToolsDataBase(1).Caption = "Update DataBase "
        '         Me.MnuToolsDataBase(2).Caption = "Change DataBase "
        Me.MnuDataBaseTools.Caption = "Data Base Tools"
        Me.UsersData.Caption = "Users"
        Me.AddUser.Caption = "Add New  User..."
        Me.DelUser.Caption = "Delete User..."
        Me.EditPw.Caption = "Change Password..."
        UserRpt.Caption = "Users Log File   "
        Me.UserAbility.Caption = "Users Premissions..."
        Me.MnuUsersScreensPremission.Caption = "Users Screens Premission"
        Me.Options.Caption = "Options"
        Me.ShortCuts.Caption = "Shortcuts"
         
        Me.MnuToolsSetPrinters0.Caption = "Set Local Printer..."
        Me.MnuToolsSetPrinters(1).Caption = "Accounts Coding"
        Me.MnuToolsSetPrinters(2).Caption = "Doc Type  "
        Me.MnuToolsSetPrinters(3).Caption = "Show Alarms "
         
        Me.MnuToolsSetPrinters(4).Caption = "Voucher Coding"
        Me.MnuToolsSetPrinters(5).Caption = "Fields Coding"
        Me.MnuToolsSetPrinters(6).Caption = " Local Messenger "
        Me.MnuToolsSetPrinters7.Caption = " SMS Settings "
        
        Me.MnuInterface.Caption = "User Interface"
        Me.MnuInterfaceSub(0).Caption = "Arabic Interface"
        Me.MnuInterfaceSub(1).Caption = "English Interface"
        Me.MnuWindowsList.Caption = "Programe Windows"
        Me.MnuWindowsListOpen.Caption = "Opened Windows"
        Me.Help.Caption = "Help"
        Me.HelpFile.Caption = "Contents..."
        Me.HelpIndex.Caption = "Index..."
        Me.SearchInHelp.Caption = "Search..."
        Me.DailyToolTip.Caption = "Daily Tool Tip..."
        Me.MnuHelpForums.Caption = "Technical Support Forums"
        Me.About.Caption = "About..."
        Me.ConnectUs.Caption = "Register..."
 
        prdo(0).Caption = "Production"

        prdo1(0).Caption = "Shifts Data"
        prdo1(1).Caption = "Equipments Data"
        prdo1(2).Caption = "Production Lines "
        prosub1(0).Caption = "Define Production Lines"
        prosub1(1).Caption = "Allocate and Trannsfer Employee "

        prdo1(3).Caption = "Production Cycle"

        prdo1(4).Caption = " Purchase Order"
        prdo1(5).Caption = "Production/Work Order"
        prdo1(6).Caption = "Issue Voucher-Row Material Items"
        prdo1(7).Caption = "Receive Voucher- Production Items"

        prdo1(8).Caption = "Typical production costs"
        prdo1(9).Caption = "Indirect Costs Distributions"
        prdo1(10).Caption = "Production Reports"
 
        PrbH(0).Caption = "Production Issue Voucher"
        PrbH(1).Caption = " Production work order"
        PrbH(2).Caption = "Production Recieve Voucher "
 
        MNUFixedAssets.Caption = "FixedAssets"
        xxxxx(0).Caption = "Fixed Assets Groups"
        xxxxx(1).Caption = "Fixed Assets Data"
        xxxxx(2).Caption = "Fixed Assets Invoice"
        xxxxx(3).Caption = "Depreciation Installments Issueing"
        xxxxx(4).Caption = " Disposal  OF F.A."
        xxxxx(5).Caption = "FA Additions"
        xxxxx(6).Caption = "Delivering and receiving assets"
        xxxxx(7).Caption = "Reports"
 
        MnuLevels.Caption = "Documents Approvals"
        MnuLevelsSub(0).Caption = "Approval Levels"
        MnuLevelsSub(1).Caption = "Approval for Documents"
 
        ArrowsBase.Caption = "Arrows Mangements"
        ArrowsFollow(0).Caption = "Capital Market Data"
        ArrowsFollow(1).Caption = "Groups of Arrows"
        ArrowsFollow(2).Caption = "Companies Data"
        ArrowsFollow(3).Caption = "Loading Prices"
        ArrowsFollow(4).Caption = "Historical prices"
        ArrowsFollow(5).Caption = "Bockets"

        ArrowsFollowBocket(0).Caption = "Bockets Data"
        ArrowsFollowBocket(1).Caption = "Arrows Purchases"
        ArrowsFollowBocket(2).Caption = "Arrows Salling"
        ArrowsFollowBocket(3).Caption = "Arrows Current Value"

        ArrowsFollow(6).Caption = "Links"
        ArrowsFollow(7).Caption = "Reports"

        '
        'Me.MnuPopItemsTreePane_Array(0).Caption = "Refresh"
        'Me.MnuPopItemsTreePane_Array(2).Caption = "Dock"
        'Me.MnuPopItemsTreePane_Array(3).Caption = "Close"
        'Me.MnuPopItemsTreePane_Array(5).Caption = "Groups Sort"
        'Me.MPITP_GSort_Option(0).Caption = "Group ID (Ascending)"
        'Me.MPITP_GSort_Option(1).Caption = "Group ID (Descending)"
        'Me.MPITP_GSort_Option(2).Caption = "-"
        'Me.MPITP_GSort_Option(3).Caption = "Group Code (Ascending)"
        'Me.MPITP_GSort_Option(4).Caption = "Group Code (Descending)"
        'Me.MPITP_GSort_Option(5).Caption = "-"
        'Me.MPITP_GSort_Option(6).Caption = "Group Name (Ascending)"
        'Me.MPITP_GSort_Option(7).Caption = "Group Name (Descending)"
        'Me.MnuPopItemsTreePane_Array(6).Caption = "-"
        'Me.MnuPopItemsTreePane_Array(7).Caption = "Items Sort"
        'Me.MPITP_ISort_Option(0).Caption = "Item ID (Ascending)"
        'Me.MPITP_ISort_Option(1).Caption = "Item ID (Descending)"
        'Me.MPITP_ISort_Option(2).Caption = "-"
        'Me.MPITP_ISort_Option(3).Caption = "Item Code (Ascending)"
        'Me.MPITP_ISort_Option(4).Caption = "Item Code (Descending)"
        'Me.MPITP_ISort_Option(5).Caption = "-"
        'Me.MPITP_ISort_Option(6).Caption = "Item Name (Ascending)"
        '            Me.MPITP_ISort_Option(7).Caption = "Item Name (Descending)"
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

    IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

    If IntRes = vbYes Then
        'End
        '    Exit Function
        AddToLogFile CInt(user_id), 0, Date, Time, "  ”ÃÌ·   «·Œ—ÊÃ „‰ «·‰Ÿ«„ ", " System LogOut", Me.name, "L", "", ""
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
    xHelpPane.Options = PaneHasMenuButton
    
    If SystemOptions.SysMantainceAllow = True Then
        Set xMantaincePane = Me.DockingPane1.CreatePane(DockingPanesIDs.MantainceID, 250, 200, DockLeftOf, Nothing)

        If SystemOptions.UserInterface = ArabicInterface Then
            xMantaincePane.Title = "«·’Ì«‰…"
        Else
            xMantaincePane.Title = "Mantaince"
        End If

        xMantaincePane.Options = PaneHasMenuButton
        '    xMantaincePane.IconId = Me.ImgLstMenuIcons.ListImages("Tools").Index
    End If

    Set xCalendarPane = Me.DockingPane1.CreatePane(DockingPanesIDs.CalendarPaneID, 250, 250, DockLeftOf, Nothing)
    '    xCalendarPane.IconId = Me.ImgLstMenuIcons.ListImages("OpenAcc").Index
    xCalendarPane.Options = PaneHasMenuButton
    
    If SystemOptions.UserInterface = ArabicInterface Then
        x.Title = "„⁄·Ê„«  «·»—‰«„Ã"
        Y.Title = "‘—Ìÿ «·√Œ ’«—« "
        xItemsTreePane.Title = "‘Ã—… «·√’‰«›"
        xInternetPane.Title = "√Œ»«— «·√‰ —‰ "
        xHelpPane.Title = "«·„”«⁄œ… «··ÕŸÌ…"
        xCalendarPane.Title = "«·”«⁄…"
    Else
        x.Title = "Information OutBar"
        Y.Title = "Shortcut OutBar"
        xItemsTreePane.Title = "Items Tree"
        xInternetPane.Title = "Internet News"
        xHelpPane.Title = "Dynamic Help"
        xCalendarPane.Title = "Calendar"
    End If

    DockingPane1.VisualTheme = ThemeVisio
    DockingPane1.HidePane x
    DockingPane1.HidePane xItemsTreePane
    DockingPane1.HidePane xInternetPane
    DockingPane1.HidePane xCalendarPane

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

    If Not FrmOutBarPane Is Nothing Then
        Unload FrmOutBarPane
    End If

    If Not FrmNewsBarPane Is Nothing Then
        Unload FrmNewsBarPane
    End If

    If Not ItemsTreePane Is Nothing Then
        Unload ItemsTreePane
    End If

    If Not FrmDynamicHelpPane Is Nothing Then
        Unload FrmDynamicHelpPane
    End If

    If Not FrmCalendarPane Is Nothing Then
        Unload FrmCalendarPane
    End If

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
        xPaneRec.PaneID = Me.DockingPane1.Panes(i).id
        xPaneRec.PanePositon = Me.DockingPane1(i).Position
        xPaneRec.PaneTitle = Me.DockingPane1(i).Title
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
    Dim i As Integer, J As Integer
    Dim Lparent As Long
    Dim BolTemp As Boolean
    Dim IntCount As Integer
    Dim StrOldFrmName As String

    If mdifrmmain.ActiveForm Is Nothing Then
        Me.PopMenu1.ClearSubMenusOfItem ("MnuWindowsListOpen")
        MnuWindowsListOpen.Enabled = False
        Exit Sub
    Else
        MnuWindowsListOpen.Enabled = True
    End If

    Me.PopMenu1.ClearSubMenusOfItem ("MnuWindowsListOpen")

    For i = 0 To Forms.count - 1

        If Forms(i).name <> "MDIFrmMain" Then
            If Forms(i).MDIChild = True Then

                With Me.PopMenu1
                    Lparent = .MenuIndex("MnuWindowsListOpen")

                    If ImgInImgList(Forms(i).name) = -1 Then
                        Dim CCCC As Long
                        'Me.ImgLstMenuIcons.ListImages.Add , Forms(I).name, Forms(I).Icon
                        'me.ImgLstMenuIcons.ListImages.Add
                        'cccc=me.ImgLstMenuIcons.ListImages(forms(i).name).
                        Dim xx As IPictureDisp
                        Set xx = Forms(i).Icon
                        Me.ilsIcons.AddFromHandle xx.Handle, IMAGE_ICON, Forms(i).name
                    End If

                    BolTemp = False

                    For J = 1 To .count

                        If StrOldFrmName <> Forms(i).name Then
                            IntCount = 0
                            StrOldFrmName = Forms(i).name
                        End If

                        If .MenuKey(J) = Forms(i).name Then
                            IntCount = IntCount + 1
                            StrOldFrmName = Forms(i).name
                            BolTemp = True
                        End If

                    Next J

                    If BolTemp = False Then
                        .AddItem Forms(i).Caption, Forms(i).name, , 1000 + .count, Lparent, Me.ilsIcons.ItemIndex(Forms(i).name) - 1, True, True
                    ElseIf BolTemp = True Then
                        .AddItem Forms(i).Caption & " " & IntCount, , , 1000 + .count, Lparent, Me.ilsIcons.ItemIndex(Forms(i).name) - 1, True, True
                    End If

                    If mdifrmmain.ActiveForm.name = Forms(i).name Then
                        .MenuDefault(Forms(i).name) = True
                    Else
                        .MenuDefault(Forms(i).name) = False
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

            mowazna.show

        Case 4
            tahlil_maly.show

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

            If checkApility("BaklanceSheet") = False Then
                Exit Sub
            End If

            BaklanceSheet.show

        Case 8

            If checkApility("BaklanceSheetvIEW") = False Then
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

            If checkApility("FrmExpenses40") = False Then
                Exit Sub
            End If
    
            'FrmExpenses40.Show
            FrmExpenses40E.show

        Case 5
            FrmExpenses40A.show

        Case 6
            FrmExpensesT.show

        Case 7

            If checkApility("ShowFixedAssets") = False Then
                Exit Sub
            End If
    
            frmFixedAsseteports.show

    End Select

End Sub

Private Sub xxy_Click(Index As Integer)

    Select Case Index

        Case 0

            If checkApility("mowazna") = False Then
                Exit Sub
            End If

            mowazna.show
 
        Case 1

            If checkApility("Cash_flow") = False Then
                Exit Sub
            End If

            Cash_flow.show

        Case 2

            If checkApility("BaklanceSheetvIEW") = False Then
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

            If checkApility("FrmStatistics") = False Then
                Exit Sub
            End If

            OpenScreen StatisticsShow

        Case 8

            If checkApility("FrmCustomersAgenda") = False Then
                Exit Sub
            End If

            FrmCustomersAgenda.show

        Case 9
            FrmBalanceSheet1.show

    End Select

End Sub
