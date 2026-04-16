VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbaListView6.ocx"
Begin VB.Form FRMPOS 
   ClientHeight    =   10950
   ClientLeft      =   4110
   ClientTop       =   3120
   ClientWidth     =   20250
   Icon            =   "frmTestListView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame5 
      Height          =   10935
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   20295
      Begin VB.PictureBox imgLarge 
         BackColor       =   &H80000005&
         Height          =   480
         Left            =   4320
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   43
         Top             =   120
         Width           =   1200
      End
      Begin VB.CheckBox chkGroupView 
         Caption         =   "&Grouped View"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   42
         Top             =   480
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   1200
         TabIndex        =   20
         Top             =   -4680
         Visible         =   0   'False
         Width           =   6495
         Begin VB.PictureBox picOptions 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2790
            Left            =   0
            ScaleHeight     =   2790
            ScaleWidth      =   12150
            TabIndex        =   21
            Top             =   0
            Width           =   12150
            Begin VB.CheckBox chkCustomDraw 
               Caption         =   "Custom Draw"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2460
               TabIndex        =   39
               Top             =   2520
               Value           =   1  'Checked
               Width           =   2775
            End
            Begin VB.CheckBox chkBorderSelect 
               Caption         =   "&Border Select (Large Icons)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   38
               Top             =   2040
               Width           =   2295
            End
            Begin VB.CheckBox chkAutoArrange 
               Caption         =   "Auto-Arran&ge"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   37
               Top             =   300
               Value           =   1  'Checked
               Width           =   2295
            End
            Begin VB.CheckBox chkHeaderDragDrop 
               Caption         =   "&Header Drag-Drop (Report)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   36
               Top             =   1800
               UseMaskColor    =   -1  'True
               Width           =   2295
            End
            Begin VB.CheckBox chkEnabled 
               Caption         =   "&Enabled"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   35
               Top             =   60
               Value           =   1  'Checked
               Width           =   2295
            End
            Begin VB.ComboBox cboBorder 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3360
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   60
               Width           =   2235
            End
            Begin VB.ComboBox cboAppearance 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3360
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   420
               Width           =   2235
            End
            Begin VB.CheckBox chkHideSelection 
               Caption         =   "&Hide Selection"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   32
               Top             =   840
               Width           =   2295
            End
            Begin VB.CheckBox chkMultiSelect 
               Caption         =   "&Multi-Select"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   31
               Top             =   1080
               Value           =   1  'Checked
               Width           =   2295
            End
            Begin VB.CheckBox chkBackground 
               Caption         =   "&Background Picture"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2460
               TabIndex        =   30
               Top             =   840
               Width           =   2235
            End
            Begin VB.CheckBox chkInfoTips 
               Caption         =   "&Info Tips"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2460
               TabIndex        =   29
               Top             =   1080
               Value           =   1  'Checked
               Width           =   2235
            End
            Begin VB.CheckBox chkLabelEdit 
               Caption         =   "Label Edi&t"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   28
               Top             =   1320
               Width           =   2295
            End
            Begin VB.CheckBox chkGridLines 
               Caption         =   "&Gridlines (Report)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2460
               TabIndex        =   27
               Top             =   1320
               Width           =   2235
            End
            Begin VB.CheckBox chkHeaderButtons 
               Caption         =   "&Header Buttons (Report)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   26
               Top             =   1560
               Value           =   1  'Checked
               Width           =   2295
            End
            Begin VB.CheckBox chkSubItemImages 
               Caption         =   "&Sub-Item Images (Report)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2460
               TabIndex        =   25
               Top             =   1560
               Width           =   2235
            End
            Begin VB.CheckBox chkCheckBoxes 
               Caption         =   "&Check Boxes"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2460
               TabIndex        =   24
               Top             =   1800
               Width           =   2235
            End
            Begin VB.CheckBox chkFlatScrollBars 
               Caption         =   "&Flat Scroll Bars"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2460
               TabIndex        =   23
               Top             =   2040
               Width           =   2235
            End
            Begin VB.CheckBox chkFullRowSelect 
               Caption         =   "F&ull Row Select (Report or Tile)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2460
               TabIndex        =   22
               Top             =   2280
               Width           =   2775
            End
            Begin VB.Label lblInfo 
               Caption         =   "BorderStyle:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   2400
               TabIndex        =   41
               Top             =   120
               Width           =   915
            End
            Begin VB.Label lblInfo 
               Caption         =   "Appearance:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   2400
               TabIndex        =   40
               Top             =   480
               Width           =   915
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4935
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   0
         Width           =   4815
         Begin VB.Image Image1 
            Height          =   4845
            Left            =   0
            Picture         =   "frmTestListView.frx":000C
            Top             =   0
            Width           =   4845
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   360
            Width           =   1965
         End
         Begin VB.Label lblqty 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   1200
            Width           =   3525
         End
         Begin VB.Label lBLnO 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   615
            Index           =   0
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label lBLnO 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   615
            Index           =   1
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label lBLnO 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   615
            Index           =   2
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label lBLnO 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   615
            Index           =   3
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label lBLnO 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   615
            Index           =   6
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label lBLnO 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   615
            Index           =   7
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lBLnO 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   615
            Index           =   8
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lBLnO 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   615
            Index           =   9
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lBLclr 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   1455
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label LBLdOT 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   735
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   4440
            Width           =   975
         End
         Begin VB.Label lBLnO 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   615
            Index           =   4
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label lBLnO 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   615
            Index           =   5
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label LblUserName 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   4440
            Width           =   3045
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   1695
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         Begin vbalIml6.vbalImageList ilsIcons32 
            Left            =   120
            Top             =   240
            _ExtentX        =   953
            _ExtentY        =   953
            IconSizeX       =   48
            IconSizeY       =   48
            ColourDepth     =   24
            Size            =   9660
            Images          =   "frmTestListView.frx":80C3
            Version         =   131072
            KeyCount        =   1
            Keys            =   ""
         End
         Begin vbalIml6.vbalImageList ilsIcons16 
            Left            =   0
            Top             =   720
            _ExtentX        =   953
            _ExtentY        =   953
            IconSizeX       =   48
            IconSizeY       =   48
            ColourDepth     =   24
            Size            =   48300
            Images          =   "frmTestListView.frx":A69F
            Version         =   131072
            KeyCount        =   5
            Keys            =   "ˇˇˇˇ"
         End
         Begin vbalIml6.vbalImageList ilsIcons48 
            Left            =   0
            Top             =   0
            _ExtentX        =   953
            _ExtentY        =   953
            IconSizeX       =   48
            IconSizeY       =   48
            ColourDepth     =   24
            Size            =   48300
            Images          =   "frmTestListView.frx":1636B
            Version         =   131072
            KeyCount        =   5
            Keys            =   "ˇˇˇˇ"
         End
         Begin VB.Label lblStatus 
            Alignment       =   1  'Right Justify
            Caption         =   "Label10"
            Height          =   495
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   960
            Width           =   135
         End
      End
      Begin vbalListViewLib6.vbalListViewCtl lvwMain 
         Height          =   8775
         Left            =   17880
         TabIndex        =   3
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   15478
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiSelect     =   -1  'True
         LabelEdit       =   0   'False
         AutoArrange     =   0   'False
         HeaderButtons   =   0   'False
         HeaderTrackSelect=   0   'False
         HideSelection   =   0   'False
         InfoTips        =   0   'False
      End
      Begin vbalListViewLib6.vbalListViewCtl lvwItems 
         Height          =   8775
         Left            =   8040
         TabIndex        =   44
         Top             =   600
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   15478
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiSelect     =   -1  'True
         LabelEdit       =   0   'False
         AutoArrange     =   0   'False
         HeaderButtons   =   0   'False
         HeaderTrackSelect=   0   'False
         HideSelection   =   0   'False
         InfoTips        =   0   'False
      End
      Begin vbalListViewLib6.vbalListViewCtl lvwTables 
         Height          =   7215
         Left            =   0
         TabIndex        =   45
         Top             =   2160
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   12726
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         MultiSelect     =   -1  'True
         LabelEdit       =   0   'False
         AutoArrange     =   0   'False
         HeaderButtons   =   0   'False
         HeaderTrackSelect=   0   'False
         HideSelection   =   0   'False
         InfoTips        =   0   'False
      End
      Begin vbalIml6.vbalImageList GrouplImageList 
         Left            =   0
         Top             =   600
         _ExtentX        =   953
         _ExtentY        =   953
         IconSizeX       =   48
         IconSizeY       =   48
         ColourDepth     =   24
         Size            =   48300
         Images          =   "frmTestListView.frx":22037
         Version         =   131072
         KeyCount        =   5
         Keys            =   "ˇˇˇˇ"
      End
      Begin VB.Image Image8 
         Height          =   675
         Left            =   0
         Stretch         =   -1  'True
         Top             =   960
         Width           =   3195
      End
      Begin VB.Image Image7 
         Height          =   435
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   3075
      End
      Begin VB.Image Image4 
         Height          =   795
         Left            =   0
         Stretch         =   -1  'True
         Top             =   120
         Width           =   3195
      End
      Begin VB.Image Image3 
         Height          =   435
         Left            =   17880
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2355
      End
      Begin VB.Image Image2 
         Height          =   435
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   120
         Width           =   9795
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ã„Ê⁄« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   0
         Left            =   18120
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   120
         Width           =   1965
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«’‰«›"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   21120
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ÿ«Ê·« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1680
         Width           =   1965
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ã„Ê⁄« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   120
         Width           =   1965
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«’‰«›"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   120
         Width           =   8205
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ÿ·»«  «·Œ«—ÃÌ…"
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
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   360
         Width           =   1965
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "Œœ„… «· Ê’Ì·"
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
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1080
         Width           =   1965
      End
   End
End
Attribute VB_Name = "FRMPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 

Private Sub cboAppearance_Click()
   lvwMain.Appearance = cboAppearance.ItemData(cboAppearance.ListIndex)
End Sub

Private Sub cboBorder_Click()
   lvwMain.BorderStyle = cboBorder.ItemData(cboBorder.ListIndex)
End Sub

Private Sub cboView_Click()
 '  If cboView.ListIndex > -1 Then
 '     lvwMain.View = cboView.ItemData(cboView.ListIndex)
 '  End If
End Sub

Private Sub chkAutoArrange_Click()
   lvwMain.AutoArrange = (chkAutoArrange.value = vbChecked)
End Sub

Private Sub chkBackground_Click()
   If chkBackground.value = Checked Then
      lvwMain.BackColor = -1
      lvwMain.BackgroundPicture = App.path & "\back.jpg"
   Else
      lvwMain.BackColor = vbWindowBackground
      lvwMain.BackgroundPicture = ""
   End If
End Sub

Private Sub chkBorderSelect_Click()
   lvwMain.ItemBorderSelect = (chkBorderSelect.value = Checked)
End Sub

Private Sub chkCheckBoxes_Click()
   lvwMain.Checkboxes = (chkCheckBoxes.value = Checked)
End Sub

Private Sub chkCustomDraw_Click()
   lvwMain.CustomDraw = (chkCustomDraw.value = Checked)
End Sub

Private Sub chkEnabled_Click()
   lvwMain.Enabled = (chkEnabled.value = Checked)
End Sub

Private Sub chkFlatScrollBars_Click()
   lvwMain.FlatScrollBar = (chkFlatScrollBars.value = Checked)
End Sub

Private Sub chkFullRowSelect_Click()
   lvwMain.FullRowSelect = (chkFullRowSelect.value = Checked)
End Sub

Private Sub chkGridLines_Click()
   lvwMain.GridLines = (chkGridLines.value = Checked)
End Sub

Private Sub chkGroupView_Click()
   
   ' very slow unless we do this
   lvwMain.Visible = False
   If (chkGroupView.value = vbChecked) Then
      Dim i As Long
      
      ' Create three groups and display them on screen:
      lvwMain.ItemGroups.Enabled = True
      
      If (lvwMain.ItemGroups.count = 0) Then
         ' Create a group and add the first five items to it:
         Dim cG As cItemGroup
         Set cG = lvwMain.ItemGroups.Add(1, "GROUP1", "First Five Items")
         Debug.Print cG.Header
         For i = 1 To 5
            lvwMain.Listitems(i).Group = cG
         Next i
         
         ' Create a group and add the next ten items:
         Set cG = lvwMain.ItemGroups.Add(5, "GROUP2", "Next Ten Items")
         For i = 6 To 15
            lvwMain.Listitems(i).Group = cG
         Next i
         
         ' And the rest:
         Set cG = lvwMain.ItemGroups.Add(15, "GROUP3", "The Remainder")
         For i = 16 To lvwMain.Listitems.count
            lvwMain.Listitems(i).Group = cG
         Next i
      End If
      
   Else
      ' Hide all the groups:
      lvwMain.ItemGroups.Enabled = False
      
   End If
   lvwMain.Visible = True
   
End Sub

Private Sub chkHeaderButtons_Click()
   lvwMain.HeaderButtons = (chkHeaderButtons.value = Checked)
End Sub

Private Sub chkHeaderDragDrop_Click()
   lvwMain.HeaderDragDrop = (chkHeaderDragDrop.value = Checked)
End Sub

Private Sub chkHideSelection_Click()
   lvwMain.HideSelection = (chkHideSelection.value = Checked)
End Sub

Private Sub chkInfoTips_Click()
   lvwMain.InfoTips = (chkInfoTips.value = Checked)
End Sub

Private Sub chkLabelEdit_Click()
   lvwMain.LabelEdit = (chkLabelEdit.value = Checked)
End Sub

Private Sub chkMultiSelect_Click()
   lvwMain.MultiSelect = (chkMultiSelect.value = Checked)
End Sub

Private Sub chkSubItemImages_Click()
Dim i As Long
   lvwMain.SubItemImages = (chkSubItemImages.value = Checked)
   If chkSubItemImages.value = Checked Then
      With lvwMain.Listitems
         For i = 1 To .count
            With .Item(i).SubItems(1)
               .IconIndex = Rnd * ilsIcons16.ImageCount
               Debug.Print .IconIndex
            End With
         Next i
      End With
   End If
End Sub



Private Sub CmdAdd_Click()
Dim sText As String
Dim sKey As String
On Error GoTo ErrorHandler
   sText = InputBox$("Please enter the caption of the item to add", , "Test Item " & lvwMain.Listitems.count + 1)
   If sText <> "" Then
      sKey = InputBox$("Please enter the key for the item:", , "C" & lvwMain.Listitems.count + 1)
      If sKey <> "" Then
         lvwMain.Listitems.Add , sKey, sText
      End If
   End If
   Exit Sub
ErrorHandler:
   MsgBox "Error: " & Err.description & " [" & Err.Number & "]", vbInformation
   Exit Sub
End Sub

Private Sub CmdInfo_Click()
On Error GoTo ErrorHandler
Dim sInfo As String
   If Not lvwMain.SelectedItem Is Nothing Then
      With lvwMain.SelectedItem
         sInfo = "Text = " & .text & vbCrLf
         sInfo = sInfo & "BackColor = " & .BackColor & vbCrLf
         sInfo = sInfo & "ForeColor = " & .ForeColor & vbCrLf
         sInfo = sInfo & "Tag = " & .Tag & vbCrLf
         sInfo = sInfo & "ToolTipText = " & .ToolTipText & vbCrLf
         sInfo = sInfo & "Checked = " & .Checked & vbCrLf
         sInfo = sInfo & "Cut = " & .Cut & vbCrLf
         sInfo = sInfo & "Selected = " & .Selected & vbCrLf
         sInfo = sInfo & "Hot = " & .Hot & vbCrLf
         sInfo = sInfo & "Indent = " & .indent & vbCrLf
         sInfo = sInfo & "ItemData = " & .ItemData & vbCrLf
         sInfo = sInfo & "Key = " & .key & vbCrLf
         sInfo = sInfo & "Left =" & .left & vbCrLf
         sInfo = sInfo & "Top = " & .top & vbCrLf
         
         MsgBox sInfo, vbInformation
      End With
   Else
      MsgBox "No item is selected.", vbInformation
   End If
   Exit Sub
ErrorHandler:
   MsgBox "Error: " & Err.description & " [" & Err.Number & "]", vbInformation
   Exit Sub
End Sub

Private Sub cmdNew_Click()
'   Dim f As New frmTestListView
'   f.Show
'   f.Move Me.left + 32 * Screen.TwipsPerPixelX, Me.top + 32 * Screen.TwipsPerPixelY
End Sub

Private Sub CmdRemove_Click()
On Error GoTo ErrorHandler
   If Not lvwMain.SelectedItem Is Nothing Then
      lvwMain.Listitems.Remove lvwMain.SelectedItem.key
   Else
      MsgBox "No item is selected.", vbInformation
   End If
   Exit Sub
ErrorHandler:
   MsgBox "Error: " & Err.description & " [" & Err.Number & "]", vbInformation
   Exit Sub
End Sub


Private Sub cmdWorkAreas_Click()
'   Dim fW As New frmTestWorkAreas
'   fW.Show
End Sub

Private Sub Command1_Click()
   With lvwItems
   lvwItems.Listitems.Clear
   End With
   
End Sub

Private Sub CmdDeleteRow_Click()
RemoveGridRow
End Sub

Private Sub Form_Resize()
   On Error Resume Next
 '  lvwMain.Move _
 '     lvwMain.Left, _
 '     lvwMain.Top, _
 '     Me.ScaleWidth - picTest.Width - Me.ScaleX(4, vbPixels, Me.ScaleMode), _
 '     Me.ScaleHeight - lvwMain.Top - picOptions.Height - picStatus.Height - Me.ScaleY(4, vbPixels, Me.ScaleMode)
End Sub

Private Sub ISButtonLW1_Click()

End Sub

Private Sub Grid_Click()
lblqty.Caption = ""
End Sub

Private Sub lBLclr_Click()
lblqty.Caption = ""
End Sub

Private Sub LBLdOT_Click()
lblqty.Caption = lblqty.Caption & "."
End Sub

Private Sub lBLnO_Click(Index As Integer)
 'With Me.Grid
 'If .Rows = 1 Then Exit Sub
 'End With
lblqty.Caption = lblqty.Caption & Index
End Sub

Function addrow(ITEMID As Integer, itemname As String, ITEMPRICE As Double, _
ItemType As Integer)
lblqty.Caption = ""
  Dim Msg As String
Dim LngRow As Long
Dim LngFindRow As Long
Dim des As String
On Error Resume Next
    Me.Grid.Rows = Me.Grid.Rows + 1
    LngRow = Me.Grid.Rows - 1
 With Me.Grid
     .TextMatrix(LngRow, .ColIndex("Code")) = ITEMID
     .TextMatrix(LngRow, .ColIndex("Name")) = itemname
      .TextMatrix(LngRow, .ColIndex("Count")) = 1
      .TextMatrix(LngRow, .ColIndex("Price")) = ITEMPRICE
       .TextMatrix(LngRow, .ColIndex("Totals")) = ITEMPRICE
      .TextMatrix(LngRow, .ColIndex("ItemType")) = ItemType
      .AutoSize 0, .Cols - 1, False
    
      .Row = .Rows - 1
End With

 
 ReLineGrid
 

End Function
Private Sub RemoveGridRow()
With Me.Grid
    If .Row <= 0 Then Exit Sub
      .RemoveItem .Row
End With
ReLineGrid
End Sub


Private Sub ReLineGrid()
On Error Resume Next
  Me.TxtTotalCash.Caption = 0
Dim IntCounter As Integer
 IntCounter = 0
 Dim i As Integer
With Me.Grid
    For i = .FixedRows To .Rows - 1
    
        If .TextMatrix(i, .ColIndex("Code")) <> "" Then
            IntCounter = IntCounter + 1
            .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
        End If
    Next i
    Me.TxtTotalCash.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Totals"), .Rows - 1, .ColIndex("Totals"))
 
      .AutoSize 0, .Cols - 1, False
End With
                
 
End Sub



Private Sub lblqty_Change()
If Val(lblqty.Caption) = 0 Then Exit Sub
 With Me.Grid
     .TextMatrix(.Row, .ColIndex("Count")) = Val(lblqty.Caption)
     .TextMatrix(.Row, .ColIndex("Totals")) = Val(lblqty.Caption) * _
     Val(.TextMatrix(.Row, .ColIndex("Price")))
ReLineGrid
   
End With

End Sub

Private Sub lvwItems_ItemClick(Item As vbalListViewLib6.cListItem)
addrow Val(Item.SubItems(2).Caption), Item.text, Val(Item.SubItems(1).Caption), Val(Item.SubItems(3).Caption)
 
End Sub

Private Sub lvwMain_AfterLabelEdit(Cancel As Boolean, NewString As String, Item As cListItem)
   Debug.Print "After Label Edit: ", NewString, Item.text
End Sub

Private Sub lvwMain_BeforeLabelEdit(Cancel As Boolean, Item As cListItem)
   Debug.Print "Before Label Edit: ", Item.text
End Sub

Private Sub lvwMain_Click()
   lblStatus.Caption = "Click"
End Sub

Private Sub lvwMain_ColumnClick(Column As cColumn)
   ' Sort according to the column type:
   Select Case Column.key
   Case "NAME"
      Column.SortType = eLVSortString
      Column.SortOrder = NewSortOrder(Column.SortOrder)
   Case "DATE"
      Column.SortType = eLVSortDate
      Column.SortOrder = NewSortOrder(Column.SortOrder)
   Case "SIZE"
      Column.SortType = eLVSortNumeric
      Column.SortOrder = NewSortOrder(Column.SortOrder)
   End Select
   lvwMain.Listitems.SortItems
End Sub

Private Function NewSortOrder(ByVal SortOrder As ESortOrderConstants) As ESortTypeConstants
   Select Case SortOrder
   Case eSortOrderNone, eSortOrderDescending
      NewSortOrder = eSortOrderAscending
   Case eSortOrderAscending
      NewSortOrder = eSortOrderDescending
   End Select
End Function

Private Sub lvwMain_DblClick()
   lblStatus.Caption = "Double Click"
End Sub

Private Sub lvwMain_ItemClick(Item As cListItem)
lblqty.Caption = ""
   lblStatus.Caption = "Clicked Item " & Item.text
On Error GoTo ErrorHandler
Dim sInfo As String
   If Not lvwMain.SelectedItem Is Nothing Then
      With lvwMain.SelectedItem
       
     '    sInfo = "Key = " & Item.key & Item.text
Label4.Caption = "«·«’‰«› «·Œ«’… » " & Item.text
FillItems (Item.key)
      End With
 
   End If
   Exit Sub
ErrorHandler:
   MsgBox "Error: " & Err.description & " [" & Err.Number & "]", vbInformation
   Exit Sub



End Sub

Private Sub lvwMain_ItemDblClick(Item As cListItem)
   lblStatus.Caption = "Double-Clicked Item " & Item.text
End Sub

Private Sub lvwMain_KeyDown(KeyCode As Integer, Shift As Integer)
   lblStatus.Caption = "KeyDown " & KeyCode & ",Shift"
End Sub

Private Sub lvwMain_KeyPress(KeyAscii As Integer)
   lblStatus.Caption = "KeyPress " & KeyAscii
End Sub

Private Sub lvwMain_KeyUp(KeyCode As Integer, Shift As Integer)
   lblStatus.Caption = "KeyUp " & KeyCode & ",Shift"
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   lblStatus.Caption = "MouseDown " & x & "," & Y
End Sub

Private Sub lvwMain_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   lblStatus.Caption = "MouseMove " & x & "," & Y
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   lblStatus.Caption = "MouseUp " & x & "," & Y
End Sub

Private Sub Form_Load()

 Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
Dim colX As cColumn
Dim itmX As cListItem
Dim i As Long
Dim J As Long
   
   Me.Show
   Me.Refresh
   
   
      With lvwItems
       lvwItems.Listitems.Clear
      .Visible = False
      .CustomDraw = True
            
      .AutoArrange = True
      
      ' Set up image lists:
      .ImageList(eLVLargeIcon) = GrouplImageList ' ilsIcons32
      .ImageList(eLVSmallIcon) = GrouplImageList ' ilsIcons16
      .ImageList(eLVTileImages) = GrouplImageList ' ilsIcons48
      .ImageList(eLVHeaderImages) = GrouplImageList ' ilsIcons16
      
      ' Add column headers
      Set colX = .Columns.Add(, "NAME", "Name")
      colX.Tag = "Stores the name of the item"
      colX.IconIndex = 0
      Set colX = .Columns.Add(, "Code", "Code")
      colX.Tag = "Stores the date of the item"
      colX.IconIndex = 1
      Set colX = .Columns.Add(, "id", "id")
      colX.Tag = "Stores the size of the item"
      colX.Alignment = eLVColumnAlignRight

 Set colX = .Columns.Add(, "ItemType", "ItemType")
      colX.Tag = "Stores the size of the item"
      colX.Alignment = eLVColumnAlignRight
      
End With
 
Image2.Picture = LoadPicture(App.path & "\Images\pos\gray.jpg")
Image3.Picture = LoadPicture(App.path & "\Images\pos\gray.jpg")
Image6.Picture = LoadPicture(App.path & "\Images\pos\gray.jpg")
Image7.Picture = LoadPicture(App.path & "\Images\pos\gray.jpg")
Image5.Picture = LoadPicture(App.path & "\Images\pos\blue.jpg")
 Image1.Picture = LoadPicture(App.path & "\Images\pos\DialPad.jpg")
 Image4.Picture = LoadPicture(App.path & "\Images\pos\takeaway.jpg")
  Image8.Picture = LoadPicture(App.path & "\Images\pos\phone.jpg")
  
   With cboBorder
      .AddItem "None"
      .ItemData(.NewIndex) = 0
      .AddItem "Fixed Single"
      .ItemData(.NewIndex) = 1
      .AddItem "Thin"
      .ItemData(.NewIndex) = 2
      .ListIndex = 1
   End With

   With cboAppearance
      .AddItem "Flat"
      .ItemData(.NewIndex) = 0
      .AddItem "3D"
      .ItemData(.NewIndex) = 1
      .ListIndex = 1
   End With
 

   
   FillGroups
FillTables
End Sub
Function FillGroups()
Dim colX As cColumn
Dim itmX As cListItem
Dim i As Long
Dim J As Long
 Dim sql As String
Dim rs As New ADODB.Recordset
Dim Balance As Double
 
 
sql = " SELECT * from  Groups where GroupID>1 "
 
rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

If rs.RecordCount = 0 Then
 GoTo XGroups
 End If
   
   
   With lvwMain
      .Visible = False
      .CustomDraw = True
            
      .AutoArrange = True
      
      ' Set up image lists:
      .ImageList(eLVLargeIcon) = GrouplImageList ' ilsIcons32
      .ImageList(eLVSmallIcon) = GrouplImageList ' ilsIcons16
      .ImageList(eLVTileImages) = GrouplImageList ' ilsIcons48
      .ImageList(eLVHeaderImages) = GrouplImageList ' ilsIcons16
      
      ' Add column headers
      Set colX = .Columns.Add(, "NAME", "Name")
      colX.Tag = "Stores the name of the item"
      colX.IconIndex = 0
      Set colX = .Columns.Add(, "DATE", "Date")
      colX.Tag = "Stores the date of the item"
      colX.IconIndex = 1
      Set colX = .Columns.Add(, "SIZE", "Size")
      colX.Tag = "Stores the size of the item"
      colX.Alignment = eLVColumnAlignRight
            
    '  For i = 1 To 3
    '     .Columns(i).ItemData = i * 100
    '  Next i
      
      With .Listitems
         For i = 0 To rs.RecordCount - 1
           Set itmX = .Add(, rs("GroupID").value, rs("GroupName").value, i, i)
      '      Set itmX = .Add(, "I" & i, "Test Item " & i, 0, 1)
            If (i Mod 2) = 0 Then
               itmX.ToolTipText = "This is a test tool tip for item " & i
            End If
            With itmX.SubItems(1)
               .Caption = DateSerial(year(Now), Rnd * Month(Now) + 1, Rnd * Day(Now) + 1)
               .ShowInTile = ((i Mod 2) = 0)
               '.IconIndex = itmX.IconIndex
            End With
            With itmX.SubItems(2)
               .Caption = CLng(Rnd * 1024 * 1024)
               .ShowInTile = True
            End With
            If (i = 1) Then
               ' test font/colours:
    '           itmX.BackColor = RGB(98, 176, 255)
    '           itmX.ForeColor = RGB(240, 248, 255)
    '           Dim sFnt As New StdFont
    '           sFnt.name = "Tahoma"
    '           sFnt.Size = 10
    '           sFnt.Bold = True
    ''           itmX.Font = sFnt
            End If
            rs.MoveNext
         Next i
      End With
      
      .TileViewItemLines = 3
               
      .Visible = True
   End With
XGroups:
   

End Function
Function FillItems(groupid As Integer)
Dim colX As cColumn
Dim itmX As cListItem
Dim i As Long
Dim J As Long
 Dim sql As String
Dim rs As New ADODB.Recordset
Dim Balance As Double
 
 
sql = " SELECT * from  TblItems where GroupID=" & groupid
 
rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

If rs.RecordCount = 0 Then
 GoTo XGroups
 End If
   
   
   With lvwItems
   lvwItems.Listitems.Clear
      .Visible = False
      .CustomDraw = True
            
      .AutoArrange = True
      
      ' Set up image lists:
      .ImageList(eLVLargeIcon) = GrouplImageList ' ilsIcons32
      .ImageList(eLVSmallIcon) = GrouplImageList ' ilsIcons16
      .ImageList(eLVTileImages) = GrouplImageList ' ilsIcons48
      .ImageList(eLVHeaderImages) = GrouplImageList ' ilsIcons16
      
 
            
    '  For i = 1 To 3
    '     .Columns(i).ItemData = i * 100
    '  Next i
      
      With .Listitems
         For i = 0 To rs.RecordCount - 1
           Set itmX = .Add(, rs("ItemID").value & i, rs("ItemName").value, i, i)
      '      Set itmX = .Add(, "I" & i, "Test Item " & i, 0, 1)
            If (i Mod 2) = 0 Then
               itmX.ToolTipText = "This is a test tool tip for item " & i
            End If
            With itmX.SubItems(1)
               .Caption = rs("SallingPrice").value    '  DateSerial(year(Now), Rnd * Month(Now) + 1, Rnd * Day(Now) + 1)
               .ShowInTile = ((i Mod 2) = 0)
               '.IconIndex = itmX.IconIndex
            End With
            With itmX.SubItems(2)
               .Caption = rs("ItemID").value  '  CLng(Rnd * 1024 * 1024)
               .ShowInTile = True
            End With
            
          With itmX.SubItems(3)
               .Caption = rs("ItemType").value  '  CLng(Rnd * 1024 * 1024)
               .ShowInTile = True
            End With
            
            If (i = 1) Then
            
               ' test font/colours:
    '           itmX.BackColor = RGB(98, 176, 255)
    '           itmX.ForeColor = RGB(240, 248, 255)
    '           Dim sFnt As New StdFont
    '           sFnt.name = "Tahoma"
    '           sFnt.Size = 10
    '           sFnt.Bold = True
    ''           itmX.Font = sFnt
            End If
            rs.MoveNext
         Next i
      End With
      
      .TileViewItemLines = 3
               
      .Visible = True
   End With
XGroups:
   
   
   
   
   
  End Function


Function FillTables()
 'fill tables
 Dim colX As cColumn
Dim itmX As cListItem
Dim i As Long
Dim J As Long

 Dim sql As String
Dim rs As New ADODB.Recordset
Dim Balance As Double
 
 
sql = " SELECT * from  Stables "
 
rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

If rs.RecordCount = 0 Then
 GoTo XTable
 End If
   



  With lvwTables
      .Visible = False
      .CustomDraw = True
            
      .AutoArrange = True
      
      ' Set up image lists:
    .ImageList(eLVLargeIcon) = ilsIcons32
      .ImageList(eLVSmallIcon) = ilsIcons16
      .ImageList(eLVTileImages) = ilsIcons48
      .ImageList(eLVHeaderImages) = ilsIcons16
      
      ' Add column headers
      Set colX = .Columns.Add(, "NAME", "Name")
      colX.Tag = "Stores the name of the item"
      colX.IconIndex = 0
      Set colX = .Columns.Add(, "DATE", "Date")
      colX.Tag = "Stores the date of the item"
      colX.IconIndex = 1
      Set colX = .Columns.Add(, "SIZE", "Size")
      colX.Tag = "Stores the size of the item"
      colX.Alignment = eLVColumnAlignRight
            
     'For i = 1 To 3
     '    .Columns(i).ItemData = i * 100
     ' Next i
  
      With .Listitems
         For i = 1 To rs.RecordCount
            Set itmX = .Add(, rs("id").value, "ÿ«Ê·… —ﬁ„ " & rs("name").value, 0, 0)
            If (i Mod 2) = 0 Then
               itmX.ToolTipText = "This is a test tool tip for item " & i
            End If
            With itmX.SubItems(1)
               .Caption = DateSerial(year(Now), Rnd * Month(Now) + 1, Rnd * Day(Now) + 1)
               .ShowInTile = ((i Mod 2) = 0)
               '.IconIndex = itmX.IconIndex
            End With
            With itmX.SubItems(2)
               .Caption = CLng(Rnd * 1024 * 1024)
               .ShowInTile = True
            End With
            If (Not IsNull(rs("Status").value)) Then
               ' test font/colours:
               itmX.BackColor = vbRed 'RGB(98, 176, 255)
               itmX.ForeColor = RGB(240, 248, 255)
            '   Dim sFnt As New StdFont
         '      sFnt.name = "Tahoma"
         '      sFnt.Size = 20
         '      sFnt.Bold = True
        
         '      itmX.Font = sFnt
         Else
           itmX.BackColor = vbGreen
            End If
            rs.MoveNext
         Next i
      End With
      
      .TileViewItemLines = 3
               
      .Visible = True
   End With
   rs.Close
XTable:
End Function
Private Sub lvwMain_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
   AllowedEffects = vbDropEffectMove
End Sub

Private Sub lvwMain_Resize()
   '
   'lvwMain.Arrange eLVAlignLeft
   '
End Sub

Private Sub lvwTables_ItemClick(Item As vbalListViewLib6.cListItem)
On Error GoTo ErrorHandler
Dim sInfo As String
   If Not lvwTables.SelectedItem Is Nothing Then
      With lvwTables.SelectedItem
       
     '    sInfo = "Key = " & Item.key & Item.text
Label8.Caption = Item.text

      End With
 
   End If
   Exit Sub
ErrorHandler:
   MsgBox "Error: " & Err.description & " [" & Err.Number & "]", vbInformation
   Exit Sub


End Sub
