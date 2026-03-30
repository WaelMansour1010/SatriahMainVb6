VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmBillCarMaintExtra 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17025
   Icon            =   "FrmCarMaintExtra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9345
   ScaleWidth      =   17025
   Begin VB.PictureBox Picture1 
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   255
      TabIndex        =   180
      Top             =   7320
      Width           =   255
   End
   Begin VB.TextBox txtAuthoOrder 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   177
      Top             =   600
      Width           =   1995
   End
   Begin VB.ComboBox DcbBasedOn 
      Height          =   315
      Left            =   4080
      TabIndex        =   152
      Top             =   720
      Width           =   1635
   End
   Begin VB.TextBox TxtPaymentValue 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   11550
      TabIndex        =   150
      Top             =   1080
      Width           =   1305
   End
   Begin VB.TextBox TxtNetDiscount 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8040
      TabIndex        =   149
      Top             =   1440
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox TxtNoteID2 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   3480
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   144
      Top             =   8880
      Width           =   1815
   End
   Begin VB.TextBox TxtSparePart 
      Alignment       =   1  'Right Justify
      Height          =   525
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   142
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox txtCusId1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   139
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox XPCboDiscountType 
      Height          =   315
      Left            =   13800
      TabIndex        =   132
      Text            =   "XPCboDiscountType"
      Top             =   1440
      Width           =   1635
   End
   Begin VB.TextBox XPTxtDiscountVal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   11550
      TabIndex        =   131
      Top             =   1440
      Width           =   1305
   End
   Begin VB.ComboBox CboPayMentType 
      Height          =   315
      Left            =   13800
      TabIndex        =   130
      Text            =   "CboPayMentType"
      Top             =   1080
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   3720
      TabIndex        =   98
      Top             =   8040
      Width           =   7935
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   5
         Left            =   1440
         Picture         =   "FrmCarMaintExtra.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   4
         Left            =   2760
         Picture         =   "FrmCarMaintExtra.frx":07D5
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   3
         Left            =   3480
         Picture         =   "FrmCarMaintExtra.frx":0D2D
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   2
         Left            =   4920
         Picture         =   "FrmCarMaintExtra.frx":11E6
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   1
         Left            =   3240
         Picture         =   "FrmCarMaintExtra.frx":16B6
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarMaintExtra.frx":1B57
         Height          =   555
         Index           =   0
         Left            =   7080
         Picture         =   "FrmCarMaintExtra.frx":8E89
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarMaintExtra.frx":9430
         Height          =   555
         Index           =   6
         Left            =   5640
         Picture         =   "FrmCarMaintExtra.frx":10762
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarMaintExtra.frx":10C03
         Height          =   555
         Index           =   7
         Left            =   4200
         Picture         =   "FrmCarMaintExtra.frx":17F35
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   8
         Left            =   2040
         Picture         =   "FrmCarMaintExtra.frx":187C5
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarMaintExtra.frx":18CAA
         Height          =   555
         Index           =   9
         Left            =   720
         Picture         =   "FrmCarMaintExtra.frx":1FFDC
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarMaintExtra.frx":204FC
         Height          =   555
         Index           =   10
         Left            =   6360
         Picture         =   "FrmCarMaintExtra.frx":20AE3
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarMaintExtra.frx":210CA
         Height          =   555
         Index           =   11
         Left            =   0
         Picture         =   "FrmCarMaintExtra.frx":283FC
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox TxtReq 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   87
      Top             =   12000
      Width           =   2055
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   21600
      TabIndex        =   49
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   21480
      TabIndex        =   48
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   13800
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   5400
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   8880
      Width           =   1815
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   21480
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   16965
      _cx             =   29924
      _cy             =   1032
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
      Caption         =   "›« Ê—… «’·«Õ "
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
      Begin VB.TextBox txtCusId 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1185
         TabIndex        =   17
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
         ButtonImage     =   "FrmCarMaintExtra.frx":28F90
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
         Left            =   120
         TabIndex        =   18
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
         ButtonImage     =   "FrmCarMaintExtra.frx":2932A
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
         Left            =   1710
         TabIndex        =   19
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
         ButtonImage     =   "FrmCarMaintExtra.frx":296C4
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
         Left            =   645
         TabIndex        =   20
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
         ButtonImage     =   "FrmCarMaintExtra.frx":29A5E
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   8280
         Picture         =   "FrmCarMaintExtra.frx":29DF8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2280
         TabIndex        =   47
         Top             =   120
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   11520
      TabIndex        =   21
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   703594497
      CurrentDate     =   38784
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   8280
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8760
      Width           =   8265
      _cx             =   14579
      _cy             =   953
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
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   -2147483630
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   23
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
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
         Left            =   6255
         TabIndex        =   24
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         Left            =   5415
         TabIndex        =   25
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         Left            =   4560
         TabIndex        =   26
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         Left            =   3705
         TabIndex        =   27
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         Left            =   120
         TabIndex        =   28
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         Height          =   375
         Left            =   975
         TabIndex        =   29
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   5
         Left            =   2880
         TabIndex        =   40
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   9
         Left            =   2040
         TabIndex        =   50
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   10
         Left            =   2880
         TabIndex        =   128
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄Â "
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
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   12540
      TabIndex        =   30
      Top             =   8400
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBoxx 
      Height          =   315
      Left            =   21360
      TabIndex        =   31
      Top             =   4080
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   21360
      TabIndex        =   42
      Top             =   3120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo Dcbranch 
      Bindings        =   "FrmCarMaintExtra.frx":2DA60
      Height          =   315
      Left            =   6780
      TabIndex        =   44
      Top             =   720
      Width           =   3795
      _ExtentX        =   6694
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
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   6255
      Left            =   0
      TabIndex        =   51
      Top             =   1800
      Width           =   16560
      _cx             =   29210
      _cy             =   11033
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
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   0
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   16711680
      Caption         =   "»Ì«‰«  ›« Ê—… «’·«Õ|Õ«·Â «·«⁄ „«œ|”‰œ«  «·’—›"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   1
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   0   'False
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   1
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Picture(0)      =   "FrmCarMaintExtra.frx":2DA75
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5790
         Left            =   17205
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   45
         Width           =   16470
         _cx             =   29051
         _cy             =   10213
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
         Align           =   0
         AutoSizeChildren=   0
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
         Begin VSFlex8UCtl.VSFlexGrid GRID2 
            Height          =   3630
            Left            =   120
            TabIndex        =   53
            Tag             =   "1"
            Top             =   240
            Width           =   13230
            _cx             =   23336
            _cy             =   6403
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
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmCarMaintExtra.frx":2DE0F
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
            Editable        =   0
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
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   4080
            Width           =   3375
         End
         Begin VB.Label Label1100 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9960
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   5790
         Index           =   15
         Left            =   45
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   45
         Width           =   16470
         _cx             =   29051
         _cy             =   10213
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   12
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
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   1
         ChildSpacing    =   1
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
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   1
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmCarMaintExtra.frx":2DF5B
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5760
            Index           =   16
            Left            =   15
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   15
            Width           =   16440
            _cx             =   28998
            _cy             =   10160
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
            Appearance      =   5
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
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
            Frame           =   0
            FrameStyle      =   3
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·œ›⁄«  «·„ﬁœ„…"
               Height          =   2970
               Left            =   -90
               RightToLeft     =   -1  'True
               TabIndex        =   171
               Top             =   2310
               Width           =   4755
               Begin VSFlex8Ctl.VSFlexGrid grdCash 
                  Height          =   2220
                  Left            =   150
                  TabIndex        =   172
                  Top             =   240
                  Width           =   4545
                  _cx             =   8017
                  _cy             =   3916
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmCarMaintExtra.frx":2DF91
                  ScrollTrack     =   0   'False
                  ScrollBars      =   2
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
               Begin ImpulseButton.ISButton ISButton1 
                  Height          =   270
                  Left            =   7920
                  TabIndex        =   173
                  Top             =   2640
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   476
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–›"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmCarMaintExtra.frx":2E06A
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbllblTotalVat 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   300
                  Left            =   270
                  RightToLeft     =   -1  'True
                  TabIndex        =   178
                  Top             =   2610
                  Width           =   1095
               End
               Begin VB.Label Label9 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì„ﬂ‰ﬂ «· ⁄œÌ· ›Ï ﬁÌ„… «·œ›⁄«  ÌœÊÌ«ı"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   255
                  Left            =   60
                  TabIndex        =   176
                  Top             =   2280
                  Visible         =   0   'False
                  Width           =   2595
               End
               Begin VB.Label Label8 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·œ›⁄«  «·„ﬁœ„…"
                  Height          =   285
                  Left            =   2940
                  RightToLeft     =   -1  'True
                  TabIndex        =   175
                  Top             =   2655
                  Width           =   1665
               End
               Begin VB.Label lblTotalPay 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   300
                  Left            =   1830
                  RightToLeft     =   -1  'True
                  TabIndex        =   174
                  Top             =   2640
                  Width           =   1095
               End
            End
            Begin VB.TextBox TxtTotalValue 
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
               Height          =   330
               Left            =   135
               Locked          =   -1  'True
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Text            =   "0"
               Top             =   5280
               Width           =   1500
            End
            Begin VB.TextBox TxtFATValue 
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
               Height          =   330
               Left            =   5520
               Locked          =   -1  'True
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   157
               Text            =   "0"
               Top             =   5280
               Width           =   1275
            End
            Begin VB.TextBox TxtFATYou 
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
               Height          =   330
               Left            =   8010
               Locked          =   -1  'True
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   156
               Text            =   "0"
               Top             =   5280
               Width           =   1275
            End
            Begin VB.CommandButton Accredit 
               Caption         =   "Command1"
               Height          =   255
               Left            =   5940
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   5640
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Frame lblExt 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„’«—Ì› «·›—⁄Ì…"
               Height          =   2970
               Left            =   4710
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   2280
               Width           =   5655
               Begin VSFlex8Ctl.VSFlexGrid fg2 
                  Height          =   2220
                  Left            =   60
                  TabIndex        =   14
                  Top             =   360
                  Width           =   5505
                  _cx             =   9710
                  _cy             =   3916
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
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmCarMaintExtra.frx":2E604
                  ScrollTrack     =   0   'False
                  ScrollBars      =   2
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
               Begin ImpulseButton.ISButton CdmDelet2 
                  Height          =   270
                  Left            =   7920
                  TabIndex        =   155
                  Top             =   2640
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   476
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–›"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmCarMaintExtra.frx":2E732
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label LbToTalExtra 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   300
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   2640
                  Width           =   1575
               End
               Begin VB.Label lblEx 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·„’«—Ì› «·›—⁄Ì…"
                  Height          =   285
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   2655
                  Width           =   1935
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì„ﬂ‰ﬂ «· ⁄œÌ· ›Ï ﬁÌ„… «·œ›⁄«  ÌœÊÌ«ı"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   255
                  Left            =   60
                  TabIndex        =   71
                  Top             =   2280
                  Visible         =   0   'False
                  Width           =   2595
               End
            End
            Begin VB.Frame LblWork 
               BackColor       =   &H00E2E9E9&
               Caption         =   "√⁄„«· «· ’·ÌÕ"
               Height          =   2970
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   2280
               Width           =   6030
               Begin VSFlex8Ctl.VSFlexGrid fg 
                  Height          =   2340
                  Left            =   270
                  TabIndex        =   13
                  Top             =   240
                  Width           =   5685
                  _cx             =   10028
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
                  Cols            =   23
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmCarMaintExtra.frx":2ECCC
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
               Begin ImpulseButton.ISButton CmdDelete 
                  Height          =   270
                  Left            =   6720
                  TabIndex        =   154
                  Top             =   2640
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   476
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–›"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmCarMaintExtra.frx":2F00E
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbTotalMenteDis 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   300
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   2640
                  Width           =   1695
               End
               Begin VB.Label lbTotalMente 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   300
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   2640
                  Width           =   1455
               End
               Begin VB.Label LblM 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «⁄„«· «· ’·ÌÕ"
                  Height          =   285
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   2640
                  Width           =   1935
               End
            End
            Begin VB.Frame lblDataCli 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·⁄„Ì·"
               Height          =   2415
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   0
               Width           =   16395
               Begin VB.TextBox TxtCarMetarOut 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   140
                  Top             =   720
                  Width           =   2775
               End
               Begin VB.ComboBox ComGranty 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   124
                  Top             =   600
                  Width           =   3375
               End
               Begin VB.Frame frmgranty 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  «·÷„«‰"
                  Height          =   1455
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   960
                  Width           =   4575
                  Begin VB.ComboBox ComMD 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   116
                     Top             =   240
                     Width           =   1815
                  End
                  Begin VB.TextBox TxtLongGranty 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   1920
                     TabIndex        =   115
                     Top             =   240
                     Width           =   1575
                  End
                  Begin VB.TextBox txtKM 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   600
                     TabIndex        =   114
                     Top             =   1080
                     Width           =   2055
                  End
                  Begin MSComCtl2.DTPicker DateStartG 
                     Height          =   315
                     Left            =   2160
                     TabIndex        =   117
                     Top             =   600
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   192217089
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker DateEndg 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   118
                     Top             =   600
                     Width           =   1455
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   192086017
                     CurrentDate     =   38784
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Ì‰ ÂÌ"
                     Height          =   285
                     Index           =   9
                     Left            =   1410
                     TabIndex        =   123
                     Top             =   615
                     Width           =   645
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Ì»œ√ „‰"
                     Height          =   285
                     Index           =   5
                     Left            =   3690
                     TabIndex        =   122
                     Top             =   615
                     Width           =   765
                  End
                  Begin VB.Label lbllong 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "„œ… «·÷„«‰"
                     Height          =   255
                     Left            =   3360
                     TabIndex        =   121
                     Top             =   240
                     Width           =   1185
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " €ÌÌ— «·“Ì  »⁄œ „—Ê—"
                     Height          =   285
                     Index           =   3
                     Left            =   2880
                     TabIndex        =   120
                     Top             =   1080
                     Width           =   1605
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ„"
                     Height          =   285
                     Index           =   16
                     Left            =   -360
                     TabIndex        =   119
                     Top             =   1080
                     Width           =   765
                  End
               End
               Begin VB.TextBox TxtCliientName 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   12600
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   960
                  Width           =   2775
               End
               Begin VB.TextBox TxtClientPhone 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   2
                  Top             =   330
                  Width           =   2775
               End
               Begin VB.ComboBox DcbyearFactor 
                  Height          =   315
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   4
                  Top             =   1170
                  Width           =   2775
               End
               Begin VB.TextBox TxtAmoutAccept 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   1530
                  Width           =   855
               End
               Begin VB.TextBox TXtCarMeter 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   1920
                  Width           =   2775
               End
               Begin VB.TextBox TxtFirstPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6720
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   1530
                  Width           =   855
               End
               Begin VB.ComboBox DcbOrderStatus 
                  Height          =   315
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   1080
                  Width           =   2775
               End
               Begin XtremeSuiteControls.CheckBox ChAccept 
                  Height          =   495
                  Left            =   120
                  TabIndex        =   12
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   1455
                  _Version        =   786432
                  _ExtentX        =   2566
                  _ExtentY        =   873
                  _StockProps     =   79
                  Caption         =   " „  „Ê«›ﬁ… «·⁄„Ì·"
                  UseVisualStyle  =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbCarModel 
                  Bindings        =   "FrmCarMaintExtra.frx":2F5A8
                  Height          =   315
                  Left            =   12600
                  TabIndex        =   3
                  Top             =   1680
                  Width           =   2775
                  _ExtentX        =   4895
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
               Begin MSDataListLib.DataCombo DcbColor 
                  Bindings        =   "FrmCarMaintExtra.frx":2F5BD
                  Height          =   315
                  Left            =   8760
                  TabIndex        =   5
                  Top             =   840
                  Width           =   2775
                  _ExtentX        =   4895
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
               Begin MSComCtl2.DTPicker TxtEndDate 
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   11
                  Top             =   1920
                  Width           =   2775
                  _ExtentX        =   4895
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   192086017
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcbCarType 
                  Bindings        =   "FrmCarMaintExtra.frx":2F5D2
                  Height          =   315
                  Left            =   12600
                  TabIndex        =   112
                  Top             =   1320
                  Width           =   2775
                  _ExtentX        =   4895
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
               Begin MSDataListLib.DataCombo DcbCustmer 
                  Bindings        =   "FrmCarMaintExtra.frx":2F5E7
                  Height          =   315
                  Left            =   10080
                  TabIndex        =   145
                  Top             =   360
                  Width           =   5295
                  _ExtentX        =   9340
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
               Begin ImpulseButton.ISButton BtnShow 
                  Height          =   315
                  Left            =   8760
                  TabIndex        =   148
                  Top             =   360
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  ButtonPositionImage=   1
                  Caption         =   "ﬂ‘› Õ”«»"
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
               Begin MSDataListLib.DataCombo DcbCar 
                  Bindings        =   "FrmCarMaintExtra.frx":2F5FC
                  Height          =   315
                  Left            =   8760
                  TabIndex        =   163
                  Top             =   1560
                  Width           =   2775
                  _ExtentX        =   4895
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
               Begin VB.TextBox TxtPlatNo 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   1560
                  Width           =   2775
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«”„ «·⁄„Ì·"
                  Height          =   255
                  Left            =   15480
                  RightToLeft     =   -1  'True
                  TabIndex        =   147
                  Top             =   360
                  Width           =   855
               End
               Begin VB.Label LblCli 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄„Ì· ‰ﬁœÌ"
                  Height          =   255
                  Left            =   15360
                  RightToLeft     =   -1  'True
                  TabIndex        =   146
                  Top             =   960
                  Width           =   855
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œ«œ «·Œ—ÊÃ"
                  Height          =   270
                  Index           =   0
                  Left            =   7830
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   720
                  Width           =   825
               End
               Begin VB.Label lblty 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "›∆… «·ÿ·»"
                  Height          =   255
                  Left            =   3240
                  TabIndex        =   125
                  Top             =   600
                  Width           =   1335
               End
               Begin VB.Label lblColor 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«··Ê‰"
                  Height          =   255
                  Left            =   11520
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   840
                  Width           =   855
               End
               Begin VB.Label LblOrderSt 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Õ«·… «·ÿ·»"
                  Height          =   255
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   1080
                  Width           =   1095
               End
               Begin VB.Label LblAmountAcc 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„»·€ «·„ ›ﬁ ⁄·ÌÂ"
                  Height          =   255
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   1560
                  Width           =   975
               End
               Begin VB.Label LblPhone 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Â« › «·⁄„Ì·"
                  Height          =   255
                  Left            =   7440
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.Label LblPla 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «··ÊÕ…"
                  Height          =   255
                  Left            =   11400
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   1560
                  Width           =   1095
               End
               Begin VB.Label LblCarMeter 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œ«œ «·œŒÊ·"
                  Height          =   270
                  Left            =   11670
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   1920
                  Width           =   825
               End
               Begin VB.Label LblPayF 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·œ›⁄… «·„ﬁœ„…"
                  Height          =   270
                  Left            =   7470
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   1560
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " √—ÌŒ «·Œ—ÊÃ"
                  Height          =   285
                  Index           =   2
                  Left            =   7650
                  TabIndex        =   72
                  Top             =   1920
                  Width           =   1005
               End
               Begin VB.Label LblYear 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "”‰… «·’‰⁄"
                  Height          =   255
                  Left            =   11400
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   1200
                  Width           =   1095
               End
               Begin VB.Label lblModel 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ÿ—«“ «·„⁄œÂ/«·”Ì«—…"
                  Height          =   255
                  Left            =   15480
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   1680
                  Width           =   855
               End
               Begin VB.Label LblCar 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰Ê⁄ «·„⁄œÂ/«·”Ì«—…"
                  Height          =   255
                  Left            =   15480
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   1320
                  Width           =   855
               End
            End
            Begin MSDataListLib.DataCombo AccountVat 
               Bindings        =   "FrmCarMaintExtra.frx":2F611
               Height          =   315
               Left            =   225
               TabIndex        =   162
               Top             =   5520
               Visible         =   0   'False
               Width           =   3450
               _ExtentX        =   6085
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
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ã„«·Ì"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   68
               Left            =   2355
               RightToLeft     =   -1  'True
               TabIndex        =   161
               Top             =   5280
               Width           =   540
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «·›« "
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   67
               Left            =   7020
               RightToLeft     =   -1  'True
               TabIndex        =   160
               Top             =   5280
               Width           =   810
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "‰”»…«·›« "
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   66
               Left            =   9510
               RightToLeft     =   -1  'True
               TabIndex        =   159
               Top             =   5280
               Width           =   675
            End
            Begin VB.Label lbldifdis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               Height          =   390
               Left            =   3300
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   5280
               Visible         =   0   'False
               Width           =   1365
            End
            Begin VB.Label LbtotalDis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               Height          =   330
               Left            =   12960
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   5280
               Width           =   1845
            End
            Begin VB.Label lblchange 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„ »ﬁÌ"
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   4260
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   5280
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.Label lbpricefirst 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·œ›⁄… «·„ﬁœ„…"
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   11730
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   5280
               Width           =   1035
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   330
               Left            =   -2175
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   -4440
               Width           =   1320
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ã„«·Ì «·⁄«„"
               Height          =   300
               Left            =   8610
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   1920
               Width           =   2535
            End
            Begin VB.Label firstprice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               Height          =   330
               Left            =   10455
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   5280
               Width           =   1275
            End
            Begin VB.Label Lbtotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   330
               Left            =   13095
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   5280
               Width           =   1350
            End
            Begin VB.Label Lbtota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ã„«·Ì «·⁄«„"
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   14805
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   5280
               Width           =   1185
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   3315
               Index           =   62
               Left            =   3165
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   1575
               Width           =   780
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5760
            Index           =   9
            Left            =   15
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   15
            Width           =   16440
            _cx             =   28998
            _cy             =   10160
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
            Appearance      =   5
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   0
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
            Begin VB.TextBox Text8 
               Alignment       =   1  'Right Justify
               Height          =   4320
               Left            =   4260
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   1245
               Width           =   945
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   2940
               Left            =   5475
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   1575
               Width           =   1320
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2940
               Index           =   67
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   1575
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   2880
               Index           =   68
               Left            =   5205
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   1965
               Width           =   45
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3420
               Index           =   69
               Left            =   3945
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   1575
               Width           =   315
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   5790
         Left            =   17505
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   45
         Width           =   16470
         _cx             =   29051
         _cy             =   10213
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
         Align           =   0
         AutoSizeChildren=   0
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
         Begin VSFlex8UCtl.VSFlexGrid vchrgrid 
            Height          =   5205
            Left            =   0
            TabIndex        =   165
            Top             =   120
            Width           =   16485
            _cx             =   29078
            _cy             =   9181
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
            Rows            =   1
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmCarMaintExtra.frx":2F626
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰œ «·’—›"
               Height          =   1050
               Index           =   51
               Left            =   0
               TabIndex        =   166
               Top             =   5880
               Width           =   1440
            End
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰œ«  «·„‰’—›… ··«„—"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   35
            Left            =   7680
            TabIndex        =   170
            Top             =   120
            Width           =   3120
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " ÕœÌÀ"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   11280
            RightToLeft     =   -1  'True
            TabIndex        =   169
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  «·”‰œ« "
            Height          =   285
            Index           =   57
            Left            =   4440
            TabIndex        =   168
            Top             =   5400
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Index           =   58
            Left            =   240
            TabIndex        =   167
            Top             =   5400
            Width           =   3765
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   315
      Index           =   8
      Left            =   1920
      TabIndex        =   97
      Top             =   8880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… «·ﬁÌœ"
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
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   6780
      TabIndex        =   133
      Top             =   1095
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbldif 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   330
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   179
      Top             =   8820
      Width           =   1260
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "»‰«¡ ⁄·Ï"
      Height          =   240
      Index           =   18
      Left            =   5760
      TabIndex        =   153
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„œ›Ê⁄"
      Height          =   330
      Index           =   17
      Left            =   12870
      TabIndex        =   151
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÿ⁄ €Ì«— „‰ ﬁ»· »«·⁄„Ì·"
      Height          =   240
      Index           =   15
      Left            =   4560
      TabIndex        =   143
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   240
      Index           =   11
      Left            =   10560
      TabIndex        =   138
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "%"
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
      Index           =   55
      Left            =   11070
      TabIndex        =   137
      Top             =   1560
      Width           =   345
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·Œ’„"
      Height          =   240
      Index           =   12
      Left            =   15480
      TabIndex        =   136
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·”œ«œ"
      Height          =   240
      Index           =   13
      Left            =   15480
      TabIndex        =   135
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„…"
      Height          =   330
      Index           =   14
      Left            =   12840
      TabIndex        =   134
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      Height          =   255
      Index           =   10
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   95
      Top             =   8880
      Width           =   855
   End
   Begin VB.Label lbreq 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "»‰«¡⁄·Ï «„— ‘€· —ﬁ„"
      Height          =   255
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   91
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·⁄„Ì·"
      Height          =   255
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   88
      Top             =   10320
      Width           =   855
   End
   Begin VB.Image img 
      Height          =   855
      Left            =   22680
      Picture         =   "FrmCarMaintExtra.frx":2F7F2
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   720
   End
   Begin VB.Image imgnul 
      Height          =   1095
      Left            =   22680
      Top             =   4800
      Width           =   735
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   15240
      Picture         =   "FrmCarMaintExtra.frx":30816
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   720
   End
   Begin VB.Label lblBr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      Height          =   255
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   780
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   20760
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·›« Ê—…"
      Height          =   285
      Index           =   4
      Left            =   15360
      TabIndex        =   39
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   12600
      TabIndex        =   38
      Top             =   735
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   15285
      TabIndex        =   37
      Top             =   8355
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2310
      TabIndex        =   36
      Top             =   8310
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   570
      TabIndex        =   35
      Top             =   8310
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   -30
      TabIndex        =   34
      Top             =   8340
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1620
      TabIndex        =   33
      Top             =   8340
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   21240
      TabIndex        =   32
      Top             =   2640
      Width           =   1005
   End
End
Attribute VB_Name = "FrmBillCarMaintExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim Employee_account As String
Public Ch As Boolean
Dim general_noteid As Long
Dim RsNotesGeneral As ADODB.Recordset
Dim mIdTrans As Integer
Sub ClculteVAT()
If Me.TxtModFlg.text <> "R" Then
Dim Percetage As Double
Dim account As String
PercentgValueAddedAccount_Transec XPDtbTrans.value, 21, 1, account, Percetage
TxtFATYou.text = Percetage
AccountVat.BoundText = account
Calculte
End If
End Sub
Sub Calculte()
Dim Valu As Double

Valu = val(LbtotalDis.Caption) - val(TxtNetDiscount)
If Me.TxtModFlg.text <> "R" Then
If val(TxtFATYou.text) > 0 Then
TxtFATValue.text = (Valu * val(TxtFATYou.text)) / 100
Else
TxtFATValue.text = 0
End If
TxtTotalValue.text = Valu + val(TxtFATValue.text)
End If
lbldifdis.Caption = val(Me.LbtotalDis.Caption) - val(firstprice.Caption) - val(TxtNetDiscount.text) - val(TxtPaymentValue.text) + val(TxtFATValue.text)
End Sub
Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter  As Integer
        Me.lbTotalMente.Caption = 0
        Me.Lbtotal.Caption = 0
        Me.LbToTalExtra.Caption = 0
        lbTotalMenteDis.Caption = 0
      '    lbTotalMenteDis.Caption = 0
            LbtotalDis.Caption = 0
    IntCounter = 0

    With Fg

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("serial")) = IntCounter
                If val(.TextMatrix(i, .ColIndex("count"))) <> 0 Then
                .TextMatrix(i, .ColIndex("nett")) = val(.TextMatrix(i, .ColIndex("valuedis"))) * val(.TextMatrix(i, .ColIndex("count")))
        .TextMatrix(i, .ColIndex("totalm")) = val(.TextMatrix(i, .ColIndex("valuedis"))) * val(.TextMatrix(i, .ColIndex("count")))
        Else
         .TextMatrix(i, .ColIndex("nett")) = .TextMatrix(i, .ColIndex("valuedis"))
        .TextMatrix(i, .ColIndex("totalm")) = .TextMatrix(i, .ColIndex("valuedis"))
        .TextMatrix(i, .ColIndex("count")) = 1
       End If
        .TextMatrix(i, .ColIndex("serial")) = IntCounter
            End If
 If .TextMatrix(i, .ColIndex("valuedis")) <> "" Then
                
                Me.lbTotalMente.Caption = val(Me.lbTotalMente.Caption) + val(Fg.TextMatrix(i, Fg.ColIndex("valuedis")))
        lbTotalMenteDis.Caption = val(lbTotalMenteDis.Caption) + val(Fg.TextMatrix(i, Fg.ColIndex("nett")))
            End If
        Next i
 
    End With
    
IntCounter = 0
    With fg2

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("serial")) = IntCounter
                 If val(.TextMatrix(i, .ColIndex("count"))) <> 0 Then
        .TextMatrix(i, .ColIndex("totalex")) = .TextMatrix(i, .ColIndex("value")) * .TextMatrix(i, .ColIndex("count"))
        Else
        .TextMatrix(i, .ColIndex("totalex")) = .TextMatrix(i, .ColIndex("value"))
        .TextMatrix(i, .ColIndex("count")) = 1
       End If
            End If

      
 If .TextMatrix(i, .ColIndex("value")) <> "" Then
                
                Me.LbToTalExtra.Caption = val(Me.LbToTalExtra.Caption) + val(fg2.TextMatrix(i, fg2.ColIndex("totalex")))
        
            End If
        Next i
    End With
Me.Lbtotal.Caption = val(Me.LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption) + val(lbl(58).Caption)
LbtotalDis.Caption = val(LbToTalExtra.Caption) + val(Me.lbTotalMenteDis.Caption) + val(lbl(58).Caption)
Calculte
FillCal
End Sub
Private Sub Accredit_Click()
    Dim BeginTrans As Boolean
    
    Cn.BeginTrans
    BeginTrans = True

    If IsNull(rs("Posted")) Then
        rs("Posted") = user_id
        rs("PostedDate") = Time
    Else
        rs("Posted") = Null
       rs("PostedDate") = Time
    End If
   
    rs.update
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If

    Cn.CommitTrans
    BeginTrans = False
FillApprovedTable
    Retrive (val(Me.XPTxtID.text))
End Sub



Private Sub BtnShow_Click()
If val(Me.DcbCustmer.BoundText) = 0 Then
Dim StrTempAccountCode As String
            Dim FirstPeriod As Date
            getFirstPeriodDateInthisYear FirstPeriod
                   StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbCustmer.BoundText))
            ShowReport StrTempAccountCode, DcbCustmer.text, FirstPeriod, Date
      End If
End Sub

Private Sub CboPayMentType_Change()
    On Error GoTo ErrTrap

    'If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    If CboPayMentType.ListIndex = 0 Then '‰ﬁœÌ
       TxtPaymentValue.Enabled = True
        DcboBox.Enabled = True
  
    Else
 TxtPaymentValue.Enabled = False
 TxtPaymentValue.text = 0
        DcboBox.BoundText = ""
        DcboBox.Enabled = False
     
  
    End If

    'End If
    Exit Sub
ErrTrap:
End Sub

Private Sub CboPayMentType_Click()

    CboPayMentType_Change
 
End Sub


Private Sub CdmDelet2_Click()
RemoveGridRow2
End Sub

Private Sub CmdDelete_Click()
RemoveGridRow
End Sub
Private Sub RemoveGridRow()
    With Me.Fg
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Private Sub RemoveGridRow2()
    With Me.fg2
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Private Sub DcbBasedOn_Change()
TxtCliientName.Enabled = False
DcbColor.Enabled = False
TxtClientPhone.Enabled = False
ComGranty.Enabled = False
frmgranty.Enabled = False
TxtEndDate.Enabled = False
DcbOrderStatus.Enabled = False
TxtCarMetarOut.Enabled = False
TxtFirstPrice.Enabled = False
TXtCarMeter.Enabled = False
TxtPlatNo.Enabled = False
DcbyearFactor.Enabled = False
DcbCarType.Enabled = False
TxtReq.Visible = False
lbreq.Visible = False
DcbCarModel.Enabled = False
TxtAmoutAccept.Enabled = False
TxtAuthoOrder.Visible = False
If val(DcbBasedOn.ListIndex) = 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                lbreq = "»‰«¡⁄·Ï «„— ‘€· —ﬁ„"
            Else
            lbreq = " Job  No"
            End If
Else
            If SystemOptions.UserInterface = ArabicInterface Then
               lbreq = "»‰«¡⁄·Ï «–‰ «’·«Õ —ﬁ„"
            Else
            lbreq = " Repair #"
            
            End If
    
End If



If val(DcbBasedOn.ListIndex) = 1 Then
    TxtReq.Visible = True
    lbreq.Visible = True
ElseIf val(DcbBasedOn.ListIndex) = 2 Then

    TxtAuthoOrder.Visible = True
Else
    TxtFirstPrice.Enabled = False
    TxtCliientName.Enabled = True
    DcbColor.Enabled = True
    TxtClientPhone.Enabled = True
    ComGranty.Enabled = True
    frmgranty.Enabled = True
    TxtEndDate.Enabled = True
    TxtAmoutAccept.Enabled = True
    DcbOrderStatus.Enabled = True
    DcbCarType.Enabled = True
    DcbCarModel.Enabled = True
    DcbyearFactor.Enabled = True
    TxtPlatNo.Enabled = True
    Me.DcbCar.Enabled = True
    TXtCarMeter.Enabled = True
    TxtCarMetarOut.Enabled = True
End If
If Me.TxtModFlg.text <> "R" Then
Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 2
            fg2.Enabled = True
End If
End Sub

Private Sub DcbBasedOn_Click()
DcbBasedOn_Change
End Sub



Private Sub DcbCustmer_Change()
DcbCustmer_Click (0)
End Sub

Private Sub DcbCustmer_Click(Area As Integer)
If Me.TxtModFlg.text <> "R" Then
     If SystemOptions.LinkCustomerWithCars = True Then
       Dim Dcombos As ClsDataCombos
       Set Dcombos = New ClsDataCombos
       Dcombos.GetCarsOfCustomer DcbCar, val(DcbCustmer.BoundText)
       DcbCar.BoundText = GetFirstCarOfCustomer(val(DcbCustmer.BoundText))
       End If
       
End If
End Sub
Function GetFirstCarOfCustomer(Optional CusID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT MIN(id) as MinID From TblCusCar where CustomerID =" & CusID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetFirstCarOfCustomer = IIf(IsNull(rs2("MinID").value), 0, rs2("MinID").value)
Else
GetFirstCarOfCustomer = 0
End If
End Function
Sub GetInformationOfCustomerCar(Optional CarID As Double)
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = "Select * from TblCusCar where ID=" & CarID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
DcbCarType.BoundText = IIf(IsNull(rs2("BrandID").value), 0, rs2("BrandID").value)
'DcbyearFactor.ListIndex = IIf(IsNull(rs2("ModelID").value), -1, rs2("ModelID").value)
TxtPlatNo.text = DcbCar.text
DcbCarModel.BoundText = IIf(IsNull(rs2("CarModelID").value), 0, rs2("CarModelID").value)
DcbColor.BoundText = IIf(IsNull(rs2("ColorID").value), 0, rs2("ColorID").value)
Else
DcbCarModel.BoundText = 0
DcbColor.BoundText = 0
DcbCarType.BoundText = 0
DcbyearFactor.ListIndex = -1
End If
End Sub
Private Sub DcbCar_Change()
DcbCar_Click (0)
End Sub

Private Sub DcbCar_Click(Area As Integer)
If Me.TxtModFlg.text <> "R" And Me.TxtModFlg.text <> "" Then
If SystemOptions.LinkCustomerWithCars = True Then
GetInformationOfCustomerCar val(DcbCar.BoundText)
End If
End If
End Sub

Private Sub Fg_KeyDown(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.text <> "R" Then
If val(DcbBasedOn.ListIndex) = 0 Then
With Fg
If KeyCode = vbKeyF3 Then
Select Case .ColKey(.Col)
Case "name"

  Unload FrmBillCarMaintExtrSearch
  FrmBillCarMaintExtrSearch.IndTyp = 1
           FrmBillCarMaintExtrSearch.Row = .Row
  Load FrmBillCarMaintExtrSearch
           FrmBillCarMaintExtrSearch.IndTyp = 1
           FrmBillCarMaintExtrSearch.Row = .Row
            FrmBillCarMaintExtrSearch.show vbModal
End Select
End If
End With
End If
End If
End Sub

Private Sub FG_KeyPress(KeyAscii As Integer)
If Me.TxtModFlg.text <> "R" Then
If val(DcbBasedOn.ListIndex) = 0 Then
With Fg
If KeyAscii = vbKeyF3 Then
Select Case .ColKey(.Col)
Case "name"

 Unload FrmBillCarMaintExtrSearch
 FrmBillCarMaintExtrSearch.IndTyp = 1
           FrmBillCarMaintExtrSearch.Row = .Row
 Load FrmBillCarMaintExtrSearch
           FrmBillCarMaintExtrSearch.IndTyp = 1
           FrmBillCarMaintExtrSearch.Row = .Row
            FrmBillCarMaintExtrSearch.show vbModal
End Select
End If
End With
End If
End If
End Sub

Private Sub fg_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.text <> "R" Then
If val(DcbBasedOn.ListIndex) = 0 Then
With Fg
If KeyCode = vbKeyF3 Then
Select Case .ColKey(.Col)
Case "name"

  Unload FrmBillCarMaintExtrSearch
  FrmBillCarMaintExtrSearch.IndTyp = 1
            FrmBillCarMaintExtrSearch.Row = .Row
   Load FrmBillCarMaintExtrSearch
            FrmBillCarMaintExtrSearch.IndTyp = 1
            FrmBillCarMaintExtrSearch.Row = .Row
             FrmBillCarMaintExtrSearch.show vbModal
End Select
End If
End With
End If
End If
End Sub

Private Sub fg_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If Me.TxtModFlg.text <> "R" Then
If val(DcbBasedOn.ListIndex) = 0 Then
With Fg
If KeyCode = vbKeyF3 Then
Select Case .ColKey(.Col)
Case "name"

  Unload FrmBillCarMaintExtrSearch
  FrmBillCarMaintExtrSearch.IndTyp = 1
           FrmBillCarMaintExtrSearch.Row = .Row
  Load FrmBillCarMaintExtrSearch
           FrmBillCarMaintExtrSearch.IndTyp = 1
           FrmBillCarMaintExtrSearch.Row = .Row
            FrmBillCarMaintExtrSearch.show vbModal
End Select
End If
End With
End If
End If
End Sub

Private Sub firstprice_Change()
If Me.TxtModFlg.text <> "R" Then
lbldifdis.Caption = val(Me.LbtotalDis.Caption) - val(firstprice.Caption) - val(TxtNetDiscount.text) - val(TxtPaymentValue.text) + val(TxtFATValue.text)
Calculte
End If
End Sub

Private Sub lbldifdis_Change()
Calculte
End Sub

Private Sub TxtNetDiscount_Change()
If Me.TxtModFlg.text <> "R" Then
lbldifdis.Caption = val(Me.LbtotalDis.Caption) - val(firstprice.Caption) - val(TxtNetDiscount.text) - val(TxtPaymentValue.text) + val(TxtFATValue.text)
End If
End Sub

Private Sub TxtPaymentValue_Change()
If Me.TxtModFlg.text <> "R" Then
lbldifdis.Caption = val(Me.LbtotalDis.Caption) - val(firstprice.Caption) - val(TxtNetDiscount.text) - val(TxtPaymentValue.text)
End If
End Sub

Private Sub TxtPaymentValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPaymentValue.text, 0)
End Sub

Private Sub XPCboDiscountType_Change()
    XPCboDiscountType_Click
    
End Sub

Private Sub XPCboDiscountType_Click()
If Me.TxtModFlg.text <> "R" Then
    On Error GoTo ErrTrap
    If val(XPCboDiscountType.ListIndex) = 2 Then
    TxtNetDiscount.text = (val(XPTxtDiscountVal.text) * val(LbtotalDis.Caption)) / 100
    If SystemOptions.UserInterface = ArabicInterface Then
    lbl(14).Caption = "‰”»…"
    Else
    lbl(14).Caption = "Percentage"
    End If
    TxtNetDiscount.text = Round(val(TxtNetDiscount.text), 2)
    ElseIf val(XPCboDiscountType.ListIndex) = 1 Then
    TxtNetDiscount.text = val(XPTxtDiscountVal.text)
    If SystemOptions.UserInterface = ArabicInterface Then
    lbl(14).Caption = "ﬁÌ„…"
    Else
    lbl(14).Caption = "Value"
    End If
    Else
    TxtNetDiscount.text = 0
    XPTxtDiscountVal.text = 0
    End If
lbldifdis.Caption = val(Me.LbtotalDis.Caption) - val(firstprice.Caption) - val(TxtNetDiscount.text) - val(TxtPaymentValue.text) + val(TxtFATValue.text)
    If XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = 3 Or XPCboDiscountType.ListIndex = -1 Then
    
        XPTxtDiscountVal.Enabled = False
        XPTxtDiscountVal.text = ""
    Else
    
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.text = ""
    End If

 

    Me.lbl(55).Visible = (Me.XPCboDiscountType.ListIndex = 2)

    'Me.lbl(21).Visible = (Me.XPCboDiscountType.ListIndex = 2)
    If XPCboDiscountType.ListIndex = 0 Then
        lbl(8).Visible = False
        XPTxtDiscountVal.Visible = False
        lbl(8).Visible = False
    Else
        lbl(8).Visible = True
        XPTxtDiscountVal.Visible = True
        lbl(8).Visible = True
    End If
    FillCal
End If
    Exit Sub
    
ErrTrap:
End Sub
Sub FillCal()
Dim i As Integer
With Fg
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("nett"))) <> 0 Then
If val(XPCboDiscountType.ListIndex) = 1 Then
.TextMatrix(i, .ColIndex("Percentage")) = val(lbTotalMenteDis.Caption) / val(.TextMatrix(i, .ColIndex("nett")))
.TextMatrix(i, .ColIndex("Percentage")) = val(.TextMatrix(i, .ColIndex("Percentage")))
.TextMatrix(i, .ColIndex("DiscValue")) = (val(.TextMatrix(i, .ColIndex("Percentage"))) * val(XPTxtDiscountVal.text))
.TextMatrix(i, .ColIndex("TotalNet")) = val(.TextMatrix(i, .ColIndex("nett"))) - val(.TextMatrix(i, .ColIndex("DiscValue")))
ElseIf val(XPCboDiscountType.ListIndex) = 2 Then
.TextMatrix(i, .ColIndex("Percentage")) = val(XPTxtDiscountVal.text)
.TextMatrix(i, .ColIndex("DiscValue")) = (val(.TextMatrix(i, .ColIndex("nett"))) * val(.TextMatrix(i, .ColIndex("Percentage")))) / 100
.TextMatrix(i, .ColIndex("TotalNet")) = val(.TextMatrix(i, .ColIndex("nett"))) - val(.TextMatrix(i, .ColIndex("DiscValue")))
Else
.TextMatrix(i, .ColIndex("Percentage")) = 0
.TextMatrix(i, .ColIndex("DiscValue")) = 0
.TextMatrix(i, .ColIndex("TotalNet")) = val(.TextMatrix(i, .ColIndex("nett")))
End If
End If
Next i
End With
Calculte
End Sub

Private Sub ChAccept_Click()
If Me.ChAccept.value = vbChecked Then
Me.DcbOrderStatus.ListIndex = 1
End If
End Sub
Sub imgg()
 'Me.img9.Picture = Me.imgnul.Picture
 '       Me.img10.Picture = Me.imgnul.Picture
 '       Me.imag1.Picture = Me.imgnul.Picture
 '       Me.imag2.Picture = Me.imgnul.Picture
 '       Me.imag3.Picture = Me.imgnul.Picture
 '       Me.imag4.Picture = Me.imgnul.Picture
 '       Me.imag5.Picture = Me.imgnul.Picture
 '       Me.img6.Picture = Me.imgnul.Picture
 '       Me.img7.Picture = Me.imgnul.Picture
 '       Me.img8.Picture = Me.imgnul.Picture
End Sub
Private Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index

Case 8
    ShowGL_cc Me.TxtNoteSerial.text, , 200, val(Me.TXTNoteID.text)
        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            XPDtbTrans.value = Date
            DcbBasedOn.ListIndex = 0
            DcbBasedOn_Change
            ClculteVAT
Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 2
            fg2.Enabled = True
            
            grdCash.Clear flexClearScrollable, flexClearEverything
            grdCash.Rows = 2
            grdCash.Enabled = True
            
            vchrgrid.Clear flexClearScrollable, flexClearEverything
            vchrgrid.Rows = 2
             Me.DCboUserName.BoundText = user_id
        imgg
        lbldifdis.Caption = 0
        LbToTalExtra.Caption = 0
        lblTotalPay.Caption = 0
        lbllblTotalVat.Caption = 0
        firstprice.Caption = 0
        lbTotalMenteDis.Caption = 0
        LbtotalDis.Caption = 0
            Me.Lbtotal.Caption = 0
            Me.LbToTalExtra.Caption = 0
            
            Me.lbTotalMente.Caption = 0
     Me.DcbOrderStatus.ListIndex = 0
   ' Me.ComGranty.ListIndex = 1
           ' TxtPaymentCounts.text = 1
dcBranch.BoundText = Current_branch
            'XPDtbTrans.SetFocus
            
          '  Accredit.Enabled = True
             '   If SystemOptions.UserInterface = ArabicInterface Then
                        '                            Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                  '                                Else
                                         '           Accredit.Caption = " send to Approval   "
                            '                   End If
                              '
          CboPayMentType.ListIndex = 0
            XPCboDiscountType.ListIndex = 0
            DcbBasedOn.ListIndex = 2
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
 Fg.Rows = Fg.Rows + 1
            Fg.Enabled = True
            fg2.Rows = fg2.Rows + 1
            fg2.Enabled = True
            TxtModFlg.text = "E"
            TxtPaymentValue_Change
            Me.DCboUserName.BoundText = user_id

        Case 2
            If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              If val(CboPayMentType.ListIndex) = 0 Then
              If val(Me.TxtPaymentValue.text) <= 0 Then
              If SystemOptions.UserInterface = ArabicInterface Then
              MsgBox "ÌÃ» «œŒ«· «·ﬁÌ„… «·„œ›Ê⁄…"
              Else
              MsgBox "Please enter value"
              End If
              Exit Sub
              End If
            If Round(val(Me.lbldifdis.Caption)) < 0 Then
              If SystemOptions.UserInterface = ArabicInterface Then
              MsgBox "«·ﬁÌ„… «·„œ›Ê⁄… «ﬂ»— „‰ «·«Ã„«·Ì"
              Else
              MsgBox "The value larger than Total"
              End If
              Exit Sub
              End If
              End If
            Dim Msg As String

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.dcBranch.BoundText

 
                       
    If TxtNoteSerial.text = "" Then
        If Notes_coding(val(my_branch), XPDtbTrans.value) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        Else
                       
            If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
            Else
                '                       TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
            End If
        End If
    End If

Dim TxtNoteSerial1str As String

    If TxtNoteSerial1.text = "" Then
    TxtNoteSerial1str = Voucher_coding(val(my_branch), XPDtbTrans.value, 50, 5050)
    
                If TxtNoteSerial1str = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›…   ›« Ê—… ’Ì«‰…  ÃœÌœ… ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                               
                    If TxtNoteSerial1str = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ›« Ê—…  «·’Ì«‰…  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                        '             txtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21, DCPreFix.text)
                    End If
                End If
    End If
    If val(DcbBasedOn.ListIndex) = -1 Then

If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «·›« Ê—…"
Else
MsgBox "Please Type Bills"
End If
DcbBasedOn.SetFocus
Exit Sub
End If
If val(DcbBasedOn.ListIndex) = 0 Then
txtCusId1.text = val(DcbCustmer.BoundText)
End If
Dim DebitAccount As String
If Round(val(lbldifdis.Caption), 2) > 0 Then
  DebitAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbCustmer.BoundText), "Account_code")
         
If val(DcbCustmer.BoundText) = 0 Or DebitAccount = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄„Ì·"
Else
MsgBox "Please select Customer"
End If
DcbCustmer.SetFocus
Exit Sub
End If
End If
If val(CboPayMentType.ListIndex) = 1 Then
If val(DcbCustmer.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄„Ì·"
Else
MsgBox "Please select Customer"
End If
DcbCustmer.SetFocus
Exit Sub
End If
End If
Dim AccountVATDept As String
If AccountVat.BoundText = "" And True = True And CheckAnyVAT = True Then
MsgBox "Ì—ÃÏ ÷»ÿ «⁄œ«œ  «·ﬁÌ„… «·„÷«›…"
Exit Sub
End If
            SaveData

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
        Me.Ch = True
     
        FrmBillCarMaintExtrSearch.IndTyp = 0
           Load FrmBillCarMaintExtrSearch
            FrmBillCarMaintExtrSearch.IndTyp = 0
             FrmBillCarMaintExtrSearch.show vbModal

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.text, , 200

        Case 8
            'CalCulateParts
            
            
                 Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
                print_report val(Me.XPTxtID.text), 0
        
        
            End If
                Case 10

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
                print_report val(Me.XPTxtID.text), 1
        
        
            End If
        
    End Select

    Exit Sub
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String, Optional indexe As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    Dim cCompanyInfo As New ClsCompanyInfo
    
             SaveQRCode "TblCarBillMentains", "ID", val(XPTxtID), TxtNoteSerial1.text, (XPDtbTrans.value), _
        (TxtTotalValue.text), Picture1, 0, (TxtFATValue.text), (TxtTotalValue.text)

    
MySQL = " SELECT      " & cCompanyInfo.VATRegNo & " as VATRegNo , TblCarBillMentains.QrCodeImage, dbo.TblCarModels.Model, dbo.TblCarModels.CarID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblBranchesData.branch_name, "
MySQL = MySQL & "                       dbo.TblColor.name AS colorname, dbo.TblColor.namee AS colornamee, dbo.TblCarBillMentainsDetils.ID2, dbo.TblCarBillMentainsDetils.ID,"
MySQL = MySQL & "                       dbo.TblCarBillMentainsDetils.Type, dbo.TblCarBillMentainsDetils.[Value], dbo.TblCarBillMentainsDetils.[count], dbo.TblCarBillMentainsDetils.comp,"
MySQL = MySQL & "                       dbo.TblCarBillMentainsDetils.bill, dbo.TblCarBillMentainsDetils.Mainte, dbo.TblCarBillMentains.RecordDate, dbo.TblCarBillMentains.ClientName,"
MySQL = MySQL & "                       dbo.TblCarBillMentains.Telephone, dbo.TblCarBillMentains.UserID, dbo.TblCarBillMentains.CarModelID, dbo.TblCarBillMentains.PlateNo,"
MySQL = MySQL & "                       dbo.TblCarBillMentains.YearFact, dbo.TblCarBillMentains.OrderStatus, dbo.TblCarBillMentains.Accept, dbo.TblCarBillMentains.EndDate,"
MySQL = MySQL & "                       dbo.TblCarBillMentains.Granty, dbo.TblCarBillMentains.Month_Day, dbo.TblCarBillMentains.CarMeter, dbo.TblCarBillMentains.PayFirst,"
MySQL = MySQL & "                       dbo.TblCarBillMentains.AmountAccept, dbo.TblCarBillMentains.Complaint, dbo.TblCarBillMentains.Noteinitial, dbo.TblCarBillMentains.CarTypeID,"
MySQL = MySQL & "                       dbo.TblCarBillMentains.BranchID, dbo.TblCarBillMentains.ID AS IdM, dbo.TblExtraExpeneses.name AS nameExt, dbo.TblExtraExpeneses.namee AS nameeExt,"
MySQL = MySQL & "                       dbo.TblMaintenanceWork.name AS nameM, dbo.TblMaintenanceWork.namee AS nameeM, dbo.TblCarBillMentains.OverKM, dbo.TblCarBillMentains.WorkOrderNO,"
MySQL = MySQL & "                       dbo.TblCarBillMentains.CusID, dbo.TblCarBillMentains.NoteID, dbo.TblCarBillMentains.NoteSerial1, dbo.TblCarBillMentains.NoteSerial,"
MySQL = MySQL & "                       dbo.TblCarBillMentains.LongGranty, dbo.TblCarBillMentains.DateStartG, dbo.TblCarBillMentains.DateEndG, dbo.TblCarBillMentainsDetils.ValiueDis,"
MySQL = MySQL & "                       dbo.TblCarBillMentainsDetils.fittervalue, dbo.TblCarBillMentains.CarMetarOut, dbo.TblCarBillMentains.Trans_Discount, dbo.TblCarBillMentains.Trans_DiscountType,"
MySQL = MySQL & "                       dbo.TblCarBillMentains.PaymentType, dbo.TblCustemers.CusName, TblCustemers.VatNo,dbo.TblCustemers.CusNamee, dbo.TblCustemers.CustGID, dbo.TblCarModels.ModelE,"
MySQL = MySQL & "                       dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Fullcode, dbo.TblCarBillMentains.PaymentValue,"
MySQL = MySQL & "                       dbo.TblCarBillMentains.difdisValue, dbo.TblCarBillMentains.NetDiscount, dbo.TblCarBillMentains.SparePart, dbo.TblCarBillMentainsDetils.Emp_ID,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS Expr1, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Name3 , dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality"
MySQL = MySQL & "  FROM         dbo.TblExtraExpeneses RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblMaintenanceWork RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCarBillMentainsDetils LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee ON dbo.TblCarBillMentainsDetils.Emp_ID = dbo.TblEmployee.Emp_ID ON dbo.TblMaintenanceWork.Id = dbo.TblCarBillMentainsDetils.Mainte ON"
MySQL = MySQL & "                       dbo.TblExtraExpeneses.Id = dbo.TblCarBillMentainsDetils.Mainte RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblColor RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCarBillMentains LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCustemers ON dbo.TblCarBillMentains.CusID = dbo.TblCustemers.CusID ON dbo.TblColor.Id = dbo.TblCarBillMentains.ColorID ON"
MySQL = MySQL & "                       dbo.TblBranchesData.branch_id = dbo.TblCarBillMentains.BranchID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCarModels ON dbo.TblCarBillMentains.CarModelID = dbo.TblCarModels.Id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TBLCarTypes ON dbo.TblCarBillMentains.CarTypeID = dbo.TBLCarTypes.id ON dbo.TblCarBillMentainsDetils.ID = dbo.TblCarBillMentains.ID"
MySQL = MySQL & "  Where (dbo.TblCarBillMentainsDetils.id = " & val(XPTxtID.text) & ")"
If indexe = 1 Then
If ComGranty.ListIndex = 0 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCarBillMainteneFGrantyDis.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCarBillMainteneFGrantyDis.rpt"
        End If
Else
  If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCarBillMainteneFDis.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCarBillMainteneFDis.rpt"
        End If
        End If
        ''''''
Else
If ComGranty.ListIndex = 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCarBillMainteneFGranty.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCarBillMainteneFGranty.rpt"
        End If
Else
  If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCarBillMainteneF.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCarBillMainteneF.rpt"
        End If
        End If
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub



Private Sub ComGranty_Change()
If val(DcbBasedOn.ListIndex) = 0 Then
If Me.ComGranty.ListIndex = 0 Then  '"»÷„«‰" Then
frmgranty.Visible = True
Else
frmgranty.Visible = False
End If
Else
frmgranty.Visible = True
End If

End Sub

Private Sub ComGranty_Click()
If Me.ComGranty.ListIndex = 0 Then  '"»÷„«‰" Then

frmgranty.Visible = True
Else

frmgranty.Visible = False
End If


End Sub

Private Sub ComMD_Change()
Dim NewDate, old As Date
NewDate = Me.DateStartG
If Me.ComMD.ListIndex = 0 Then
old = DateAdd("m", val(Me.TxtLongGranty.text), NewDate)
Else
old = DateAdd("d", val(Me.TxtLongGranty.text), NewDate)
End If
Me.DateEndg = old
End Sub

Private Sub ComMD_Click()
Dim NewDate, old As Date
NewDate = Me.DateStartG
If Me.ComMD.ListIndex = 0 Then
old = DateAdd("m", val(Me.TxtLongGranty.text), NewDate)
Else
old = DateAdd("d", val(Me.TxtLongGranty.text), NewDate)
End If
Me.DateEndg = old
End Sub

Private Sub DcbCarType_Change()
Dim Dcombos As ClsDataCombos
      Set Dcombos = New ClsDataCombos
    
      If val(Me.DcbCarType.BoundText) <> 0 Then
      
   Dcombos.GetTblCarModels Me.DcbCarModel, , val(Me.DcbCarType.BoundText)
   End If
End Sub

Private Sub DcbCarType_Click(Area As Integer)
 DcbCarType_Change
   
   
   
End Sub


Function newret()
  Dim RsDetails1 As New ADODB.Recordset
Dim StrSQL As String
Dim i As Integer
vchrgrid.Clear flexClearScrollable, flexClearEverything
            vchrgrid.Rows = 2
      ReLineGrid2
   If val(DcbBasedOn.ListIndex) = 0 Then Exit Function
StrSQL = "SELECT     dbo.Transactions.Transaction_ID,Transactions.StoreID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, "
StrSQL = StrSQL & "                      dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_HijriDate, dbo.Transactions.TransactionComment, dbo.Transactions.OpOrderID,"
StrSQL = StrSQL & "                      dbo.Transactions.OldOpOrderID, dbo.Transaction_Details.UnitId,dbo.Transaction_Details.OperPrice, dbo.Transaction_Details.ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.Item_ID,"
StrSQL = StrSQL & "                      dbo.TblItems.itemname ,dbo.Transaction_Details.UnitId, dbo.TblItems.ItemNamee, dbo.TblItems.fullcode , dbo.Transaction_Details.showPrice"
StrSQL = StrSQL & " FROM         dbo.TblItems RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_Type = 19) "

If val(DcbBasedOn.ListIndex) = 1 Then
            
   StrSQL = StrSQL & "  And (dbo.Transactions.OpOrderID = " & val(TxtReq.text) & ")"
ElseIf val(DcbBasedOn.ListIndex) = 2 Then

   StrSQL = StrSQL & "  And (dbo.Transactions.RepairOrder = " & val(TxtAuthoOrder) & ")"
End If
    RsDetails1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
Dim mPrice As Double
    If Not (RsDetails1.BOF Or RsDetails1.EOF) Then
       With Me.vchrgrid
      '  RsDetails1.MoveFirst
        .Rows = .FixedRows + RsDetails1.RecordCount
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Ser")) = i
            .TextMatrix(i, .ColIndex("Transaction_Date")) = (IIf(IsNull(RsDetails1("Transaction_Date").value), "", RsDetails1("Transaction_Date").value))
            .TextMatrix(i, .ColIndex("NoteSerial1")) = val(IIf(IsNull(RsDetails1("NoteSerial1").value), "", RsDetails1("NoteSerial1").value))
            .TextMatrix(i, .ColIndex("TransactionComment")) = (IIf(IsNull(RsDetails1("TransactionComment").value), "", RsDetails1("TransactionComment").value))
            .TextMatrix(i, .ColIndex("Transaction_ID")) = (IIf(IsNull(RsDetails1("Transaction_ID").value), 0, RsDetails1("Transaction_ID").value))
            .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(RsDetails1("ID").value), 0, RsDetails1("ID").value)
            .TextMatrix(i, .ColIndex("ShowQty")) = (IIf(IsNull(RsDetails1("ShowQty").value), 0, RsDetails1("ShowQty").value))
            .TextMatrix(i, .ColIndex("Item_ID")) = (IIf(IsNull(RsDetails1("Item_ID").value), 0, RsDetails1("Item_ID").value))
            
            mPrice = val(RsDetails1!OperPrice & "")
            
            If mPrice = 0 Then
                mPrice = GetItemPrice(val(.TextMatrix(i, .ColIndex("Item_ID"))), 1, IIf(IsNull(RsDetails1("UnitId").value), 0, RsDetails1("UnitId").value)) '''(IIf(IsNull(RsDetails1("ShowPrice").value), 0, RsDetails1("ShowPrice").value))
            End If
            If mPrice = 0 Then
                mPrice = ModItemCostPrice.GetCostItemPrice(val(RsDetails1!Item_ID & ""), 0, "", , SystemOptions.SysMainStockCostMethod, , , XPDtbTrans.value, val(RsDetails1!Transaction_ID & ""), val(RsDetails1!UnitID & ""), val(RsDetails1!StoreID & ""))
                mPrice = mPrice * (1 + 0.3)
            End If
      
            .TextMatrix(i, .ColIndex("OperPrice")) = mPrice
            
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("ItemName")) = (IIf(IsNull(RsDetails1("ItemName").value), "", RsDetails1("ItemName").value))
            Else
            .TextMatrix(i, .ColIndex("ItemName")) = (IIf(IsNull(RsDetails1("ItemNamee").value), "", RsDetails1("ItemNamee").value))
            End If
            RsDetails1.MoveNext
         
        Next i
    ReLineGrid2
    
End With
End If
End Function
Private Sub ReLineGrid2()
    Dim i As Integer
    Dim IntCounter As Integer
    Dim summ As Double
   ''''///
    summ = 0
IntCounter = 0
lbl(58).Caption = 0
        With Me.vchrgrid
        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Transaction_ID")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("ShowQty"))) * val(.TextMatrix(i, .ColIndex("OperPrice")))
                summ = summ + val(.TextMatrix(i, .ColIndex("Total")))
             
                  End If
        Next i
    End With
    lbl(58).Caption = summ
    'salimhere
    lbl(58).Caption = 0
Me.Lbtotal.Caption = val(Me.LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption) + val(lbl(58).Caption)
LbtotalDis.Caption = val(LbToTalExtra.Caption) + val(Me.lbTotalMenteDis.Caption) + val(lbl(58).Caption)

    End Sub
    
Private Sub ReLineGrid3()
    Dim i As Integer
    Dim IntCounter As Integer
    Dim summ As Double
     Dim summ2 As Double
   ''''///
    summ = 0
    summ2 = 0
IntCounter = 0
lblTotalPay = 0
        With Me.grdCash
        For i = .FixedRows To .Rows - 1
            summ = summ + val(.TextMatrix(i, .ColIndex("Note_Value")))
            summ2 = summ2 + val(.TextMatrix(i, .ColIndex("Vat")))
        Next i
    End With
    lblTotalPay.Caption = summ

    lbllblTotalVat.Caption = summ2
    End Sub
    
    Private Sub vchrgrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
ReLineGrid2
End Sub
Private Sub vchrgrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With vchrgrid
Select Case .ColKey(Col)
Case "NoteSerial1"
Cancel = True
Case "Transaction_Date"
Cancel = True
Case "ItemName"
Cancel = True
Case "ShowQty"
Cancel = True
Case "Total"
Cancel = True
Case "TransactionComment"
Cancel = True
End Select
End With
End Sub
Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
Dim StrAccountCode1 As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim StrComboList As String
            
    
    With Fg

        Select Case .ColKey(Col)
         Case "Emp_Name"
                 StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Emp_ID"), False, True)
                .TextMatrix(Row, .ColIndex("Emp_ID")) = StrAccountCode
           Case "DepartmentName"
                 StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Deptid"), False, True)
                .TextMatrix(Row, .ColIndex("Deptid")) = StrAccountCode
                StrSQL = "select AccountCode from  TblEmpDepartments where DeparmentID=" & val(.TextMatrix(Row, .ColIndex("Deptid"))) & ""
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                 If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("AccountCode")) = IIf(IsNull(rs("AccountCode").value), "", rs("AccountCode").value)
                Else
                    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                End If
            Case "name"
                 StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("cod"), False, True)
                .TextMatrix(Row, .ColIndex("cod")) = StrAccountCode
                 StrSQL = "select * from TblMaintenanceWork where Id=" & val(StrAccountCode)
                 rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("value")) = IIf(IsNull(rs("InitialPrice").value), 0, rs("InitialPrice").value)
                Else
                    .TextMatrix(Row, .ColIndex("value")) = ""
                End If
      End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With Fg

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
        
            If val(DcbBasedOn.ListIndex) = 0 Then
            Select Case .ColKey(Col)
            Case "name"
         If .TextMatrix(Row, .ColIndex("AccountCode")) = "" Then
         Cancel = True
            MsgBox "Ì—ÃÏ «Œ Ì— «·ﬁ”„ «Ê·« «Ê «· «ﬂœ „‰ —»ÿ «·ﬁ”„ »«·Õ”«»« "
            Exit Sub
          Else
          Cancel = False
            End If
              Case "TotalNet"
             Cancel = True
             
             Case "nett"
             Cancel = True
              Case "valuedis"
             Fg.ComboList = ""
             
            Case "cod"
               Fg.ComboList = ""
            Case "value"
             Fg.ComboList = ""
             '  fg.ComboList = ""
               
               Case "count"
               Fg.ComboList = ""
               Case "totalm"
               Fg.ComboList = ""
           End Select

            Else
            Select Case .ColKey(Col)
            Case "TotalNet"
             Cancel = True
              Case "nett"
              Cancel = True
              Case "valuedis"
             Cancel = True
             
                Case "cod"
                Cancel = True
            Case "value"
             Cancel = True
             '  fg.ComboList = ""
               Case "name"
                Cancel = True
               Case "count"
               Cancel = True
               Case "totalm"
               Cancel = True
               'fg.ComboList = ""
             '    Cancel = True
             End Select

            End If
        
    End With

    
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg

        Select Case .ColKey(Col)
 Case "DepartmentName"
      StrSQL = "SELECT  DISTINCT    dbo.TblEmpDepartments.DeparmentID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpDepartments.Dpeterial, "
      StrSQL = StrSQL & "                    dbo.TblEmpDepartments.DeptBr"
      StrSQL = StrSQL & "  FROM         dbo.SuperTech INNER JOIN"
      StrSQL = StrSQL & "                     dbo.TblEmpDepartments ON dbo.SuperTech.DeparmentID = dbo.TblEmpDepartments.DeparmentID"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "DepartmentName", "DeparmentID")
                Else
                    StrComboList = .BuildComboList(rs, "DepartmentNamee", "DeparmentID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
         
         
            Case "name"
      
                StrSQL = "select * from TblMaintenanceWork"
                If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = StrSQL & " order by name"
                Else
                StrSQL = StrSQL & " order by namE"
                End If
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = Fg.BuildComboList(rs, "namE", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList

      Case "Emp_Name"
                 StrSQL = " SELECT     dbo.Technicians1.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.Technicians1.Emp_ID1"
                 StrSQL = StrSQL & "              FROM         dbo.Technicians1 LEFT OUTER JOIN"
                 StrSQL = StrSQL & "      dbo.TblEmployee ON dbo.Technicians1.Emp_ID1 = dbo.TblEmployee.Emp_ID"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

             If SystemOptions.UserInterface = ArabicInterface Then
                   StrComboList = Fg.BuildComboList(rs, "Emp_Name", "Emp_ID1")
             Else
                  StrComboList = Fg.BuildComboList(rs, "Emp_Namee", "Emp_ID1")
             End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub FG2_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With fg2

        Select Case .ColKey(Col)
 
            Case "name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("cod"), False, True)
                .TextMatrix(Row, .ColIndex("cod")) = StrAccountCode
                   End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub

Private Sub FG2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With fg2

     
        If val(DcbBasedOn.ListIndex) = 0 Then
        Select Case .ColKey(Col)
            
            Case "cod"
               .ComboList = ""
            Case "value"
             .ComboList = ""
             '  fg.ComboList = ""
             
               Case "count"
               .ComboList = ""
               Case "totalex"
               Cancel = True
               Case "comp"
               .ComboList = ""
               Case "bill"
               .ComboList = ""
        End Select
      Else
              Select Case .ColKey(Col)
            
            Case "cod"
               Cancel = True
            Case "value"
             Cancel = True
             '  fg.ComboList = ""
               Case "name"
               Cancel = True
               Case "count"
               Cancel = True
               Case "totalex"
               Cancel = True
               Case "comp"
               Cancel = True
               Case "bill"
               Cancel = True
        End Select
      End If

    End With
   

    fg2.ComboList = ""
End Sub

 Sub Enable()
 Me.DcbOrderStatus.Enabled = False
    Me.TxtCliientName.Enabled = False
    Me.DcbCarModel.Enabled = False
    Me.DcbCarType.Enabled = False
    Me.DcbColor.Enabled = False
    Me.TXtCarMeter.Enabled = False
    Me.TxtClientPhone.Enabled = False
    Me.DcbOrderStatus.Enabled = False
    Me.TxtPlatNo.Enabled = False
    Me.TxtAmoutAccept.Enabled = False
    Me.ChAccept.Enabled = False
    Me.TxtEndDate.Enabled = False
    Me.DcbyearFactor.Enabled = False
    Me.TxtFirstPrice.Enabled = False
    Me.dcBranch.Enabled = True
    Me.XPDtbTrans.Enabled = True
    Me.DcbCar.Enabled = False
    
    End Sub

Private Sub FG2_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With fg2

        Select Case .ColKey(Col)

            Case "name"
                StrSQL = "select * from TblExtraExpeneses"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = fg2.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = fg2.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub


Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub Lbtotal_Change()

lbldif.Caption = val(Me.Lbtotal.Caption) - val(firstprice.Caption)
End Sub



Private Sub LbtotalDis_Change()
If Me.TxtModFlg.text <> "R" Then
lbldifdis.Caption = val(Me.LbtotalDis.Caption) - val(firstprice.Caption) - val(TxtNetDiscount.text) - val(TxtPaymentValue.text) + val(TxtFATValue.text)
Calculte
End If
End Sub



Private Sub menue_Click(Index As Integer)
showsforms Index
End Sub

Private Sub TxtFirstPrice_Change()
Me.Lbtotal.Caption = val(Me.LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
firstprice.Caption = TxtFirstPrice.text
End Sub

Private Sub TxtLongGranty_Change()
Dim NewDate, old As Date
NewDate = Me.DateStartG
If Me.ComMD.ListIndex = 0 Then
old = DateAdd("m", val(Me.TxtLongGranty.text), NewDate)
Else
old = DateAdd("d", val(Me.TxtLongGranty.text), NewDate)
End If
Me.DateEndg = old
End Sub

Sub returnfileds(i As Integer)
Me.retrive1 (i)
Enable
End Sub

Private Sub TxtReq_KeyDown(KeyCode As Integer, Shift As Integer)
If val(DcbBasedOn.ListIndex) = 1 Then
If KeyCode = vbKeyF3 Then
 Me.Ch = True
 
  FrmCarAuthoSearch.GetData1
   Load FrmCarAuthoSearch
            FrmCarAuthoSearch.show
           
End If
End If

ClculteVAT
End Sub

Private Sub txtAuthoOrder_KeyUp(KeyCode As Integer, Shift As Integer)





If val(DcbBasedOn.ListIndex) = 2 Then
Dim StrSQL As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
If KeyCode = vbKeyReturn Then
            If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
            
            
            
            If DcbBasedOn.ListIndex = 2 Then
                StrSQL = "select * From TblCardAuthorizationReform  WHERE WorkOrder=" & val(TxtAuthoOrder.text) & "   Order By ID"
            End If
            Rs4.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Rs4.RecordCount > 0 Then
                mIdTrans = (Rs4("id").value)
                returnfileds (Rs4("id").value)
                Else
                MsgBox "·«Ì ÊÃœ »Ì«‰«  »Â–« «·«„—"
                            clear_all Me
            
Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 2
            fg2.Enabled = True
        imgg
            Me.Lbtotal.Caption = 0
            Me.LbToTalExtra.Caption = 0
            
            Me.lbTotalMente.Caption = 0
     Me.DcbOrderStatus.ListIndex = 0
     DisplayCashInvoice
            End If
               Me.Ch = True
            End If
End If
End If

ClculteVAT
End Sub


Private Sub TxtReq_KeyUp(KeyCode As Integer, Shift As Integer)


If val(DcbBasedOn.ListIndex) = 1 Then
Dim StrSQL As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
If KeyCode = vbKeyReturn Then
            If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
            
            
            If DcbBasedOn.ListIndex = 1 Then
                StrSQL = "select * From TblCardAuthorizationReform  WHERE WorkOrder=" & val(TxtReq.text) & "   Order By ID"
            ElseIf DcbBasedOn.ListIndex = 2 Then
                StrSQL = "select * From TblCardAuthorizationReform  WHERE WorkOrder=" & val(TxtAuthoOrder.text) & "   Order By ID"
            End If
            Rs4.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Rs4.RecordCount > 0 Then
                returnfileds (Rs4("id").value)
                Else
                MsgBox "·«Ì ÊÃœ »Ì«‰«  »Â–« «·«„—"
                            clear_all Me
            
Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 2
            fg2.Enabled = True
        imgg
            Me.Lbtotal.Caption = 0
            Me.LbToTalExtra.Caption = 0
            
            Me.lbTotalMente.Caption = 0
     Me.DcbOrderStatus.ListIndex = 0
     
            End If
               Me.Ch = True
            End If
End If
End If

ClculteVAT
End Sub
Private Sub DisplayCashInvoice()
If DcbBasedOn.ListIndex <> 2 Then Exit Sub
Dim s As String
Dim Rs4 As New ADODB.Recordset

    s = "select * From TblCardAuthorizationReform  WHERE AuthoOrder=" & val(TxtAuthoOrder.text) & "   Order By ID"

Rs4.Open s, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
    mIdTrans = (Rs4("id").value)
End If
                

s = "select * From Notes where NoteType=4    AND branch_no in(" & Current_branchSql & ")"
s = s & " and akarid is Null"
s = s & " And IsNull(CashingType,0) = 10 "
s = s & " And BillMaintID = " & mIdTrans

s = s & " and  displayed is null Order By NoteID"

loadgrid s, grdCash, True, False
ReLineGrid3
ReLineGrid
End Sub
Private Sub XPDtbTrans_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
ClculteVAT
End Sub

Private Sub Dcbranch_Click(Area As Integer)
 
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub
Sub bill(Optional X As Integer = 0)
 If SystemOptions.ShowBillCommisions = 1 Then
 lbTotalMenteDis.Visible = False
 lbTotalMente.Visible = True
 Lbtotal.Visible = True
 LbtotalDis.Visible = False
 lbldifdis.Visible = False
 lbldif.Visible = True
 'Cmd(9).Visible = False
 Cmd(10).Visible = True
 Fg.ColHidden(10) = False
 Fg.ColHidden(11) = True
 Fg.ColHidden(13) = False
 Fg.ColHidden(12) = True
 Else
  lbTotalMenteDis.Visible = True
 lbTotalMente.Visible = False
 Lbtotal.Visible = False
 LbtotalDis.Visible = True
 lbldifdis.Visible = True
 lbldif.Visible = False
 Cmd(9).Visible = True
 Cmd(10).Visible = False
  Fg.ColHidden(10) = True
 Fg.ColHidden(11) = False
 Fg.ColHidden(13) = True
 Fg.ColHidden(12) = False
 End If
End Sub
Private Sub Form_Load()
Dim X As Integer
    Dim Dcombos As ClsDataCombos
      
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

 On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic
If SystemOptions.LinkCustomerWithCars = True Then
TxtPlatNo.Visible = False
DcbCar.Visible = True
Else
TxtPlatNo.Visible = True
DcbCar.Visible = False
End If
bill

  '  With Me.Fg
  '      .RowHeightMin = 300
  '      .WallPaper = GrdBack.Picture
  '      .AutoSize 0, .Cols - 1, False
  '  End With

 
  If SystemOptions.UserInterface = EnglishInterface Then
    
        Me.DcbOrderStatus.AddItem "New"
        Me.DcbOrderStatus.AddItem "Accept Customer"
        Me.DcbOrderStatus.AddItem "Final Maintenance"
        Me.DcbOrderStatus.AddItem "Under Wait"
        DcbOrderStatus.AddItem "Not Accept"
        Me.DcbOrderStatus.AddItem "Was Recognized"
      Me.ComMD.AddItem "Month"
        Me.ComMD.AddItem "Day"
           With XPCboDiscountType
            .Clear
            .AddItem "NA"
            .AddItem "Discount Val"
            .AddItem "Discount %"
        End With

        With CboPayMentType
            .Clear
            .AddItem "Cash"
            .AddItem "Credit"
        End With
 With DcbBasedOn
 .Clear
 .AddItem "None"
 .AddItem "Order No"
 .AddItem "Repair #"
 End With
             Else
 With DcbBasedOn
 .Clear
 .AddItem "»·«"
 .AddItem "«„— ‘€·"
 .AddItem "«„— «’·«Õ"
 End With
 
 DcbOrderStatus.AddItem "ÃœÌœ"
DcbOrderStatus.AddItem " „ „Ê«›ﬁ… «·⁄„Ì·"
DcbOrderStatus.AddItem " „ «‰Â«¡ «·«’·«Õ"
Me.DcbOrderStatus.AddItem " Õ  «·«‰ Ÿ«—"
DcbOrderStatus.AddItem "⁄œ„ „Ê«›ﬁ… «·⁄„Ì·"
DcbOrderStatus.AddItem " „ «’œ«— ›« Ê—…"
 Me.ComMD.AddItem "‘Â—"
        Me.ComMD.AddItem "ÌÊ„"
   
   If SystemOptions.UserInterface = EnglishInterface Then
   
     
   Else
       With XPCboDiscountType
                .Clear
                .AddItem "·«ÌÊÃœ Œ’„"
                .AddItem "Œ’„ »ﬁÌ„…"
                .AddItem "Œ’„ »‰”»…"
            End With
    
            With CboPayMentType
                .Clear
                .AddItem "‰ﬁœ«"
                .AddItem "¬Ã·"
            End With
        End If
    End If
     If SystemOptions.UserInterface = EnglishInterface Then
        Me.ComGranty.AddItem "Granty"
        Me.ComGranty.AddItem "With out Granty"
        'Me.ComGranty.AddItem "Re Maintenance"
      
             Else
            
         Me.ComGranty.AddItem "»÷„«‰"
        Me.ComGranty.AddItem "»œÊ‰ ÷„«‰"
' Me.ComGranty.AddItem "≈⁄«œ… «’·«Õ"
 
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
    Resize_Form Me
    AddTip
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetUsers Me.DCboUserName
Dcombos.GetBoxes Me.DcboBox
Dcombos.GetCustomersSuppliers 1, Me.DcbCustmer
  'Dcombos.GetTblyearsData Me.DcbyearFactor
 

    Dcombos.GetUsers Me.DCboUserName
  Dcombos.GetTblCarsDataGroup Me.DcbCarType
    Dcombos.GetTblColor Me.DcbColor
    Dim i As Integer
      For i = 1995 To 2100
      Me.DcbyearFactor.AddItem (i)
      Next i
   
      
   Dcombos.GetTblCarModels Me.DcbCarModel
   
  ' Dcombos.GetTblCarModels Me.DcbCarModel
   Enable
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
   ' YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblCarBillMentains     Order By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
        Me.TxtModFlg.text = "R"
    Retrive


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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Label1(66).Caption = "VAT%"
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
   ' Label1.Visible = False
   lbl(17).Caption = "Payment"
   Cmd(10).Caption = "Print Customer"
   Cmd(8).Caption = "Print Reg"
   lbl(10).Caption = "Reg No"
   lbl(11).Caption = "Box"
   'Label1.Caption = "Meter Out"
   Label4.Caption = "Customer"
lblty.Caption = "Order Type"
lbl(14).Caption = "Value"
lbl(12).Caption = "Type"
BtnShow.Caption = "Statement"
lbl(13).Caption = "Payment"
    Cmd(0).Caption = "New"
    lbl(15).Caption = "Spare Part"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
Me.lblchange.Caption = "Change"

XPTab301.CurrTab = 0
XPTab301.Caption = "ID card repair data|Reform work|Bills of exchange"
    
    Me.Caption = "Invoice"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
   Me.lblBr.Caption = "Branch"
   Me.lblDataCli.Caption = "Data of Client"
  Me.LblCli.Caption = "Client Name"
  Me.lblModel.Caption = "Models"
  Me.LblPhone.Caption = "Telephone"
  Me.LblCar.Caption = "Type of Car"
  Me.LblOrderSt.Caption = "Oreder Status"
  Me.lblColor.Caption = "Color"
  Me.LblWork.Caption = "Maintenance Work"
  Me.lblExt.Caption = "ExtraExpenes"
  Me.LblPla.Caption = "Plate No"
  Me.LblYear.Caption = "Year Manfac"
  Me.ChAccept.Caption = "Has the consent of the client"
  Me.lblEx.Caption = "Total of ExtraExpeneses"
  Me.LblM.Caption = "Total of MaintenanceWork"
  Me.Lbtota.Caption = "Total"
  lbl(2).Caption = "End Date"
  Me.lbreq.Caption = "Based on "
  lbl(5).Caption = "Start"
  lbl(9).Caption = "End"
  lbllong.Caption = "Long Granty"
  lbl(3).Caption = "Oil Change after"
  lbl(16).Caption = "KM"
' Me.lblty.Caption = "Type"
  Me.frmgranty.Caption = "Data Guarantee"
'  Me.lbllong.Caption = "Duration"
  Me.LblPayF.Caption = "Pay First"
  lbpricefirst.Caption = "Pay First"
  Me.LblAmountAcc.Caption = "Rate"
  Me.LblCarMeter.Caption = "Car Meter"
  Me.ChAccept.RightToLeft = False
    
    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"
lbl(57).Caption = "Total"
lbreq.Caption = "Based On"
Label8.Caption = "Total Payment"
     With Me.Fg
        .TextMatrix(0, .ColIndex("serial")) = "NO"
        .TextMatrix(0, .ColIndex("valuedis")) = "Value"
        .TextMatrix(0, .ColIndex("value")) = "Value"
        .TextMatrix(0, .ColIndex("name")) = "Name"
         .TextMatrix(0, .ColIndex("cod")) = "Code"
        .TextMatrix(0, .ColIndex("totalm")) = "Total"
         .TextMatrix(0, .ColIndex("nett")) = "Total"
       .TextMatrix(0, .ColIndex("count")) = "Count"
       .TextMatrix(0, .ColIndex("emp_Name")) = "Technical"
              .TextMatrix(0, .ColIndex("discValue")) = "disc Value"
              .TextMatrix(0, .ColIndex("TotalNet")) = "Total Net"
                     .TextMatrix(0, .ColIndex("Fullcode")) = "Technical Code"
       
       
    End With
      With Me.fg2
        .TextMatrix(0, .ColIndex("serial")) = "NO"
        .TextMatrix(0, .ColIndex("value")) = "Value"
        .TextMatrix(0, .ColIndex("name")) = "Name"
         .TextMatrix(0, .ColIndex("cod")) = "Code"
.TextMatrix(0, .ColIndex("bill")) = "Invoice No"
.TextMatrix(0, .ColIndex("count")) = "Count"
.TextMatrix(0, .ColIndex("comp")) = " Seller"
.TextMatrix(0, .ColIndex("totalex")) = "Total"
    End With

   With vchrgrid
    
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "Reciept No"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Date"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
        .TextMatrix(0, .ColIndex("ShowQty")) = "Quantity"
         .TextMatrix(0, .ColIndex("OperPrice")) = "Price"
        .TextMatrix(0, .ColIndex("Total")) = "Total"
       .TextMatrix(0, .ColIndex("View")) = "View"
       .TextMatrix(0, .ColIndex("TransactionComment")) = "Comment"


    End With
Label1(0).Caption = "Exit counter"
lbl(18).Caption = "Based On"
Frame2.Caption = "Advance payments"
Label8.Caption = "Advance payments Total"
Label1(68).Caption = "Total"
Label1(67).Caption = "Vat"
Lbtota.Caption = "General Total"
   With grdCash
    
        .TextMatrix(0, .ColIndex("NoteDate")) = "Date"
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "Order No"
        .TextMatrix(0, .ColIndex("Note_Value")) = "Value"
        .TextMatrix(0, .ColIndex("Vat")) = "Vat"



    End With

End Sub

'Private Sub YearMonth()

  '  Dim i As Integer
  '  Dim IntDefIndex As Integer

   ' CmbMonth.Clear

   ' For i = 1 To 12
     '   CmbMonth.AddItem MonthName(i)
  '  Next

   ' CmbMonth.ListIndex = Month(Date) - 1
    'CboYear.Clear

 '   For i = 2010 To 2050
 '       CboYear.AddItem i

      '  If i = year(Date) Then
         '   IntDefIndex = CboYear.NewIndex
       ' End If

    'Next

    'CboYear.ListIndex = IntDefIndex
'End Sub

Private Sub Form_Paint()
    TTD.Destroy
End Sub

Private Sub Form_Resize()
    TTD.Destroy
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
        Set rs = Nothing
    End If

    Set TTP = Nothing
    'Set EmpReport = Nothing
    TTD.Destroy
    Exit Sub
ErrTrap:
End Sub


Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰"
            Me.ChAccept.Enabled = True
   
     Me.DcbOrderStatus.Enabled = True
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
         '   TxtAdvanceValue.Locked = True
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
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
               Me.ChAccept.value = xtpUnchecked
            Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 2
            fg2.Enabled = True
    Me.ChAccept.Enabled = False
     Me.DcbOrderStatus.ListIndex = 0
     'Me.ComGranty.ListIndex = 0
     Me.DcbOrderStatus.Enabled = False
            Me.DCboUserName.BoundText = user_id
            '      Me.XPBtnMove(0).Enabled = False
            '      Me.XPBtnMove(1).Enabled = False
            '      Me.XPBtnMove(2).Enabled = False
            '      Me.XPBtnMove(3).Enabled = False
           ' TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰(  ⁄œÌ· )"
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
           ' TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            Me.ChAccept.Enabled = True
'     Me.DcbOrderStatus.ListIndex = 0
     Me.DcbOrderStatus.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
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
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
    Me.ChAccept = xtpUnchecked
Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 2
            fg2.Enabled = True
    'On Error GoTo ErrTrap
     
     
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
 imgg
    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
'''''''///////////////////
    Me.lbldifdis.Caption = IIf(IsNull(rs.Fields("difdisValue").value), "", rs.Fields("difdisValue").value)
    Me.TxtPaymentValue.text = IIf(IsNull(rs("PaymentValue").value), 0, rs("PaymentValue").value)
    XPTxtID.text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    TxtNetDiscount.text = IIf(IsNull(rs("NetDiscount").value), 0, rs("NetDiscount").value)
    TxtNoteID2.text = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
    TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
    TxtCusID.text = IIf(IsNull(rs("CusId").value), 0, rs("CusId").value)
   ' txtCusId1.text = IIf(IsNull(rs("CusId").value), "", rs("CusId").value)
    TxtCarMetarOut.text = IIf(IsNull(rs("CarMetarOut").value), "", rs("CarMetarOut").value)
    TxtReq.text = IIf(IsNull(rs("WorkOrderNO").value), "", rs("WorkOrderNO").value)
    DcbCustmer.BoundText = IIf(IsNull(rs("CusID").value), 0, (rs("CusID").value))
    XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), -1, (rs("Trans_DiscountType").value))
    CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
    XPTxtDiscountVal.text = IIf(IsNull(rs("Trans_Discount").value), "", (rs("Trans_Discount").value))
    DcbBasedOn.ListIndex = IIf(IsNull(rs("BasedOn").value), 1, rs("BasedOn").value)
    TxtFATYou.text = IIf(IsNull(rs("FATYou").value), "", rs("FATYou").value)
    TxtFATValue.text = IIf(IsNull(rs("FATValue").value), "", rs("FATValue").value)
    TxtTotalValue.text = IIf(IsNull(rs("TotalValue").value), "", rs("TotalValue").value)
    Me.AccountVat.BoundText = IIf(IsNull(rs("AccountCodeVat").value), "", rs("AccountCodeVat").value)
        TxtAuthoOrder.text = IIf(IsNull(rs("AuthoOrder").value), "", (rs("AuthoOrder").value))
    If val(TxtAuthoOrder) <> 0 Then
        DisplayCashInvoice
    End If
   If IsNull(rs("BoxID").value) Then
        Me.DcboBox.BoundText = ""
    Else
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
    End If
         
     Dim mmm As String
    
    If Not (IsNull(rs("QrCodeImage").value)) Then
        LoadPictureFromDB Picture1, rs, "QrCodeImage", mmm
    Else
     Set Picture1.Picture = Nothing
    End If


    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    Me.TxtEndDate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    DcbCarType.BoundText = IIf(IsNull(rs("CarTypeID").value), "", rs("CarTypeID").value)
    DcbCarModel.BoundText = IIf(IsNull(rs("CarModelID").value), "", rs("CarModelID").value)
    DcbColor.BoundText = IIf(IsNull(rs("ColorID").value), "", rs("ColorID").value)
    DcbyearFactor.text = IIf(IsNull(rs("YearFact").value), "", rs("YearFact").value)
    TxtClientPhone.text = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
    TxtCliientName.text = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
    TxtPlatNo.text = IIf(IsNull(rs("PlateNo").value), "", rs("PlateNo").value)
    DcbOrderStatus.ListIndex = val(IIf(IsNull(rs("OrderStatus").value), 0, rs("OrderStatus").value))
    TXtCarMeter.text = IIf(IsNull(rs("CarMeter").value), "", rs("CarMeter").value)
    TxtSparePart.text = IIf(IsNull(rs("SparePart").value), "", rs("SparePart").value)
 
   TxtFirstPrice.text = val(IIf(IsNull(rs("PayFirst").value), 0, rs("PayFirst").value))
   Me.TxtAmoutAccept.text = val(IIf(IsNull(rs("AmountAccept").value), 0, rs("AmountAccept").value))
     Me.ComGranty.ListIndex = IIf(IsNull(rs("Granty").value), 1, rs("Granty").value)
   DateStartG.value = IIf(IsNull(rs("DateStartG").value), Date, rs("DateStartG").value)
   DateEndg.value = IIf(IsNull(rs("DateEndG").value), Date, rs("DateEndG").value)
 '  Me.TxtComplaint.text = IIf(IsNull(rs("Complaint").value), "", rs("Complaint").value)
 '  Me.TxtNoteIntial.text = IIf(IsNull(rs("Noteinitial").value), "", rs("Noteinitial").value)
 txtKM.text = IIf(IsNull(rs("OverKM").value), "", rs("OverKM").value)
    TxtLongGranty.text = IIf(IsNull(rs("LongGranty").value), "", rs("LongGranty").value)
   ''
   
 If rs("Month_Day").value = True Then
   Me.ComMD.ListIndex = 0
   Else
   Me.ComMD.ListIndex = 1
   End If
 '  If rs("Granty").value = True Then
 '  Me.ComGranty.ListIndex = 0
 '  Me.frmgranty.Visible = True
 '  Else
 ''  Me.ComGranty.ListIndex = 1
  ' Me.frmgranty.Visible = False
  ' End If
   'If rs("Month_Day").value = True Then
   'Me.ComMD.ListIndex = 0
   'Else
   'Me.ComMD.ListIndex = 1
   'End If
    If rs("Accept").value = True Then
     Me.ChAccept.value = vbChecked
     Me.DcbOrderStatus.ListIndex = 1
     Else
      Me.ChAccept.value = vbUnchecked
      End If
      'If rs("subcar1").value = True Then
      '    Me.imag1.Picture = Me.img.Picture
'Else '
' Me.imag1.Picture = Me.imgnul.Picture
''
 '          End If
 '           If rs("subcar2").value = True Then
 '          Me.imag2.Picture = Me.img.Picture
'Else
' Me.imag2.Picture = Me.imgnul.Picture
'           End If
'            If rs("subcar3").value = True Then
'        Me.imag3.Picture = Me.img.Picture
'Else
'' Me.imag3.Picture = Me.imgnul.Picture
'           End If
'            If rs("subcar4").value = True Then
'
'Me.imag4.Picture = Me.img.Picture
'Else
' Me.imag4.Picture = Me.imgnul.Picture
'           End If
'            If rs("subcar5").value = True Then
'          Me.imag5.Picture = Me.img.Picture
'Else
' Me.imag5.Picture = Me.imgnul.Picture
'           End If
'            If rs("subcar6").value = True Then
'       Me.img6.Picture = Me.img.Picture
'Else
' Me.img6.Picture = Me.imgnul.Picture
'           End If
'            If rs("subcar7").value = True Then
''           Me.img7.Picture = Me.img.Picture
'Else
' Me.img7.Picture = Me.imgnul.Picture
''           End If
 '           If rs("subcar8").value = True Then
 '         Me.img8.Picture = Me.img.Picture
''Else
 'Me.img8.Picture = Me.imgnul.Picture
 '          End If
 '           If rs("subcar9").value = True Then
 '         Me.img9.Picture = Me.img.Picture
''Else
 'Me.img9.Picture = Me.imgnul.Picture
 '          End If
 '           If rs("subcar10").value = True Then
 '          Me.img10.Picture = Me.img.Picture
'Else
' Me.img10.Picture = Me.imgnul.Picture
'           End If
   ' TxtAdvanceValue.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
  '  Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
   ' Me.TxtPaymentCounts.text = IIf(IsNull(rs("PaymentCounts").value), "", rs("PaymentCounts").value)
 
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
     ' If IsNull(rs("posted").value) Then
       '                                            If SystemOptions.UserInterface = ArabicInterface Then
        '                                            Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
          '                                        Else
                                                 '   Accredit.Caption = " send to Approval   "
               ''                                End If
                                     '          Accredit.Enabled = True
'  Else
                                      '             If SystemOptions.UserInterface = ArabicInterface Then
                                        '            Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
                                        '          Else
                                                  '  Accredit.Caption = " sent to Approval   "
                                             '  End If
                                             '  Accredit.Enabled = False
  ' End If
     If SystemOptions.LinkCustomerWithCars = True Then
       Dim Dcombos As ClsDataCombos
       Set Dcombos = New ClsDataCombos
       Dcombos.GetCarsOfCustomer DcbCar, val(DcbCustmer.BoundText)
       End If
    Me.DcbCar.BoundText = IIf(IsNull(rs("CarID").value), "", rs("CarID").value)
    Set RsDetails = New ADODB.Recordset
StrSQL = " SELECT      dbo.TblCarBillMentainsDetils.ID, dbo.TblCarBillMentainsDetils.Type, dbo.TblCarBillMentainsDetils.ValiueDis, "
StrSQL = StrSQL & "                      dbo.TblCarBillMentainsDetils.[Value], dbo.TblMaintenanceWork.name, dbo.TblMaintenanceWork.namee, dbo.TblCarBillMentainsDetils.Mainte,"
StrSQL = StrSQL & "                      dbo.TblCarBillMentainsDetils.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
StrSQL = StrSQL & "                      dbo.TblCarBillMentainsDetils.[count], dbo.TblCarBillMentainsDetils.AccountCode, dbo.TblCarBillMentainsDetils.Deptid, dbo.TblEmpDepartments.DepartmentName,"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments.DepartmentNamee ,dbo.TblCarBillMentainsDetils.TotalNet,dbo.TblCarBillMentainsDetils.Percentage,dbo.TblCarBillMentainsDetils.DiscValue"
StrSQL = StrSQL & " FROM         dbo.TblCarBillMentainsDetils INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblMaintenanceWork ON dbo.TblCarBillMentainsDetils.Mainte = dbo.TblMaintenanceWork.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.TblCarBillMentainsDetils.Deptid = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblCarBillMentainsDetils.Emp_ID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & "  Where (dbo.TblCarBillMentainsDetils.id =" & val(XPTxtID.text) & ") And (dbo.TblCarBillMentainsDetils.Type = 0)"
StrSQL = StrSQL & "   ORDER BY dbo.TblCarBillMentainsDetils.Mainte"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RsDetails.BOF Or RsDetails.EOF) Then
       With Me.Fg
      '  RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails.RecordCount
        For i = .FixedRows To .Rows - 1
    
            .TextMatrix(i, .ColIndex("serial")) = i
            .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDetails("AccountCode").value), "", RsDetails("AccountCode").value)
            .TextMatrix(i, .ColIndex("Deptid")) = IIf(IsNull(RsDetails("Deptid").value), 0, RsDetails("Deptid").value)
            .TextMatrix(i, .ColIndex("valuedis")) = IIf(IsNull(RsDetails("ValiueDis").value), "", RsDetails("ValiueDis").value)
            .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDetails("Value").value), "", RsDetails("Value").value)
             .TextMatrix(i, .ColIndex("cod")) = IIf(IsNull(RsDetails("Mainte").value), "", RsDetails("Mainte").value)
            .TextMatrix(i, .ColIndex("count")) = IIf(IsNull(RsDetails("count").value), "", RsDetails("count").value)
            .TextMatrix(i, .ColIndex("Fullcode")) = (IIf(IsNull(RsDetails("Fullcode").value), "", RsDetails("Fullcode").value))
            .TextMatrix(i, .ColIndex("Emp_ID")) = (IIf(IsNull(RsDetails("Emp_ID").value), 0, RsDetails("Emp_ID").value))
     If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Emp_Name")) = (IIf(IsNull(RsDetails("Emp_Name").value), "", RsDetails("Emp_Name").value))
            .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("name").value), "", RsDetails("name").value)
            .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(RsDetails("DepartmentName").value), "", RsDetails("DepartmentName").value)
      Else
            .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("namee").value), "", RsDetails("namee").value)
            .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(RsDetails("DepartmentNamee").value), "", RsDetails("DepartmentNamee").value)
            .TextMatrix(i, .ColIndex("Emp_Name")) = (IIf(IsNull(RsDetails("Emp_Namee").value), "", RsDetails("Emp_Namee").value))
     End If
     .TextMatrix(i, .ColIndex("TotalNet")) = IIf(IsNull(RsDetails("TotalNet").value), 0, RsDetails("TotalNet").value)
     .TextMatrix(i, .ColIndex("Percentage")) = IIf(IsNull(RsDetails("Percentage").value), 0, RsDetails("Percentage").value)
     .TextMatrix(i, .ColIndex("DiscValue")) = IIf(IsNull(RsDetails("DiscValue").value), 0, RsDetails("DiscValue").value)
     
            RsDetails.MoveNext
         
        Next i
End With
    End If

    RsDetails.Close
    Set RsDetails = Nothing
    '//////////////////////////////////////////
    Set RsDetails1 = New ADODB.Recordset
 'StrSQL = " SELECT     TOP 100 PERCENT dbo.TblCarBillMentainsDetils.ID,dbo.TblCarBillMentainsDetils.comp,dbo.TblCarBillMentainsDetils.bill,dbo.TblCarBillMentainsDetils.count, dbo.TblCarBillMentainsDetils.Type, dbo.TblCarBillMentainsDetils.[Value],"
 '           StrSQL = StrSQL & "          dbo.TblExtraExpeneses.name , dbo.TblExtraExpeneses.namee, dbo.TblCarBillMentainsDetils.Mainte"
 '         StrSQL = StrSQL & "   FROM         dbo.TblCarBillMentainsDetils INNER JOIN"
 '      StrSQL = StrSQL & "               dbo.TblExtraExpeneses ON dbo.TblCarBillMentainsDetils.Mainte = dbo.TblExtraExpeneses.Id"
 'StrSQL = StrSQL & "  Where (dbo.TblCarBillMentainsDetils.id =" & val(XPTxtID.text) & ") And (dbo.TblCarBillMentainsDetils.Type = 1)"
'StrSQL = StrSQL & "   ORDER BY dbo.TblCarBillMentainsDetils.Mainte"
StrSQL = " SELECT      dbo.TblCarBillMentainsDetils.ID, dbo.TblCarBillMentainsDetils.Type, dbo.TblCarBillMentainsDetils.Mainte, dbo.TblCarBillMentainsDetils.[Value],"
StrSQL = StrSQL & "                      dbo.TblCarBillMentainsDetils.[count], dbo.TblCarBillMentainsDetils.comp, dbo.TblCarBillMentainsDetils.bill, dbo.TblExtraExpeneses.name,"
StrSQL = StrSQL & "                      dbo.TblExtraExpeneses.namee"
StrSQL = StrSQL & " FROM         dbo.TblCarBillMentainsDetils INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblExtraExpeneses ON dbo.TblCarBillMentainsDetils.Type = dbo.TblExtraExpeneses.Id"
StrSQL = StrSQL & " Where (dbo.TblCarBillMentainsDetils.id = " & val(XPTxtID.text) & ") And (dbo.TblCarBillMentainsDetils.Type = 1)"
    RsDetails1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
     If Not (RsDetails1.BOF Or RsDetails1.EOF) Then
       With Me.fg2
      '  RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails1.RecordCount

        For i = .FixedRows To .Rows - 1
    
            .TextMatrix(i, .ColIndex("serial")) = i
             
            .TextMatrix(i, .ColIndex("value")) = RsDetails1("Value").value
             .TextMatrix(i, .ColIndex("cod")) = RsDetails1("Mainte").value
            .TextMatrix(i, .ColIndex("count")) = RsDetails1("count").value
           .TextMatrix(i, .ColIndex("comp")) = RsDetails1("comp").value
           .TextMatrix(i, .ColIndex("bill")) = RsDetails1("bill").value
               If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails1("name").value), "", RsDetails1("name").value)
                Else
                   .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails1("namee").value), "", RsDetails1("namee").value)
             End If
            RsDetails1.MoveNext
         
        Next i
End With
    End If

    RsDetails1.Close
    Set RsDetails1 = Nothing
    

    newret
    fillapprovData
    ReLineGrid
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Public Sub retrive1(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
   Dim Rs1 As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
      Set Rs1 = New ADODB.Recordset
    StrSQL = "select * From TblCardAuthorizationReform  WHERE id=" & Lngid & "   Order By ID"
    Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.ChAccept = xtpUnchecked
            ' clear_all Me
            
Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 2
            fg2.Enabled = True
        imgg
            Me.Lbtotal.Caption = 0
            Me.LbToTalExtra.Caption = 0
            
            Me.lbTotalMente.Caption = 0
     Me.DcbOrderStatus.ListIndex = 0
     
 

    'On Error GoTo ErrTrap
     
     
    If Rs1.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
 imgg
    If Rs1.EOF Or Rs1.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            Rs1.find "id=" & Lngid, , adSearchForward, adBookmarkFirst

            If Rs1.EOF Or Rs1.BOF Then
                Exit Sub
            End If
        End If
    End If
    If Lngid <> 0 Then
        
    
  '  If val(DcbBasedOn.ListIndex) = 1 Then
     
    If Not SystemOptions.IsMaintItemMode Then
     If Rs1("OrderStatus").value = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "·«Ì„ﬂ‰  ‰›Ì– «·«„— «·« »⁄œ  «·«‰ Â«¡"
     Else
     MsgBox "Must Finish All maintenance Firstly"
     End If
     End If
              clear_all Me
        
            
Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 2
            fg2.Enabled = True
        imgg
            Me.Lbtotal.Caption = 0
            Me.LbToTalExtra.Caption = 0
            
            Me.lbTotalMente.Caption = 0
     Me.DcbOrderStatus.ListIndex = 0
     Exit Sub
    
     End If
     If Rs1("OrderStatus").value = 1 Then
     MsgBox "·«Ì„ﬂ‰  ‰›Ì– «·«„— «·« »⁄œ  «·«‰ Â«¡"
              clear_all Me
              Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 2
            fg2.Enabled = True
        imgg
            Me.Lbtotal.Caption = 0
            Me.LbToTalExtra.Caption = 0
            
            Me.lbTotalMente.Caption = 0
     Me.DcbOrderStatus.ListIndex = 0
     Exit Sub
    
     End If
     
               If Rs1("OrderStatus").value = 4 Then
     MsgBox "·«Ì„ﬂ‰  ‰›Ì– «·«„—   ·⁄œ„ „Ê«›ﬁ… «·⁄„Ì·"
              clear_all Me
              Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 2
            fg2.Enabled = True
        imgg
            Me.Lbtotal.Caption = 0
            Me.LbToTalExtra.Caption = 0
            
            Me.lbTotalMente.Caption = 0
     Me.DcbOrderStatus.ListIndex = 0
     Exit Sub
    
     End If
      If Rs1("OrderStatus").value = 30 Then
     MsgBox "·«Ì„ﬂ‰  ‰›Ì– «·«„—`  ﬁœ  „ «’œ«— ·Â ›« Ê—…    "
              clear_all Me
              Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 2
            fg2.Enabled = True
        imgg
            Me.Lbtotal.Caption = 0
            Me.LbToTalExtra.Caption = 0
            
            Me.lbTotalMente.Caption = 0
     Me.DcbOrderStatus.ListIndex = 0
     Exit Sub
    
     End If
                         If Rs1("OrderStatus").value = 50 Then
     MsgBox " „ «·«‰ Â«¡ „‰ «·«„—    "
              clear_all Me
              Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 2
            fg2.Enabled = True
        imgg
            Me.Lbtotal.Caption = 0
            Me.LbToTalExtra.Caption = 0
            
            Me.lbTotalMente.Caption = 0
     Me.DcbOrderStatus.ListIndex = 0
     Exit Sub
    
     End If

     End If
 
       DcbOrderStatus.ListIndex = 5
  '     " IIf(IsNull(Rs1("OrderStatus").value), "", Rs1("OrderStatus").value)
     ' DcbOrderStatus.text = IIf(IsNull(Rs1("OrderStatus").value), "", Rs1("OrderStatus").value)
 ' Me.TxtCusID.text = IIf(IsNull(Rs1("ClientCode").value), "", val(Rs1("ClientCode").value))
            
           Me.TxtCusID.text = IIf(IsNull(Rs1("CusID").value), 0, (Rs1("CusID").value))
 

Me.DcbCustmer.BoundText = IIf(IsNull(Rs1("CusID").value), 0, (Rs1("CusID").value))

  'Me.TxtCusID.text = IIf(IsNull(Rs1("ClientCode").value), "", (Rs1("ClientCode").value))
     
     
    TxtReq.text = IIf(IsNull(Rs1("AuthoOrder").value), "", val(Rs1("AuthoOrder").value))
    TxtAuthoOrder.text = IIf(IsNull(Rs1("WorkOrder").value), "", val(Rs1("WorkOrder").value))
    
    If val(TxtAuthoOrder) <> 0 Then
        DisplayCashInvoice
    End If
    XPDtbTrans.value = IIf(IsNull(Rs1("RecordDate").value), Date, Rs1("RecordDate").value)
    Me.TxtEndDate.value = IIf(IsNull(Rs1("EndDate").value), Date, Rs1("EndDate").value)
    dcBranch.BoundText = IIf(IsNull(Rs1("BranchID").value), "", Rs1("BranchID").value)
    
    DcbCarType.BoundText = IIf(IsNull(Rs1("CarTypeID").value), "", Rs1("CarTypeID").value)
    DcbCarModel.BoundText = IIf(IsNull(Rs1("CarModelID").value), "", Rs1("CarModelID").value)
   ' DcboSpecifications.BoundText = IIf(IsNull(rs1("gradeID").value), "", rs1("gradeID").value)
    DcbColor.BoundText = IIf(IsNull(Rs1("ColorID").value), "", Rs1("ColorID").value)
    DcbyearFactor.text = IIf(IsNull(Rs1("YearFact").value), "", Rs1("YearFact").value)
   TxtClientPhone.text = IIf(IsNull(Rs1("Telephone").value), "", Rs1("Telephone").value)
   TxtCliientName.text = IIf(IsNull(Rs1("ClientName").value), "", Rs1("ClientName").value)
   TxtPlatNo.text = IIf(IsNull(Rs1("PlateNo").value), "", Rs1("PlateNo").value)
   TxtSparePart.text = IIf(IsNull(Rs1("SparePart").value), "", Rs1("SparePart").value)
     
     Dim mmm As String
    
    If Not (IsNull(Rs1("QrCodeImage").value)) Then
        LoadPictureFromDB Picture1, Rs1, "QrCodeImage", mmm
    Else
     Set Picture1.Picture = Nothing
    End If


   TXtCarMeter.text = IIf(IsNull(Rs1("CarMeter").value), "", Rs1("CarMeter").value)
 '  TxtLongGranty.text = IIf(IsNull(rs1("LongGranty").value), "", rs1("LongGranty").value)
   TxtFirstPrice.text = val(IIf(IsNull(Rs1("PayFirst").value), 0, Rs1("PayFirst").value))
   Me.TxtAmoutAccept.text = val(IIf(IsNull(Rs1("AmountAccept").value), 0, Rs1("AmountAccept").value))
        Me.ComGranty.ListIndex = IIf(IsNull(Rs1("Granty").value), 1, Rs1("Granty").value)
   DateStartG.value = IIf(IsNull(Rs1("DateStartG").value), Date, Rs1("DateStartG").value)
   DateEndg.value = IIf(IsNull(Rs1("DateEndG").value), Date, Rs1("DateEndG").value)
 '  Me.TxtComplaint.text = IIf(IsNull(rs("Complaint").value), "", rs("Complaint").value)
 '  Me.TxtNoteIntial.text = IIf(IsNull(rs("Noteinitial").value), "", rs("Noteinitial").value)
 txtKM.text = IIf(IsNull(Rs1("OverKM").value), "", Rs1("OverKM").value)
    TxtLongGranty.text = IIf(IsNull(Rs1("LongGranty").value), "", Rs1("LongGranty").value)
   ''
   
 If Rs1("Month_Day").value = True Then
   Me.ComMD.ListIndex = 0
   Else
   Me.ComMD.ListIndex = 1
   End If
     If SystemOptions.LinkCustomerWithCars = True Then
       Dim Dcombos As ClsDataCombos
       Set Dcombos = New ClsDataCombos
       Dcombos.GetCarsOfCustomer DcbCar, val(DcbCustmer.BoundText)
       End If
    Me.DcbCar.BoundText = IIf(IsNull(Rs1("CarID").value), "", Rs1("CarID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(Rs1("UserID").value), "", Rs1("UserID").value)

       Set RsDetails = New ADODB.Recordset
 StrSQL = " SELECT     dbo.TblCardAuthorizationReformDetails.Type, dbo.TblCardAuthorizationReformDetails.PriceFitter, dbo.TblCardAuthorizationReformDetails.finish, "
 StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.[Value], dbo.TblCardAuthorizationReform.ID, dbo.TblCardAuthorizationReformDetails.ID2, dbo.TblEmployee.Emp_ID,"
 StrSQL = StrSQL & "                     dbo.TblEmployee.PerceTage, dbo.TblCardAuthorizationReformDetails.ID AS IDDet, dbo.TblCardAuthorizationReformDetails.[count], dbo.TblEmployee.Emp_Name,"
 StrSQL = StrSQL & "                     dbo.TblEmployee.fullcode , dbo.TblEmployee.Emp_Namee,TblEmpDepartments.AccountCode"
 StrSQL = StrSQL & " FROM         dbo.TblCardAuthorizationReform RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblCardAuthorizationReformDetails LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.TblCardAuthorizationReformDetails.EmpID = dbo.TblEmployee.Emp_ID ON"
 StrSQL = StrSQL & "                     dbo.TblCardAuthorizationReform.ID = dbo.TblCardAuthorizationReformDetails.ID"
  StrSQL = StrSQL & "                     LEFT OUTER JOIN TblEmpDepartments ON TblEmpDepartments.DeparmentID = dbo.TblCardAuthorizationReformDetails.Deptid"
StrSQL = StrSQL & " Where (dbo.TblCardAuthorizationReformDetails.Type = 0) And (dbo.TblCardAuthorizationReformDetails.finish = 1)"
  StrSQL = StrSQL & "                    AND (dbo.TblCardAuthorizationReformDetails.ID = " & val(Lngid) & " )"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
   

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
       With Me.Fg
       Dim pre As Integer
       Dim d As Integer
       Dim precount As Integer
       Dim dfi As Integer
      '  RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .Rows - 1
    
    pre = val(RsDetails("Value").value) * val(RsDetails("count").value)
           
            .TextMatrix(i, .ColIndex("total")) = pre
            .TextMatrix(i, .ColIndex("valuedis")) = pre
             dfi = pre - val(RsDetails("PriceFitter").value)
             d = (val(IIf(IsNull(RsDetails("PerceTage").value), 0, RsDetails("PerceTage").value) / 100) * (dfi))
          ' d = dfi * (0.5)
      '    If d = 0 Then
      '    d = val(IIf(IsNull(RsDetails("PriceFitter").value), 0, RsDetails("PriceFitter").value))
      '    End If
         If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Emp_Name")) = (IIf(IsNull(RsDetails("Emp_Name").value), "", RsDetails("Emp_Name").value))
         Else
            .TextMatrix(i, .ColIndex("Emp_Name")) = (IIf(IsNull(RsDetails("Emp_Namee").value), "", RsDetails("Emp_Namee").value))
         End If
            .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDetails("AccountCode").value), "", RsDetails("AccountCode").value)
            .TextMatrix(i, .ColIndex("Fullcode")) = (IIf(IsNull(RsDetails("Fullcode").value), "", RsDetails("Fullcode").value))
            .TextMatrix(i, .ColIndex("Emp_ID")) = (IIf(IsNull(RsDetails("Emp_ID").value), 0, RsDetails("Emp_ID").value))
            .TextMatrix(i, .ColIndex("fittervalue ")) = (IIf(IsNull(RsDetails("PriceFitter").value), 0, RsDetails("PriceFitter").value))
            .TextMatrix(i, .ColIndex("rate")) = IIf(IsNull(RsDetails("PerceTage").value), "", RsDetails("PerceTage").value)
            .TextMatrix(i, .ColIndex("valueprce")) = d
             .TextMatrix(i, .ColIndex("netprce")) = d + val(RsDetails("PriceFitter").value)
            .TextMatrix(i, .ColIndex("nett")) = pre - d + val(RsDetails("PriceFitter").value)
            precount = (dfi) - d + val(RsDetails("PriceFitter").value)
            d = precount / val(RsDetails("count").value)
             .TextMatrix(i, .ColIndex("valuedis")) = d
            RsDetails.MoveNext
         
        Next i
End With
    End If

    RsDetails.Close
    Set RsDetails = Nothing
   '''''''''''llllllllllllllllllllllllllllll
   
    Set RsDetails = New ADODB.Recordset
StrSQL = "SELECT     TOP 100 PERCENT dbo.TblCardAuthorizationReformDetails.ID, dbo.TblCardAuthorizationReformDetails.Type, dbo.TblCardAuthorizationReformDetails.[Value], "
StrSQL = StrSQL & "                      dbo.TblMaintenanceWork.name, dbo.TblMaintenanceWork.namee, dbo.TblCardAuthorizationReformDetails.Mainte, dbo.TblEmpDepartments.AccountCode,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.Deptid, dbo.TblCardAuthorizationReformDetails.[count], dbo.TblEmpDepartments.DepartmentName,"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments.DepartmentNamee"
StrSQL = StrSQL & " FROM         dbo.TblCardAuthorizationReformDetails LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblMaintenanceWork ON dbo.TblCardAuthorizationReformDetails.Mainte = dbo.TblMaintenanceWork.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.TblCardAuthorizationReformDetails.Deptid = dbo.TblEmpDepartments.DeparmentID"
StrSQL = StrSQL & "  Where (dbo.TblCardAuthorizationReformDetails.id =" & val(Lngid) & ") And (dbo.TblCardAuthorizationReformDetails.Type = 0)"
StrSQL = StrSQL & "   ORDER BY dbo.TblCardAuthorizationReformDetails.Mainte"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   lbTotalMenteDis.Caption = 0
   

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
       With Me.Fg
      '  RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .Rows - 1
    
    
            .TextMatrix(i, .ColIndex("serial")) = i
            .TextMatrix(i, .ColIndex("valuedis")) = IIf(IsNull(RsDetails("Value").value), 0, RsDetails("Value").value)
            .TextMatrix(i, .ColIndex("cod")) = IIf(IsNull(RsDetails("Mainte").value), 0, RsDetails("Mainte").value)
            .TextMatrix(i, .ColIndex("count")) = IIf(IsNull(RsDetails("count").value), 0, RsDetails("count").value)
            .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDetails("AccountCode").value), "", RsDetails("AccountCode").value)
            .TextMatrix(i, .ColIndex("Deptid")) = IIf(IsNull(RsDetails("Deptid").value), 0, RsDetails("Deptid").value)
             .TextMatrix(i, .ColIndex("nett")) = .TextMatrix(i, .ColIndex("valuedis")) * .TextMatrix(i, .ColIndex("count"))
             lbTotalMenteDis.Caption = val(lbTotalMenteDis.Caption) + val(Fg.TextMatrix(i, Fg.ColIndex("nett")))
             
               If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(RsDetails("DepartmentName").value), "", RsDetails("DepartmentName").value)
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("name").value), "", RsDetails("name").value)
                Else
                   .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("namee").value), "", RsDetails("namee").value)
                   .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(RsDetails("DepartmentNamee").value), "", RsDetails("DepartmentNamee").value)
             End If
             
             
            RsDetails.MoveNext
         
        Next i
End With
    End If

    RsDetails.Close
    Set RsDetails = Nothing
    
    '//////////////////////////////////////////
   Set RsDetails1 = New ADODB.Recordset
   StrSQL = " SELECT     TOP 100 PERCENT dbo.TblCardAuthorizationReformDetails.ID,dbo.TblCardAuthorizationReformDetails.comp,dbo.TblCardAuthorizationReformDetails.bill,dbo.TblCardAuthorizationReformDetails.count, dbo.TblCardAuthorizationReformDetails.Type, dbo.TblCardAuthorizationReformDetails.[Value],"
            StrSQL = StrSQL & "          dbo.TblExtraExpeneses.name , dbo.TblExtraExpeneses.namee, dbo.TblCardAuthorizationReformDetails.Mainte"
          StrSQL = StrSQL & "   FROM         dbo.TblCardAuthorizationReformDetails INNER JOIN"
       StrSQL = StrSQL & "               dbo.TblExtraExpeneses ON dbo.TblCardAuthorizationReformDetails.Mainte = dbo.TblExtraExpeneses.Id"
 StrSQL = StrSQL & "  Where (dbo.TblCardAuthorizationReformDetails.id =" & val(Lngid) & ") And (dbo.TblCardAuthorizationReformDetails.Type = 1)"
StrSQL = StrSQL & "   ORDER BY dbo.TblCardAuthorizationReformDetails.Mainte"
    RsDetails1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
     If Not (RsDetails1.BOF Or RsDetails1.EOF) Then
       With Me.fg2
      '  RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails1.RecordCount

        For i = .FixedRows To .Rows - 1
  If RsDetails1("name").value <> "" Then
            .TextMatrix(i, .ColIndex("serial")) = i
            .TextMatrix(i, .ColIndex("value")) = RsDetails1("Value").value
             .TextMatrix(i, .ColIndex("cod")) = RsDetails1("Mainte").value
            .TextMatrix(i, .ColIndex("count")) = RsDetails1("count").value
             
           .TextMatrix(i, .ColIndex("comp")) = RsDetails1("comp").value
           .TextMatrix(i, .ColIndex("bill")) = RsDetails1("bill").value
               If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails1("name").value), "", RsDetails1("name").value)
                Else
                   .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails1("namee").value), "", RsDetails1("namee").value)
             End If
             End If
            RsDetails1.MoveNext
         
        Next i
End With
    End If

    RsDetails1.Close
    Set RsDetails1 = Nothing
    newret
    fillapprovData
    ReLineGrid
    XPTxtCurrent.Caption = Rs1.AbsolutePosition
    XPTxtCount.Caption = Rs1.RecordCount
    FillCal
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "   ›« Ê—… ’Ì«‰… »—ﬁ„ " & TxtNoteSerial1.text & "   ··⁄„Ì· " & Me.DcbCustmer.text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
tablename = "TblCarBillMentains"
Filedname = "ID"
NoteSerial1 = val(XPTxtID.text)
Notevalue = 0
 notytype = 8074
 Notevalue = val(lbldifdis.Caption) + val(TxtPaymentValue.text)
 BranchID = val(dcBranch.BoundText)
NoteDate = (XPDtbTrans.value)
 
 
 
If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                     CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, val(XPTxtID.text), des
                                              TxtNoteID2.text = NoteID
                                                     TxtNoteSerial.text = NoteSerial
                                    Else
                                                 If TxtNoteID2.text = "" Or TxtNoteSerial.text = "" Then
                                         CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des ', recordDateH.value
                                                              TxtNoteID2.text = NoteID
                                                             TxtNoteSerial.text = NoteSerial
                                                 Else
                                                              sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                              sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                 sql = sql & " where NoteID=" & val(TxtNoteID2.text)
                                                                 Cn.Execute sql
                                        
                                                End If
                                       
                                End If

CREATE_VOUCHER_GE val(TxtNoteID2.text), BranchID, user_id, NoteDate
rs.Resync adAffectCurrent
  updateNotesValueAndNobytext (val(TxtNoteID2.text))

     End If

End Function

Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
Dim valuee As Double
 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim Account_Code_Expen As String
    Dim AccountDept As String
    Dim DebitAccount  As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim i As Integer
 Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
 LngDevNO = 0
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
         StrTempDes = "   ›« Ê—… ’Ì«‰… »—ﬁ„ " & TxtNoteSerial1.text & "   ··⁄„Ì·  " & Me.DcbCustmer.text

    '«·ÿ—› «·„Ì‰
    lbldifdis.Visible = True
        my_branch = BranchID
        
     If CboPayMentType.ListIndex = 1 Then '«Ã·
      valuee = val(LbtotalDis.Caption) + val(val(TxtFATValue.text))
      DebitAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbCustmer.BoundText), "Account_code")
        LngDevNO = LngDevNO + 1
        'valuee = val(LbtotalDis.Caption)
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, DebitAccount, valuee, 0, StrTempDes & "      Õ”«» «·‘—ﬂ… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
        End If
   End If
   
   If CboPayMentType.ListIndex = 0 Then '‰ﬁœÌ
        If val(TxtPaymentValue.text) > 0 Then
             valuee = val(TxtPaymentValue)
        Else
             valuee = val(TxtPaymentValue)
        End If
        DebitAccount = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText)) '‰ﬁœÌ
        LngDevNO = LngDevNO + 1
        valuee = val(LbtotalDis.Caption) + val(val(TxtFATValue.text))
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, DebitAccount, valuee, 0, StrTempDes & "      Õ”«» «·⁄Âœ… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
         End If
    
    
        
      End If
       If val(TxtFATValue.text) > 0 And Me.AccountVat.BoundText <> "" Then
        valuee = val(TxtFATValue.text)
        LngDevNO = LngDevNO + 1
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, Me.AccountVat.BoundText, valuee, 1, StrTempDes & "      Õ”«» «·ﬁÌ„… «·„÷«›… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
            GoTo ErrTrap
        End If
    End If
    Account_Code_Expen = get_account_code_branch(77, my_branch)
    LngDevNO = LngDevNO + 1
    valuee = val(LbtotalDis.Caption)
         If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code_Expen, valuee, 1, StrTempDes & "      Õ”«» «Ì—«œ«  «·’Ì«‰…", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
            GoTo ErrTrap
        End If
  
  Exit Function
  
    With Fg
    For i = 1 To .Rows - 1
          If .TextMatrix(i, .ColIndex("AccountCode")) <> "" And val(.TextMatrix(i, .ColIndex("TotalNet"))) <> 0 Then
              AccountDept = .TextMatrix(i, .ColIndex("AccountCode"))
              valuee = val(.TextMatrix(i, .ColIndex("TotalNet")))
              LngDevNO = LngDevNO + 1
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountDept, valuee, 1, StrTempDes & "    Õ”«» „»Ì⁄«   ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                  GoTo ErrTrap
                  
               End If
          End If
      Next i
    
        End With
    
          
          If val(lblTotalPay) <> 0 Then
            valuee = val(lblTotalPay.Caption)
            If SystemOptions.CustomerhavethreeAccounts = True Then
                DebitAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbCustmer.BoundText), "Account_code2") '‘Ìﬂ«   Õ  «· Õ’Ì·
                If DebitAccount = "" Then
                    DebitAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbCustmer.BoundText))
                End If
            Else
                Dim Account_Code_dynamic82 As String
                  
                Account_Code_dynamic82 = get_account_code_branch(158, my_branch)
                 If Account_Code_dynamic82 = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» œ›⁄«  „ﬁœ„… ··⁄„·«¡   ", vbCritical
                    Else
                        MsgBox "Please Select  Account", vbCritical
                    End If
            
                    GoTo ErrTrap
                End If
                DebitAccount = Account_Code_dynamic82
            End If
             
             
            LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, DebitAccount, valuee, 0, StrTempDes & "      Õ”«» «·œ›⁄… «·„ﬁœ„… ··⁄„Ì· ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
            End If
            
            ' valuee = val(LbtotalDis.Caption)
            DebitAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbCustmer.BoundText), "Account_code")
            LngDevNO = LngDevNO + 1
            
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, DebitAccount, valuee, 1, StrTempDes & "      Õ”«» «·⁄„Ì· ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            If val(lbllblTotalVat) <> 0 Then
                LngDevNO = LngDevNO + 1
                valuee = val(lbllblTotalVat.Caption)
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, DebitAccount, valuee, 1, StrTempDes & "      Õ”«» «·⁄„Ì· ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                LngDevNO = LngDevNO + 1
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, Me.AccountVat.BoundText, val(lbllblTotalVat), 0, StrTempDes & "      Õ”«» «·ﬁÌ„… «·„÷«›… „‰ «·œ›⁄… «·„ﬁœ„… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
            End If
            
          End If
       ' End If
  '  End If
       
    If val(TxtFATValue.text) > 0 And Me.AccountVat.BoundText <> "" Then
        valuee = val(TxtFATValue.text)
        LngDevNO = LngDevNO + 1
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, Me.AccountVat.BoundText, valuee, 1, StrTempDes & "      Õ”«» «·ﬁÌ„… «·„÷«›… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
            GoTo ErrTrap
        End If
    End If
  

 ''/////////////«·„’—Ê›« 
  If val(LbToTalExtra.Caption) + val(lbl(58).Caption) > 0 Then
   Account_Code_Expen = get_account_code_branch(77, my_branch)
   valuee = val(LbToTalExtra.Caption) + val(lbl(58).Caption)
             LngDevNO = LngDevNO + 1
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code_Expen, valuee, 1, StrTempDes & "      Õ”«» «·„’—Ê›«  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  End If


ErrTrap:
End Function


Private Sub SaveData()
Dim boolcreatevoucher    As Boolean
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim i As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String

    'On Error GoTo ErrTrap



    If Me.TxtModFlg.text <> "R" Then
        
If Not SystemOptions.IsMaintItemMode Then
        If Me.DcbCarType.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄  «·„⁄œÂ/«·”Ì«—…!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            'Me.DcbCarType.SetFocus
       '  SendKeys "{F4}"
            Exit Sub
       End If
End If
 ' If Me.TxtCliientName.text = "" Then
 '           Msg = "ÌÃ» «œŒ«· «”„ «·⁄„Ì·!! "
 '           MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 '           'Me.TxtCliientName.SetFocus
 '           SendKeys "{F4}"
 '           Exit Sub
 '       End If

       If CboPayMentType.ListIndex = 0 And val(DcboBox.BoundText) = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÃ»  ÕœÌœ «·Œ“Ì‰…"
                    Else
                        Msg = "Specify   Box"
                    End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcboBox.SetFocus
       ' SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If



Dim Account_Code_dynamic As String

            Account_Code_dynamic = get_account_code_branch(77, my_branch)

            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "branch Not Created", vbCritical
               End If

                GoTo ErrTrap
            Else

               If Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»    „»Ì⁄«  ’Ì«‰… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                    End If

                    GoTo ErrTrap
'
                End If
            End If
If val(CboPayMentType.ListIndex) = 1 Then
           If val(DcbCustmer.BoundText) = 0 Then
                     If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "ÌÊÃœ Œÿ√ ›Ì Õ”«» Â–« «·⁄„Ì·", vbCritical
                    Else
                        MsgBox "Customer  Account Have an Error", vbCritical
                    End If

                    GoTo ErrTrap
           End If
   End If
        Cn.BeginTrans
        BeginTrans = True
Dim sql As String
Dim sq2 As String
   sql = "update TblCardAuthorizationReform set   OrderStatus=5  where WorkOrder=" & val(Me.TxtReq.text) & ""
  
                                    Cn.Execute sql
    sq2 = "update TblCardAuthorizationReform set   Payed=0 where WorkOrder =" & val(Me.TxtReq.text) & ""
                                    Cn.Execute sq2
                                    
              
                                    
        If TxtModFlg.text = "N" Then

            XPTxtID.text = CStr(new_id("TblCarBillMentains", "ID", "", True))

        
            rs.AddNew
       ElseIf Me.TxtModFlg.text = "E" Then
           StrSQL = "Delete From TblCarBillMentainsDetils Where ID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
           StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
          StrSQL = " Delete From TblCarOrderVouchers2 where  ORderID =" & val(Me.TxtReq.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
           
           '    StrSQL = "delete From Notes where NoteID=" & val(Me.TXTNoteID.text) ' Val(rs("Transaction_ID").value)
       ' Cn.Execute StrSQL, , adExecuteNoRecords
      
      
'MsgBox Me.DcbOrderStatus.AddItem(, 1)
        End If
      '  Dim discountvalue As Double
      '   If XPCboDiscountType.ListIndex = 0 Then
      '           discountvalue = 0
      '   ElseIf XPCboDiscountType.ListIndex = 1 Then
      '          discountvalue = val(XPTxtDiscountVal.text)
      '   ElseIf XPCboDiscountType.ListIndex = 2 Then
      '          discountvalue = val(XPTxtDiscountVal.text) / 100 * val(lbldifdis.Caption)
'
'         End If
'         discountvalue = Round(discountvalue, 2)
            
            
        
    Set RsNotesGeneral = New ADODB.Recordset
'    RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       StrSQL = "SELECT     *  from dbo.Notes Where (1 = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


' RsNotesGeneral.AddNew
'    RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
'    general_noteid = RsNotesGeneral("NoteID").value
'    TxtNoteID.text = general_noteid
'
'    RsNotesGeneral("NoteDate").value = XPDtbTrans.value
'    RsNotesGeneral("NoteType").value = 5050
'    RsNotesGeneral("Note_Value").value = val(lbTotalMente.Caption)
    my_branch = val(Me.dcBranch.BoundText)
'
'    If TxtNoteSerial.text = "" Then
'        TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
'    End If
'
    If TxtNoteSerial1.text = "" Then
        TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 50, 5050)
    End If
        rs("NoteID").value = val(TXTNoteID.text)
        rs("difdisValue").value = IIf(Me.lbldifdis.Caption = "", 0, val(Me.lbldifdis.Caption))
        rs("PaymentValue").value = IIf(Me.TxtPaymentValue.text = "", Null, val(Me.TxtPaymentValue.text))
        rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        rs("BranchID").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
        rs("FATYou").value = IIf(Me.TxtFATYou.text = "", Null, val(Me.TxtFATYou.text))
        rs("FATValue").value = IIf(Me.TxtFATValue.text = "", Null, val(Me.TxtFATValue.text))
        rs("TotalValue").value = IIf(Me.TxtTotalValue.text = "", Null, val(Me.TxtTotalValue.text))
        rs("AccountCodeVat").value = (Me.AccountVat.BoundText)
        rs("ID").value = val(XPTxtID.text)
        rs("CusId").value = val(TxtCusID.text)
        rs("CarID").value = val(Me.DcbCar.BoundText)

     rs("WorkOrderNO").value = IIf(Trim(Me.TxtReq.text) = "", Null, Trim(Me.TxtReq.text))
     rs("AuthoOrder").value = IIf(Trim(Me.TxtAuthoOrder.text) = "", Null, Trim(Me.TxtAuthoOrder.text))
     
   sql = "update TblCardAuthorizationReform set   NoteSerial='" & Me.TxtNoteSerial1.text & "'   where id=" & val(Me.TxtReq.text) & ""
  
                                    Cn.Execute sql
     
        rs("RecordDate").value = XPDtbTrans.value
         rs("EndDate").value = Me.TxtEndDate.value
        rs("ClientName").value = Me.TxtCliientName.text
        rs("CusID").value = val(Me.DcbCustmer.BoundText)
        rs("NetDiscount").value = val(Me.TxtNetDiscount.text)
        
        rs("Telephone").value = Me.TxtClientPhone.text
        rs("CarTypeID").value = val(Me.DcbCarType.BoundText)
        rs("CarModelID").value = val(Me.DcbCarModel.BoundText)
        rs("PlateNo").value = Me.TxtPlatNo.text
       ' DcbOrderStatus.text = IIf(IsNull(rs("OrderStatus").value), "", rs("OrderStatus").value)
        rs("OrderStatus").value = 5
        rs("ColorID").value = val(Me.DcbColor.BoundText)
        rs("YearFact").value = val(Me.DcbyearFactor.text)
        rs("CarMetarOut").value = Me.TxtCarMetarOut.text
        rs("SparePart").value = Me.TxtSparePart.text
    
        rs("LongGranty").value = Me.TxtLongGranty.text
        rs("CarMeter").value = Me.TXtCarMeter.text
        rs("DateStartG").value = Me.DateStartG.value
        rs("DateEndG").value = Me.DateEndg.value
        rs("PayFirst").value = val(Me.TxtFirstPrice.text)
       ' rs("Noteinitial").value = Me.TxtNoteIntial.text
       ' rs("Complaint").value = Me.TxtComplaint.text
        rs("AmountAccept").value = val(Me.TxtAmoutAccept.text)
      rs("OverKM").value = Me.txtKM.text
       If Me.ComMD.ListIndex = 0 Then
         rs("Month_Day").value = 1
         Else
          rs("Month_Day").value = 0
         End If
        rs("Granty").value = val(Me.ComGranty.ListIndex)

        rs("UserID").value = val(Me.DCboUserName.BoundText)

 If CboPayMentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else
        rs("PaymentType").value = val(CboPayMentType.ListIndex)
    End If
      If XPCboDiscountType.ListIndex = -1 Then
        rs("Trans_DiscountType").value = 0
    Else
        rs("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If
    rs("BasedOn").value = val(DcbBasedOn.ListIndex)
    
        rs("Trans_Discount").value = val(XPTxtDiscountVal.text)
  
    

    If CboPayMentType.ListIndex = 0 Then
        rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
    Else
        rs("BoxID").value = Null
      
    End If
    
        rs.update
        '''''''''/////////////////////////////////
        
      Set RsDetails = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblCarBillMentainsDetils Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
       'RsDetails.Open "TblCarBillMentainsDetils", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
'If Fg.Rows > 2 Then
'                 Fg.Rows = Fg.Rows - 1
'            End If
       For i = Me.Fg.FixedRows To Fg.Rows - 1
       If Fg.TextMatrix(i, Fg.ColIndex("name")) <> "" Then
           RsDetails.AddNew
           RsDetails("ID").value = val(XPTxtID.text)
           RsDetails("AccountCode").value = (Fg.TextMatrix(i, Fg.ColIndex("AccountCode")))
           RsDetails("Deptid").value = val(Fg.TextMatrix(i, Fg.ColIndex("Deptid")))
           RsDetails("Emp_ID").value = val(Fg.TextMatrix(i, Fg.ColIndex("Emp_ID")))
           RsDetails("Value").value = val(Fg.TextMatrix(i, Fg.ColIndex("Value")))
           RsDetails("ValiueDis").value = val(Fg.TextMatrix(i, Fg.ColIndex("valuedis")))
           RsDetails("fittervalue").value = val(Fg.TextMatrix(i, Fg.ColIndex("fittervalue ")))
           RsDetails("Type").value = 0
           RsDetails("Mainte").value = val(Fg.TextMatrix(i, Fg.ColIndex("cod")))
           If val(Fg.TextMatrix(i, Fg.ColIndex("count"))) <> 0 Then
           RsDetails("count").value = val(Fg.TextMatrix(i, Fg.ColIndex("count")))
           Else
           RsDetails("count").value = 1
           End If
           RsDetails("DiscValue").value = val(Fg.TextMatrix(i, Fg.ColIndex("DiscValue")))
           RsDetails("Percentage").value = val(Fg.TextMatrix(i, Fg.ColIndex("Percentage")))
           RsDetails("TotalNet").value = val(Fg.TextMatrix(i, Fg.ColIndex("TotalNet")))
         RsDetails.update
         End If
        Next i
        
        '''''''''''''''//////////////////////////
        
      Set RsDetails1 = New ADODB.Recordset
       RsDetails1.Open "TblCarBillMentainsDetils", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

       For i = Me.fg2.FixedRows To fg2.Rows - 1
        If fg2.TextMatrix(i, fg2.ColIndex("name")) <> "" Then
           RsDetails1.AddNew
          RsDetails1("ID").value = val(XPTxtID.text)
          RsDetails1("ValiueDis").value = val(fg2.TextMatrix(i, fg2.ColIndex("Value")))
        RsDetails1("Value").value = val(fg2.TextMatrix(i, fg2.ColIndex("Value")))
        RsDetails1("comp").value = fg2.TextMatrix(i, fg2.ColIndex("comp"))
        RsDetails1("bill").value = fg2.TextMatrix(i, fg2.ColIndex("bill"))
            RsDetails1("Type").value = 1
           RsDetails1("Mainte").value = val(fg2.TextMatrix(i, fg2.ColIndex("cod")))
           If val(fg2.TextMatrix(i, fg2.ColIndex("count"))) <> 0 Then
           RsDetails1("count").value = val(fg2.TextMatrix(i, fg2.ColIndex("count")))
           Else
           RsDetails1("count").value = 1
           End If
           
         RsDetails1.update
         End If
        Next i
      
        '''''''''''''''//////////////////////////
                '''''''''/////////////////////////////////
    Set RsDetails1 = New ADODB.Recordset
    StrSQL = "SELECT     *  from dbo.TblCarOrderVouchers2 Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
With vchrgrid
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("ID"))) <> 0 Then
RsDetails1.AddNew
RsDetails1("ORderID").value = val(TxtReq.text)
RsDetails1("Transaction_IDDet").value = val(.TextMatrix(i, .ColIndex("ID")))
RsDetails1("Transaction_ID").value = val(.TextMatrix(i, .ColIndex("Transaction_ID")))
RsDetails1.update
If val(.TextMatrix(i, .ColIndex("OperPrice"))) <> 0 Then
StrSQL = " update  Transaction_Details  set OperPrice =" & val(.TextMatrix(i, .ColIndex("OperPrice"))) & " where id =" & val(.TextMatrix(i, .ColIndex("ID"))) & ""
Cn.Execute StrSQL
End If
End If
Next i
End With
        
'            If Me.TxtModFlg.text = "E" Then
 
'                StrSQL = "Delete notes where NoteID=" & val(Me.TxtNoteID.text)
'                Cn.Execute StrSQL, , adExecuteNoRecords

'            End If

'            RsNotes.AddNew
'            NoteID = CStr(TxtNoteID.text)
'            RsNotes("NoteID").value = CStr(TxtNoteID.text)
'            RsNotes("NoteType").value = 8032
'            RsNotes("NoteDate").value = XPDtbTrans.value
'            RsNotes("UserID").value = user_id
'            RsNotes("NoteSerial").value = Trim$(Me.TxtNoteSerial.text) '„”·”· «·ﬁÌœ
'            RsNotes("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text) '„”·”· «–‰ «·’—›
'            RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
'            RsNotes("numbering_type1").value = sand_numbering_type(32) ' ”ÃÌ· «·”·›'‰Ê⁄  —ﬁÌ„    
'            RsNotes("sanad_year").value = year(XPDtbTrans.value)
'            RsNotes("sanad_month").value = Month(XPDtbTrans.value)
'            RsNotes("note_value_by_characters").value = WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
            '     RsNotes("remark").value = txtRemarks.text & bankDes
'            RsNotes("Branch_no").value = val(Me.Dcbranch.BoundText)
                
'            RsNotes.update
                
'            line_no = 1
        
'            Msg = "”·› „ÊŸ›Ì‰ —ﬁ„ " & val(Me.XPTxtID.text)
'            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
'
'            Employee_account = get_EMPLOYEE_Account(val(Me.DcboEmpName.BoundText), "Account_Code")
'            StrAccountCode = Employee_account
'
            '        StrAccountCode = "a1a3a4" 'Õ”«» “„„ «·„ÊŸ›Ì‰
'            If ModAccounts.AddNewDev(LngDevID, 1, StrAccountCode, val(Me.TxtAdvanceValue.text), 0, Msg, NoteID, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , val(Me.XPTxtID.text), , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                GoTo ErrTrap
'            End If

'            StrAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))

'            If ModAccounts.AddNewDev(LngDevID, 2, StrAccountCode, val(Me.TxtAdvanceValue.text), 1, Msg, NoteID, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , val(Me.XPTxtID.text), , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                GoTo ErrTrap
'            End If
        
        End If
    
        Cn.CommitTrans
        BeginTrans = False
        
        createVoucher
        SaveQRCode "TblCarBillMentains", "ID", val(XPTxtID), TxtNoteSerial1.text, (XPDtbTrans.value), _
        (TxtTotalValue.text), Picture1, 0, (TxtFATValue.text), (TxtTotalValue.text)

    '    RsDetails.Close
         Set RsDetails = Nothing
         Set RsDetails1 = Nothing
         Dim DebitAccount As String
         Dim CreditAccount As String
         Dim des  As String
      

        
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.text

            Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
            
               Msg = " Saved Success " & CHR(13)
                Msg = Msg + "Do you want Enter New Transaction "
            End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                MsgBox "Update Success", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
        End Select
Retrive
        TxtModFlg.text = "R"
   ' End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
       
            rs.find "ID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
  Dim StrSQL1 As String
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else

    Msg = Msg + " Confirm Delete"
End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
             StrSQL1 = "Delete From TblCarBillMentainsDetils Where ID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL1, , adExecuteNoRecords
            Cn.Execute "update  TblCardAuthorizationReform  set OrderStatus=2  WHERE WorkOrder=" & val(TxtReq.text) & ""
            StrSQL1 = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
            Cn.Execute StrSQL1, , adExecuteNoRecords
        
            StrSQL = "delete From Notes where NoteID=" & val(Me.TXTNoteID.text) ' Val(rs("Transaction_ID").value)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblCarOrderVouchers2 where  ORderID =" & val(Me.TxtReq.text)
            Cn.Execute StrSQL1, , adExecuteNoRecords
      
      
                rs.delete
  
                rs.MoveFirst

      
                If rs.RecordCount < 1 Then
                 Me.ChAccept.value = xtpUnchecked
            Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 2
            fg2.Enabled = True
            vchrgrid.Clear flexClearScrollable, flexClearEverything
            vchrgrid.Rows = 2
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub



Function FillApprovedTable()
 Dim RSApproval  As New ADODB.Recordset
   Set RSApproval = New ADODB.Recordset
   Dim currentdate As Date
   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable


 Dim sql As String
  Dim Rs1 As New ADODB.Recordset
 Dim i As Integer
    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.Name & "')"
sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
            currentdate = Now
            For i = 1 To Rs1.RecordCount
              RSApproval.AddNew
                RSApproval("ScreenName").value = Me.Name
                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
               RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
                  RSApproval("Transaction_ID").value = val(Me.XPTxtID.text)
                   RSApproval("NoteSerial").value = val(Me.XPTxtID.text)
                RSApproval("Transaction_Date").value = Date
                
                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
               RSApproval("SendTime").value = currentdate

                 If i = 1 Then
                        RSApproval("Currcursor").value = 1
                         RSApproval("FromUser").value = user_name
                End If
                
                RSApproval.update
                Rs1.MoveNext
            Next i

    End If
    
    

End Function



Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        GRID2.Rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label11.Caption = "Approved"
                                 End If
                            Label11.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label11.Caption = "Currently required Approve"
                            End If
                 Label11.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 GRID2.Rows = 1
    End If
RsDetails.Close

End Function


Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.text = "R" Then
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
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip
With TTP

        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘… «· ”·Ì„ ··⁄„Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(3), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
          .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘… «·«Ê«„— «·„› ÊÕÂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(7), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…  «· ‰»ÌÂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(4), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
     With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…  «· ﬁ«—Ì—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(5), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
      With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…  ’—› ﬁÿ⁄ «·€Ì«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(2), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

       With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘… ÿ·» ›Õ’ ﬂ„»ÌÊ —  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(6), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
         With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    ÿ·» ’Ì«‰…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(0), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
        With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…  «·⁄„Ê·«  «·„” Õﬁ…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(9), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
       With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   „·› «·⁄„·«¡ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(10), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
        With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   ﬁ«—Ì— «·⁄„Ê·«  «·„” Õﬁ… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(11), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With



    Exit Sub
ErrTrap:
End Sub
Private Sub AddTip1()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "›« Ê—… ≈’·«Õ ”Ì«—… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ≈’·«Õ ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ≈’·«Õ ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ≈’·«Õ ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ≈’·«Õ ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ≈’·«Õ ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ≈’·«Õ ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ≈’·«Õ ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ≈’·«Õ ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ≈’·«Õ ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ≈’·«Õ ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

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

                ' btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub



Private Sub XPTxtDiscountVal_Change()
If Me.TxtModFlg.text <> "R" Then
    If val(XPCboDiscountType.ListIndex) = 2 Then
    TxtNetDiscount.text = (val(XPTxtDiscountVal.text) * val(LbtotalDis.Caption)) / 100
      If SystemOptions.UserInterface = ArabicInterface Then
    lbl(14).Caption = "‰”»…"
    Else
    lbl(14).Caption = "Percentage"
    End If
    TxtNetDiscount.text = Round(val(TxtNetDiscount.text), 2)
    ElseIf val(XPCboDiscountType.ListIndex) = 1 Then
    TxtNetDiscount.text = val(XPTxtDiscountVal.text)
        If SystemOptions.UserInterface = ArabicInterface Then
    lbl(14).Caption = "ﬁÌ„…"
    Else
    lbl(14).Caption = "Value"
    End If
    Else
    TxtNetDiscount.text = 0
    XPTxtDiscountVal.text = 0
    End If
lbldifdis.Caption = val(Me.LbtotalDis.Caption) - val(firstprice.Caption) - val(TxtNetDiscount.text) - val(TxtPaymentValue.text) + val(TxtFATValue.text)
FillCal
End If
End Sub
