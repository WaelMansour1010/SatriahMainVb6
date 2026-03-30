VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmCarAuthontication 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   19050
   Icon            =   "FrmCarAuthontication.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   19050
   Begin VB.CommandButton cmdOpenCard 
      Caption         =   "› Õ «·ﬂ«—  "
      Height          =   375
      Left            =   7560
      TabIndex        =   264
      Top             =   7140
      Width           =   2445
   End
   Begin VB.CommandButton cmdEndAll 
      Caption         =   "«ﬁ›«· «·ﬂ«—  ‰Â«∆Ï"
      Height          =   375
      Left            =   9960
      TabIndex        =   263
      Top             =   7140
      Width           =   2445
   End
   Begin VB.Frame gimage 
      BackColor       =   &H80000005&
      Height          =   6615
      Left            =   4440
      TabIndex        =   148
      Top             =   1680
      Width           =   9855
      Begin VB.CommandButton bClose 
         BackColor       =   &H000000FF&
         Caption         =   "X"
         Height          =   375
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   120
         Width           =   375
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   615
         Left            =   3600
         Top             =   4560
         Width           =   735
      End
      Begin VB.Shape Shape13 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   615
         Left            =   5160
         Top             =   4560
         Width           =   735
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   615
         Left            =   2160
         Top             =   4440
         Width           =   735
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   615
         Left            =   720
         Top             =   1680
         Width           =   735
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   615
         Left            =   2520
         Top             =   1800
         Width           =   735
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   615
         Left            =   4200
         Top             =   1680
         Width           =   735
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   615
         Left            =   600
         Top             =   4560
         Width           =   735
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   615
         Left            =   6960
         Top             =   4680
         Width           =   735
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   615
         Left            =   7920
         Top             =   4440
         Width           =   735
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   615
         Left            =   8760
         Top             =   4680
         Width           =   735
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   615
         Left            =   5640
         Top             =   1440
         Width           =   735
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   615
         Left            =   7080
         Top             =   1920
         Width           =   735
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   615
         Left            =   7920
         Top             =   1680
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   615
         Left            =   8760
         Top             =   1920
         Width           =   735
      End
      Begin VB.Image img9 
         Height          =   615
         Left            =   7920
         Picture         =   "FrmCarAuthontication.frx":038A
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   705
      End
      Begin VB.Image img10 
         Height          =   615
         Left            =   6960
         Picture         =   "FrmCarAuthontication.frx":0938
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   705
      End
      Begin VB.Image img8 
         Height          =   615
         Left            =   8760
         Picture         =   "FrmCarAuthontication.frx":0EE6
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   705
      End
      Begin VB.Image img13 
         Height          =   615
         Left            =   2160
         Picture         =   "FrmCarAuthontication.frx":1494
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   705
      End
      Begin VB.Image img11 
         Height          =   615
         Left            =   5160
         Picture         =   "FrmCarAuthontication.frx":1A42
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   705
      End
      Begin VB.Image img12 
         Height          =   615
         Left            =   3600
         Picture         =   "FrmCarAuthontication.frx":1FF0
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   705
      End
      Begin VB.Image img14 
         Height          =   615
         Left            =   600
         Picture         =   "FrmCarAuthontication.frx":259E
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   705
      End
      Begin VB.Image imag1 
         Height          =   615
         Left            =   8760
         Picture         =   "FrmCarAuthontication.frx":2B4C
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   705
      End
      Begin VB.Image imag2 
         Height          =   615
         Left            =   7920
         Picture         =   "FrmCarAuthontication.frx":30FA
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   705
      End
      Begin VB.Image imag3 
         Height          =   615
         Left            =   7080
         Picture         =   "FrmCarAuthontication.frx":36A8
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   705
      End
      Begin VB.Image imag4 
         Height          =   615
         Left            =   5640
         Picture         =   "FrmCarAuthontication.frx":3C56
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   705
      End
      Begin VB.Image imag5 
         Height          =   615
         Left            =   4200
         Picture         =   "FrmCarAuthontication.frx":4204
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   705
      End
      Begin VB.Image img6 
         Height          =   615
         Left            =   2520
         Picture         =   "FrmCarAuthontication.frx":47B2
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   705
      End
      Begin VB.Image img7 
         Height          =   615
         Left            =   720
         Picture         =   "FrmCarAuthontication.frx":4D60
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   705
      End
      Begin VB.Image Image6 
         Height          =   5775
         Left            =   240
         Picture         =   "FrmCarAuthontication.frx":530E
         Stretch         =   -1  'True
         Top             =   630
         Width           =   9735
      End
   End
   Begin VB.TextBox TxtAuthoOrder 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   14520
      Locked          =   -1  'True
      TabIndex        =   203
      TabStop         =   0   'False
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox TxtWorkOrder 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   201
      TabStop         =   0   'False
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox TxtShowPriceOrder 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   16680
      Locked          =   -1  'True
      TabIndex        =   198
      TabStop         =   0   'False
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox DcbScreen 
      Height          =   315
      Left            =   0
      TabIndex        =   187
      Top             =   735
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   120
      Top             =   8010
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   5520
      TabIndex        =   139
      Top             =   7620
      Width           =   13335
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarAuthontication.frx":22A5E
         Height          =   555
         Index           =   11
         Left            =   0
         Picture         =   "FrmCarAuthontication.frx":29D90
         Style           =   1  'Graphical
         TabIndex        =   186
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarAuthontication.frx":2A924
         Height          =   555
         Index           =   10
         Left            =   6360
         Picture         =   "FrmCarAuthontication.frx":2AF0B
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarAuthontication.frx":2B4F2
         Height          =   555
         Index           =   9
         Left            =   720
         Picture         =   "FrmCarAuthontication.frx":32824
         Style           =   1  'Graphical
         TabIndex        =   184
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   8
         Left            =   2040
         Picture         =   "FrmCarAuthontication.frx":32D44
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarAuthontication.frx":33229
         Height          =   555
         Index           =   7
         Left            =   4200
         Picture         =   "FrmCarAuthontication.frx":3A55B
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarAuthontication.frx":3ADEB
         Height          =   555
         Index           =   6
         Left            =   5640
         Picture         =   "FrmCarAuthontication.frx":4211D
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarAuthontication.frx":425BE
         Height          =   555
         Index           =   0
         Left            =   7080
         Picture         =   "FrmCarAuthontication.frx":498F0
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   1
         Left            =   3240
         Picture         =   "FrmCarAuthontication.frx":49E97
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   2
         Left            =   4920
         Picture         =   "FrmCarAuthontication.frx":4A338
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   3
         Left            =   3480
         Picture         =   "FrmCarAuthontication.frx":4A808
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   4
         Left            =   2760
         Picture         =   "FrmCarAuthontication.frx":4ACC1
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   5
         Left            =   1440
         Picture         =   "FrmCarAuthontication.frx":4B219
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   120
         Width           =   735
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   8880
         TabIndex        =   205
         Top             =   240
         Visible         =   0   'False
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»Ê«”ÿ…"
         Height          =   255
         Index           =   20
         Left            =   11760
         TabIndex        =   206
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.CommandButton Accredit 
      Caption         =   "Command1"
      Height          =   375
      Left            =   0
      TabIndex        =   96
      Top             =   7710
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   0
      TabIndex        =   93
      Top             =   11760
      Width           =   2055
   End
   Begin VB.ComboBox ComGranty 
      Height          =   315
      Left            =   3240
      TabIndex        =   31
      Top             =   735
      Width           =   1815
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   22200
      TabIndex        =   70
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   22200
      TabIndex        =   69
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   21480
      TabIndex        =   67
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   22200
      TabIndex        =   63
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   16680
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   735
      Visible         =   0   'False
      Width           =   1335
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   0
      Width           =   19065
      _cx             =   33629
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
      Caption         =   "»ÿ«ﬁ… «–‰ «’·«Õ  "
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
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1185
         TabIndex        =   43
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
         ButtonImage     =   "FrmCarAuthontication.frx":4B664
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
         TabIndex        =   44
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
         ButtonImage     =   "FrmCarAuthontication.frx":4B9FE
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
         TabIndex        =   45
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
         ButtonImage     =   "FrmCarAuthontication.frx":4BD98
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
         TabIndex        =   46
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
         ButtonImage     =   "FrmCarAuthontication.frx":4C132
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
         Left            =   11760
         Picture         =   "FrmCarAuthontication.frx":4C4CC
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
         Left            =   2160
         TabIndex        =   68
         Top             =   120
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   10020
      TabIndex        =   47
      Top             =   735
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   219086849
      CurrentDate     =   38784
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   510
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   8370
      Width           =   17625
      _cx             =   31089
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
         Index           =   1
         Left            =   15840
         TabIndex        =   49
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   15015
         TabIndex        =   50
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
         Left            =   14280
         TabIndex        =   51
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
         Left            =   13425
         TabIndex        =   52
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
         Left            =   0
         TabIndex        =   53
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
         Left            =   735
         TabIndex        =   54
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
         Left            =   12600
         TabIndex        =   62
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
         Left            =   1680
         TabIndex        =   72
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄…√„— ‘€·"
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
         Index           =   0
         Left            =   16680
         TabIndex        =   97
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
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   10
         Left            =   2760
         TabIndex        =   147
         Top             =   60
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄… »ÿ«ﬁ… ≈–‰ «’·«Õ"
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
         Index           =   11
         Left            =   4560
         TabIndex        =   155
         Top             =   60
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄… ⁄—÷ ”⁄—"
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
         Index           =   12
         Left            =   9840
         TabIndex        =   196
         Top             =   60
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   " ÕÊÌ· „‰ ⁄—÷ ”⁄— «·Ï «–‰ «’·«Õ"
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorButton     =   16761024
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
         Index           =   13
         Left            =   7080
         TabIndex        =   197
         Top             =   60
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   " ÕÊÌ· „‰ «–‰ «’·«Õ «·Ï «„— ‘€·"
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorButton     =   16761024
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
         Index           =   14
         Left            =   5760
         TabIndex        =   214
         Top             =   60
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "«‰Â«¡ «·ﬂ·"
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorButton     =   16761024
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   22440
      TabIndex        =   55
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
      Left            =   22200
      TabIndex        =   64
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
      Bindings        =   "FrmCarAuthontication.frx":50134
      Height          =   315
      Left            =   5880
      TabIndex        =   39
      Top             =   735
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   8640
      TabIndex        =   150
      Top             =   735
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   219086850
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   19320
      TabIndex        =   171
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   219086849
      CurrentDate     =   38784
   End
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   6525
      Left            =   120
      TabIndex        =   73
      Top             =   960
      Width           =   19050
      _cx             =   33602
      _cy             =   11509
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
      Caption         =   "»Ì«‰«  »ÿ«ﬁ… ≈–‰ «’·«Õ|«⁄„«· «· ’·ÌÕ|Õ«·Â «·«⁄ „«œ|”‰œ«  «·’—›|«·ﬁÿ⁄ «·„ﬁœ—…"
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
      Picture(0)      =   "FrmCarAuthontication.frx":50149
      Flags(2)        =   2
      Begin VB.CommandButton Command3 
         Caption         =   "Command1"
         Height          =   375
         Left            =   0
         TabIndex        =   262
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   6060
         Left            =   20595
         TabIndex        =   232
         TabStop         =   0   'False
         Top             =   45
         Width           =   18960
         _cx             =   33443
         _cy             =   10689
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
         Begin VB.TextBox txtVat2 
            Height          =   285
            Left            =   6330
            TabIndex        =   261
            Top             =   5520
            Width           =   795
         End
         Begin VB.TextBox txtTotalAfterDiscount 
            Height          =   285
            Left            =   9690
            TabIndex        =   259
            Top             =   5520
            Width           =   1335
         End
         Begin VB.TextBox txtVatyo 
            Height          =   285
            Left            =   7890
            TabIndex        =   257
            Top             =   5520
            Width           =   795
         End
         Begin VB.TextBox txtDiscPercent 
            Height          =   285
            Left            =   12180
            TabIndex        =   255
            Top             =   5520
            Width           =   1005
         End
         Begin VB.TextBox txtDiscValue 
            Height          =   285
            Left            =   14280
            TabIndex        =   254
            Top             =   5520
            Width           =   1095
         End
         Begin VB.TextBox TxtTotal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4440
            MaxLength       =   5
            TabIndex        =   246
            Top             =   900
            Width           =   930
         End
         Begin VB.TextBox TxtQty 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6480
            MaxLength       =   5
            TabIndex        =   244
            Top             =   900
            Width           =   930
         End
         Begin VB.TextBox TxtItemCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   15960
            TabIndex        =   234
            Top             =   900
            Width           =   1635
         End
         Begin VB.TextBox TxtItemPrice 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8040
            MaxLength       =   5
            TabIndex        =   233
            Top             =   900
            Width           =   930
         End
         Begin VSFlex8UCtl.VSFlexGrid FG22 
            Height          =   3480
            Left            =   150
            TabIndex        =   235
            Top             =   1260
            Width           =   18345
            _cx             =   32359
            _cy             =   6138
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   18
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmCarAuthontication.frx":504E3
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
            WallPaperAlignment=   0
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin MSDataListLib.DataCombo DcboItems 
            Height          =   315
            Left            =   10080
            TabIndex        =   236
            Top             =   900
            Width           =   5760
            _ExtentX        =   10160
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   300
            Index           =   15
            Left            =   3045
            TabIndex        =   237
            Top             =   870
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   529
            Caption         =   "≈÷«›…"
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
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   300
            Index           =   16
            Left            =   2130
            TabIndex        =   238
            Top             =   870
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   529
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
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„»·€ ﬁ.„"
            Height          =   285
            Index           =   39
            Left            =   6840
            TabIndex        =   260
            Top             =   5520
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»⁄œ «·Œ’„"
            Height          =   285
            Index           =   38
            Left            =   10950
            TabIndex        =   258
            Top             =   5520
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… ﬁ.„"
            Height          =   285
            Index           =   37
            Left            =   8460
            TabIndex        =   256
            Top             =   5520
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ‰”»… «·Œ’„"
            Height          =   285
            Index           =   34
            Left            =   13050
            TabIndex        =   253
            Top             =   5520
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„»·€ «·Œ’„"
            Height          =   285
            Index           =   32
            Left            =   15270
            TabIndex        =   252
            Top             =   5520
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Index           =   31
            Left            =   16500
            TabIndex        =   251
            Top             =   5520
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  «·ﬁÿ⁄ "
            Height          =   285
            Index           =   29
            Left            =   17700
            TabIndex        =   250
            Top             =   5520
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "«·«Ã„«·Ì"
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
            Height          =   300
            Index           =   28
            Left            =   5040
            TabIndex        =   247
            Top             =   900
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "«·ﬂ„Ì…"
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
            Height          =   300
            Index           =   26
            Left            =   6840
            TabIndex        =   245
            Top             =   900
            Width           =   1155
         End
         Begin VB.Label Label8 
            Caption         =   "«·’‰›"
            Height          =   165
            Left            =   18000
            TabIndex        =   243
            Top             =   960
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "«·”⁄—"
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
            Height          =   300
            Index           =   21
            Left            =   8760
            TabIndex        =   242
            Top             =   930
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  «·ﬁÿ⁄ «·„ﬁœ—…"
            Height          =   285
            Index           =   22
            Left            =   4380
            TabIndex        =   241
            Top             =   5520
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Index           =   23
            Left            =   2460
            TabIndex        =   240
            Top             =   5520
            Width           =   1845
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·ﬁÿ⁄ «·„ﬁœ—…"
            Height          =   285
            Index           =   24
            Left            =   17040
            TabIndex        =   239
            Top             =   600
            Width           =   1485
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   6060
         Left            =   19995
         TabIndex        =   213
         TabStop         =   0   'False
         Top             =   45
         Width           =   18960
         _cx             =   33443
         _cy             =   10689
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6060
         Left            =   19695
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   45
         Width           =   18960
         _cx             =   33443
         _cy             =   10689
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
         Begin VB.Frame lblExt 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‘ —Ì«  Ê«·«⁄„«· «·Œ«—ÃÌÂ"
            Height          =   2655
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   132
            Top             =   2880
            Width           =   18735
            Begin VSFlex8Ctl.VSFlexGrid fg2 
               Height          =   1980
               Left            =   0
               TabIndex        =   133
               Top             =   240
               Width           =   18675
               _cx             =   32941
               _cy             =   3492
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
               Cols            =   11
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmCarAuthontication.frx":50780
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   270
               Index           =   8
               Left            =   17760
               TabIndex        =   134
               Top             =   2280
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
               ButtonImage     =   "FrmCarAuthontication.frx":50901
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label LbToTalExtra 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BackStyle       =   0  'Transparent
               Height          =   300
               Left            =   120
               TabIndex        =   136
               Top             =   2280
               Width           =   1095
            End
            Begin VB.Label lblEx 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì «·„‘ —Ì«  Ê«·«⁄„«· «·Œ«—ÃÌ…"
               Height          =   285
               Left            =   1560
               TabIndex        =   135
               Top             =   2280
               Width           =   2655
            End
         End
         Begin VB.Frame LblWork 
            BackColor       =   &H00E2E9E9&
            Caption         =   "√⁄„«· «· ’·ÌÕ"
            Height          =   2775
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   120
            Width           =   18735
            Begin VSFlex8Ctl.VSFlexGrid fg 
               Height          =   2220
               Left            =   0
               TabIndex        =   128
               Top             =   240
               Width           =   18675
               _cx             =   32941
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
               Cols            =   24
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmCarAuthontication.frx":50E9B
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   270
               Index           =   21
               Left            =   17640
               TabIndex        =   129
               Top             =   2400
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
               ButtonImage     =   "FrmCarAuthontication.frx":51213
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbTotalMente 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BackStyle       =   0  'Transparent
               Height          =   300
               Left            =   240
               TabIndex        =   131
               Top             =   2520
               Width           =   1095
            End
            Begin VB.Label LblM 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì «⁄„«· «· ’·ÌÕ"
               Height          =   285
               Left            =   1320
               TabIndex        =   130
               Top             =   2520
               Width           =   1935
            End
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9000
            TabIndex        =   85
            Top             =   7080
            Width           =   3375
         End
         Begin VB.Label Label1100 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9120
            TabIndex        =   75
            Top             =   6240
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6060
         Index           =   15
         Left            =   45
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   45
         Width           =   18960
         _cx             =   33443
         _cy             =   10689
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
         _GridInfo       =   $"FrmCarAuthontication.frx":517AD
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6030
            Index           =   16
            Left            =   15
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   15
            Width           =   18930
            _cx             =   33390
            _cy             =   10636
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
            Begin VB.TextBox TxtNoteIntial 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   40
               Top             =   1440
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Frame lblDataCli 
               BackColor       =   &H00E2E9E9&
               Caption         =   "  »Ì«‰«  «·⁄„Ì·"
               Height          =   6090
               Left            =   -420
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   0
               Width           =   19080
               Begin VB.PictureBox Picture1 
                  Height          =   1815
                  Left            =   0
                  ScaleHeight     =   1755
                  ScaleWidth      =   3435
                  TabIndex        =   273
                  Top             =   6120
                  Width           =   3495
               End
               Begin VB.Frame Frame2 
                  BackColor       =   &H80000005&
                  Height          =   3615
                  Left            =   4080
                  TabIndex        =   265
                  Top             =   120
                  Width           =   6735
                  Begin VB.TextBox TxtCodeItem 
                     Alignment       =   1  'Right Justify
                     Height          =   405
                     Left            =   2280
                     TabIndex        =   272
                     Top             =   600
                     Width           =   2895
                  End
                  Begin VB.TextBox txtSalesInvoiceOrder 
                     Height          =   375
                     Left            =   2280
                     TabIndex        =   270
                     Top             =   1680
                     Width           =   2895
                  End
                  Begin VB.CommandButton Command4 
                     BackColor       =   &H000000FF&
                     Caption         =   "X"
                     Height          =   375
                     Left            =   6120
                     Style           =   1  'Graphical
                     TabIndex        =   266
                     Top             =   120
                     Width           =   375
                  End
                  Begin MSDataListLib.DataCombo cmbItems 
                     Height          =   315
                     Left            =   1680
                     TabIndex        =   268
                     Top             =   1080
                     Width           =   3525
                     _ExtentX        =   6218
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777152
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
                  Begin VB.Label Label9 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "—ﬁ„ «·›« Ê—…"
                     Height          =   255
                     Index           =   1
                     Left            =   5040
                     TabIndex        =   271
                     Top             =   1800
                     Width           =   1065
                  End
                  Begin VB.Label Label9 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·’‰›"
                     Height          =   255
                     Index           =   0
                     Left            =   5040
                     TabIndex        =   269
                     Top             =   1080
                     Width           =   825
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "»Ì«‰«  «·’‰›"
                     Height          =   270
                     Index           =   33
                     Left            =   1080
                     TabIndex        =   267
                     Top             =   240
                     Width           =   2625
                  End
               End
               Begin VB.Frame codecar 
                  BackColor       =   &H80000005&
                  Height          =   2295
                  Left            =   11040
                  TabIndex        =   222
                  Top             =   360
                  Width           =   6615
                  Begin VB.CommandButton Command2 
                     BackColor       =   &H000000FF&
                     Caption         =   "X"
                     Height          =   375
                     Left            =   6120
                     Style           =   1  'Graphical
                     TabIndex        =   226
                     Top             =   120
                     Width           =   375
                  End
                  Begin VB.TextBox txtCodeReg 
                     Alignment       =   1  'Right Justify
                     Height          =   375
                     Left            =   120
                     TabIndex        =   225
                     Top             =   1080
                     Width           =   5415
                  End
                  Begin VB.TextBox TxtCodeDoor 
                     Alignment       =   1  'Right Justify
                     Height          =   375
                     Left            =   120
                     TabIndex        =   224
                     Top             =   600
                     Width           =   5415
                  End
                  Begin VB.TextBox TxtCodeComputer 
                     Alignment       =   1  'Right Justify
                     Height          =   375
                     Left            =   120
                     TabIndex        =   223
                     Top             =   1680
                     Width           =   5415
                  End
                  Begin VB.Label lblCodeReg 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "—ﬁ„ «·”Ã·"
                     Height          =   270
                     Left            =   5520
                     TabIndex        =   230
                     Top             =   1200
                     Width           =   945
                  End
                  Begin VB.Label LblCodeDoor 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "—ﬁ„ «·»«»"
                     Height          =   270
                     Left            =   5550
                     TabIndex        =   229
                     Top             =   720
                     Width           =   945
                  End
                  Begin VB.Label lblcomputer 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "—ﬁ„ «·ﬂ„»ÌÊ —"
                     Height          =   270
                     Left            =   5520
                     TabIndex        =   228
                     Top             =   1800
                     Width           =   945
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "‘‹‹‹‹‹‹‹‹Ì›—… «·”‹‹‹‹‹Ì«—…"
                     Height          =   270
                     Index           =   17
                     Left            =   1080
                     TabIndex        =   227
                     Top             =   240
                     Width           =   2625
                  End
               End
               Begin VB.TextBox TxtSparePart 
                  Alignment       =   1  'Right Justify
                  Height          =   645
                  Left            =   4200
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   210
                  Top             =   3840
                  Width           =   4815
               End
               Begin VB.TextBox TxtCarMetarOut 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   12120
                  TabIndex        =   208
                  Top             =   5520
                  Width           =   1935
               End
               Begin VB.TextBox TxtLastWorOrder 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   15360
                  TabIndex        =   207
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.TextBox TxtClientCode 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   10800
                  TabIndex        =   199
                  Top             =   240
                  Width           =   2895
               End
               Begin VB.Frame Fra 
                  BackColor       =   &H00E2E9E9&
                  BorderStyle     =   0  'None
                  Height          =   375
                  Index           =   0
                  Left            =   14520
                  RightToLeft     =   -1  'True
                  TabIndex        =   192
                  Top             =   120
                  Width           =   4305
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«„— ‘€·"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   360
                     Index           =   0
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   195
                     ToolTipText     =   "«ﬂ»— „‰"
                     Top             =   0
                     Width           =   1065
                  End
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "≈–‰ «’·«Õ"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   360
                     Index           =   1
                     Left            =   1680
                     RightToLeft     =   -1  'True
                     TabIndex        =   194
                     ToolTipText     =   "Ì”«ÊÏ"
                     Top             =   0
                     Value           =   -1  'True
                     Width           =   975
                  End
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄—÷ ”⁄—"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   360
                     Index           =   2
                     Left            =   2400
                     RightToLeft     =   -1  'True
                     TabIndex        =   193
                     ToolTipText     =   "«’€— „‰"
                     Top             =   0
                     Width           =   1515
                  End
               End
               Begin VB.TextBox TxtCusID 
                  Height          =   285
                  Left            =   9720
                  TabIndex        =   191
                  Text            =   "Text1"
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.CommandButton menue 
                  Height          =   405
                  Index           =   12
                  Left            =   14520
                  Picture         =   "FrmCarAuthontication.frx":517E3
                  Style           =   1  'Graphical
                  TabIndex        =   189
                  Top             =   600
                  Width           =   375
               End
               Begin VB.TextBox TxtTypeCustomer 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   4920
                  TabIndex        =   183
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   23400
                  TabIndex        =   179
                  Top             =   4800
                  Width           =   855
               End
               Begin VB.TextBox Text2 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   10440
                  TabIndex        =   178
                  Top             =   12960
                  Width           =   2295
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "‘Ì‹‹‹‹‹‹‹‹›—… «·”‹‹‹Ì«—…"
                  Height          =   1035
                  Left            =   4200
                  Picture         =   "FrmCarAuthontication.frx":51C47
                  TabIndex        =   17
                  Top             =   2010
                  Width           =   1935
               End
               Begin VB.TextBox txtnotacept 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   3840
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   176
                  Top             =   5400
                  Width           =   2895
               End
               Begin VB.CommandButton menuet 
                  Caption         =   "⁄—÷ «· Ê’Ì«  «·”«»ﬁÂ"
                  Height          =   405
                  Left            =   4200
                  Picture         =   "FrmCarAuthontication.frx":55B8E
                  TabIndex        =   172
                  Top             =   3120
                  Width           =   1935
               End
               Begin VB.TextBox txtprivate 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   10800
                  TabIndex        =   18
                  Top             =   3090
                  Width           =   6975
               End
               Begin VB.TextBox txtrecomment 
                  Alignment       =   1  'Right Justify
                  Height          =   645
                  Left            =   9240
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   22
                  Top             =   3840
                  Width           =   3135
               End
               Begin VB.Frame frmgranty 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  «·÷„«‰"
                  Height          =   1455
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   161
                  Top             =   0
                  Width           =   3855
                  Begin VB.TextBox txtKM 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   480
                     TabIndex        =   180
                     Top             =   960
                     Width           =   1215
                  End
                  Begin VB.TextBox TxtLongGranty 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   1200
                     TabIndex        =   163
                     Top             =   240
                     Width           =   1095
                  End
                  Begin VB.ComboBox ComMD 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   162
                     Top             =   240
                     Width           =   1095
                  End
                  Begin MSComCtl2.DTPicker DateStartG 
                     Height          =   315
                     Left            =   2040
                     TabIndex        =   164
                     Top             =   600
                     Width           =   1215
                     _ExtentX        =   2143
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   193855489
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker DateEndg 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   165
                     Top             =   600
                     Width           =   1455
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   193855489
                     CurrentDate     =   38784
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ„"
                     Height          =   285
                     Index           =   16
                     Left            =   -360
                     TabIndex        =   182
                     Top             =   960
                     Width           =   765
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " €ÌÌ— «·“Ì  »⁄œ „—Ê—"
                     Height          =   285
                     Index           =   2
                     Left            =   1920
                     TabIndex        =   181
                     Top             =   960
                     Width           =   1605
                  End
                  Begin VB.Label lbllong 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "„œ… «·÷„«‰"
                     Height          =   255
                     Left            =   2520
                     TabIndex        =   168
                     Top             =   240
                     Width           =   1185
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Ì»œ√ „‰"
                     Height          =   285
                     Index           =   3
                     Left            =   3090
                     TabIndex        =   167
                     Top             =   615
                     Width           =   645
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Ì‰ ÂÌ"
                     Height          =   285
                     Index           =   5
                     Left            =   1410
                     TabIndex        =   166
                     Top             =   615
                     Width           =   645
                  End
               End
               Begin VB.Frame FrReturnMaint 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«⁄«œ… «’·«Õ"
                  Height          =   735
                  Left            =   0
                  TabIndex        =   157
                  Top             =   720
                  Width           =   3975
                  Begin VB.TextBox TxtOrder 
                     Alignment       =   1  'Right Justify
                     Height          =   405
                     Left            =   120
                     TabIndex        =   158
                     Top             =   240
                     Width           =   1575
                  End
                  Begin VB.Label lbEOrder 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«œŒ· «„— »ÿ«ﬁ… «·«’·«Õ"
                     Height          =   255
                     Left            =   1800
                     TabIndex        =   159
                     Top             =   360
                     Width           =   1665
                  End
               End
               Begin VB.CommandButton BtImage 
                  Caption         =   " ÕœÌœ «·„·«ÕŸ«  "
                  Height          =   495
                  Left            =   120
                  Picture         =   "FrmCarAuthontication.frx":560E6
                  TabIndex        =   24
                  Top             =   4560
                  Width           =   3975
               End
               Begin VB.TextBox TxtRemarkCar 
                  Alignment       =   1  'Right Justify
                  Height          =   2445
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   23
                  Top             =   2040
                  Width           =   3975
               End
               Begin VB.ComboBox DcbyearFactor 
                  Height          =   315
                  Left            =   7260
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   1560
                  Width           =   2295
               End
               Begin VB.TextBox TxtAmoutAccept 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   15120
                  TabIndex        =   28
                  Top             =   5520
                  Width           =   2415
               End
               Begin VB.TextBox TxtCliientName 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   14880
                  TabIndex        =   0
                  Top             =   600
                  Width           =   2895
               End
               Begin VB.ComboBox DcbOrderStatus 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   9840
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   5520
                  Width           =   1335
               End
               Begin VB.TextBox TxtClientPhone 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   14880
                  TabIndex        =   2
                  Top             =   1530
                  Width           =   2895
               End
               Begin VB.TextBox txtmobile 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   14880
                  TabIndex        =   1
                  Top             =   1080
                  Width           =   2895
               End
               Begin VB.TextBox TxtBox 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   10800
                  TabIndex        =   6
                  Top             =   1560
                  Width           =   2895
               End
               Begin VB.TextBox TxtComplaint 
                  Alignment       =   1  'Right Justify
                  Height          =   645
                  Left            =   15720
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   20
                  Top             =   3840
                  Width           =   3135
               End
               Begin VB.TextBox txtresonwait 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   3840
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   153
                  Top             =   5400
                  Width           =   2895
               End
               Begin VB.TextBox TxtNoteIntial1 
                  Alignment       =   1  'Right Justify
                  Height          =   645
                  Left            =   12480
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   21
                  Top             =   3840
                  Width           =   3135
               End
               Begin VB.TextBox TxtDriver 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   14880
                  TabIndex        =   3
                  Top             =   2040
                  Width           =   2895
               End
               Begin VB.OptionButton RdCompany 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘—ﬂ« "
                  Height          =   195
                  Left            =   7320
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.OptionButton RdPerson 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«›—«œ"
                  Height          =   195
                  Left            =   9360
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.Frame FramAccount 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ê÷⁄ «·„«·Ì"
                  Height          =   495
                  Left            =   2430
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   7695
                  Begin VB.OptionButton rdCredit 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ã·"
                     Height          =   195
                     Left            =   1920
                     RightToLeft     =   -1  'True
                     TabIndex        =   120
                     Top             =   240
                     Width           =   1335
                  End
                  Begin VB.OptionButton Rdacco 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«»"
                     Height          =   195
                     Left            =   3960
                     RightToLeft     =   -1  'True
                     TabIndex        =   119
                     Top             =   240
                     Width           =   1335
                  End
                  Begin VB.OptionButton RdCash 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰ﬁœ«"
                     Height          =   195
                     Left            =   6600
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   240
                     Width           =   975
                  End
               End
               Begin VB.TextBox TxtTtpeReg 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   4200
                  TabIndex        =   16
                  Top             =   1560
                  Width           =   1935
               End
               Begin VB.TextBox txtboxzip 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   10800
                  TabIndex        =   7
                  Top             =   2040
                  Width           =   2895
               End
               Begin VB.TextBox txtEmail 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   10800
                  TabIndex        =   5
                  Top             =   1080
                  Width           =   2895
               End
               Begin VB.TextBox txtFax 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   10800
                  TabIndex        =   4
                  Top             =   600
                  Width           =   2895
               End
               Begin VB.TextBox txtAddres 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   10800
                  TabIndex        =   8
                  Top             =   2610
                  Width           =   6975
               End
               Begin VB.TextBox TXtShaseh 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   4200
                  TabIndex        =   15
                  Top             =   1080
                  Width           =   1935
               End
               Begin VB.TextBox TXtCarMeter 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   4200
                  TabIndex        =   14
                  Top             =   600
                  Width           =   1935
               End
               Begin VB.TextBox TxtFirstPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   15120
                  TabIndex        =   27
                  Top             =   5160
                  Width           =   2415
               End
               Begin XtremeSuiteControls.CheckBox ChAccept 
                  Height          =   495
                  Left            =   8400
                  TabIndex        =   33
                  Top             =   5400
                  Width           =   1575
                  _Version        =   786432
                  _ExtentX        =   2778
                  _ExtentY        =   873
                  _StockProps     =   79
                  Caption         =   " „ „Ê«›ﬁ…  «·⁄„Ì· "
                  UseVisualStyle  =   -1  'True
               End
               Begin MSComCtl2.DTPicker DTPEnterDate 
                  Height          =   315
                  Left            =   16320
                  TabIndex        =   30
                  Top             =   4680
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   211353601
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker DTPTimeExptExit 
                  Height          =   315
                  Left            =   10080
                  TabIndex        =   26
                  Top             =   4680
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   211353602
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker DTPDateExptExit 
                  Height          =   315
                  Left            =   13080
                  TabIndex        =   25
                  Top             =   4680
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   211353601
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker DTPDateAcutExite 
                  Height          =   315
                  Left            =   7080
                  TabIndex        =   36
                  Top             =   4680
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   211353601
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker DTPTimeAcutExite 
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   37
                  Top             =   4680
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   211353602
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcbCarType 
                  Bindings        =   "FrmCarAuthontication.frx":5A02D
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   9
                  Top             =   600
                  Width           =   2295
                  _ExtentX        =   4048
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
               Begin MSDataListLib.DataCombo DcbCarModel 
                  Bindings        =   "FrmCarAuthontication.frx":5A042
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   10
                  Top             =   1080
                  Width           =   2295
                  _ExtentX        =   4048
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
                  Bindings        =   "FrmCarAuthontication.frx":5A057
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   12
                  Top             =   2040
                  Width           =   2295
                  _ExtentX        =   4048
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
               Begin MSDataListLib.DataCombo DcboFitter 
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   19
                  Top             =   3120
                  Width           =   2295
                  _ExtentX        =   4048
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox CheckBox1 
                  Height          =   495
                  Left            =   7200
                  TabIndex        =   173
                  Top             =   5400
                  Width           =   1095
                  _Version        =   786432
                  _ExtentX        =   1931
                  _ExtentY        =   873
                  _StockProps     =   79
                  Caption         =   " Õ  «·«‰ Ÿ«—"
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox CheckBox2 
                  Height          =   495
                  Left            =   5640
                  TabIndex        =   174
                  Top             =   5400
                  Width           =   1455
                  _Version        =   786432
                  _ExtentX        =   2566
                  _ExtentY        =   873
                  _StockProps     =   79
                  Caption         =   "⁄œ„ „Ê«›ﬁ…«·⁄„Ì·"
                  UseVisualStyle  =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbCar 
                  Bindings        =   "FrmCarAuthontication.frx":5A06C
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   212
                  Top             =   2610
                  Width           =   2295
                  _ExtentX        =   4048
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
                  Left            =   7200
                  TabIndex        =   13
                  Top             =   2610
                  Width           =   2295
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÿ⁄ €Ì«— „‰ ﬁ»· «·⁄„Ì·"
                  Height          =   270
                  Left            =   4920
                  TabIndex        =   211
                  Top             =   3600
                  Width           =   2835
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œ«œ «·Œ—ÊÃ"
                  Height          =   255
                  Left            =   13890
                  TabIndex        =   209
                  Top             =   5550
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·⁄„Ì·"
                  Height          =   285
                  Index           =   19
                  Left            =   13560
                  TabIndex        =   200
                  Top             =   255
                  Width           =   885
               End
               Begin VB.Label lbltypecus 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄„Ì·"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C000C0&
                  Height          =   375
                  Left            =   4200
                  TabIndex        =   190
                  Top             =   120
                  Width           =   2895
               End
               Begin VB.Label lblnotacept 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "”»» «·—›÷"
                  Height          =   255
                  Left            =   5400
                  TabIndex        =   175
                  Top             =   5160
                  Width           =   975
               End
               Begin VB.Label lblprivatecopm 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ«  Œ«’Â"
                  Height          =   375
                  Left            =   17760
                  TabIndex        =   170
                  Top             =   3120
                  Width           =   1065
               End
               Begin VB.Label lblrecomentclient 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· Ê’Ì«  ··⁄„Ì·"
                  Height          =   270
                  Left            =   9960
                  TabIndex        =   169
                  Top             =   3600
                  Width           =   1515
               End
               Begin VB.Label lblmarks 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„·«ÕŸ«  ⁄·Ï «·„⁄œÂ/«·”Ì«—…"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   160
                  Top             =   1800
                  Width           =   2055
               End
               Begin VB.Label LblYear 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„ÊœÌ· «·„⁄œÂ/«·”Ì«—…"
                  Height          =   255
                  Left            =   9570
                  TabIndex        =   154
                  Top             =   1560
                  Width           =   1095
               End
               Begin VB.Label lblresonwaite 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "”»» «·«‰ Ÿ«—"
                  Height          =   255
                  Left            =   4440
                  TabIndex        =   149
                  Top             =   5160
                  Width           =   975
               End
               Begin VB.Label lblcodeclient 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬂÊœ «·⁄„Ì·"
                  Height          =   270
                  Left            =   20400
                  TabIndex        =   138
                  Top             =   4920
                  Visible         =   0   'False
                  Width           =   945
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘ﬂÊÏ «·⁄„Ì·"
                  Height          =   435
                  Index           =   15
                  Left            =   16860
                  TabIndex        =   137
                  Top             =   3600
                  Width           =   945
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Êﬁ  «·Œ—ÊÃ «·›⁄·Ì"
                  Height          =   195
                  Index           =   14
                  Left            =   5640
                  TabIndex        =   126
                  Top             =   4680
                  Width           =   1320
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Êﬁ  «·Œ—ÊÃ «·„ Êﬁ⁄"
                  Height          =   195
                  Index           =   13
                  Left            =   11640
                  TabIndex        =   125
                  Top             =   4710
                  Width           =   1380
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—Œ «·Œ—ÊÃ «·„ Êﬁ⁄"
                  Height          =   195
                  Index           =   12
                  Left            =   14640
                  TabIndex        =   124
                  Top             =   4710
                  Width           =   1395
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " √—ÌŒ «·Œ—ÊÃ «·›⁄·Ì"
                  Height          =   195
                  Index           =   11
                  Left            =   8520
                  TabIndex        =   123
                  Top             =   4680
                  Width           =   1410
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·”«∆ﬁ"
                  Height          =   195
                  Index           =   10
                  Left            =   18060
                  TabIndex        =   122
                  Top             =   2160
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " √—ÌŒ «·œŒÊ·"
                  Height          =   195
                  Index           =   9
                  Left            =   17925
                  TabIndex        =   121
                  Top             =   4710
                  Width           =   900
               End
               Begin VB.Label lblremrk 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ«  «·„‘—›"
                  Height          =   615
                  Left            =   13440
                  TabIndex        =   118
                  Top             =   3600
                  Width           =   1215
               End
               Begin VB.Label LblFitter 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‘—› «·«” ﬁ»«·"
                  Height          =   270
                  Left            =   9480
                  TabIndex        =   117
                  Top             =   3150
                  Width           =   1185
               End
               Begin VB.Label lblTypeReg 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰Ê⁄ «· ”ÃÌ·"
                  Height          =   255
                  Left            =   6000
                  TabIndex        =   116
                  Top             =   1590
                  Width           =   1095
               End
               Begin VB.Label LblCarMeter 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁ—«∆… «·⁄œ«œ"
                  Height          =   255
                  Left            =   6000
                  TabIndex        =   115
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.Label lblboxzib 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·—„“ «·»—ÌœÌ"
                  Height          =   255
                  Left            =   13695
                  TabIndex        =   114
                  Top             =   2070
                  Width           =   975
               End
               Begin VB.Label lblemail 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«Ì„Ì·"
                  Height          =   255
                  Left            =   13575
                  TabIndex        =   113
                  Top             =   1200
                  Width           =   855
               End
               Begin VB.Label lblfax 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·›«ﬂ”"
                  Height          =   255
                  Left            =   13575
                  TabIndex        =   112
                  Top             =   750
                  Width           =   855
               End
               Begin VB.Label lblAdress 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄‰Ê«‰"
                  Height          =   255
                  Left            =   17970
                  TabIndex        =   111
                  Top             =   2640
                  Width           =   855
               End
               Begin VB.Label lblbox 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "’‰œÊﬁ »—Ìœ"
                  Height          =   255
                  Left            =   13680
                  TabIndex        =   110
                  Top             =   1590
                  Width           =   855
               End
               Begin VB.Label lblMobile 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·ÃÊ«·"
                  Height          =   255
                  Left            =   17970
                  TabIndex        =   109
                  Top             =   1080
                  Width           =   855
               End
               Begin VB.Label LblCodeShaseh 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·‘«”ÌÂ"
                  Height          =   270
                  Left            =   6150
                  TabIndex        =   108
                  Top             =   1080
                  Width           =   945
               End
               Begin VB.Label LblAmountAcc 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· ﬂ·›… «· ﬁœÌ—Ì…"
                  Height          =   255
                  Left            =   17730
                  TabIndex        =   107
                  Top             =   5190
                  Width           =   1095
               End
               Begin VB.Label LblPhone 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Â« › «·⁄„Ì·"
                  Height          =   255
                  Left            =   17970
                  TabIndex        =   106
                  Top             =   1560
                  Width           =   855
               End
               Begin VB.Label LblPla 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «··ÊÕ…"
                  Height          =   255
                  Left            =   9570
                  TabIndex        =   105
                  Top             =   2670
                  Width           =   1095
               End
               Begin VB.Label LblOrderSt 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Õ«·… «·ÿ·»"
                  Height          =   255
                  Left            =   10920
                  TabIndex        =   104
                  Top             =   5550
                  Width           =   1095
               End
               Begin VB.Label LblPayF 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·œ›⁄… «·„ﬁœ„…"
                  Height          =   270
                  Left            =   17520
                  TabIndex        =   103
                  Top             =   5550
                  Width           =   1185
               End
               Begin VB.Label lblColor 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "·Ê‰ «·„⁄œÂ/«·”Ì«—…"
                  Height          =   255
                  Left            =   9810
                  TabIndex        =   102
                  Top             =   2070
                  Width           =   855
               End
               Begin VB.Label lblModel 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ÿ—«“ "
                  Height          =   255
                  Left            =   9810
                  TabIndex        =   101
                  Top             =   1080
                  Width           =   855
               End
               Begin VB.Label lbltycar 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰Ê⁄ «·„⁄œÂ/«·”Ì«—…"
                  Height          =   255
                  Left            =   9810
                  TabIndex        =   100
                  Top             =   600
                  Width           =   855
               End
               Begin VB.Label LblCli 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄„Ì·"
                  Height          =   255
                  Left            =   17970
                  TabIndex        =   99
                  Top             =   630
                  Width           =   855
               End
            End
            Begin VB.Label lbtechnical 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„·«ÕŸ… «·„»œ∆Ì… ··›‰Ì"
               Height          =   450
               Left            =   990
               TabIndex        =   95
               Top             =   2385
               Width           =   300
            End
            Begin VB.Label lbldif 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   390
               Left            =   150
               TabIndex        =   92
               Top             =   5610
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   420
               Left            =   -225
               TabIndex        =   91
               Top             =   -4785
               Width           =   150
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ã„«·Ì «·⁄«„"
               Height          =   270
               Left            =   2430
               TabIndex        =   90
               Top             =   2115
               Width           =   450
            End
            Begin VB.Label firstprice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   390
               Left            =   1065
               TabIndex        =   89
               Top             =   5610
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.Label lbtotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   390
               Left            =   1740
               TabIndex        =   87
               Top             =   5610
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.Label LbToTa 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ã„«·Ì «·⁄«„"
               Height          =   375
               Left            =   1815
               TabIndex        =   86
               Top             =   5805
               Visible         =   0   'False
               Width           =   450
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   3540
               Index           =   62
               Left            =   450
               TabIndex        =   78
               Top             =   1725
               Width           =   75
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6030
            Index           =   9
            Left            =   15
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   15
            Width           =   18930
            _cx             =   33390
            _cy             =   10636
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
               Height          =   4665
               Left            =   600
               MaxLength       =   4
               TabIndex        =   81
               Top             =   1305
               Width           =   150
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   3090
               Left            =   750
               TabIndex        =   80
               Top             =   1725
               Width           =   240
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   3090
               Index           =   67
               Left            =   375
               TabIndex        =   84
               Top             =   1725
               Width           =   75
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   3075
               Index           =   68
               Left            =   750
               TabIndex        =   83
               Top             =   2160
               Width           =   90
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
               Height          =   3615
               Index           =   69
               Left            =   450
               TabIndex        =   82
               Top             =   1725
               Width           =   150
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   6060
         Left            =   20295
         TabIndex        =   215
         TabStop         =   0   'False
         Top             =   45
         Width           =   18960
         _cx             =   33443
         _cy             =   10689
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
            Height          =   5385
            Left            =   240
            TabIndex        =   216
            Top             =   0
            Width           =   18765
            _cx             =   33099
            _cy             =   9499
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
            FormatString    =   $"FrmCarAuthontication.frx":5A081
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
               TabIndex        =   217
               Top             =   5880
               Width           =   1440
            End
         End
         Begin MSDataListLib.DataCombo dcItemunit 
            Height          =   315
            Left            =   0
            TabIndex        =   231
            Top             =   0
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰œ«  «·„‰’—›… ··«„—"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   35
            Left            =   7680
            TabIndex        =   221
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
            TabIndex        =   220
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
            TabIndex        =   219
            Top             =   5520
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Index           =   58
            Left            =   240
            TabIndex        =   218
            Top             =   5520
            Width           =   3765
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   17
      Left            =   3330
      TabIndex        =   248
      Top             =   7920
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… ⁄—÷ ”⁄— 2"
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
      Index           =   18
      Left            =   1590
      TabIndex        =   249
      Top             =   7950
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… «·›« Ê—…"
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«„— ‘€·"
      Height          =   285
      Index           =   8
      Left            =   13560
      TabIndex        =   204
      Top             =   720
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«–‰ «’·«Õ"
      Height          =   285
      Index           =   7
      Left            =   15720
      TabIndex        =   202
      Top             =   720
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«” Œœ«„ «·‘«‘Â"
      Height          =   255
      Index           =   18
      Left            =   1800
      TabIndex        =   188
      Top             =   750
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "”»» «·«‰ Ÿ«—"
      Height          =   255
      Left            =   14040
      TabIndex        =   151
      Top             =   4230
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·⁄„Ì·"
      Height          =   255
      Left            =   6600
      TabIndex        =   94
      Top             =   10320
      Width           =   855
   End
   Begin VB.Label lblty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "›∆… «·ÿ·»"
      Height          =   255
      Left            =   4920
      TabIndex        =   88
      Top             =   750
      Width           =   855
   End
   Begin VB.Image img 
      Height          =   855
      Left            =   22680
      Picture         =   "FrmCarAuthontication.frx":5A24D
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
      Left            =   15600
      Picture         =   "FrmCarAuthontication.frx":5A7FB
      Stretch         =   -1  'True
      Top             =   10920
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Â–Â «·’Ê—…  ”„Õ ·ﬂ » ÕœÌœ «·«Ã“«¡ «·„—«œ «’·«ÕÂ«"
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
      Height          =   540
      Index           =   25
      Left            =   11640
      TabIndex        =   71
      Top             =   9120
      Width           =   4575
   End
   Begin VB.Label lblBr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      Height          =   255
      Left            =   7920
      TabIndex        =   66
      Top             =   750
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   20760
      TabIndex        =   65
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄—÷ ”⁄—"
      Height          =   285
      Index           =   4
      Left            =   17910
      TabIndex        =   61
      Top             =   750
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ Ê«·Êﬁ "
      Height          =   285
      Index           =   1
      Left            =   11070
      TabIndex        =   60
      Top             =   750
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   1650
      TabIndex        =   59
      Top             =   7110
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1050
      TabIndex        =   58
      Top             =   7140
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   2700
      TabIndex        =   57
      Top             =   7140
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   21240
      TabIndex        =   56
      Top             =   2640
      Width           =   1005
   End
End
Attribute VB_Name = "FrmCarAuthontication"
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
Dim ide As Integer
Public bo As Boolean
Public LngRow As Long
Public chektab As Boolean
Public LngCol As Long
Public chpo As Boolean
Public screenData As Boolean
Dim WorkOrder As Double
Dim ShowPriceOrder As Double
Dim AuthoOrder As Double
Public mFromCustomerForm As Boolean

Private Sub cmbItems_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

      Load FrmItemSearch
            FrmItemSearch.RetrunType = 310
            FrmItemSearch.show vbModal
            
End If

End Sub


Private Sub cmdOpenCard_Click()
Dim s As String
Dim Msg As String
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ê› Ì „ › Õ «·ﬂ«—  „⁄ «‰  Â–« «·ﬂ«—  „€·ﬁ ‰Â«∆Ì«  .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
        Else
            Msg = "This card will be open"
        End If
        If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
            s = "Update TblCardAuthorizationReform Set IsEndAll = 0,OrderStatus =1 where Id = " & val(XPTxtID.text)
            Me.DcbOrderStatus.ListIndex = 1
            Cn.Execute s
            rs.Resync adAffectCurrent
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ › Õ «·ﬂ«—   "
                cmdOpenCard.Caption = " „ › Õ «·ﬂ«— "
                cmdEndAll.Caption = "«ﬁ›«· «·ﬂ«—  ‰Â«∆Ì«"
            Else
                MsgBox "The card has been opened"
                cmdOpenCard.Caption = "The card has been opened"
                cmdEndAll.Caption = "Close the card"
            
            End If
            cmdOpenCard.Enabled = False
            
            cmdOpenCard.Enabled = True
        End If
            
        

End Sub

Private Sub cmdEndAll_Click()
Dim s As String
Dim Msg As String
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ê› Ì „ «ﬁ›«· Â–« «·ﬂ«—  ‰Â«∆Ì«  .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
        Else
            Msg = "This card will be permanently closed"
        End If
        If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
            s = "Update TblCardAuthorizationReform Set IsEndAll = 1,OrderStatus =2 where Id = " & val(XPTxtID.text)
            Me.DcbOrderStatus.ListIndex = 2
            Cn.Execute s
            rs.Resync adAffectCurrent
            
            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "The card has been permanently locked "
                cmdEndAll.Caption = "The card has been permanently locked "
                cmdOpenCard.Caption = "Open Card"
            Else
            
                MsgBox " „ «ﬁ›«· «·ﬂ«—  ‰Â«∆Ì« "
                
                cmdEndAll.Caption = " „ «ﬁ›«· «·ﬂ«— "
                
                
                
                cmdOpenCard.Caption = " › Õ «·ﬂ«— "
            End If
           
           cmdOpenCard.Enabled = True
           
           cmdEndAll.Enabled = False
        End If
            
        

End Sub

Private Sub DcboItems_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

      Load FrmItemSearch
            FrmItemSearch.RetrunType = 310
            FrmItemSearch.show vbModal
End If

If KeyCode = vbKeyF5 Then
    Dim Dcombos As New ClsDataCombos
   
    Dcombos.GetItemsNames Me.DcboItems
    
End If

End Sub

Private Sub DcboItems_Change()
    Dim UnitID As Long
    Dim UnitName As String
     Me.TxtItemCode.text = GetItemCode(val(Me.DcboItems.BoundText))
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsUnits·byitemid Me.dcItemunit, val(Me.DcboItems.BoundText)
    GetDefaultItemUnit val(Me.DcboItems.BoundText), UnitID, UnitName
    dcItemunit.text = UnitName
    dcItemunit.BoundText = UnitID
    Me.TxtItemPrice.text = GetItemPrice(val(Me.DcboItems.BoundText), 1, UnitID)

End Sub

Private Sub DcboItems_Click(Area As Integer)
'    DcboItems_Change
End Sub

Private Sub FG22_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim k As Integer
    Dim StrComboList As String
    Dim UnitID As Long
    Dim UnitName As String
    With FG22
        Select Case .ColKey(Col)

              Case "itemname"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ItemID"), False, True)
               .TextMatrix(Row, .ColIndex("ItemID")) = StrAccountCode
               .TextMatrix(Row, .ColIndex("ItemCode")) = GetItemCode(val(.TextMatrix(Row, .ColIndex("ItemID"))))
             
             Case "ItemCode"
                Set rs = New ADODB.Recordset
                StrSQL = " SELECT        TOP (100) PERCENT ItemID, ItemName, ItemNamee, Fullcode"
                StrSQL = StrSQL & "            From dbo.TblItems"
                StrSQL = StrSQL & "          WHERE        (Fullcode = N'" & .TextMatrix(Row, .ColIndex("ItemCode")) & "')"
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Row, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                Else
                    .TextMatrix(Row, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemNamee").value), "", rs("ItemNamee").value)
                End If
             Else
                .TextMatrix(Row, .ColIndex("ItemID")) = 0
                .TextMatrix(Row, .ColIndex("ItemName")) = ""
             End If
      

           
         Case "DiscValue"
            .TextMatrix(Row, .ColIndex("Price")) = val(.TextMatrix(Row, .ColIndex("PriceBDisc"))) - val(.TextMatrix(Row, Col))
            If val(.TextMatrix(Row, .ColIndex("PriceBDisc"))) <> 0 Then
                .TextMatrix(Row, .ColIndex("DiscPercent")) = Round(val(.TextMatrix(Row, Col)) / val(.TextMatrix(Row, .ColIndex("PriceBDisc"))) * 100, 2)
            Else
                .TextMatrix(Row, .ColIndex("DiscPercent")) = ""
                .TextMatrix(Row, Col) = ""
            End If
         Case "DiscPercent"
            .TextMatrix(Row, .ColIndex("DiscValue")) = val(.TextMatrix(Row, .ColIndex("PriceBDisc"))) * val(.TextMatrix(Row, Col)) / 100
            .TextMatrix(Row, .ColIndex("Price")) = val(.TextMatrix(Row, .ColIndex("PriceBDisc"))) - val(.TextMatrix(Row, .ColIndex("DiscValue")))
         Case "PriceBDisc"
            If val(.TextMatrix(Row, .ColIndex("PriceBDisc"))) <> 0 Then
                .TextMatrix(Row, .ColIndex("Price")) = val(.TextMatrix(Row, .ColIndex("PriceBDisc"))) - val(.TextMatrix(Row, .ColIndex("DiscValue")))
            Else
                .TextMatrix(Row, .ColIndex("Price")) = ""
                .TextMatrix(Row, .ColIndex("DiscPercent")) = ""
                .TextMatrix(Row, .ColIndex("DiscValue")) = ""
            End If
         End Select
         Dim i As Long

'        For i = 1 To FG22.Rows - 1
'
'        Next
                ReLineGrid2
                
    End With
    'RelinFg

End Sub

Private Sub FG22_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With FG22
        Select Case .ColKey(Col)
            Case "ActualQty", "isReplaced", "ForUnit", "TotalWithVat", "PriceBDisc", "DiscPercent", "DiscValue", "ItemName2", "Price"
            .ComboList = ""
            Case "Amount", "Vatyo", "Vat2"
            Cancel = True
            Case "ItemQty", "Remark"
            .ComboList = ""
              Case "ItemCode"
            .ComboList = ""
              Case "QtyPerfect", "Calories"
            .ComboList = ""
            Case "ItemPrice", "BeforeVat"
            Cancel = True
        End Select

    End With
End Sub

Private Sub FG22_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim StrAccountCode As String
Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With FG22

    Select Case .ColKey(Col)
   Case "itemname"
     StrSQL = " SELECT     ItemID, ItemName, ItemNamee"
     StrSQL = StrSQL & "  From dbo.TblItems"
     Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "ItemName", "ItemID")
                Else
                    StrComboList = .BuildComboList(rs, "ItemNamee", "ItemID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
    
       

        End Select

    End With


End Sub

Private Sub Text1_Change()

End Sub

Private Sub TxtCodeItem_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

      Load FrmItemSearch
            FrmItemSearch.RetrunType = 310
            FrmItemSearch.show vbModal
End If

End Sub

Private Sub txtDiscPercent_Validate(Cancel As Boolean)
    txtDiscValue = val(lbl(31)) * val(txtDiscPercent) / 100
    txtTotalAfterDiscount = val(val(lbl(31))) - val(txtDiscValue)
End Sub

Private Sub txtDiscValue_Change()

    txtTotalAfterDiscount = val(lbl(31)) - val(txtDiscValue)
    If val(lbl(31)) <> 0 Then
        txtDiscPercent = Round(val(txtDiscValue) / val(lbl(31)) * 100, 2)
    End If
     CalculteValueAdded 1, 21
End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If TxtItemCode.text = "" Then
            Me.DcboItems.BoundText = ""
        Else
            Me.DcboItems.BoundText = GetItemID(Trim$(Me.TxtItemCode.text))
        End If
    End If

End Sub
Public Sub Retrive3(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
  clear_all Me
       If rs.State = adStateOpen Then
   rs.Close
   
   Else
'rs.Open
   
   End If

     StrSQL = "select * From dbo.TblCardAuthorizationReform     where id=" & Lngid & ""
       rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
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

    XPTxtID.text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    'Me.TxtEndDate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    DcbCarType.BoundText = IIf(IsNull(rs("CarTypeID").value), "", rs("CarTypeID").value)
    DcbCarModel.BoundText = IIf(IsNull(rs("CarModelID").value), "", rs("CarModelID").value)
      TxtTypeCustomer.text = val(IIf(IsNull(rs("TypeCustomer").value), 0, rs("TypeCustomer").value))
     txtKM.text = IIf(IsNull(rs("OverKM").value), "", rs("OverKM").value)
    
    cmbItems.BoundText = IIf(IsNull(rs("ItemID33").value), "", rs("ItemID33").value)
   txtSalesInvoiceOrder = IIf(IsNull(rs("SalesInvoiceOrder").value), "", rs("SalesInvoiceOrder").value)
    

   ' DcboSpecifications.BoundText = IIf(IsNull(rs("gradeID").value), "", rs("gradeID").value)
   ' Me.TxtRemarkCar.text = IIf(IsNull(rs("Remarkcar").value), "", rs("Remarkcar").value)
    DcbColor.BoundText = IIf(IsNull(rs("ColorID").value), "", rs("ColorID").value)
    DcbyearFactor.text = IIf(IsNull(rs("YearFact").value), "", rs("YearFact").value)
   TxtClientPhone.text = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
   TxtCliientName.text = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
    cmdEndAll.Tag = val(rs!IsEndAll & "")
   If val(rs!IsEndAll & "") = 1 Then
    cmdEndAll.Enabled = False
    cmdEndAll.Caption = " „ «ﬁ›«· «·ﬂ«— "
    
    cmdOpenCard.Enabled = True
    cmdOpenCard.Caption = " › Õ «·ﬂ«— "
            
   Else
     cmdEndAll.Enabled = True
     cmdEndAll.Caption = "«ﬁ›«· «·ﬂ«—  ‰Â«∆Ï"
     
    cmdOpenCard.Enabled = False
    cmdOpenCard.Caption = " › Õ «·ﬂ«— "

   End If
    
        TxtCusID.text = val(IIf(IsNull(rs("CusID").value), 0, rs("CusID").value))
        If val(TxtCusID.text) = 0 Then
            Dim ss As String
            ss = "Select cusId From TblCustemers Where Code = N'" & Trim(TxtClientCode) & "'"
            Dim rsDummy As New ADODB.Recordset
            rsDummy.Open ss, Cn, adOpenStatic, adLockReadOnly
            If Not rsDummy.EOF Then
                TxtCusID.text = rsDummy!CusID & ""
            End If
        End If
        
    If SystemOptions.UserInterface = EnglishInterface Then
        StrSQL = "Select CusNamee ClientName FROM TblCustemers where CusId = " & val(TxtCusID.text)
        Dim rsDu As New ADODB.Recordset
        rsDu.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
        If Not rsDu.EOF Then
            TxtCliientName.text = IIf(IsNull(rsDu("ClientName").value), "", rsDu("ClientName").value)
        End If
        If Trim(TxtCliientName.text) = "" Then
            TxtCliientName.text = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
        End If
    End If

   TxtPlatNo.text = IIf(IsNull(rs("PlateNo").value), "", rs("PlateNo").value)
   Me.TxtCodeComputer.text = IIf(IsNull(rs("CodeComputer").value), "", rs("CodeComputer").value)
        Me.TxtWorkOrder.text = IIf(IsNull(rs("WorkOrder").value), "", rs("WorkOrder").value)
     Me.TxtShowPriceOrder.text = IIf(IsNull(rs("ShowPriceOrder").value), "", rs("ShowPriceOrder").value)
     Me.TxtAuthoOrder.text = IIf(IsNull(rs("AuthoOrder").value), "", rs("AuthoOrder").value)
     
   'DcbOrderStatus.ListIndex = IIf(IsNull(rs("OrderStatus").value), 0, rs("OrderStatus").value)
   'TXtCarMeter.text = IIf(IsNull(rs("CarMeter").value), "", rs("CarMeter").value)
   'TxtLongGranty.text = IIf(IsNull(rs("LongGranty").value), "", rs("LongGranty").value)
   'TxtFirstPrice.text = val(IIf(IsNull(rs("PayFirst").value), 0, rs("PayFirst").value))
   Me.TXtShaseh.text = IIf(IsNull(rs("Shaseh").value), "", rs("Shaseh").value)
   '    Me.DcboFitter.BoundText = IIf(IsNull(rs("FitterID").value), "", rs("FitterID").value)
       Me.TxtMobile.text = IIf(IsNull(rs("mobile").value), "", rs("mobile").value) ' rs("mobile").value
        Me.TxtBox.text = IIf(IsNull(rs("box").value), "", rs("box").value) 'rs("box").value
        Me.TxtFax.text = IIf(IsNull(rs("fax").value), "", rs("fax").value) 'rs("fax").value
        Me.TxtEmail.text = IIf(IsNull(rs("email").value), "", rs("email").value) ' rs("email").value
         Me.TxtAddres.text = IIf(IsNull(rs("address").value), "", rs("address").value) ' rs("address").value
         Me.txtboxzip.text = IIf(IsNull(rs("boxzip").value), "", rs("boxzip").value) 'rs("boxzip").value
         Me.txtCodeReg.text = IIf(IsNull(rs("codereg").value), "", rs("codereg").value) 'rs("codereg").value
         Me.TxtTtpeReg.text = IIf(IsNull(rs("typereg").value), "", rs("typereg").value) 'rs("typereg").value
       If SystemOptions.LinkCustomerWithCars = True Then
       Dim Dcombos As ClsDataCombos
       Set Dcombos = New ClsDataCombos
       Dcombos.GetCarsOfCustomer DcbCar, val(TxtCusID.text)
       End If
Me.DcbCar.BoundText = IIf(IsNull(rs("CarID").value), "", rs("CarID").value)
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter  As Integer
        Me.lbTotalMente.Caption = 0
        Me.Lbtotal.Caption = 0
        Me.LbToTalExtra.Caption = 0
    IntCounter = 0

    With Fg

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("serial")) = IntCounter
                If val(.TextMatrix(i, .ColIndex("count"))) <> 0 Then
        .TextMatrix(i, .ColIndex("totalm")) = val(.TextMatrix(i, .ColIndex("value"))) * .TextMatrix(i, .ColIndex("count"))
        Else
        .TextMatrix(i, .ColIndex("totalm")) = val(.TextMatrix(i, .ColIndex("value")))
        .TextMatrix(i, .ColIndex("count")) = 1
       End If
        .TextMatrix(i, .ColIndex("serial")) = IntCounter
            End If
 If .TextMatrix(i, .ColIndex("value")) <> "" Then
                
                Me.lbTotalMente.Caption = val(Me.lbTotalMente.Caption) + val(Fg.TextMatrix(i, Fg.ColIndex("totalm")))
        
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
        .TextMatrix(i, .ColIndex("totalex")) = val(.TextMatrix(i, .ColIndex("value"))) * .TextMatrix(i, .ColIndex("count"))
        Else
        .TextMatrix(i, .ColIndex("totalex")) = val(.TextMatrix(i, .ColIndex("value")))
        .TextMatrix(i, .ColIndex("count")) = 1
       End If
            End If

      
 If .TextMatrix(i, .ColIndex("value")) <> "" Then
                
                Me.LbToTalExtra.Caption = val(Me.LbToTalExtra.Caption) + val(fg2.TextMatrix(i, fg2.ColIndex("totalex")))
        
            End If
        Next i
    End With
Me.Lbtotal.Caption = val(Me.LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
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

Private Sub bClose_Click()
BtImage.Visible = True
gimage.Visible = False
lblmarks.Visible = True
Me.TxtRemarkCar.Visible = True
End Sub

Private Sub BtImage_Click()
'Dim val As Integer
'val = 3000
'gimage.Visible = True
'Me.gimage.Width = 8500
'Image6.Width = 8500
BtImage.Visible = False
'imwidth 0
gimage.Visible = True
lblmarks.Visible = False
Me.TxtRemarkCar.Visible = False
'Me.imag1.left = Me.imag1.left + 3565
'Me.imag1.left = Me.imag1.left + 3565

End Sub
Sub imwidth(Optional Lngid As Long = 0)
If Lngid = 0 Then
Me.imag1.left = Me.imag1.left + 3350
Me.imag2.left = Me.imag2.left + 3000
Me.imag3.left = Me.imag3.left + 2800
Me.imag4.left = Me.imag4.left + 2300
Me.imag5.left = Me.imag5.left + 1500
Me.img6.left = Me.img6.left + 1000
Me.img7.left = Me.img7.left + 500
Me.img8.left = Me.img8.left + 3500
Me.img9.left = Me.img9.left + 3100
Me.img10.left = Me.img10.left + 2800
Me.img11.left = Me.img11.left + 2300
Me.img12.left = Me.img12.left + 1200
Me.img13.left = Me.img13.left + 800
Me.img14.left = Me.img14.left + 300
Else
Me.imag1.left = Me.imag1.left - 3500
Me.imag2.left = Me.imag2.left - 3000
Me.imag3.left = Me.imag3.left - 2800
Me.imag4.left = Me.imag4.left - 2300
Me.imag5.left = Me.imag5.left - 1500
Me.img6.left = Me.img6.left - 1000
Me.img7.left = Me.img7.left - 500
Me.img8.left = Me.img8.left - 3500
Me.img9.left = Me.img9.left - 3100
Me.img10.left = Me.img10.left - 2800
Me.img11.left = Me.img11.left - 2300
Me.img12.left = Me.img12.left - 1200
Me.img13.left = Me.img13.left - 800
Me.img14.left = Me.img14.left - 300
End If
End Sub

Private Sub ChAccept_Click()
If chpo = False Then
If Me.ChAccept.value = vbChecked Then
If Me.CheckBox1.value = vbChecked Then
Cmd_Click (1)
Me.DcbOrderStatus.ListIndex = 3
Me.CheckBox1.value = xtpChecked
txtnotacept.Visible = False
lblnotacept.Visible = False

Me.CheckBox2.value = xtpUnchecked
Cmd_Click (2)
Else
Cmd_Click (1)
txtnotacept.Visible = False
lblnotacept.Visible = False
Me.DcbOrderStatus.ListIndex = 1
Me.CheckBox1.value = xtpUnchecked
Me.CheckBox2.value = xtpUnchecked
Cmd_Click (2)
End If
End If
End If
chpo = False
End Sub
Sub imgg()
 Me.img9.Picture = Me.imgnul.Picture
        Me.img10.Picture = Me.imgnul.Picture
        Me.imag1.Picture = Me.imgnul.Picture
        Me.imag2.Picture = Me.imgnul.Picture
        Me.imag3.Picture = Me.imgnul.Picture
        Me.imag4.Picture = Me.imgnul.Picture
        Me.imag5.Picture = Me.imgnul.Picture
        Me.img6.Picture = Me.imgnul.Picture
        Me.img7.Picture = Me.imgnul.Picture
        Me.img8.Picture = Me.imgnul.Picture
         Me.img11.Picture = Me.imgnul.Picture
        Me.img12.Picture = Me.imgnul.Picture
        Me.img13.Picture = Me.imgnul.Picture
        Me.img14.Picture = Me.imgnul.Picture
End Sub




Private Sub CheckBox1_Click()
If chpo = False Then

If Me.CheckBox1.value = vbChecked Then
'Cmd_Click (1)
Me.DcbOrderStatus.ListIndex = 3
txtnotacept.Visible = False
lblnotacept.Visible = False
'Me.ChAccept.value = xtpUnchecked
Me.CheckBox2.value = xtpUnchecked
txtresonwait.Visible = True
lblresonwaite.Visible = True
'Cmd_Click (2)
Else
txtresonwait.Visible = False
lblresonwaite.Visible = False
End If
End If
chpo = False
End Sub

Private Sub CheckBox2_Click()
If chpo = False Then
If Me.CheckBox2.value = vbChecked Then
txtnotacept.Visible = True
lblnotacept.Visible = True

Cmd_Click (1)
Me.DcbOrderStatus.ListIndex = 4
Me.CheckBox1.value = xtpUnchecked
Me.ChAccept.value = xtpUnchecked
Cmd_Click (2)
End If
End If
chpo = False


End Sub
Function Checked(Optional WorkOrder As Double = 0, Optional ShowPriceOrder As Double = 0, Optional AuthoOrder As Double = 0) As Boolean
     Checked = False
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
      Set RsDev = New ADODB.Recordset
    If WorkOrder <> 0 Then
   StrSQL = " select * from TblCardAuthoSerial where WorkOrder=" & WorkOrder & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
    If ShowPriceOrder <> 0 Then
  StrSQL = " select * from TblCardAuthoSerial where ShowPriceOrder=" & ShowPriceOrder & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If

    If AuthoOrder <> 0 Then
  StrSQL = " select * from TblCardAuthoSerial where AuthoOrder=" & AuthoOrder & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
End Function

Sub maxx(Optional ByRef WorkOrder As Double = 0, Optional ByRef ShowPriceOrder As Double = 0, Optional ByRef AuthoOrder As Double = 0)
     
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
  Set RsDev = New ADODB.Recordset
    If WorkOrder <> 0 Then
   StrSQL = " select max(WorkOrder) as mx from TblCardAuthoSerial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   WorkOrder = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "TblCardAuthoSerial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("WorkOrder").value = WorkOrder
RsDev.update
End If
'''''''''''/////
    If ShowPriceOrder <> 0 Then
   StrSQL = " select max(ShowPriceOrder) as mx from TblCardAuthoSerial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   ShowPriceOrder = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "TblCardAuthoSerial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("ShowPriceOrder").value = ShowPriceOrder
RsDev.update
End If
'''''''''''/////
    If AuthoOrder <> 0 Then
   StrSQL = " select max(AuthoOrder) as mx from TblCardAuthoSerial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   AuthoOrder = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "TblCardAuthoSerial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("AuthoOrder").value = AuthoOrder
RsDev.update
End If
End Sub
Sub FinishServ()
Dim i As Integer
 Dim b As Boolean
If TxtModFlg.text <> "N" Then
With Fg
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("finish")) = True
chektab = True
cheh b
If b = True Then
Me.DcbOrderStatus.ListIndex = 2
End If
Next i
 Cmd_Click (1)
 Cmd_Click (2)
End With
End If
End Sub
Public Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index

        Case 0
screenData = False
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            opt(2).value = True
            Me.DcbScreen.ListIndex = 0
            
            cmdEndAll.Tag = 0
            cmdEndAll.Enabled = True
            
            DcbScreen_Click
            
        imgg
        Me.RdCash.value = True
        RdPerson.value = True
        DTPicker1.value = Time
            Me.Lbtotal.Caption = 0
            Me.LbToTalExtra.Caption = 0
            
            Me.lbTotalMente.Caption = 0
     Me.DcbOrderStatus.ListIndex = 0
    Me.ComGranty.ListIndex = 1
    FG22.Rows = 1
           ' TxtPaymentCounts.text = 1
        
dcBranch.BoundText = Current_branch
            'XPDtbTrans.SetFocus
            
          '  Accredit.Enabled = True
             '   If SystemOptions.UserInterface = ArabicInterface Then
                        '                            Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                  '                                Else
                                         '           Accredit.Caption = " send to Approval   "
                            '                   End If
            chektab = False
      XPTab301.CurrTab = 0                       '
DcboFitter.BoundText = user_id ' 272727
Me.TxtCliientName.SetFocus
     vchrgrid.Clear flexClearScrollable, flexClearEverything
       vchrgrid.Rows = 2
        Case 1
        If val(cmdEndAll.Tag) = 1 Then
            MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· ⁄·Ï Â–« «·ﬂ«—  ·«‰Â „€·ﬁ ‰Â«∆Ì«"
            Exit Sub
        End If


Dim StrSQL As String
'----------------------------------------------------------------
    '·√‰‰« ﬁ„‰« »≈÷«›… Õ—ﬂ… „‰ ‰Ê⁄ „Œ ·›…
   '
   ' Set rs = New ADODB.Recordset
   ' StrSQL = "select * From TblCardAuthorizationReform     Order By ID"
   ' rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText


     
   ' Me.Retrive val(Me.XPTxtID.text)
    '----------------------------------------------------------------




'If (DcbOrderStatus.ListIndex = 5) Then
'           If SystemOptions.UserInterface = ArabicInterface Then
'              MsgBox "·« Ì„ﬂ‰  ⁄œÌ· «·”‰œ"
'            Else
'            MsgBox "·« Ì„ﬂ‰  ⁄œÌ· «·”‰œ"
'          End If
'Exit Sub
'End If
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
               If chektab = False Then
               Fg.Rows = Fg.Rows + 1
            Fg.Enabled = True
            fg2.Rows = fg2.Rows + 1
            fg2.Enabled = True
  
   
    End If
             Me.DcbScreen.ListIndex = 0
            DcbScreen_Click
            
            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id
DcboFitter.BoundText = user_id ' 272727

        Case 2
        If val(cmdEndAll.Tag) = 1 Then
            MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· ⁄·Ï Â–« «·ﬂ«—  ·«‰Â „€·ﬁ ‰Â«∆Ì«"
            Exit Sub
        End If


        DcboFitter.BoundText = user_id ' 272727
     If chektab = False Then
    XPTab301.CurrTab = 0
    Else
    XPTab301.CurrTab = 1
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
           '     SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.dcBranch.BoundText

            SaveData
          
            chektab = False

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5
       bo = True
           Load FrmCarAutoMSearch
             FrmCarAutoMSearch.show

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.text, , 200

        'Case 8
         '   CalCulateParts
            Case 15
          Dim i As Long
            AddNewFgRow
                        For i = 1 To FG22.Rows - 1
                 CalculteValueAdded i, 21
             Next
             
             
            
     
             ReLineGrid2
            Case 16
            DeleteFgRow
                   For i = 1 To FG22.Rows - 1
                 CalculteValueAdded i, 21
             Next
             
            ReLineGrid2
            
                 Case 9
                         If val(TxtWorkOrder.text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«  ÊÃœ »Ì«‰«   «„— ‘€·"
Else
MsgBox "Not Found Data of work order "
End If
Exit Sub
End If
Dim Respons As String
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
 Respons = MsgBox("·ÿ»«⁄Â »«·⁄—»Ì «Œ — „Ê«›ﬁ «Ê ·« ·ÿ»«⁄Â »«·«‰Ã·Ì“Ì", vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)
            If val(Me.XPTxtID.text) <> 0 Then
            If Respons = vbNo Then
                print_report val(Me.XPTxtID.text), 9
                Else
         print_report val(Me.XPTxtID.text), 1
        End If
            End If
        
        Case 10
        If val(TxtAuthoOrder.text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«  ÊÃœ »Ì«‰«   ≈–‰ ≈’·«Õ"
Else
MsgBox "Not Found Data "
End If
Exit Sub
End If
          If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
           If ComGranty.ListIndex = 0 Then
                print_report val(Me.XPTxtID.text), 0
                Else
                print_report val(Me.XPTxtID.text), 3
                End If
        
        
            End If
Case 11
If val(TxtShowPriceOrder.text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«  ÊÃœ »Ì«‰«  ⁄—÷ ”⁄—"
Else
MsgBox "Not Found Data of Show Price"
End If
Exit Sub
End If
 If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
                print_report val(Me.XPTxtID.text), 2
        
        
            End If
            
            
Case 17
If val(TxtShowPriceOrder.text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«  ÊÃœ »Ì«‰«  ⁄—÷ ”⁄—"
Else
MsgBox "Not Found Data of Show Price"
End If
Exit Sub
End If
 If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
                print_report val(Me.XPTxtID.text), 5
        
        
            End If
Case 18
If val(TxtShowPriceOrder.text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«  ÊÃœ »Ì«‰«  ›« Ê—… «„— «’·«Õ"
Else
MsgBox "Not Found Data of Show Price"
End If
Exit Sub
End If
 If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
                print_report val(Me.XPTxtID.text), 6
        
        
            End If
            
            
       Case 21
       
       RemoveGridRow
       Case 8
       RemoveGridRow1
       Case 12
       AuthoOrder = val(TxtAuthoOrder.text)
  If Me.Checked(, , AuthoOrder) = True Then
   Else
    AuthoOrder = 1
     maxx , , AuthoOrder
     TxtAuthoOrder.text = AuthoOrder
  End If
  Cmd_Click (1)
  Cmd_Click (2)
          Case 13
            WorkOrder = val(TxtWorkOrder.text)
  If Me.Checked(WorkOrder, 0, 0) = True Then
   Else
     WorkOrder = 1
     maxx WorkOrder
     TxtWorkOrder.text = WorkOrder
  End If
  Cmd_Click (1)
  Cmd_Click (2)
 Case 14
 FinishServ
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub RemoveGridRow()

    With Me.Fg

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub RemoveGridRow1()

    With Me.fg2

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Function print_report(Optional NoteSerial As String, Optional X As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim s As String
    
 Dim RsData2  As New ADODB.Recordset
        Dim RsData3  As New ADODB.Recordset
        Dim RsDetails1 As New ADODB.Recordset
        Dim StrSQL As String
    
 MySQL = " SELECT     dbo.TblCardAuthorizationReformDetails.Type, dbo.TblCardAuthorizationReformDetails.Mainte, dbo.TblMaintenanceWork.name, dbo.TblMaintenanceWork.namee,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.EmpID, TblEmployee_2.Emp_Name AS fiter, TblEmployee_2.Emp_Namee AS fitere,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.empsuper, TblEmployee_1.Emp_Name AS NameSuper, TblEmployee_1.Emp_Namee AS NamesuperE,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.Deptid, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.Dpeterial, dbo.TblCardAuthorizationReformDetails.DeptBr, dbo.TblCardAuthorizationReformDetails.DeptColor,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.PriceFitter, dbo.TblCardAuthorizationReformDetails.payed, dbo.TblCardAuthorizationReformDetails.allocation,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.TimOut, dbo.TblCardAuthorizationReformDetails.TimeEnter, dbo.TblCardAuthorizationReformDetails.DateExit,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.DateEnter, dbo.TblCardAuthorizationReformDetails.finish, dbo.TblCardAuthorizationReformDetails.nohours,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.bill, dbo.TblCardAuthorizationReformDetails.comp, dbo.TblCardAuthorizationReformDetails.[count],"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.[Value], dbo.TblCardAuthorizationReform.RecordDate, dbo.TblCardAuthorizationReform.ClientName,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.Telephone, dbo.TblCardAuthorizationReform.Posted, dbo.TblCardAuthorizationReform.CarTypeID,"
 MySQL = MySQL & "                     dbo.TBLCarTypes.name AS CarName, dbo.TBLCarTypes.namee AS CarNameE, dbo.TblCardAuthorizationReform.CarModelID, dbo.TblCarModels.Model,"
 MySQL = MySQL & "                     dbo.TblCarModels.ModelE, dbo.TblCardAuthorizationReform.PlateNo, dbo.TblCardAuthorizationReform.BranchID, dbo.TblBranchesData.branch_name,"
 MySQL = MySQL & "                     dbo.TblBranchesData.branch_namee, dbo.TblCardAuthorizationReform.ColorID, dbo.TblColor.name AS Color, dbo.TblColor.namee AS ColorE,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.YearFact, dbo.TblCardAuthorizationReform.OrderStatus, dbo.TblCardAuthorizationReform.Accept,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.EndDate, dbo.TblCardAuthorizationReform.subcar1, dbo.TblCardAuthorizationReform.subcar2,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.subcar3, dbo.TblCardAuthorizationReform.subcar4, dbo.TblCardAuthorizationReform.subcar5,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.subcar6, dbo.TblCardAuthorizationReform.subcar7, dbo.TblCardAuthorizationReform.subcar8,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.subcar9, dbo.TblCardAuthorizationReform.subcar10, dbo.TblCardAuthorizationReform.Month_Day,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.Granty, dbo.TblCardAuthorizationReform.DateStartG, dbo.TblCardAuthorizationReform.DateEndG,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.CarMeter, dbo.TblCardAuthorizationReform.LongGranty, dbo.TblCardAuthorizationReform.PayFirst,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.AmountAccept, dbo.TblCardAuthorizationReform.Complaint, dbo.TblCardAuthorizationReform.Noteinitial,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.Shaseh, dbo.TblCardAuthorizationReform.NotAccept, dbo.TblCardAuthorizationReform.EmpID2,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.EmpID1, dbo.TblCardAuthorizationReform.EmpID AS EmPPID, dbo.TblCardAuthorizationReform.typerequest,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.FitterID, dbo.TblUsers.UserName, dbo.TblCardAuthorizationReform.ClientCode, dbo.TblCardAuthorizationReform.mobile,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.Cash, dbo.TblCardAuthorizationReform.Accoun, dbo.TblCardAuthorizationReform.credit, dbo.TblCardAuthorizationReform.box,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.fax, dbo.TblCardAuthorizationReform.email, dbo.TblCardAuthorizationReform.address, dbo.TblCardAuthorizationReform.boxzip,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.codereg, dbo.TblCardAuthorizationReform.codedoor, dbo.TblCardAuthorizationReform.typereg,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.DateEnter AS DateEnterR, dbo.TblCardAuthorizationReform.persons, dbo.TblCardAuthorizationReform.Companies,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.driver, dbo.TblCardAuthorizationReform.DateAcutExite, dbo.TblCardAuthorizationReform.DateExptExit,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.TimeAcutExite, dbo.TblCardAuthorizationReform.TimeExptExit, dbo.TblCardAuthorizationReform.DateExit AS DateExitR,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.subcar11, dbo.TblCardAuthorizationReform.subcar12, dbo.TblCardAuthorizationReform.subcar13,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.subcar14, dbo.TblCardAuthorizationReform.ResonUnderWait, dbo.TblCardAuthorizationReform.Remarkcar,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.Payed AS PayedR, dbo.TblCardAuthorizationReform.finish AS finishR, dbo.TblCardAuthorizationReform.PrivateCop,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.ReComentClient, dbo.TblCardAuthorizationReform.wait, dbo.TblCardAuthorizationReform.notAcepted,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.NoteSerial, dbo.TblCardAuthorizationReform.CodeComputer, dbo.TblCardAuthorizationReform.ID,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.TypeCustomer, dbo.TblCardAuthorizationReform.OverKM, dbo.TblCustemers.CusName, TblCustemers.VATNO, dbo.TblCustemers.CusNamee,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.SendSMS, dbo.TblCardAuthorizationReform.TypeOrder, dbo.TblCardAuthorizationReform.WorkOrder,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.ShowPriceOrder, dbo.TblCardAuthorizationReform.AuthoOrder, dbo.TblCardAuthorizationReform.LastWorOrder,"
 MySQL = MySQL & "                     dbo.TblCustemers.Fullcode, dbo.TblCustemers.CustGID, dbo.TblCustemers.ExpireDateH, dbo.TblCardAuthorizationReform.RecordeTime,"
 MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.CarMetarOut"

 MySQL = MySQL & "                     FROM            TBLCarTypes RIGHT OUTER JOIN"
  MySQL = MySQL & "                                             TblColor RIGHT OUTER JOIN"
 MySQL = MySQL & "                                              TblCardAuthorizationReform LEFT OUTER JOIN"
 MySQL = MySQL & "                                              TblCustemers ON TblCardAuthorizationReform.ClientCode = TblCustemers.Fullcode LEFT OUTER JOIN"
 MySQL = MySQL & "                                              TblUsers ON TblCardAuthorizationReform.FitterID = TblUsers.UserID ON TblColor.Id = TblCardAuthorizationReform.ColorID LEFT OUTER JOIN"
 MySQL = MySQL & "                                              TblBranchesData ON TblCardAuthorizationReform.BranchID = TblBranchesData.branch_id LEFT OUTER JOIN"
 MySQL = MySQL & "                                              TblCarModels ON TblCardAuthorizationReform.CarModelID = TblCarModels.Id LEFT OUTER JOIN"
 MySQL = MySQL & "                                              TblEmpDepartments RIGHT OUTER JOIN"
 MySQL = MySQL & "                                              TblCardAuthorizationReformDetails LEFT OUTER JOIN"
 MySQL = MySQL & "                                              TblMaintenanceWork ON TblCardAuthorizationReformDetails.Mainte = TblMaintenanceWork.Id ON"
 MySQL = MySQL & "                                              TblEmpDepartments.DeparmentID = TblCardAuthorizationReformDetails.Deptid LEFT OUTER JOIN"
  MySQL = MySQL & "                                             TblEmployee AS TblEmployee_1 ON TblCardAuthorizationReformDetails.empsuper = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
  MySQL = MySQL & "                                             TblEmployee AS TblEmployee_2 ON TblCardAuthorizationReformDetails.EmpID = TblEmployee_2.Emp_ID ON"
  MySQL = MySQL & "                                             TblCardAuthorizationReform.ID = TblCardAuthorizationReformDetails.ID ON TBLCarTypes.id = TblCardAuthorizationReform.CarTypeID"
 MySQL = MySQL & "  Where (dbo.TblCardAuthorizationReform.id =  " & val(XPTxtID.text) & ") "
 'and (dbo.TblCardAuthorizationReformDetails.type=0)"
If X = 9 Then
If (Me.ChAccept.value = xtpChecked) Then
  If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCardAutintcation1.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCardAutintcation1.rpt"
        End If
        
Else
MsgBox "·«Ì„ﬂ‰ «·ÿ»«⁄… «„— ‘€· «·« ›Ì Õ«·… „Ê«›ﬁ… «·⁄„Ì·"
   Exit Function
End If
End If





If (X = 1) Then
    If (Me.ChAccept.value = xtpChecked) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCardAutintcation1A.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCardAutintcation1A.rpt"
        End If
    Else
        MsgBox "·«Ì„ﬂ‰ «·ÿ»«⁄… «„— ‘€· «·« ›Ì Õ«·… „Ê«›ﬁ… «·⁄„Ì·"
       Exit Function
    End If
Else
If X = 2 Or X = 5 Or X = 6 Then
   
   If X = 5 Then
        If SystemOptions.UserInterface = ArabicInterface Then
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCardAutintcationShow2.rpt"
         Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCardAutintcationShow2.rpt"
         End If
ElseIf X = 2 Then
    If SystemOptions.UserInterface = ArabicInterface Then
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCardAutintcationShow.rpt"
         Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCardAutintcationShow.rpt"
         End If
ElseIf X = 6 Then
    If SystemOptions.UserInterface = ArabicInterface Then
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCardAutintcationInvoices2.rpt"
         Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCardAutintcationInvoices2.rpt"
         End If
End If


 
            
                
    StrSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, "
    StrSQL = StrSQL & "                      dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_HijriDate, dbo.Transactions.TransactionComment, dbo.Transactions.OpOrderID,"
    StrSQL = StrSQL & "                      dbo.Transactions.OldOpOrderID, dbo.Transaction_Details.UnitId,dbo.Transaction_Details.OperPrice, dbo.Transaction_Details.ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.Item_ID,"
    StrSQL = StrSQL & "                      dbo.TblItems.itemname , dbo.TblItems.ItemNamee, dbo.TblItems.fullcode , dbo.Transaction_Details.showPrice"
    StrSQL = StrSQL & " ,ShowPrice = (SELECT Top 1 TblItemsUnits.UnitSalesPrice"
    StrSQL = StrSQL & "                 From TblItemsUnits"
    StrSQL = StrSQL & "                 Where ItemID = Transaction_Details.Item_ID"
    StrSQL = StrSQL & "                        AND UnitID           = Transaction_Details.UnitId  )"
    StrSQL = StrSQL & " FROM         dbo.TblItems RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
    StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_Type = 19) And (dbo.Transactions.RepairOrder = " & val(TxtWorkOrder.text) & ")"
        RsData2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
         

s = "Select  TblCardAuthorizationReformItems.beforeVat  ,TblCardAuthorizationReformItems.ItemName2 ,TblCardAuthorizationReformItems.qty ShowQty, IsNull(TblCardAuthorizationReformItems.DiscValue,0) as ItemDiscountValue,IsNull(TblCardAuthorizationReformItems.PriceBDisc,TblCardAuthorizationReformItems.Price)   ShowPrice,TblCardAuthorizationReformItems.TotalWithVat localprice,tblItems.ItemCode,tblItems.ItemName,tblItems.ItemNamee from TblCardAuthorizationReformItems Left Outer Join tblItems On tblItems.ItemID =TblCardAuthorizationReformItems.ItemID "
s = s & "  Where (dbo.TblCardAuthorizationReformItems.id =" & val(XPTxtID.text) & ")"
       
 RsData3.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
 End If
 
 

 
 
 If X = 3 Then
   If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCardAutintcationWithOutGranty.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCardAutintcationWithOutGranty.rpt"
        End If
 End If
 If X = 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCardAutintcationGranty.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCardAutintcationGranty.rpt"
        End If
 End If
 End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation  RepCardAutintcationShow
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
       ' xReport.ParameterFields(15).AddCurrentValue Me.DcboFitter.text
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
    
      '  xReport.ParameterFields(15).AddCurrentValue Me.DcboFitter.text
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
    If X = 2 Or X = 5 Or X = 6 Then
        xReport.OpenSubreport("Out").Database.SetDataSource RsData3
       ' xReport.OpenSubreport("RepCar").Database.SetDataSource RsData3
            Dim i As Integer
             xReport.EnableParameterPrompting = False
         For i = 1 To xReport.ParameterFields.count
             Select Case xReport.ParameterFields.Item(i).ParameterFieldName
             
            Case "TotalNet"
                xReport.ParameterFields.Item(i).AddCurrentValue "" & WriteNo(Format(val(val(LbToTalExtra.Caption) + (val(lbl(23))) + (val(Me.lbTotalMente.Caption)) * 1.05), "0.00"), 0, True, ".") & ""
            Case "TotalNet2"
                xReport.ParameterFields.Item(i).AddCurrentValue "" & (val(val(LbToTalExtra.Caption) + (val(lbl(23))) + (val(Me.lbTotalMente.Caption)) * 1.05)) & ""
                
             Case "TotalVat"
                 xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtVat2) + val(val(Me.lbTotalMente.Caption) * 0.05) & ""
             Case "DisckPercent"
                 xReport.ParameterFields.Item(i).AddCurrentValue "" & txtDiscPercent & ""
             Case "TotalPriceBeDisk"
                 xReport.ParameterFields.Item(i).AddCurrentValue lbl(31).Caption
             Case "TotalAfterDisc"
                 xReport.ParameterFields.Item(i).AddCurrentValue "" & txtTotalAfterDiscount & ""
                 
            Case "TotalHand"
                 xReport.ParameterFields.Item(i).AddCurrentValue CStr(val(Me.lbTotalMente.Caption))
            Case "VATRegNo"
                If SystemOptions.VATNoAccordActivity = False Then
                    xReport.ParameterFields(i).AddCurrentValue cCompanyInfo.VATRegNo
                Else
                    xReport.ParameterFields(i).AddCurrentValue GetRegVATNo(val(dcBranch.BoundText))
                End If
             End Select
         Next i
        
    End If
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
  Dim total As String
   Dim dif As String
  Dim totl As Double
  totl = val(LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
  total = totl
  dif = val(totl) - val(TxtAmoutAccept)
   xReport.ParameterFields(12).AddCurrentValue CStr(val(Me.lbTotalMente.Caption) * 1.05)
      xReport.ParameterFields(13).AddCurrentValue CStr(LbToTalExtra.Caption + val(lbl(23)))
        xReport.ParameterFields(14).AddCurrentValue total
        xReport.ParameterFields(15).AddCurrentValue dif
       
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



Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub


Private Sub ComGranty_Click()
If Me.ComGranty.ListIndex = 0 Then  '"»÷„«‰" Then

frmgranty.Visible = True
Else

frmgranty.Visible = False
End If

'Else
''Me.frmgranty.Visible = False
'End If
If Me.ComGranty.ListIndex = 2 Then  '"«⁄«œ… «’·«Õ" Then
FrReturnMaint.Visible = True
'Me.FrReturnMaint.Visible = True
Else
Me.FrReturnMaint.Visible = False
End If
End Sub



Private Sub Command1_Click()
codecar.Visible = True
End Sub


Private Sub Command2_Click()
codecar.Visible = False
End Sub

Private Sub Command3_Click()
    XPTab301.CurrTab = 0
    XPTab301.left = 0
    XPTab301.Width = Me.Width
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

Private Sub DcbCarModel_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub DcbCarType_Change()
Dim Dcombos As ClsDataCombos
      Set Dcombos = New ClsDataCombos
    
      If val(Me.DcbCarType.BoundText) <> 0 Then
      
   Dcombos.GetTblCarModels Me.DcbCarModel, , val(Me.DcbCarType.BoundText)
   End If
End Sub

Private Sub DcbCarType_Click(Area As Integer)
Dim Dcombos As ClsDataCombos
      Set Dcombos = New ClsDataCombos
    
      If val(Me.DcbCarType.BoundText) <> 0 Then
      
   Dcombos.GetTblCarModels Me.DcbCarModel, , val(Me.DcbCarType.BoundText)
   End If
End Sub





Private Sub DcbCarType_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

 Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub DcbColor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub



Private Sub DcboFitter_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Sub che()

If Me.DcbOrderStatus.ListIndex = 0 Then
Me.ChAccept.value = xtpUnchecked
Me.CheckBox1.value = xtpUnchecked
Me.CheckBox2.value = xtpUnchecked
End If
If Me.DcbOrderStatus.ListIndex = 2 Then
'Me.ChAccept.value = xtpUnchecked
Me.CheckBox1.value = xtpUnchecked
Me.CheckBox2.value = xtpUnchecked
End If
If Me.DcbOrderStatus.ListIndex = 5 Or Me.DcbOrderStatus.ListIndex = 6 Then
'Me.ChAccept.value = xtpUnchecked
Me.CheckBox1.value = xtpUnchecked
Me.CheckBox2.value = xtpUnchecked
End If
If Me.DcbOrderStatus.ListIndex = 1 Then
Me.ChAccept.value = xtpChecked
'Me.CheckBox1.value = xtpUnchecked
Me.CheckBox2.value = xtpUnchecked

End If
If Me.DcbOrderStatus.ListIndex = 3 Then
'Me.ChAccept.value = xtpUnchecked
Me.CheckBox1.value = xtpChecked

Me.CheckBox2.value = xtpUnchecked

End If
If Me.DcbOrderStatus.ListIndex = 4 Then
Me.ChAccept.value = xtpUnchecked
Me.CheckBox1.value = xtpUnchecked
Me.CheckBox2.value = xtpChecked

End If
End Sub



Private Sub DcbScreen_Change()
DcbScreen_Click
End Sub

Private Sub DcbScreen_Click()
Dim StrSQL As String
Set rs = New ADODB.Recordset
     If Me.DcbScreen.ListIndex = 1 Then
     
        StrSQL = "select * From TblCardAuthorizationReform     Order By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    Else
        StrSQL = "select * From TblCardAuthorizationReform   where id =" & val(Me.XPTxtID.text) & " "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    End If
     
End Sub


Private Sub DcbyearFactor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub DTPEnterDate_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Public Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
Dim StrAccountCode1 As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim k As Integer
Dim StrComboList As String
            
    
    With Fg

        Select Case .ColKey(Col)
              Case "supervisor"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("empsuper"), False, True)
                .TextMatrix(Row, .ColIndex("empsuper")) = StrAccountCode
                 'StrSQL = " SELECT     dbo.Technicians1.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.Technicians1.Emp_ID1, "
    'StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_ID AS Expr1, dbo.SuperTech.ID, dbo.SuperTech.DeparmentID"
 'StrSQL = StrSQL & " FROM         dbo.Technicians1 INNER JOIN"
 'StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Technicians1.Emp_ID = dbo.TblEmployee.Emp_ID INNER JOIN"
 '  StrSQL = StrSQL & "                    dbo.SuperTech ON dbo.Technicians1.DeparmentID = dbo.SuperTech.ID"
'StrSQL = StrSQL & " Where (dbo.Technicians1.Emp_id1 = " & val(StrAccountCode) & ") And (dbo.SuperTech.DeparmentID =" & val(.TextMatrix(Row, .ColIndex("Deptid"))) & ")"
' StrSQL = StrSQL & " Where (dbo.Technicians1.DeparmentID =" & val(.TextMatrix(Row, .ColIndex("Deptid"))) & " ) And (dbo.Technicians1.Emp_id1 = " & val(StrAccountCode) & ")"
'StrSQL = " SELECT     dbo.Technicians1.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.Technicians1.Emp_ID1,"
'StrSQL = StrSQL & "                      dbo.SuperTech.id , dbo.SuperTech.DeparmentID"
'StrSQL = StrSQL & " FROM         dbo.Technicians1 INNER JOIN"
'StrSQL = StrSQL & "                      dbo.SuperTech ON dbo.Technicians1.DeparmentID = dbo.SuperTech.ID INNER JOIN"
'StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Technicians1.Emp_ID1 = dbo.TblEmployee.Emp_ID"
'StrSQL = StrSQL & " Where (dbo.SuperTech.DeparmentID =" & val(.TextMatrix(Row, .ColIndex("Deptid"))) & ") And (dbo.Technicians1.Emp_id =" & val(StrAccountCode) & ")"
'
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                  If rs.RecordCount > 0 Then
'                  If SystemOptions.UserInterface = ArabicInterface Then
'                    .TextMatrix(Row, .ColIndex("fitter")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
'                Else
'                    .TextMatrix(Row, .ColIndex("fitter")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
'                End If
'                ' If SystemOptions.UserInterface = ArabicInterface Then
                '    StrComboList = Fg.BuildComboList(rs, "Emp_Name", "Emp_ID")
                'Else
                '    StrComboList = Fg.BuildComboList(rs, "Emp_Namee", "Emp_ID")
               ' End If
'
'                If StrComboList <> "" Then
'                    StrComboList = "|" & StrComboList
               'End If
'                 .ComboList = StrComboList
                
                '   StrAccountCode = .ComboData
                '   LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("EmpID"), False, True)
                '.TextMatrix(Row, .ColIndex("EmpID")) = StrAccountCode
'                End If
             Case "fitter"
        
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("EmpID"), False, True)
                .TextMatrix(Row, .ColIndex("EmpID")) = StrAccountCode
                
        Case "workshop"
        Dim s As String
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Deptid"), False, True)
                .TextMatrix(Row, .ColIndex("Deptid")) = StrAccountCode
                 StrSQL = "select * from TblEmpDepartments where DeparmentID =" & val(StrAccountCode)
             
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                 s = IIf(IsNull(rs.Fields("Dpeterial").value), "", rs.Fields("Dpeterial").value)
  
    If s <> "" Then
    s = val(s) + 1
   Else
   s = 0
    
    End If
                    .TextMatrix(Row, .ColIndex("Dpeterial")) = s 'IIf(IsNull(rs("Dpeterial").value), 0, rs("Dpeterial").value)
                
                    .TextMatrix(Row, .ColIndex("DeptColor")) = IIf(IsNull(rs("DeptColor").value), 0, rs("DeptColor").value)
                End If
            Case "name"
               
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("cod"), False, True)
                .TextMatrix(Row, .ColIndex("cod")) = StrAccountCode
                StrSQL = "select * from TblMaintenanceWork where Id=" & val(StrAccountCode)
             
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("value")) = IIf(IsNull(rs("InitialPrice").value), 0, rs("InitialPrice").value)
                    Dim HDWM As Double
                HDWM = val(IIf(IsNull(rs("HDWM").value), 0, rs("HDWM").value))
                Else
                    .TextMatrix(Row, .ColIndex("value")) = ""
                End If
                
                k = 0
                
                Select Case HDWM
                Case 0
                k = HDWM
                Case 1
                k = HDWM * 24
                Case 2
                k = HDWM * 24 * 7
                Case 3
                k = HDWM * 24 * 30
                End Select
                .TextMatrix(Row, .ColIndex("dateenter")) = Date
                .TextMatrix(Row, .ColIndex("timEnter")) = Time
                .TextMatrix(Row, .ColIndex("nohours")) = k
    ' .TextMatrix(Row, .ColIndex("finish")) = 0
                                  Case "finish"
             
                                  If TxtModFlg.text <> "N" Then
                                    chektab = True
                                    Dim b As Boolean
cheh b
If b = True Then

Me.DcbOrderStatus.ListIndex = 2
End If

                          Cmd_Click (1)
                          Cmd_Click (2)
                           End If
                           'End If
'SaveData1end if
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

      
        Select Case .ColKey(Col)
            
            Case "cod"
               Cancel = True
            Case "value"
            
               Fg.ComboList = ""
               Case "count"
               Fg.ComboList = ""
               Case "nohours"
               Fg.ComboList = ""
             '    Cancel = True
              Case "dateout"
              Cancel = True
              Case "TimOut"
              Cancel = True
               Case "PriceFitter"
    Fg.ComboList = ""
               'Fg.ComboList = ""
               Case "finish"
               Fg.ComboList = ""
               
        End Select

    End With

    
End Sub

Private Sub FG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
   Dim LngItemID As Long
    Dim LngStoreID As Long
    Dim rdate As Date
  ' Dim frm As FrmGridAddItemComment
    Dim Frm1 As FromRegisterDateTime

    'On Error GoTo ErrTrap

    With Me.Fg

        Select Case .ColKey(Col)

                 Case "dateenter"
                  LngRow = Row

 LngCol = Col
             ' ItemProductionDate Row, Col, , 1
                Load FromRegisterDateTime
                FromRegisterDateTime.show
                  Case "timEnter"
                                    LngRow = Row

 LngCol = Col
                  Load FromRegisterDateTime
                FromRegisterDateTime.show
                    
                End Select
                End With
End Sub







Private Sub fg_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.text <> "R" Then

With Fg
If KeyCode = vbKeyF3 Then
Select Case .ColKey(.Col)
Case "name"
 Unload FrmBillCarMaintExtrSearch
 FrmBillCarMaintExtrSearch.IndTyp = 2
           FrmBillCarMaintExtrSearch.Row = .Row
  Load FrmBillCarMaintExtrSearch
           FrmBillCarMaintExtrSearch.IndTyp = 2
           FrmBillCarMaintExtrSearch.Row = .Row
            FrmBillCarMaintExtrSearch.show vbModal
End Select
End If
End With
End If

'If KeyCode = vbKeyF3 Then

'Load FrmMaintnSearch
'            FrmMaintnSearch.show
'
'
'End If
End Sub

Private Sub fg_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If Me.TxtModFlg.text <> "R" Then

With Fg
If KeyCode = vbKeyF3 Then
Select Case .ColKey(.Col)
Case "name"
 Unload FrmBillCarMaintExtrSearch
 FrmBillCarMaintExtrSearch.IndTyp = 2
           FrmBillCarMaintExtrSearch.Row = .Row
 Load FrmBillCarMaintExtrSearch
           FrmBillCarMaintExtrSearch.IndTyp = 2
           FrmBillCarMaintExtrSearch.Row = .Row
            FrmBillCarMaintExtrSearch.show vbModal
End Select
End If
End With
End If
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
'Dim StrComboList As String
    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg

        Select Case .ColKey(Col)
 Case "workshop"
            StrSQL = " SELECT    DeparmentID, DepartmentName ,DepartmentNamee,Dpeterial "
            StrSQL = StrSQL & "  FROM         TblEmpDepartments"
            StrSQL = StrSQL & "                    where not (Dpeterial is null)"
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg.BuildComboList(rs, "DepartmentName", "DeparmentID")
                Else
                    StrComboList = Fg.BuildComboList(rs, "DepartmentNamee", "DeparmentID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Deptid"), False, True)
                .TextMatrix(Row, .ColIndex("name")) = ""
            Case "name"
            If val(.TextMatrix(Row, .ColIndex("Deptid"))) = 0 Then
                   StrSQL = "select * from TblMaintenanceWork where DeptID =-1 "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = Fg.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
            MsgBox "Ì—ÃÏ «Œ Ì«— «·ﬁ”„ «Ê·«"
            Exit Sub
            Else
            
                StrSQL = "select * from TblMaintenanceWork where DeptID =" & val(.TextMatrix(Row, .ColIndex("Deptid"))) & " "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = Fg.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
            End If
    Case "fitter"

                  If Fg.TextMatrix(Row, Fg.ColIndex("workshop")) = "" Then
  MsgBox "ÌÃ» «Œ Ì«— «·ﬁ”„ «Ê·«"
  Exit Sub
  Else
                 If Fg.TextMatrix(Row, Fg.ColIndex("supervisor")) = "" Then
  MsgBox "ÌÃ» «Œ Ì«— «·„‘—› «Ê·«"
  Exit Sub
  Else
  
StrSQL = " SELECT     dbo.Technicians1.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.Technicians1.Emp_ID1,"
StrSQL = StrSQL & "                       dbo.SuperTech.id , dbo.SuperTech.DeparmentID"
StrSQL = StrSQL & "  FROM         dbo.Technicians1 INNER JOIN"
 StrSQL = StrSQL & "                      dbo.SuperTech ON dbo.Technicians1.DeparmentID = dbo.SuperTech.ID INNER JOIN"
 StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Technicians1.Emp_ID1 = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & "  Where (dbo.SuperTech.DeparmentID =" & val(Fg.TextMatrix(Row, Fg.ColIndex("Deptid"))) & ") And (dbo.Technicians1.Emp_id =" & val(Fg.TextMatrix(Row, Fg.ColIndex("empsuper"))) & ")"
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg.BuildComboList(rs, "Emp_Name", "Emp_ID1")
                Else
                    StrComboList = Fg.BuildComboList(rs, "Emp_Namee", "Emp_ID1")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                
                   StrAccountCode = .ComboData
               LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("EmpID"), False, True)
               .TextMatrix(Row, .ColIndex("EmpID")) = StrAccountCode
End If
End If
                  Case "supervisor"
                   If Fg.TextMatrix(Row, Fg.ColIndex("workshop")) = "" Then
  MsgBox "ÌÃ» «Œ Ì«— «·ﬁ”„ «Ê·«"
  Exit Sub
  Else
        '  StrSQL = " SELECT     dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TBLSalesRepData3.GroupID, dbo.TblEmployee.Emp_ID"
'StrSQL = StrSQL & " FROM         dbo.TBLSalesRepData3 LEFT OUTER JOIN"
' StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.TBLSalesRepData3.EmpID = dbo.TblEmployee.Emp_ID"
'StrSQL = StrSQL & " Where (dbo.TBLSalesRepData3.GroupID = 2)"
            StrSQL = " SELECT DISTINCT "
     StrSQL = StrSQL & "                 dbo.Technicians1.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_ID AS Expr1,"
      StrSQL = StrSQL & "                dbo.SuperTech.id , dbo.SuperTech.DeparmentID"
StrSQL = StrSQL & " FROM         dbo.Technicians1 INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Technicians1.Emp_ID = dbo.TblEmployee.Emp_ID INNER JOIN"
 StrSQL = StrSQL & "                     dbo.SuperTech ON dbo.Technicians1.DeparmentID = dbo.SuperTech.ID"
'Where (dbo.SuperTech.DeparmentID = 16)
StrSQL = StrSQL & " Where (dbo.SuperTech.DeparmentID =" & val(.TextMatrix(Row, .ColIndex("Deptid"))) & ")"
' StrSQL = StrSQL & " Where (dbo.Technicians1.DeparmentID =" & val(.TextMatrix(Row, .ColIndex("Deptid"))) & " ) And (dbo.Technicians1.Emp_id1 = " & val(StrAccountCode) & ")"
             
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If rs.RecordCount > 0 Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg.BuildComboList(rs, "Emp_Name", "Emp_ID")
                Else
                    StrComboList = Fg.BuildComboList(rs, "Emp_Namee", "Emp_ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 End If
             ' LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("empsuper"), False, True)
                ' MsgBox LngRow
                End If
                 .ComboList = StrComboList
                
             Case "dateenter"
            .ColComboList(.ColIndex("dateenter")) = "..."
               Case "timEnter"
            .ColComboList(.ColIndex("timEnter")) = "..."
 
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
 Case "typeexpen"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Codtype"), False, True)
                .TextMatrix(Row, .ColIndex("Codtype")) = StrAccountCode
            Case "name"
         '.TextMatrix(Row, .ColIndex("userid")) = user_id
       
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("cod"), False, True)
             .TextMatrix(Row, .ColIndex("cod")) = StrAccountCode
       '         StrSQL = "select * from TblExtraExpeneses where Id=" & val(StrAccountCode)
       '         rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
       '          If rs.RecordCount > 0 Then
       '             .TextMatrix(Row, .ColIndex("typeexpen")) = IIf(IsNull(rs("TypeExtrExpen").value), 0, rs("TypeExtrExpen").value)
       '         Else
          '          .TextMatrix(Row, .ColIndex("typeexpen")) = ""
          '      End If
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

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
        Select Case .ColKey(Col)
            
            Case "cod"
               Cancel = True
                Case "typeexpen"
               
            Case "value"
                fg2.ComboList = ""
                 Case "count"
                fg2.ComboList = ""
                 Case "comp"
                fg2.ComboList = ""
                 Case "bill"
                fg2.ComboList = ""
        End Select

    End With

    fg2.ComboList = ""
End Sub



Private Sub FG2_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim StrAccountCode As String
    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With fg2

        Select Case .ColKey(Col)
Case "typeexpen"
  StrSQL = "select * from TblTypeExtraExpeneses"
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
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Codtype"), False, True)
            Case "name"
            If .TextMatrix(Row, .ColIndex("typeexpen")) = "" Then
            MsgBox "ÌÃ» √Œ Ì«— «·‰Ê⁄ «·«Ê·"
            Exit Sub
            Else
                StrSQL = "select * from TblExtraExpeneses where TypeID =" & val(.TextMatrix(Row, .ColIndex("Codtype"))) & " "
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
                End If
        End Select

    End With

End Sub



Private Sub imag1_Click()

If Me.imag1.Picture = Me.imgnul.Picture Then
Me.imag1.Picture = Me.Img.Picture
Else
 Me.imag1.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub imag2_Click()

If Me.imag2.Picture = Me.imgnul.Picture Then
Me.imag2.Picture = Me.Img.Picture
Else
 Me.imag2.Picture = Me.imgnul.Picture
 End If
End Sub



Private Sub imag3_Click()

If Me.imag3.Picture = Me.imgnul.Picture Then
Me.imag3.Picture = Me.Img.Picture
Else
 Me.imag3.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub imag4_Click()

If Me.imag4.Picture = Me.imgnul.Picture Then
Me.imag4.Picture = Me.Img.Picture
Else
 Me.imag4.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub imag5_Click()

If Me.imag5.Picture = Me.imgnul.Picture Then
Me.imag5.Picture = Me.Img.Picture
Else
 Me.imag5.Picture = Me.imgnul.Picture
 End If
End Sub







Private Sub img10_Click()
If Me.img10.Picture = Me.imgnul.Picture Then
Me.img10.Picture = Me.Img.Picture
Else
 Me.img10.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub img11_Click()
If Me.img11.Picture = Me.imgnul.Picture Then
Me.img11.Picture = Me.Img.Picture
Else
 Me.img11.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub img12_Click()
If Me.img12.Picture = Me.imgnul.Picture Then
Me.img12.Picture = Me.Img.Picture
Else
 Me.img12.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub img13_Click()
If Me.img13.Picture = Me.imgnul.Picture Then
Me.img13.Picture = Me.Img.Picture
Else
 Me.img13.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub img14_Click()
If Me.img14.Picture = Me.imgnul.Picture Then
Me.img14.Picture = Me.Img.Picture
Else
 Me.img14.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub img6_Click()

If Me.img6.Picture = Me.imgnul.Picture Then
Me.img6.Picture = Me.Img.Picture
Else
 Me.img6.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub img7_Click()

If Me.img7.Picture = Me.imgnul.Picture Then
Me.img7.Picture = Me.Img.Picture
Else
 Me.img7.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub img8_Click()

If Me.img8.Picture = Me.imgnul.Picture Then
Me.img8.Picture = Me.Img.Picture
Else
 Me.img8.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub img9_Click()

If Me.img9.Picture = Me.imgnul.Picture Then
Me.img9.Picture = Me.Img.Picture
Else
 Me.img9.Picture = Me.imgnul.Picture
 End If
End Sub







Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub Lbtotal_Change()
lbldif.Caption = val(Me.Lbtotal.Caption) - val(firstprice.Caption)
End Sub



Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub menue_Click(Index As Integer)
If Index = 12 Then
    Load FrmCustemers
    FrmCustemers.FormNamee = "FrmCarAuthontication"
    FrmCustemers.show

Else
    showsforms Index
End If

End Sub

Private Sub menuet_Click()
Load FrmCarShowRecomen
FrmCarShowRecomen.show

End Sub


Private Sub Timer1_Timer()
Dim StrSQL As String
If Me.TxtModFlg.text = "R" Then
If DcbScreen.ListIndex <> 1 Then
   If rs.State = adStateOpen Then
   rs.Close

   
   End If
    StrSQL = "select * From TblCardAuthorizationReform   where id ='" & Me.XPTxtID.text & " '"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
Me.Retrive val(Me.XPTxtID.text)
End If
End If
End Sub

Private Sub txtAddres_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub TxtAmoutAccept_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub TxtBox_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub txtboxzip_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub TXtCarMeter_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub
Public Sub retInfoCustomer(Optional Fullcode2 As String)
 
'   If mFromCustomerForm Then
'            Cmd_Click (0)
'       End If
 Dim EmpID As Integer
Dim Name As String
Dim Mobile As String
Dim phone As String
Dim boxmail As String
Dim fax As String
Dim mail As String
Dim adress As String
Dim ZipCode As String
Dim DigCus As String
    Dim Fullcode As String
    Dim CusID As Integer
    If Fullcode2 <> "" Then
         GetCustomerIDFromCode TxtClientCode.text, CusID, , Fullcode, , Name, Mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus    ', CusID
         Else
         GetCustomerIDFromCode TxtClientCode.text, CusID, , Fullcode, Me.TxtCliientName, Name, Mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus    ', CusID
     End If
         Me.TxtClientCode = Fullcode
        TxtCliientName = Name
        Me.TxtMobile.text = Mobile
        Me.TxtClientPhone.text = phone
        Me.TxtBox.text = boxmail
        Me.TxtFax.text = fax
        Me.TxtEmail.text = mail
        Me.TxtAddres.text = adress
        Me.txtboxzip.text = ZipCode
        TxtCusID.text = CusID
        Me.TxtTypeCustomer.text = val(DigCus) + 1
       ' DcboEmpName.BoundText = EmpID
       If SystemOptions.LinkCustomerWithCars = True Then
       Dim Dcombos As ClsDataCombos
       Set Dcombos = New ClsDataCombos
       Dcombos.GetCarsOfCustomer DcbCar, val(TxtCusID.text)
       DcbCar.BoundText = GetFirstCarOfCustomer(val(TxtCusID.text))
       
       
       End If
    
       mFromCustomerForm = False
      
End Sub
Function GetFirstCarOfCustomer(Optional CusID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Dim Rs3 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT Max(id) as MinID From TblCusCar where CustomerID =" & CusID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
    
    GetFirstCarOfCustomer = IIf(IsNull(rs2("MinID").value), 0, rs2("MinID").value)
    sql = " SELECT ChasisNo,ModelID,BrandID,CarModelID,ColorID FROM TblCusCar Where Id = " & val(rs2!MinId & "")
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs3.EOF Then
        MsgBox "Â–« «·⁄„Ì· €Ì— „ÊÃÊœ"
        GetFirstCarOfCustomer = 0
        Exit Function
    End If
    DcbCarType.BoundText = val(Rs3!BrandID & "")
    DcbyearFactor.ListIndex = val(Rs3!ModelID & "")
    DcbCarModel.BoundText = val(Rs3!CarModelID & "")
    TXtShaseh = (Rs3!ChasisNo & "")
    DcbColor.BoundText = val(Rs3!ColorID & "")
Else
GetFirstCarOfCustomer = 0
End If
End Function

Private Sub TxtClientCode_Change()
'If Me.TxtModFlg.Text <> "R" Then
'RetInfoCustomer
'End If
End Sub

Private Sub TxtClientCode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.TxtCliientName.text = ""
retInfoCustomer
End If
End Sub

Private Sub TxtClientCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False

  Load FrmFilCustomerSearch
 FrmFilCustomerSearch.show
            
End If
End Sub

Private Sub TxtClientPhone_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
            
End Sub

Public Sub TxtCliientName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.TxtClientCode.text = ""
retInfoCustomer
End If
End Sub

Private Sub TxtCliientName_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False

' Load FrmFilCustomerSearch
'            FrmFilCustomerSearch.Show
            
End If
End Sub

Private Sub TxtCodeDoor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False
Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub txtCodeReg_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False
Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub TxtComplaint_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False
Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub TxtDriver_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False
Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub txtEmail_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False
Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub



Private Sub txtFax_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False
Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub TxtFirstPrice_Change()
Me.Lbtotal.Caption = val(Me.LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption) - val(Me.TxtFirstPrice.text)
firstprice.Caption = TxtFirstPrice.text
End Sub

Private Sub TxtFirstPrice_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False
Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub TxtItemCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

      Load FrmItemSearch
            FrmItemSearch.RetrunType = 310
            FrmItemSearch.show vbModal
End If


End Sub

Private Sub TxtItemCode_Validate(Cancel As Boolean)
    TxtItemPrice = GetItemPrice(val(Me.DcboItems.BoundText))
    
End Sub

Private Sub TxtItemPrice_Change()
txtTotal.text = val(TxtItemPrice.text) * val(Txtqty.text)
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

Private Sub txtmobile_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False
Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub TxtNoteIntial1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False
Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub TxtOrder_Change()
If Me.TxtOrder.text <> "" And Me.TxtOrder.text <> "TxtOrder" Then
FrmBillCarMaintExtra.Ch = False

Me.Retrive2 (GetID(val(Me.TxtOrder.text)))
'Load FrmCarAuthSearch
'            FrmCarAuthSearch.show
    Me.TxtAmoutAccept.text = 0
    Me.TxtFirstPrice.text = 0
    Me.TXtCarMeter.text = ""
    Me.DcbOrderStatus.ListIndex = 0
ComGranty.ListIndex = 2

End If
End Sub

  'DB_updateField "foxy", "id1 ", "nvarchar(255) not null  "
' KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)

Private Sub TxtOrder_KeyPress(KeyAscii As Integer)

KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOrder.text, 1)



End Sub

Private Sub TxtOrder_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
'Me.TxtOrder.text = ""
Load FrmCarAuthSearch
            FrmCarAuthSearch.show
            FrmBillCarMaintExtra.Ch = False
End If
End Sub

Private Sub TxtPlatNo_Change()
If Me.TxtModFlg.text = "N" Then
If Me.TxtPlatNo.text <> "" Then
TxtLastWorOrder.text = GetLastWorkOrder()
End If
End If
End Sub

Private Sub TxtPlatNo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False
Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub TxtQty_Change()
txtTotal.text = val(TxtItemPrice.text) * val(Txtqty.text)
End Sub

Private Sub txtSalesInvoiceOrder_Change()
If txtSalesInvoiceOrder = "" Then Exit Sub
Dim s As String
Dim rsDummy As New ADODB.Recordset
s = "SELECT td.GranteeType,td.GranteeStartDate,td.GranteeEndDate,td.guaranteeTime, * FROM Transaction_Details td where Transaction_Id In "
s = s & " (Select Top 1 transactions.Transaction_Id from transactions where transactions.NoteSerial1 = N'" & Trim(txtSalesInvoiceOrder) & "')"
s = s & " and Item_Id = " & val(cmbItems.BoundText)


rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
If Not rsDummy.EOF Then
    DateStartG.value = IIf(IsNull(rsDummy("GranteeStartDate").value), Date, rsDummy("GranteeStartDate").value)
    
    DateEndg.value = IIf(IsNull(rsDummy("GranteeEndDate").value), Date, rsDummy!GranteeEndDate & "")
    TxtLongGranty.text = rsDummy!guaranteeTime & ""
    frmgranty.Visible = True
    ComGranty.ListIndex = 0
Else
    ComGranty.ListIndex = 1
End If
End Sub

Private Sub txtSalesInvoiceOrder_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim m_FrmSearch  As New FrmBuySearch
           Set m_FrmSearch = New FrmBuySearch
                 m_FrmSearch.DealingForm = InvoiceTransaction
                 m_FrmSearch.Index = 310
                 m_FrmSearch.mmItemId = val(cmbItems.BoundText)
                
                                 If SystemOptions.UserInterface = ArabicInterface Then
                    m_FrmSearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄"
                 Else
               m_FrmSearch.Caption = "Search About Sales Invoice"
                 End If
                
                
                 Set m_FrmSearch.RetrunFrm = Me
                 m_FrmSearch.show vbModal
        
End Sub

Private Sub txtSalesInvoiceOrder_Validate(Cancel As Boolean)
'Dim s As String
'Dim rsDummy As New ADODB.Recordset
's = "SELECT td.GranteeType,td.GranteeStartDate,td.GranteeEndDate,td.guaranteeTime, * FROM Transaction_Details td where Transaction_Id In "
's = s & " (Select Top 1 transactions.Transaction_Id from transactions where transactions.NoteSerial1 = N'" & Trim(txtSalesInvoiceOrder) & "')"
's = s & " and Item_Id = " & val(cmbItems.BoundText)
'
'rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
'If Not rsDummy.EOF Then
'    DateStartG.value = rsDummy!GranteeStartDate & ""
'    DateEndg.value = rsDummy!GranteeEndDate & ""
'    TxtLongGranty.Text = rsDummy!guaranteeTime & ""
'
'End If
End Sub

Private Sub TXtShaseh_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False
Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
            
End Sub

Private Sub TxtTtpeReg_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False
Load FrmCarAutoMSearch
            FrmCarAutoMSearch.show
            
End If
End Sub

Private Sub TxtTypeCustomer_Change()
Dim i As Integer
Me.lbltypecus.Caption = " "
For i = 1 To val(Me.TxtTypeCustomer.text)
lbltypecus.Caption = lbltypecus.Caption + "*" + " "
Next i
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
Function newret()
  Dim RsDetails1 As New ADODB.Recordset
Dim StrSQL As String
Dim i As Integer
vchrgrid.Clear flexClearScrollable, flexClearEverything
            vchrgrid.Rows = 2
            
            
StrSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, "
StrSQL = StrSQL & "                      dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_HijriDate, dbo.Transactions.TransactionComment, dbo.Transactions.OpOrderID,"
StrSQL = StrSQL & "                      dbo.Transactions.OldOpOrderID, dbo.Transaction_Details.UnitId,dbo.Transaction_Details.OperPrice, dbo.Transaction_Details.ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.Item_ID,"
StrSQL = StrSQL & "                      dbo.TblItems.itemname , dbo.TblItems.ItemNamee, dbo.TblItems.fullcode , dbo.Transaction_Details.showPrice"
StrSQL = StrSQL & " FROM         dbo.TblItems RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_Type = 19) And (dbo.Transactions.RepairOrder = " & val(TxtWorkOrder.text) & ")"
    RsDetails1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    

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
            .TextMatrix(i, .ColIndex("OperPrice")) = GetItemPrice(val(.TextMatrix(i, .ColIndex("Item_ID"))), 1, IIf(IsNull(RsDetails1("UnitId").value), 0, RsDetails1("UnitId").value))  '''(IIf(IsNull(RsDetails1("ShowPrice").value), 0, RsDetails1("ShowPrice").value))
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

Private Sub vchrgrid_Click()
    With vchrgrid

        Select Case .Col
            Case 10
If val(.TextMatrix(.Row, .ColIndex("Transaction_ID"))) = 0 Then Exit Sub
           If checkApility("FrmOut") = False Then
                        Exit Sub
                    End If
               
               FrmOut.Retrive val(.TextMatrix(.Row, .ColIndex("Transaction_ID")))

         
        End Select

    End With
End Sub

Private Sub XPDtbTrans_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    CalculteValueAdded 1, 21
End Sub

Private Sub Dcbranch_Click(Area As Integer)
 
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub
Function GetLastWorkOrder() As Double
Dim Rs8 As ADODB.Recordset
Dim sql As String
Set Rs8 = New ADODB.Recordset
sql = " SELECT     MAX(WorkOrder) AS MaxOrder"
sql = sql & " From dbo.TblCardAuthorizationReform"
sql = sql & " WHERE  (PlateNo = N'" & TxtPlatNo.text & "')"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GetLastWorkOrder = IIf(IsNull(Rs8("MaxOrder").value), 0, Rs8("MaxOrder").value)
Else
GetLastWorkOrder = 0
End If
End Function
Private Sub Form_Load()
DcbOrderStatus.Enabled = False
gimage.Visible = False
codecar.Visible = False
chektab = False
screenData = False
menue(12).Visible = False
txtnotacept.Visible = False
txtresonwait.Visible = False
lblresonwaite.Visible = False
'MKDataGrid1.DateControl 2, True
'XPDtbtimeTrans.value = Time
 lblnotacept.Visible = False
DTPTimeExptExit.value = Time
If SystemOptions.LinkCustomerWithCars = True Then
TxtPlatNo.Visible = False
DcbCar.Visible = True
Else
TxtPlatNo.Visible = True
DcbCar.Visible = False


End If

If SystemOptions.IsMaintItemMode Then
          Dim LOcalCBO As String
           LOcalCBO = " SELECT     dbo.TblItems.ItemID, dbo.TblItems.ItemName  FROM         dbo.Groups INNER JOIN dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID  "
           'Where ( Groups.BranchId = " & my_branch & " )"
            fill_combo cmbItems, LOcalCBO
            Frame2.Visible = True
Else
    Frame2.Visible = False
End If
        
If SystemOptions.CanOpenWorkOrder = True Then
    cmdOpenCard.Visible = True
Else
    cmdOpenCard.Visible = False
End If

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    Dim Dcombos As ClsDataCombos
      
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

  '  With Me.Fg
  '      .RowHeightMin = 300
  '      .WallPaper = GrdBack.Picture
  '      .AutoSize 0, .Cols - 1, False
  '  End With
  
  
  If SystemOptions.UserInterface = EnglishInterface Then
        Me.ComGranty.AddItem "Granty"
        Me.ComGranty.AddItem "With out Granty"
        Me.ComGranty.AddItem "Re Maintenance"
        Me.DcbOrderStatus.AddItem "New"
        Me.DcbOrderStatus.AddItem "Accept Customer"
        Me.DcbOrderStatus.AddItem "Final Maintenance"
        Me.DcbOrderStatus.AddItem "Under wait"
        Me.DcbOrderStatus.AddItem " Customer Not Accept"
        DcbOrderStatus.AddItem "Been issued bill"
        Me.ComMD.AddItem "Month"
        Me.ComMD.AddItem "Day"
        DcbScreen.AddItem "Data Entry"
        DcbScreen.AddItem "Show Data"
     Else
        DcbScreen.AddItem "«œÕ«· »Ì«‰« "
        DcbScreen.AddItem "«” ⁄—«÷ «·»Ì«‰« "
        Me.ComGranty.AddItem "»÷„«‰"
        Me.ComGranty.AddItem "»œÊ‰ ÷„«‰"
        Me.ComGranty.AddItem "≈⁄«œ… «’·«Õ"
        DcbOrderStatus.AddItem "ÃœÌœ"
        DcbOrderStatus.AddItem " „ „Ê«›ﬁ… «·⁄„Ì·"
        DcbOrderStatus.AddItem " „ «‰Â«¡ «·«’·«Õ"
        DcbOrderStatus.AddItem " Õ  «·«‰ Ÿ«—"
        DcbOrderStatus.AddItem "⁄œ„ „Ê«›ﬁ… «·⁄„Ì·"
        DcbOrderStatus.AddItem " „ «’œ«— ›« Ê—…"
        Me.ComMD.AddItem "‘Â—"
        Me.ComMD.AddItem "ÌÊ„"
    End If
    
DcbOrderStatus.AddItem ""
  
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
  DcbOrderStatus.AddItem ""
    
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
    Dcombos.GetUsers Me.DcboFitter
    
        Dcombos.GetItemsUnits Me.dcItemunit
    
    
    Dcombos.GetItemsNames Me.DcboItems
  'Dcombos.GetTblYearFact Me.DcbyearFactor
 
DcbScreen.ListIndex = 0
Dim year As Integer
 
    Dcombos.GetTblCarsDataGroup Me.DcbCarType
    Dcombos.GetTblColor Me.DcbColor
    Dim i As Integer
      For i = 1900 To 2100
      Me.DcbyearFactor.AddItem (i)
      Next i
      
   Dcombos.GetTblCarModels Me.DcbCarModel
   
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
   ' Me.XPDtbtimeTrans.value = sysdate
   ' YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblCardAuthorizationReform     where 1=-1"

       rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    'rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
   'rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
    'XPDtbtimeTrans.value = Time
       Me.TxtModFlg.text = "R"
        chpo = True
        
   ' Retrive



    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    
    Exit Sub

ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
   ' Label1.Visible = False
   Me.lblrecomentclient.Caption = "RecommentClient"
   Me.lblprivatecopm.Caption = "Private Company"
lbl(9).Caption = "Date Entry"
Cmd(14).Caption = "Final All"
lblnotacept.Caption = "Reason Not Accept"
lblresonwaite.Caption = "Reason Wait"
lbl(12).Caption = "Date of exit expected"
lbl(18).Caption = "Use Screen"
lbl(13).Caption = "Time out expected"
lbl(11).Caption = "Actual date of exit"
lbl(14).Caption = "Actual time out"
Me.CheckBox2.RightToLeft = False
lbl(19).Caption = "Code"
Label1.Caption = "Meter out"
Label5.Caption = "Spare Part "
lbl(24).Caption = "Estimated pieces"
Label8.Caption = "Item"
'lbl(57).Caption "Total"
lbl(28).Caption = "Total"
lbl(26).Caption = "Qty"
lbl(21).Caption = "Price"
lbl(29).Caption = "Total"
lbl(32).Caption = "Discount"
lbl(38).Caption = "after discount"
lbl(57).Caption = "Total"
lbl(34).Caption = "Disc %"
lbl(37).Caption = "Vat %"
lbl(39).Caption = "Vat"
lbl(22).Caption = "Net"

Cmd(15).Caption = "Add"
Cmd(16).Caption = "Delete"

Me.CheckBox2.Caption = "Final"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    menuet.Caption = "Show Recommendations"
    Me.bClose.Caption = "X Close "
   Cmd(11).Caption = "Print quote"
    Cmd(10).Caption = "Print permission reform"
    Cmd(9).Caption = "Print is filled"
    Cmd(17).Caption = "Print Offer price 2"
    Cmd(18).Caption = "Print Invoice Order 2"
 Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
    XPTab301.CurrTab = 0
    XPTab301.Caption = "ID card repair data|Reform work|CC|Bills of exchange|Estimated pieces"
    Me.Caption = " WorkOrder"
    Me.lblcomputer.Caption = "ComputerCode"
    Me.BtImage.Caption = "Show Image"
    Me.lblmarks.Caption = "Notes From the Car"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "ShNo#"
    lbl(7).Caption = "AuthNo#"
    lbl(8).Caption = "Repair#"
    lbl(1).Caption = "Date Time"
   Me.lblBr.Caption = "Branch"
   Me.lblDataCli.Caption = "Data of Client"
  ' Me.LblTypReq.Caption = "Type Request"
  lbl(17).Caption = "Car Code"
  Me.Command1.Caption = "Car Code"
   lbl(16).Caption = "KM"
   lbl(2).Caption = "Oil Change after"
   Me.CheckBox1.Caption = "Under Wait"
   Me.FramAccount.Caption = "Financial Situation"
   Me.RdCash.Caption = "Cash"
   Me.RdCash.RightToLeft = False
   Me.rdCredit.Caption = "Credit"
   Me.rdCredit.RightToLeft = False
   Me.Rdacco.Caption = "Account"
   Me.Rdacco.RightToLeft = False
   Me.RdCompany.Caption = "Companies"
   Me.RdCompany.RightToLeft = False
   Me.RdPerson.Caption = "Persons"
   Me.RdPerson.RightToLeft = False
  Me.LblCli.Caption = "Client Name"
  Me.lblModel.Caption = "Models"
  lbl(25).Caption = "This image allows you to select the parts that you want to maintain"
  Me.LblPhone.Caption = "Telephone"
  Me.lblMobile.Caption = "Mobile"
  Me.lblbox.Caption = "Mailbox"
  Me.lblfax.Caption = "Fax"
  Me.lblemail.Caption = "Email"
  Me.lblAdress.Caption = "Address"
  Me.lblboxzib.Caption = "Postcode"
  Me.lblremrk.Caption = "Initial observations of the art"
  
  Me.lbltycar.Caption = "Type of Car"
  Me.LblOrderSt.Caption = "Oreder Status"
  Me.lblColor.Caption = "Color"
  Me.LblWork.Caption = "Maintenance Work"
  Me.lblExt.Caption = "Purchases and external works"
  Me.LblPla.Caption = "Plate No"
  Me.LblYear.Caption = "Year Manfac"
  Me.ChAccept.Caption = "Has the consent of the client"
  Me.lblEx.Caption = "Total of Purchas external works"
  Me.LblM.Caption = "Total of MaintenanceWork"
  Me.Lbtota.Caption = "Total"
 ' lbl(2).Caption = "End Date"
  lbl(3).Caption = "Start"
  Cmd(8).Caption = "Delete"
  Cmd(21).Caption = "Delete"
  lbl(5).Caption = "End"
'  GroupBox1.RightToLeft = False
  Me.lbEOrder.Caption = "Enter Order"
  
lblty.Caption = "Category demand "
  Me.FrReturnMaint.Caption = "Re Maintenance"
  Me.lbtechnical.Caption = "The initial observation of Technical"
  Me.lbltycar.Caption = "Type"
  Me.frmgranty.Caption = "Data Guaranty"
  Me.lbllong.Caption = "Duration"
  Me.LblPayF.Caption = "Pay First"

  Me.LblAmountAcc.Caption = "Estimated Cost"
  Me.LblCarMeter.Caption = "Car Meter"
  Me.LblCodeShaseh.Caption = "Shaseh No"
   Me.lblCodeReg.Caption = "Record No"
     Me.LblCodeDoor.Caption = "Door No"
   Me.lblTypeReg.Caption = "Type recording "
  Me.ChAccept.RightToLeft = False
    lbl(10).Caption = "Driver"
    Me.LblFitter.Caption = "Fitter"
    Me.lblcodeclient.Caption = "Client Code"
    lbl(15).Caption = "Customer complaint"
    lbl(20).Caption = "By"
'    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"
    opt(2).Caption = "Show Price"
    opt(1).Caption = "Autho"
    opt(0).Caption = "Order"
    Cmd(12).Caption = "Convert  From Show TO Auth"
    Cmd(13).Caption = "Convert  From Auth TO Order"
''PriceFitter

With FG22
.TextMatrix(0, .ColIndex("itemcode")) = "item code"
.TextMatrix(0, .ColIndex("itemname")) = "item name"
.TextMatrix(0, .ColIndex("ItemName2")) = "Item Name2"
.TextMatrix(0, .ColIndex("Price")) = "Unit Price"
.TextMatrix(0, .ColIndex("Qty")) = "Qty"
.TextMatrix(0, .ColIndex("BeforeVat")) = "Total"
.TextMatrix(0, .ColIndex("Remark")) = "Remark"

End With
    With Me.Fg
    .TextMatrix(0, .ColIndex("PriceFitter")) = "PriceFitter"
        .TextMatrix(0, .ColIndex("serial")) = "NO"
        .TextMatrix(0, .ColIndex("value")) = "Value"
        .TextMatrix(0, .ColIndex("name")) = "Name"
         .TextMatrix(0, .ColIndex("cod")) = "Code"
        .TextMatrix(0, .ColIndex("totalm")) = "Count Mainte"
       .TextMatrix(0, .ColIndex("count")) = "Count"
       .TextMatrix(0, .ColIndex("fitter")) = "Fitter"
        .TextMatrix(0, .ColIndex("supervisor")) = "Dept"
        .TextMatrix(0, .ColIndex("workshop")) = "WorkShop"
         .TextMatrix(0, .ColIndex("nohours")) = "No Hours"
        .TextMatrix(0, .ColIndex("finish")) = "Finished"
       .TextMatrix(0, .ColIndex("dateenter")) = "DateEnter"
       .TextMatrix(0, .ColIndex("dateout")) = "Date Exit"
       .TextMatrix(0, .ColIndex("timEnter")) = "Time Enter"
       .TextMatrix(0, .ColIndex("TimOut")) = "Time Exit"
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
.TextMatrix(0, .ColIndex("typeexpen")) = "TypeExpen"
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
            menue(12).Visible = False
            Me.ChAccept.Enabled = True
   Me.CheckBox1.Enabled = True
    Me.CheckBox2.Enabled = True
     Me.DcbOrderStatus.Enabled = False
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
            menue(12).Visible = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
               Me.ChAccept.value = xtpUnchecked
               Me.CheckBox1.value = xtpUnchecked
                Me.CheckBox2.value = xtpUnchecked
            Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 2
            fg2.Enabled = True
    Me.ChAccept.Enabled = False
    Me.CheckBox1.Enabled = False
     Me.DcbOrderStatus.ListIndex = 0
     Me.ComGranty.ListIndex = 0
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
            menue(12).Visible = False
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
             Me.CheckBox1.Enabled = True
'     Me.DcbOrderStatus.ListIndex = 0
     Me.DcbOrderStatus.Enabled = False
    End Select

    Exit Sub
ErrTrap:
End Sub

  

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap
    screenData = True
 
    
    
chpo = True
    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case Index

        Case 0

            If Not (rs.EOF Or rs.BOF) Then
                rs.MovePrevious
chpo = True
                If rs.BOF Then rs.MoveFirst
            End If

        Case 1

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                chpo = True
            End If

        Case 2

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
                chpo = True
            End If

        Case 3

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveNext
chpo = True

                If rs.EOF Then rs.MoveLast
            End If

    End Select

    Retrive
    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)

DcbOrderStatus.Enabled = False
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
  Dim ContactTime As Date
  vchrgrid.Clear flexClearScrollable, flexClearEverything
            vchrgrid.Rows = 1
 
'If screenData = False Then
'   If rs.State = adStateOpen Then
'   rs.Close
'
'   Else
'rs.Open
   
'   End If
'Set rs = Nothing
'     strsql = "select * From dbo.TblCardAuthorizationReform     where id='" & Lngid & "'"
'       rs.Open strsql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    End If
    
       
Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 1
            Fg.Enabled = True
           fg2.Clear flexClearScrollable, flexClearEverything
            fg2.Rows = 1
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

    XPTxtID.text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    TxtCusID.text = val(IIf(IsNull(rs("CusID").value), 0, rs("CusID").value))
        If Not IsNull(rs("RecordeTime").value) Then
          ContactTime = FormatDateTime(rs("RecordeTime").value, vbShortTime)
          Me.DTPicker1.value = ContactTime
        End If
    cmdEndAll.Tag = val(rs!IsEndAll & "")
    
If val(rs!IsEndAll & "") = 1 Then
    cmdEndAll.Enabled = False
    If SystemOptions.UserInterface = ArabicInterface Then
        cmdEndAll.Caption = " „ «ﬁ›«· «·ﬂ«— "
    Else
        cmdEndAll.Caption = "The card has been permanently locked"
    End If
    cmdOpenCard.Enabled = True
   Else
     cmdEndAll.Enabled = True
     If SystemOptions.UserInterface = ArabicInterface Then
        cmdEndAll.Caption = "«ﬁ›«· «·ﬂ«—  ‰Â«∆Ï"
    Else
        cmdEndAll.Caption = "Close Card"
    End If
    cmdOpenCard.Enabled = False
   End If
    
   
    TxtCarMetarOut.text = val(IIf(IsNull(rs("CarMetarOut").value), 0, rs("CarMetarOut").value))
    TxtLastWorOrder.text = val(IIf(IsNull(rs("LastWorOrder").value), 0, rs("LastWorOrder").value))
    TxtTypeCustomer.text = val(IIf(IsNull(rs("TypeCustomer").value), 0, rs("TypeCustomer").value))
    txtKM.text = IIf(IsNull(rs("OverKM").value), "", rs("OverKM").value)
    txtnotacept.text = IIf(IsNull(rs("NotAccept").value), "", rs("NotAccept").value)
    cmbItems.BoundText = IIf(IsNull(rs("ItemID33").value), "", rs("ItemID33").value)
   txtSalesInvoiceOrder = IIf(IsNull(rs("SalesInvoiceOrder").value), "", rs("SalesInvoiceOrder").value)
    
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    DcbCarType.BoundText = IIf(IsNull(rs("CarTypeID").value), "", rs("CarTypeID").value)
    DcbCarModel.BoundText = IIf(IsNull(rs("CarModelID").value), "", rs("CarModelID").value)
    TxtSparePart.text = IIf(IsNull(rs("SparePart").value), "", rs("SparePart").value)
    DcbColor.BoundText = IIf(IsNull(rs("ColorID").value), "", rs("ColorID").value)
    DcbyearFactor.text = IIf(IsNull(rs("YearFact").value), "", rs("YearFact").value)
    TxtClientPhone.text = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
    Me.TxtRemarkCar.text = IIf(IsNull(rs("Remarkcar").value), "", rs("Remarkcar").value)
    TxtCliientName.text = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
        DcbCarType.BoundText = IIf(IsNull(rs("CarTypeID").value), "", rs("CarTypeID").value)
        
    TxtPlatNo.text = IIf(IsNull(rs("PlateNo").value), "", rs("PlateNo").value)
    Me.TxtNoteIntial1 = IIf(IsNull(rs("Noteinitial").value), "", rs("Noteinitial").value) 's("Noteinitial").value
    Me.TxtComplaint.text = IIf(IsNull(rs("Complaint").value), "", rs("Complaint").value) 's("Complaint").value
    Me.txtresonwait.text = IIf(IsNull(rs("ResonUnderWait").value), "", rs("ResonUnderWait").value) 's("ResonUnderWait").value
    Me.TxtCodeComputer.text = IIf(IsNull(rs("CodeComputer").value), "", rs("CodeComputer").value) 'rs("CodeComputer").value
  ''/////////////////////////////////////
    Me.txtprivate.text = IIf(IsNull(rs("PrivateCop").value), "", rs("PrivateCop").value) ' rs("PrivateCop").value
    Me.txtrecomment.text = IIf(IsNull(rs("ReComentClient").value), "", rs("ReComentClient").value) ' rs("ReComentClient").value
    
   Me.txtDiscValue.text = IIf(IsNull(rs("DiscValue").value), "", rs("DiscValue").value) ' rs("ReComentClient").value
   Me.txtDiscPercent.text = IIf(IsNull(rs("DiscPercent").value), "", rs("DiscPercent").value) ' rs("ReComentClient").value
   Me.txtTotalAfterDiscount.text = IIf(IsNull(rs("TotalAfterDiscount").value), "", rs("TotalAfterDiscount").value) ' rs("ReComentClient").value
   Me.txtVatyo.text = IIf(IsNull(rs("Vatyo").value), "", rs("Vatyo").value) ' rs("ReComentClient").value
   Me.txtVat2.text = IIf(IsNull(rs("Vat2").value), "", rs("Vat2").value) ' rs("ReComentClient").value
    
                        
                        
    '    Me.combtypereq.ListIndex = IIf(IsNull(rs("typerequest").value), "", rs("typerequest").value)
 
        If rs("Cash").value = True Then
        Me.RdCash.value = True
        Else
         Me.RdCash.value = False
         End If
         If rs("Accoun").value = True Then
        Me.Rdacco.value = True
        Else
         Me.Rdacco.value = False
         End If
     If rs("credit").value = True Then
        Me.rdCredit.value = True
        Else
         Me.rdCredit.value = False
         End If
       Me.DcboFitter.BoundText = IIf(IsNull(rs("FitterID").value), "", rs("FitterID").value)
       DTPicker1.value = IIf(IsNull(rs("RecordeTime").value), Null, rs("RecordeTime").value) 'rs("RecordeTime").value
       Me.TxtMobile.text = IIf(IsNull(rs("mobile").value), "", rs("mobile").value) 'rs("mobile").value
       Me.TxtBox.text = IIf(IsNull(rs("box").value), "", rs("box").value) 'rs("box").value
       TxtClientCode.text = IIf(IsNull(rs("ClientCode").value), "", rs("ClientCode").value) 'rs("ClientCode").value
       Me.TxtFax.text = IIf(IsNull(rs("fax").value), "", rs("fax").value) 'rs("fax").value
       Me.TxtEmail.text = IIf(IsNull(rs("email").value), "", rs("email").value) 'rs("email").value
       Me.TxtAddres.text = IIf(IsNull(rs("address").value), "", rs("address").value) ' rs("address").value
       Me.txtboxzip.text = IIf(IsNull(rs("boxzip").value), "", rs("boxzip").value) 'rs("boxzip").value
       Me.txtCodeReg.text = IIf(IsNull(rs("codereg").value), "", rs("codereg").value) 'rs("codereg").value
       Me.TxtTtpeReg.text = IIf(IsNull(rs("typereg").value), "", rs("typereg").value) 'rs("typereg").value
       Me.TxtCodeDoor.text = IIf(IsNull(rs("codedoor").value), "", rs("codedoor").value) 'rs("codedoor").value
       Me.TxtDriver.text = IIf(IsNull(rs("driver").value), "", rs("driver").value) 'rs("driver").value
       Me.DTPEnterDate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
       DTPDateAcutExite.value = IIf(IsNull(rs("DateAcutExite").value), Date, rs("DateAcutExite").value) ' rs("DateAcutExite").value
       DTPDateExptExit.value = IIf(IsNull(rs("DateExptExit").value), Date, rs("DateExptExit").value) 'rs("DateExptExit").value
       DTPTimeAcutExite.value = IIf(IsNull(rs("TimeAcutExite").value), Null, rs("TimeAcutExite").value) 'rs("TimeAcutExite").value
       DTPTimeExptExit.value = IIf(IsNull(rs("TimeExptExit").value), Null, rs("TimeExptExit").value) 'rs("TimeExptExit").value
       
      If rs("persons").value = 1 Then
        Me.RdPerson.value = True
        Else
        Me.RdPerson.value = False
      End If
   
      If rs("Companies").value = 1 Then
        Me.RdCompany.value = True
        Else
         Me.RdCompany.value = False
         End If
   

   
   DcbOrderStatus.ListIndex = val(IIf(IsNull(rs("OrderStatus").value), 0, rs("OrderStatus").value))


   TXtCarMeter.text = IIf(IsNull(rs("CarMeter").value), "", rs("CarMeter").value)
   Me.TXtShaseh.text = IIf(IsNull(rs("Shaseh").value), "", rs("Shaseh").value)
  ' Me.TXtNotAccept.text = IIf(IsNull(rs("NotAccept").value), "", rs("NotAccept").value)
   TxtLongGranty.text = IIf(IsNull(rs("LongGranty").value), "", rs("LongGranty").value)
   TxtFirstPrice.text = val(IIf(IsNull(rs("PayFirst").value), 0, rs("PayFirst").value))
   Me.TxtAmoutAccept.text = val(IIf(IsNull(rs("AmountAccept").value), 0, rs("AmountAccept").value))
   DateStartG.value = IIf(IsNull(rs("DateStartG").value), Date, rs("DateStartG").value)
   DateEndg.value = IIf(IsNull(rs("DateEndG").value), Date, rs("DateEndG").value)
   Me.TxtComplaint.text = IIf(IsNull(rs("Complaint").value), "", rs("Complaint").value)
   Me.TxtNoteIntial1.text = IIf(IsNull(rs("Noteinitial").value), "", rs("Noteinitial").value)
   
       
        
        TxtCusID.text = val(IIf(IsNull(rs("CusID").value), 0, rs("CusID").value))
        If val(TxtCusID.text) = 0 Then
            Dim ss As String
            ss = "Select cusId From TblCustemers Where Code = N'" & Trim(TxtClientCode) & "'"
            Dim rsDummy As New ADODB.Recordset
            rsDummy.Open ss, Cn, adOpenStatic, adLockReadOnly
            If Not rsDummy.EOF Then
                TxtCusID.text = rsDummy!CusID & ""
            End If
        End If
        
    If SystemOptions.UserInterface = EnglishInterface Then
        StrSQL = "Select CusNamee ClientName FROM TblCustemers where CusId = " & val(TxtCusID.text)
        Dim rsDu As New ADODB.Recordset
        rsDu.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
        If Not rsDu.EOF Then
            TxtCliientName.text = IIf(IsNull(rsDu("ClientName").value), "", rsDu("ClientName").value)
        End If
        If Trim(TxtCliientName.text) = "" Then
            TxtCliientName.text = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
        End If
    End If



  ' If rs("Granty").value = True Then
   Me.ComGranty.ListIndex = IIf(IsNull(rs("Granty").value), 0, rs("Granty").value)
 '  Me.frmgranty.Visible = True
  ' Else
  ' Me.ComGranty.ListIndex = 1
'   Me.frmgranty.Visible = False
  ' End If
   If rs("Month_Day").value = True Then
   Me.ComMD.ListIndex = 0
   Else
   Me.ComMD.ListIndex = 1
   End If
   chpo = True
     If rs("wait").value = True Then
     Me.CheckBox1.value = vbChecked
  
     Else
     Me.CheckBox1.value = vbUnchecked
     End If
     chpo = True
     '''///////////////16 11 2015
     Me.TxtWorkOrder.text = IIf(IsNull(rs("WorkOrder").value), "", rs("WorkOrder").value)
     Me.TxtShowPriceOrder.text = IIf(IsNull(rs("ShowPriceOrder").value), "", rs("ShowPriceOrder").value)
     Me.TxtAuthoOrder.text = IIf(IsNull(rs("AuthoOrder").value), "", rs("AuthoOrder").value)
    If rs("TypeOrder").value = 0 Then
    Me.opt(0).value = True
    ElseIf rs("TypeOrder").value = 1 Then
    Me.opt(1).value = True
    ElseIf rs("TypeOrder").value = 2 Then
    Me.opt(2).value = True
    End If
  '
  '   End If
    chpo = True
       If rs("notAcepted").value = True Then
     Me.CheckBox2.value = vbChecked
  Me.ChAccept.value = vbUnchecked
     Else
     Me.CheckBox2.value = vbUnchecked
     End If
      If rs("subcar1").value = True Then
          Me.imag1.Picture = Me.Img.Picture
Else
 Me.imag1.Picture = Me.imgnul.Picture
            
           End If
            If rs("subcar2").value = True Then
           Me.imag2.Picture = Me.Img.Picture
Else
 Me.imag2.Picture = Me.imgnul.Picture
           End If
            If rs("subcar3").value = True Then
        Me.imag3.Picture = Me.Img.Picture
Else
 Me.imag3.Picture = Me.imgnul.Picture
           End If
            If rs("subcar4").value = True Then
 
Me.imag4.Picture = Me.Img.Picture
Else
 Me.imag4.Picture = Me.imgnul.Picture
           End If
            If rs("subcar5").value = True Then
          Me.imag5.Picture = Me.Img.Picture
Else
 Me.imag5.Picture = Me.imgnul.Picture
           End If
            If rs("subcar6").value = True Then
       Me.img6.Picture = Me.Img.Picture
Else
 Me.img6.Picture = Me.imgnul.Picture
           End If
            If rs("subcar7").value = True Then
           Me.img7.Picture = Me.Img.Picture
Else
 Me.img7.Picture = Me.imgnul.Picture
           End If
            If rs("subcar8").value = True Then
          Me.img8.Picture = Me.Img.Picture
Else
 Me.img8.Picture = Me.imgnul.Picture
           End If
            If rs("subcar9").value = True Then
          Me.img9.Picture = Me.Img.Picture
Else
 Me.img9.Picture = Me.imgnul.Picture
           End If
            If rs("subcar10").value = True Then
           Me.img10.Picture = Me.Img.Picture
Else
 Me.img10.Picture = Me.imgnul.Picture
           End If
           ''''''''//////////7/5/2014
            If rs("subcar11").value = True Then
           Me.img11.Picture = Me.Img.Picture
Else
 Me.img11.Picture = Me.imgnul.Picture
           End If
                      If rs("subcar12").value = True Then
           Me.img12.Picture = Me.Img.Picture
Else
 Me.img12.Picture = Me.imgnul.Picture
           End If
                      If rs("subcar13").value = True Then
           Me.img13.Picture = Me.Img.Picture
Else
 Me.img13.Picture = Me.imgnul.Picture
           End If
                      If rs("subcar14").value = True Then
           Me.img14.Picture = Me.Img.Picture
Else
 Me.img14.Picture = Me.imgnul.Picture
           End If
       If SystemOptions.LinkCustomerWithCars = True Then
       'Dim Dcombos As ClsDataCombos
       Dim Dcombos As New ClsDataCombos
       Dcombos.GetCarsOfCustomer DcbCar, val(TxtCusID.text)
       End If
Me.DcbCar.BoundText = IIf(IsNull(rs("CarID").value), "", rs("CarID").value)
    Set RsDetails = New ADODB.Recordset
StrSQL = " SELECT     dbo.TblCardAuthorizationReformDetails.ID, dbo.TblCardAuthorizationReformDetails.Type, dbo.TblCardAuthorizationReformDetails.Mainte,"
StrSQL = StrSQL & "                      dbo.TblMaintenanceWork.name, dbo.TblMaintenanceWork.namee, dbo.TblCardAuthorizationReformDetails.EmpID, dbo.TblEmployee.Emp_Name AS fiter,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee AS fitere, dbo.TblCardAuthorizationReformDetails.empsuper, TblEmployee_1.Emp_Name AS NameSuper,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Namee AS NamesuperE, dbo.TblCardAuthorizationReformDetails.Deptid, dbo.TblEmpDepartments.DepartmentName,"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments.DepartmentNamee, dbo.TblCardAuthorizationReformDetails.Dpeterial, dbo.TblCardAuthorizationReformDetails.DeptBr,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.DeptColor, dbo.TblCardAuthorizationReformDetails.PriceFitter, dbo.TblCardAuthorizationReformDetails.payed,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.allocation, dbo.TblCardAuthorizationReformDetails.TimOut, dbo.TblCardAuthorizationReformDetails.TimeEnter,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.DateExit, dbo.TblCardAuthorizationReformDetails.DateEnter, dbo.TblCardAuthorizationReformDetails.finish,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.nohours, dbo.TblCardAuthorizationReformDetails.bill, dbo.TblCardAuthorizationReformDetails.comp,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.[count] , dbo.TblCardAuthorizationReformDetails.[value]"
StrSQL = StrSQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblMaintenanceWork ON dbo.TblCardAuthorizationReformDetails.Mainte = dbo.TblMaintenanceWork.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.TblCardAuthorizationReformDetails.Deptid = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblCardAuthorizationReformDetails.empsuper = TblEmployee_1.Emp_ID ON"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_id = dbo.TblCardAuthorizationReformDetails.EmpID"
StrSQL = StrSQL & "  Where (dbo.TblCardAuthorizationReformDetails.id  = " & val(XPTxtID.text) & ") And (dbo.TblCardAuthorizationReformDetails.Type = 0)"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
 

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
       With Me.Fg
      ' Deptid RsDetails.MoveFirst  RsDetails("PriceFitter").value = val(IIf((fg.TextMatrix(i, fg.ColIndex("PriceFitter"))), fg.TextMatrix(i, fg.ColIndex("PriceFitter")), 0))
        .Rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .Rows - 1
    
            .TextMatrix(i, .ColIndex("serial")) = i
            Fg.TextMatrix(i, Fg.ColIndex("dateenter")) = IIf(IsNull(RsDetails("DateEnter").value), "", RsDetails("DateEnter").value)
     
            Fg.TextMatrix(i, Fg.ColIndex("dateout")) = IIf(IsNull(RsDetails("DateExit").value), "", RsDetails("DateExit").value)
            Fg.TextMatrix(i, Fg.ColIndex("timEnter")) = IIf(IsNull(RsDetails("TimeEnter").value), "", RsDetails("TimeEnter").value)
            Fg.TextMatrix(i, Fg.ColIndex("TimOut")) = IIf(IsNull(RsDetails("TimOut").value), "", RsDetails("TimOut").value)
             .TextMatrix(i, .ColIndex("nohours")) = IIf(IsNull(RsDetails("nohours").value), "", RsDetails("nohours").value) 'Details("nohours").value
               .TextMatrix(i, .ColIndex("supervisor")) = IIf(IsNull(RsDetails("NameSuper").value), "", RsDetails("NameSuper").value) 'sDetails("supervisor").value
                 .TextMatrix(i, .ColIndex("workshop")) = IIf(IsNull(RsDetails("DepartmentName").value), "", RsDetails("DepartmentName").value) 'sDetails("workshop").value
                   .TextMatrix(i, .ColIndex("fitter")) = IIf(IsNull(RsDetails("fiter").value), "", RsDetails("fiter").value) 'sDetails("fitter").value
            .TextMatrix(i, .ColIndex("value")) = val(IIf(IsNull(RsDetails("Value").value), 0, RsDetails("Value").value)) 'RsDetails("Value").value
            ' .TextMatrix(i, .ColIndex("PriceFitter")) = val(IIf(IsNull(RsDetails("PriceFitter").value), 0, RsDetails("PriceFitter").value))
            .TextMatrix(i, .ColIndex("finish")) = IIf(IsNull(RsDetails("finish").value), "", RsDetails("finish").value) 'RsDetails("finish").value
             .TextMatrix(i, .ColIndex("cod")) = IIf(IsNull(RsDetails("Mainte").value), "", RsDetails("Mainte").value) 'RsDetails("Mainte").value
          .TextMatrix(i, .ColIndex("count")) = IIf(IsNull(RsDetails("count").value), "", RsDetails("count").value) 'RsDetails("count").value
           .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(RsDetails("EmpID").value), "", RsDetails("EmpID").value)
           .TextMatrix(i, .ColIndex("empsuper")) = IIf(IsNull(RsDetails("empsuper").value), "", RsDetails("empsuper").value)
            .TextMatrix(i, .ColIndex("Deptid")) = IIf(IsNull(RsDetails("Deptid").value), "", RsDetails("Deptid").value)
          .TextMatrix(i, .ColIndex("Dpeterial")) = IIf(IsNull(RsDetails("Dpeterial").value), "", RsDetails("Dpeterial").value) 'RsDetails("count").value
           .TextMatrix(i, .ColIndex("DeptColor")) = IIf(IsNull(RsDetails("DeptColor").value), "", RsDetails("DeptColor").value)
            .TextMatrix(i, .ColIndex("payed")) = IIf(IsNull(RsDetails("payed").value), 0, RsDetails("payed").value)
            .TextMatrix(i, .ColIndex("allocation")) = IIf(IsNull(RsDetails("allocation").value), 0, RsDetails("allocation").value)
               If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("name").value), "", RsDetails("name").value)
                Else
                   .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("namee").value), "", RsDetails("namee").value)
             End If
             .TextMatrix(i, .ColIndex("PriceFitter")) = val(IIf(IsNull(RsDetails("PriceFitter").value), 0, RsDetails("PriceFitter").value))
            RsDetails.MoveNext
         
        Next i
End With
    End If
      RsDetails.Close
    Set RsDetails = Nothing
  
    '//////////////////////////////////////////
    Set RsDetails1 = New ADODB.Recordset
 StrSQL = " SELECT      dbo.TblCardAuthorizationReformDetails.ID, dbo.TblCardAuthorizationReformDetails.comp, dbo.TblCardAuthorizationReformDetails.Type, "
  StrSQL = StrSQL & "                    dbo.TblCardAuthorizationReformDetails.[Value], dbo.TblExtraExpeneses.name, dbo.TblExtraExpeneses.namee, dbo.TblExtraExpeneses.TypeExtrExpen,"
   StrSQL = StrSQL & "                   dbo.TblCardAuthorizationReformDetails.Mainte, dbo.TblCardAuthorizationReformDetails.[count], dbo.TblCardAuthorizationReformDetails.bill,"
   StrSQL = StrSQL & "                   dbo.TblExtraExpeneses.Id AS Expr1, dbo.TblTypeExtraExpeneses.name AS nameTy, dbo.TblTypeExtraExpeneses.namee AS nameeTy,"
   StrSQL = StrSQL & "                   dbo.TblExtraExpeneses.typeid"
 StrSQL = StrSQL & " FROM         dbo.TblCardAuthorizationReformDetails INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblExtraExpeneses ON dbo.TblCardAuthorizationReformDetails.Mainte = dbo.TblExtraExpeneses.Id INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblTypeExtraExpeneses ON dbo.TblExtraExpeneses.TypeID = dbo.TblTypeExtraExpeneses.Id"
StrSQL = StrSQL & "  Where (dbo.TblCardAuthorizationReformDetails.id =" & val(XPTxtID.text) & ") And (dbo.TblCardAuthorizationReformDetails.Type = 1)"

    RsDetails1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
     If Not (RsDetails1.BOF Or RsDetails1.EOF) Then
       With Me.fg2
      '  RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails1.RecordCount

        For i = .FixedRows To .Rows - 1
    
            .TextMatrix(i, .ColIndex("serial")) = i
            .TextMatrix(i, .ColIndex("value")) = val(IIf(IsNull(RsDetails1("Value").value), 0, RsDetails1("Value").value)) 'RsDetails1("Value").value
             .TextMatrix(i, .ColIndex("cod")) = val(IIf(IsNull(RsDetails1("Mainte").value), 0, RsDetails1("Mainte").value)) 'RsDetails1("Mainte").value
            .TextMatrix(i, .ColIndex("count")) = val(IIf(IsNull(RsDetails1("count").value), 0, RsDetails1("count").value)) 'RsDetails1("count").value
             .TextMatrix(i, .ColIndex("Codtype")) = val(IIf(IsNull(RsDetails1("TypeID").value), 0, RsDetails1("TypeID").value))
           .TextMatrix(i, .ColIndex("comp")) = IIf(IsNull(RsDetails1("comp").value), "", RsDetails1("comp").value) ' RsDetails1("comp").value
           .TextMatrix(i, .ColIndex("bill")) = IIf(IsNull(RsDetails1("bill").value), "", RsDetails1("bill").value) 'RsDetails1("bill").value
            .TextMatrix(i, .ColIndex("typeexpen")) = IIf(IsNull(RsDetails1("nameTy").value), "", RsDetails1("nameTy").value)
          ' .TextMatrix(i, .ColIndex("typeexpen")) = IIf(IsNull(RsDetails1("TypeExtrExpen").value), "", RsDetails1("TypeExtrExpen").value)
               If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails1("name").value), "", RsDetails1("name").value)
                Else
                   .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails1("namee").value), "", RsDetails1("namee").value)
             End If
            RsDetails1.MoveNext
         
        Next i
End With
    End If
   

   '
    
RsDetails1.Close
 Set RsDetails1 = Nothing


Dim s As String
s = "Select *,tblItems.ItemCode,tblItems.ItemName from TblCardAuthorizationReformItems Left Outer Join tblItems On tblItems.ItemID =TblCardAuthorizationReformItems.ItemID "
s = s & "  Where (dbo.TblCardAuthorizationReformItems.id =" & val(XPTxtID.text) & ")"
loadgrid s, FG22, True, True
  '  fillapprovData
    ReLineGrid
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
       chpo = True
       
che
newret
ReLineGrid2
    Exit Sub
    

ErrTrap:

End Sub
Private Sub ReLineGrid2()
    Dim i As Integer
    Dim IntCounter As Integer
    Dim summ As Double
   ''''///
    summ = 0
IntCounter = 0
lbl(58).Caption = 0
lbl(23).Caption = 0

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

Dim mPriceBeDisc As Double
Dim mDiscValue As Double
Dim mDiscPercent As Double
Dim mVat As Double

summ = 0
    With Me.FG22
        For i = .FixedRows To .Rows - 1

            If (.TextMatrix(i, .ColIndex("itemcode"))) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
              
                If val(.TextMatrix(i, .ColIndex("PriceBDisc"))) = 0 Then
                    .TextMatrix(i, .ColIndex("PriceBDisc")) = .TextMatrix(i, .ColIndex("Price"))
                End If
                
                mPriceBeDisc = mPriceBeDisc + (val(.TextMatrix(i, .ColIndex("Price"))) * val(.TextMatrix(i, .ColIndex("Qty"))))
                mDiscValue = mDiscValue + (val(.TextMatrix(i, .ColIndex("DiscValue"))) * val(.TextMatrix(i, .ColIndex("Qty"))))
                mVat = mVat + (val(.TextMatrix(i, .ColIndex("Vat2")))) '* val(.TextMatrix(i, .ColIndex("Qty"))))
                
                summ = summ + val(.TextMatrix(i, .ColIndex("TotalWithVat")))
             
                  End If
        Next i
    End With
    If mPriceBeDisc <> 0 Then
        mDiscPercent = Round(val(mDiscValue) / val(mPriceBeDisc) * 100, 2)
    Else
        mDiscPercent = 0
    End If
    'lbl(23).Caption = summ
     lbl(31).Caption = mPriceBeDisc
     txtTotalAfterDiscount = mPriceBeDisc - val(txtDiscValue)
     'lbl(33).Caption = mDiscValue
     'lbl(36).Caption = mDiscPercent
     'lbl(38).Caption = mVat
     CalculteValueAdded 1
    
   
    End Sub
Public Sub retrive1(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
    clear_all Me
   If rs.State = adStateOpen Then
   rs.Close
   
   Else
'rs.Open
   
   End If

     StrSQL = "select * From dbo.TblCardAuthorizationReform     where id=" & Lngid & ""
       rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
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

    XPTxtID.text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
     TxtTypeCustomer.text = val(IIf(IsNull(rs("TypeCustomer").value), 0, rs("TypeCustomer").value))
     txtKM.text = IIf(IsNull(rs("OverKM").value), "", rs("OverKM").value)
   ' Me.TxtEndDate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    DcbCarType.BoundText = IIf(IsNull(rs("CarTypeID").value), "", rs("CarTypeID").value)
    DcbCarModel.BoundText = IIf(IsNull(rs("CarModelID").value), "", rs("CarModelID").value)
    TxtClientCode.text = IIf(IsNull(rs("ClientCode").value), "", rs("ClientCode").value) 'rs("ClientCode").value
    
        cmbItems.BoundText = IIf(IsNull(rs("ItemID33").value), "", rs("ItemID33").value)
   txtSalesInvoiceOrder = IIf(IsNull(rs("SalesInvoiceOrder").value), "", rs("SalesInvoiceOrder").value)
    

   cmdEndAll.Tag = val(rs!IsEndAll & "")
   If val(rs!IsEndAll & "") = 1 Then
    
    cmdEndAll.Enabled = False
    
    cmdEndAll.Caption = " „ «ﬁ›«· «·ﬂ«— "
    
   Else
     cmdEndAll.Enabled = True
     cmdEndAll.Caption = "«ﬁ›«· «·ﬂ«—  ‰Â«∆Ï"
   End If
    
   
   ' DcboSpecifications.BoundText = IIf(IsNull(rs("gradeID").value), "", rs("gradeID").value)
    DcbColor.BoundText = IIf(IsNull(rs("ColorID").value), "", rs("ColorID").value)
    DcbyearFactor.text = IIf(IsNull(rs("YearFact").value), "", rs("YearFact").value)
   TxtClientPhone.text = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
    Me.TxtRemarkCar.text = IIf(IsNull(rs("Remarkcar").value), "", rs("Remarkcar").value)
   TxtCliientName.text = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)

        TxtCusID.text = val(IIf(IsNull(rs("CusID").value), 0, rs("CusID").value))
        If val(TxtCusID.text) = 0 Then
            Dim ss As String
            ss = "Select cusId From TblCustemers Where Code = N'" & Trim(TxtClientCode) & "'"
            Dim rsDummy As New ADODB.Recordset
            rsDummy.Open ss, Cn, adOpenStatic, adLockReadOnly
            If Not rsDummy.EOF Then
                TxtCusID.text = rsDummy!CusID & ""
            End If
        End If
        
    If SystemOptions.UserInterface = EnglishInterface Then
        StrSQL = "Select CusNamee ClientName FROM TblCustemers where CusId = " & val(TxtCusID.text)
        Dim rsDu As New ADODB.Recordset
        rsDu.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
        If Not rsDu.EOF Then
            TxtCliientName.text = IIf(IsNull(rsDu("ClientName").value), "", rsDu("ClientName").value)
        End If
        If Trim(TxtCliientName.text) = "" Then
            TxtCliientName.text = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
        End If
    End If


   TxtPlatNo.text = IIf(IsNull(rs("PlateNo").value), "", rs("PlateNo").value)
    Me.TxtNoteIntial1 = IIf(IsNull(rs("Noteinitial").value), "", rs("Noteinitial").value) 'rs("Noteinitial").value
       Me.TxtComplaint.text = IIf(IsNull(rs("Complaint").value), "", rs("Complaint").value) 'rs("Complaint").value
     Me.txtresonwait.text = IIf(IsNull(rs("ResonUnderWait").value), "", rs("ResonUnderWait").value) 'rs("ResonUnderWait").value
          Me.TxtWorkOrder.text = IIf(IsNull(rs("WorkOrder").value), "", rs("WorkOrder").value)
     Me.TxtShowPriceOrder.text = IIf(IsNull(rs("ShowPriceOrder").value), "", rs("ShowPriceOrder").value)
     Me.TxtAuthoOrder.text = IIf(IsNull(rs("AuthoOrder").value), "", rs("AuthoOrder").value)
  ''/////////////////////////////////////
        Me.TxtCodeComputer.text = IIf(IsNull(rs("CodeComputer").value), "", rs("CodeComputer").value)
       ' Me.combtypereq = IIf(IsNull(rs("typerequest").value), "", rs("typerequest").value)
 
        If rs("Cash").value = True Then
        Me.RdCash.value = True
        Else
         Me.RdCash.value = False
         End If
         If rs("Accoun").value = True Then
        Me.Rdacco.value = True
        Else
         Me.Rdacco.value = False
         End If
     If rs("credit").value = True Then
        Me.rdCredit.value = True
        Else
         Me.rdCredit.value = False
         End If
      Me.DcboFitter.BoundText = IIf(IsNull(rs("FitterID").value), "", rs("FitterID").value)
   DTPicker1.value = IIf(IsNull(rs("RecordeTime").value), Null, rs("RecordeTime").value) 'rs("RecordeTime").value
       Me.TxtMobile.text = IIf(IsNull(rs("mobile").value), "", rs("mobile").value) 'rs("mobile").value
        Me.TxtBox.text = IIf(IsNull(rs("box").value), "", rs("box").value) 'rs("box").value
        
        Me.TxtFax.text = IIf(IsNull(rs("fax").value), "", rs("fax").value) 'rs("fax").value
        Me.TxtEmail.text = IIf(IsNull(rs("email").value), "", rs("email").value) 'rs("email").value
         Me.TxtAddres.text = IIf(IsNull(rs("address").value), "", rs("address").value) ' rs("address").value
         Me.txtboxzip.text = IIf(IsNull(rs("boxzip").value), "", rs("boxzip").value) 'rs("boxzip").value
         Me.txtCodeReg.text = IIf(IsNull(rs("codereg").value), "", rs("codereg").value) 'rs("codereg").value
         Me.TxtTtpeReg.text = IIf(IsNull(rs("typereg").value), "", rs("typereg").value) 'rs("typereg").value
        Me.TxtCodeDoor.text = IIf(IsNull(rs("codedoor").value), "", rs("codedoor").value) 'rs("codedoor").value
        Me.TxtDriver.text = IIf(IsNull(rs("driver").value), "", rs("driver").value) 'rs("driver").value
        Me.DTPEnterDate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
         DTPDateAcutExite.value = IIf(IsNull(rs("DateAcutExite").value), Date, rs("DateAcutExite").value) ' rs("DateAcutExite").value
         DTPDateExptExit.value = IIf(IsNull(rs("DateExptExit").value), Date, rs("DateExptExit").value) 'rs("DateExptExit").value
         DTPTimeAcutExite.value = IIf(IsNull(rs("TimeAcutExite").value), Null, rs("TimeAcutExite").value) 'rs("TimeAcutExite").value
        DTPTimeExptExit.value = IIf(IsNull(rs("TimeExptExit").value), Null, rs("TimeExptExit").value) 'rs("TimeExptExit").value
             
          If rs("persons").value = 1 Then
        Me.RdPerson.value = True
        Else
        Me.RdPerson.value = False
         End If
   
      If rs("Companies").value = 1 Then
        Me.RdCompany.value = True
        Else
         Me.RdCompany.value = False
         End If
   

   
   DcbOrderStatus.ListIndex = val(IIf(IsNull(rs("OrderStatus").value), 0, rs("OrderStatus").value))


   'If rs("OrderStatus").value = 3 Then
'TXtNotAccept.Visible = True
'LbNotAccept.Visible = True
'ChAccept.Visible = False
'Else

'ChAccept.Visible = True
'TXtNotAccept.Visible = False
'LbNotAccept.Visible = False
'If rs("OrderStatus").value = 1 Then
'Me.ChAccept.value = xtpChecked
'Else

'Me.ChAccept.value = xtpUnchecked
'End If
'End If
   TXtCarMeter.text = IIf(IsNull(rs("CarMeter").value), "", rs("CarMeter").value)
   Me.TXtShaseh.text = IIf(IsNull(rs("Shaseh").value), "", rs("Shaseh").value)
  ' Me.TXtNotAccept.text = IIf(IsNull(rs("NotAccept").value), "", rs("NotAccept").value)
   TxtLongGranty.text = IIf(IsNull(rs("LongGranty").value), "", rs("LongGranty").value)
   TxtFirstPrice.text = val(IIf(IsNull(rs("PayFirst").value), 0, rs("PayFirst").value))
   Me.TxtAmoutAccept.text = val(IIf(IsNull(rs("AmountAccept").value), 0, rs("AmountAccept").value))
   DateStartG.value = IIf(IsNull(rs("DateStartG").value), Date, rs("DateStartG").value)
   DateEndg.value = IIf(IsNull(rs("DateEndG").value), Date, rs("DateEndG").value)
   Me.TxtComplaint.text = IIf(IsNull(rs("Complaint").value), "", rs("Complaint").value)
   Me.TxtNoteIntial1.text = IIf(IsNull(rs("Noteinitial").value), "", rs("Noteinitial").value)
   TxtSparePart.text = IIf(IsNull(rs("SparePart").value), "", rs("SparePart").value)
   
  ' If rs("Granty").value = True Then
   Me.ComGranty.ListIndex = IIf(IsNull(rs("Granty").value), 0, rs("Granty").value)
 '  Me.frmgranty.Visible = True
  ' Else
  ' Me.ComGranty.ListIndex = 1
'   Me.frmgranty.Visible = False
  ' End If
   If rs("Month_Day").value = True Then
   Me.ComMD.ListIndex = 0
   Else
   Me.ComMD.ListIndex = 1
   End If
   ' If rs("Accept").value = True Then
   '  Me.ChAccept.value = vbChecked
   '  Me.DcbOrderStatus.ListIndex = 1
   '  Else
   '   Me.ChAccept.value = vbUnchecked
   '   End If
      If rs("subcar1").value = True Then
          Me.imag1.Picture = Me.Img.Picture
Else
 Me.imag1.Picture = Me.imgnul.Picture

           End If
            If rs("subcar2").value = True Then
          Me.imag2.Picture = Me.Img.Picture
Else
 Me.imag2.Picture = Me.imgnul.Picture
           End If
           If rs("subcar3").value = True Then
        Me.imag3.Picture = Me.Img.Picture
Else
Me.imag3.Picture = Me.imgnul.Picture
           End If
            If rs("subcar4").value = True Then

Me.imag4.Picture = Me.Img.Picture
Else
 Me.imag4.Picture = Me.imgnul.Picture
          End If
           If rs("subcar5").value = True Then
          Me.imag5.Picture = Me.Img.Picture
Else
 Me.imag5.Picture = Me.imgnul.Picture
           End If
            If rs("subcar6").value = True Then
       Me.img6.Picture = Me.Img.Picture
Else
 Me.img6.Picture = Me.imgnul.Picture
           End If
           If rs("subcar7").value = True Then
          Me.img7.Picture = Me.Img.Picture
Else
 Me.img7.Picture = Me.imgnul.Picture
           End If
            If rs("subcar8").value = True Then
          Me.img8.Picture = Me.Img.Picture
Else
 Me.img8.Picture = Me.imgnul.Picture
           End If
            If rs("subcar9").value = True Then
         Me.img9.Picture = Me.Img.Picture
Else
 Me.img9.Picture = Me.imgnul.Picture
           End If
            If rs("subcar10").value = True Then
           Me.img10.Picture = Me.Img.Picture
Else
 Me.img10.Picture = Me.imgnul.Picture
           End If
'          ''''''''//////////7/5/2014
           If rs("subcar11").value = True Then
          Me.img11.Picture = Me.Img.Picture
Else
Me.img11.Picture = Me.imgnul.Picture
          End If
                    If rs("subcar12").value = True Then
          Me.img12.Picture = Me.Img.Picture
Else
 Me.img12.Picture = Me.imgnul.Picture
           End If
                      If rs("subcar13").value = True Then
           Me.img13.Picture = Me.Img.Picture
Else
 Me.img13.Picture = Me.imgnul.Picture
           End If
                      If rs("subcar14").value = True Then
           Me.img14.Picture = Me.Img.Picture
Else
 Me.img14.Picture = Me.imgnul.Picture
           End If
            If SystemOptions.LinkCustomerWithCars = True Then
       Dim Dcombos As ClsDataCombos
       Set Dcombos = New ClsDataCombos
       Dcombos.GetCarsOfCustomer DcbCar, val(TxtCusID.text)
       End If
Me.DcbCar.BoundText = IIf(IsNull(rs("CarID").value), "", rs("CarID").value)
           'Me.txtresonwait.text = IIf(IsNull(rs("ResonUnderWait").value), "", rs("ResonUnderWait").value) 'rs("ResonUnderWait").value
    'TxtAdvanceValue.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
    'Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
  ' Me.TxtPaymentCounts.text = IIf(IsNull(rs("PaymentCounts").value), "", rs("PaymentCounts").value)
     Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
      If IsNull(rs("posted").value) Then
                                                  If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                               Else
                                                  Accredit.Caption = " send to Approval   "
                                             End If
                                              Accredit.Enabled = True
  Else
                                                  If SystemOptions.UserInterface = ArabicInterface Then
                                                   Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
                                                 Else
                                                  Accredit.Caption = " sent to Approval   "
                                              End If
                                              Accredit.Enabled = False
  End If
    
  
    Set RsDetails = New ADODB.Recordset


StrSQL = " SELECT     dbo.TblCardAuthorizationReformDetails.ID, dbo.TblCardAuthorizationReformDetails.Type, dbo.TblCardAuthorizationReformDetails.Mainte,"
StrSQL = StrSQL & "                      dbo.TblMaintenanceWork.name, dbo.TblMaintenanceWork.namee, dbo.TblCardAuthorizationReformDetails.EmpID, dbo.TblEmployee.Emp_Name AS fiter,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee AS fitere, dbo.TblCardAuthorizationReformDetails.empsuper, TblEmployee_1.Emp_Name AS NameSuper,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Namee AS NamesuperE, dbo.TblCardAuthorizationReformDetails.Deptid, dbo.TblEmpDepartments.DepartmentName,"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments.DepartmentNamee, dbo.TblCardAuthorizationReformDetails.Dpeterial, dbo.TblCardAuthorizationReformDetails.DeptBr,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.DeptColor, dbo.TblCardAuthorizationReformDetails.PriceFitter, dbo.TblCardAuthorizationReformDetails.payed,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.allocation, dbo.TblCardAuthorizationReformDetails.TimOut, dbo.TblCardAuthorizationReformDetails.TimeEnter,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.DateExit, dbo.TblCardAuthorizationReformDetails.DateEnter, dbo.TblCardAuthorizationReformDetails.finish,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.nohours, dbo.TblCardAuthorizationReformDetails.bill, dbo.TblCardAuthorizationReformDetails.comp,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.[count] , dbo.TblCardAuthorizationReformDetails.[value]"
StrSQL = StrSQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblMaintenanceWork ON dbo.TblCardAuthorizationReformDetails.Mainte = dbo.TblMaintenanceWork.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.TblCardAuthorizationReformDetails.Deptid = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblCardAuthorizationReformDetails.empsuper = TblEmployee_1.Emp_ID ON"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_id = dbo.TblCardAuthorizationReformDetails.EmpID"
StrSQL = StrSQL & "  Where (dbo.TblCardAuthorizationReformDetails.id  = " & val(XPTxtID.text) & ") And (dbo.TblCardAuthorizationReformDetails.Type = 0)"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
 

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
       With Me.Fg
      ' Deptid RsDetails.MoveFirst  RsDetails("PriceFitter").value = val(IIf((fg.TextMatrix(i, fg.ColIndex("PriceFitter"))), fg.TextMatrix(i, fg.ColIndex("PriceFitter")), 0))
        .Rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .Rows - 1
    
            .TextMatrix(i, .ColIndex("serial")) = i
            Fg.TextMatrix(i, Fg.ColIndex("dateenter")) = IIf(IsNull(RsDetails("DateEnter").value), "", RsDetails("DateEnter").value)
     
            Fg.TextMatrix(i, Fg.ColIndex("dateout")) = IIf(IsNull(RsDetails("DateExit").value), "", RsDetails("DateExit").value)
            Fg.TextMatrix(i, Fg.ColIndex("timEnter")) = IIf(IsNull(RsDetails("TimeEnter").value), "", RsDetails("TimeEnter").value)
            Fg.TextMatrix(i, Fg.ColIndex("TimOut")) = IIf(IsNull(RsDetails("TimOut").value), "", RsDetails("TimOut").value)
             .TextMatrix(i, .ColIndex("nohours")) = IIf(IsNull(RsDetails("nohours").value), "", RsDetails("nohours").value) 'Details("nohours").value
               .TextMatrix(i, .ColIndex("supervisor")) = IIf(IsNull(RsDetails("NameSuper").value), "", RsDetails("NameSuper").value) 'sDetails("supervisor").value
                 .TextMatrix(i, .ColIndex("workshop")) = IIf(IsNull(RsDetails("DepartmentName").value), "", RsDetails("DepartmentName").value) 'sDetails("workshop").value
                   .TextMatrix(i, .ColIndex("fitter")) = IIf(IsNull(RsDetails("fiter").value), "", RsDetails("fiter").value) 'sDetails("fitter").value
            .TextMatrix(i, .ColIndex("value")) = val(IIf(IsNull(RsDetails("Value").value), 0, RsDetails("Value").value)) 'RsDetails("Value").value
            ' .TextMatrix(i, .ColIndex("PriceFitter")) = val(IIf(IsNull(RsDetails("PriceFitter").value), 0, RsDetails("PriceFitter").value))
            .TextMatrix(i, .ColIndex("finish")) = IIf(IsNull(RsDetails("finish").value), "", RsDetails("finish").value) 'RsDetails("finish").value
             .TextMatrix(i, .ColIndex("cod")) = IIf(IsNull(RsDetails("Mainte").value), "", RsDetails("Mainte").value) 'RsDetails("Mainte").value
          .TextMatrix(i, .ColIndex("count")) = IIf(IsNull(RsDetails("count").value), "", RsDetails("count").value) 'RsDetails("count").value
           .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(RsDetails("EmpID").value), "", RsDetails("EmpID").value)
           .TextMatrix(i, .ColIndex("empsuper")) = IIf(IsNull(RsDetails("empsuper").value), "", RsDetails("empsuper").value)
            .TextMatrix(i, .ColIndex("Deptid")) = IIf(IsNull(RsDetails("Deptid").value), "", RsDetails("Deptid").value)
          .TextMatrix(i, .ColIndex("Dpeterial")) = IIf(IsNull(RsDetails("Dpeterial").value), "", RsDetails("Dpeterial").value) 'RsDetails("count").value
           .TextMatrix(i, .ColIndex("DeptColor")) = IIf(IsNull(RsDetails("DeptColor").value), "", RsDetails("DeptColor").value)
            .TextMatrix(i, .ColIndex("payed")) = IIf(IsNull(RsDetails("payed").value), 0, RsDetails("payed").value)
            .TextMatrix(i, .ColIndex("allocation")) = IIf(IsNull(RsDetails("allocation").value), 0, RsDetails("allocation").value)
               If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("name").value), "", RsDetails("name").value)
                Else
                   .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("namee").value), "", RsDetails("namee").value)
             End If
             .TextMatrix(i, .ColIndex("PriceFitter")) = val(IIf(IsNull(RsDetails("PriceFitter").value), 0, RsDetails("PriceFitter").value))
            RsDetails.MoveNext
         
        Next i
End With
    End If
      RsDetails.Close
    Set RsDetails = Nothing
  
    '//////////////////////////////////////////
    Set RsDetails1 = New ADODB.Recordset
 StrSQL = " SELECT     TOP 100 PERCENT dbo.TblCardAuthorizationReformDetails.ID, dbo.TblCardAuthorizationReformDetails.comp, dbo.TblCardAuthorizationReformDetails.Type,"
 StrSQL = StrSQL & "                     dbo.TblCardAuthorizationReformDetails.[Value], dbo.TblExtraExpeneses.name, dbo.TblExtraExpeneses.namee,dbo.TblExtraExpeneses.TypeExtrExpen, dbo.TblCardAuthorizationReformDetails.Mainte,"
StrSQL = StrSQL & "                       dbo.TblCardAuthorizationReformDetails.[count] , dbo.TblCardAuthorizationReformDetails.bill, dbo.TblExtraExpeneses.id"
StrSQL = StrSQL & "  FROM         dbo.TblCardAuthorizationReformDetails INNER JOIN"
StrSQL = StrSQL & "                       dbo.TblExtraExpeneses ON dbo.TblCardAuthorizationReformDetails.Mainte = dbo.TblExtraExpeneses.Id"
'StrSQL = StrSQL & "   Where (dbo.TblCardAuthorizationReformDetails.Type = 1) And (dbo.TblCardAuthorizationReformDetails.id =1)"
StrSQL = StrSQL & "  Where (dbo.TblCardAuthorizationReformDetails.id =" & val(XPTxtID.text) & ") And (dbo.TblCardAuthorizationReformDetails.Type = 1)"

    RsDetails1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
     If Not (RsDetails1.BOF Or RsDetails1.EOF) Then
       With Me.fg2
      '  RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails1.RecordCount

        For i = .FixedRows To .Rows - 1
    
            .TextMatrix(i, .ColIndex("serial")) = i
            .TextMatrix(i, .ColIndex("value")) = val(IIf(IsNull(RsDetails1("Value").value), 0, RsDetails1("Value").value)) 'RsDetails1("Value").value
             .TextMatrix(i, .ColIndex("cod")) = IIf(IsNull(RsDetails1("Mainte").value), "", RsDetails1("Mainte").value) 'RsDetails1("Mainte").value
            .TextMatrix(i, .ColIndex("count")) = IIf(IsNull(RsDetails1("count").value), "", RsDetails1("count").value) 'RsDetails1("count").value
             
           .TextMatrix(i, .ColIndex("comp")) = IIf(IsNull(RsDetails1("comp").value), "", RsDetails1("comp").value) 'RsDetails1("comp").value
           .TextMatrix(i, .ColIndex("bill")) = IIf(IsNull(RsDetails1("bill").value), "", RsDetails1("bill").value) ' RsDetails1("bill").value
           .TextMatrix(i, .ColIndex("typeexpen")) = IIf(IsNull(RsDetails1("TypeExtrExpen").value), "", RsDetails1("TypeExtrExpen").value)
               If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails1("name").value), "", RsDetails1("name").value)
                Else
                   .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails1("namee").value), "", RsDetails1("namee").value)
             End If
            RsDetails1.MoveNext
         
        Next i
End With
    End If
   

Dim s As String
s = "Select *,tblItems.ItemCode,tblItems.ItemName from TblCardAuthorizationReformItems Left Outer Join tblItems On tblItems.ItemID =TblCardAuthorizationReformItems.ItemID "
s = s & "  Where (dbo.TblCardAuthorizationReformItems.id =" & val(XPTxtID.text) & ")"
loadgrid s, FG22, True, True
  '  fillapprovData
    ReLineGrid

   '
 
RsDetails1.Close
 Set RsDetails1 = Nothing

  '  fillapprovData
    ReLineGrid
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
   ' rs.Close
    Exit Sub
ErrTrap:
End Sub
Function GetID(Optional Lngid As Double) As Double
Dim StrSQL As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
     StrSQL = "select ID From dbo.TblCardAuthorizationReform     where WorkOrder=" & Lngid & ""
       rs2.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If rs2.RecordCount > 0 Then
   GetID = IIf(IsNull(rs2("ID").value), 0, rs2("ID"))
   Else
   GetID = 0
   End If
End Function
Public Sub Retrive2(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
  clear_all Me
      If rs.State = adStateOpen Then
   rs.Close
   
   Else
'rs.Open
   
   End If

     
     StrSQL = "select * From dbo.TblCardAuthorizationReform     where ID=" & Lngid & ""
       rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
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
    'If TxtModFlg.text = "N" Then 'And ComGranty.ListIndex = 2 Then
   
If rs("OrderStatus").value <> 5 Then

MsgBox " ·«Ì„ﬂ‰ ≈⁄«œ… ⁄„·Ì… «·«’·«Õ ·«‰Â ·„  ’œ— ›« Ê—… „‰ ﬁ»· "
 clear_all Me
        imgg
            Me.Lbtotal.Caption = 0
            Me.LbToTalExtra.Caption = 0
            
            Me.lbTotalMente.Caption = 0
     Me.DcbOrderStatus.ListIndex = 0
    Me.ComGranty.ListIndex = 1
Exit Sub
'End If
End If
    XPTxtID.text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    'Me.TxtEndDate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    DcbCarType.BoundText = IIf(IsNull(rs("CarTypeID").value), "", rs("CarTypeID").value)
    DcbCarModel.BoundText = IIf(IsNull(rs("CarModelID").value), "", rs("CarModelID").value)
      TxtTypeCustomer.text = val(IIf(IsNull(rs("TypeCustomer").value), 0, rs("TypeCustomer").value))
     txtKM.text = IIf(IsNull(rs("OverKM").value), "", rs("OverKM").value)
         cmbItems.BoundText = IIf(IsNull(rs("ItemID33").value), "", rs("ItemID33").value)
   txtSalesInvoiceOrder = IIf(IsNull(rs("SalesInvoiceOrder").value), "", rs("SalesInvoiceOrder").value)
    

   ' DcboSpecifications.BoundText = IIf(IsNull(rs("gradeID").value), "", rs("gradeID").value)
   ' Me.TxtRemarkCar.text = IIf(IsNull(rs("Remarkcar").value), "", rs("Remarkcar").value)
    DcbColor.BoundText = IIf(IsNull(rs("ColorID").value), "", rs("ColorID").value)
    DcbyearFactor.text = IIf(IsNull(rs("YearFact").value), "", rs("YearFact").value)
   TxtClientPhone.text = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
   TxtCliientName.text = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
   cmdEndAll.Tag = val(rs!IsEndAll & "")
    If val(rs!IsEndAll & "") = 1 Then
    cmdEndAll.Enabled = False
    cmdEndAll.Caption = " „ «ﬁ›«· «·ﬂ«— "
    
   Else
     cmdEndAll.Enabled = True
     cmdEndAll.Caption = "«ﬁ›«· «·ﬂ«—  ‰Â«∆Ï"
   End If
    
   

   TxtPlatNo.text = IIf(IsNull(rs("PlateNo").value), "", rs("PlateNo").value)
   Me.TxtCodeComputer.text = IIf(IsNull(rs("CodeComputer").value), "", rs("CodeComputer").value)
   'DcbOrderStatus.ListIndex = IIf(IsNull(rs("OrderStatus").value), 0, rs("OrderStatus").value)
   'TXtCarMeter.text = IIf(IsNull(rs("CarMeter").value), "", rs("CarMeter").value)
   'TxtLongGranty.text = IIf(IsNull(rs("LongGranty").value), "", rs("LongGranty").value)
   'TxtFirstPrice.text = val(IIf(IsNull(rs("PayFirst").value), 0, rs("PayFirst").value))
   Me.TXtShaseh.text = IIf(IsNull(rs("Shaseh").value), "", rs("Shaseh").value)
  ' Me.TXtNotAccept.text = IIf(IsNull(rs("NotAccept").value), "", rs("NotAccept").value)
   'Me.TxtAmoutAccept.text = val(IIf(IsNull(rs("AmountAccept").value), 0, rs("AmountAccept").value))
   'DateStartG.value = IIf(IsNull(rs("DateStartG").value), Date, rs("DateStartG").value)
   'DateEndg.value = IIf(IsNull(rs("DateEndG").value), Date, rs("DateEndG").value)
   'Me.TxtComplaint.text = IIf(IsNull(rs("Complaint").value), "", rs("Complaint").value)
   'Me.TxtNoteIntial1.text = IIf(IsNull(rs("Noteinitial").value), "", rs("Noteinitial").value)
 '''   /////////////////////////////////////
 
    '    If rs("Cash").value = True Then
    '    Me.RdCash.value = True
    '    Else
    '     Me.RdCash.value = False
    '     End If
    '     If rs("Accoun").value = True Then
    '    Me.Rdacco.value = True
    ''    Else
    '     Me.Rdacco.value = False
    ''     End If
    ' If rs("credit").value = True Then
   '     Me.rdCredit.value = True
    '    Else
    '     Me.rdCredit.value = False
   '      End If
    '  Me.DcboFitter.BoundText = IIf(IsNull(rs("FitterID").value), "", rs("FitterID").value)
   'M 'e.XPDtbtimeTrans.value = rs("RecordeTime").value
    '   Me.txtmobile.text = rs("mobile").value
   ' '    Me.TxtBox.text = rs("box").value
   '     TxtClientCode.text = rs("ClientCode").value
   '     Me.txtFax.text = rs("fax").value
   '     Me.txtEmail.text = rs("email").value
   '      Me.txtAddres.text = rs("address").value
   '      Me.txtboxzip.text = rs("boxzip").value
   '      Me.txtCodeReg.text = rs("codereg").value
   '      Me.TxtTtpeReg.text = rs("typereg").value
   '     Me.TxtCodeDoor.text = rs("codedoor").value
   '     Me.TxtDriver.text = rs("driver").value
   '     Me.DTPEnterDate.value = rs("DateEnter").value
   ''      DTPDateAcutExite.value = rs("DateAcutExite").value
   '      DTPDateExptExit.value = rs("DateExptExit").value
   '      DTPTimeAcutExite.value = rs("TimeAcutExite").value
   '     DTPTimeExptExit.value = rs("TimeExptExit").value
             
         ' If rs("persons").value = 1 Then
        'Me.RdPerson.value = True
        'Else
        'Me.RdPerson.value = False
        ' End If
   
      'If rs("Companies").value = 1 Then
      '  Me.RdCompany.value = True
      '  Else
      '   Me.RdCompany.value = False
      '   End If
 '''''''''''///////////
        
        'Me.combtypereq.ListIndex = rs("typerequest").value
   '     If rs("Cash").value = True Then
   '     Me.RdCash.value = True
   '     Else
   '      Me.RdCash.value = False
   '      End If
   '      If rs("Accoun").value = True Then
   '     Me.Rdacco.value = True
   '     Else
   '      Me.Rdacco.value = False
   '      End If
   '  If rs("credit").value = True Then
   '     Me.rdCredit.value = True
   '     Else
   '      Me.rdCredit.value = False
   '      End If
   '    Me.DcboFitter.BoundText = IIf(IsNull(rs("FitterID").value), "", rs("FitterID").value)
       Me.TxtMobile.text = IIf(IsNull(rs("mobile").value), "", rs("mobile").value) ' rs("mobile").value
        Me.TxtBox.text = IIf(IsNull(rs("box").value), "", rs("box").value) 'rs("box").value
        Me.TxtFax.text = IIf(IsNull(rs("fax").value), "", rs("fax").value) 'rs("fax").value
        Me.TxtEmail.text = IIf(IsNull(rs("email").value), "", rs("email").value) ' rs("email").value
         Me.TxtAddres.text = IIf(IsNull(rs("address").value), "", rs("address").value) ' rs("address").value
         Me.txtboxzip.text = IIf(IsNull(rs("boxzip").value), "", rs("boxzip").value) 'rs("boxzip").value
         Me.txtCodeReg.text = IIf(IsNull(rs("codereg").value), "", rs("codereg").value) 'rs("codereg").value
         Me.TxtTtpeReg.text = IIf(IsNull(rs("typereg").value), "", rs("typereg").value) 'rs("typereg").value
         
        TxtCusID.text = val(IIf(IsNull(rs("CusID").value), 0, rs("CusID").value))
        If val(TxtCusID.text) = 0 Then
            Dim ss As String
            ss = "Select cusId From TblCustemers Where Code = N'" & Trim(TxtClientCode) & "'"
            Dim rsDummy As New ADODB.Recordset
            rsDummy.Open ss, Cn, adOpenStatic, adLockReadOnly
            If Not rsDummy.EOF Then
                TxtCusID.text = rsDummy!CusID & ""
            End If
        End If
        
    If SystemOptions.UserInterface = EnglishInterface Then
        StrSQL = "Select CusNamee ClientName FROM TblCustemers where CusId = " & val(TxtCusID.text)
        Dim rsDu As New ADODB.Recordset
        rsDu.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
        If Not rsDu.EOF Then
            TxtCliientName.text = IIf(IsNull(rsDu("ClientName").value), "", rsDu("ClientName").value)
        End If
        If Trim(TxtCliientName.text) = "" Then
            TxtCliientName.text = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
        End If
    End If
         
         
   '     Me.TxtCodeDoor.text = IIf(IsNull(rs("codedoor").value), "", rs("codedoor").value) 'rs("codedoor").value
   '     Me.TxtDriver.text = IIf(IsNull(rs("driver").value), "", rs("driver").value) 'rs("driver").value
   '     Me.DTPEnterDate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value) ' rs("EndDate").value
   '      DTPDateAcutExite.value = IIf(IsNull(rs("DateAcutExite").value), "", rs("DateAcutExite").value) ' rs("DateAcutExite").value
   '      DTPDateExptExit.value = IIf(IsNull(rs("DateExptExit").value), "", rs("DateExptExit").value) ' rs("DateExptExit").value
   '      DTPTimeAcutExite.value = IIf(IsNull(rs("TimeAcutExite").value), "", rs("TimeAcutExite").value) ' rs("TimeAcutExite").value
   '     DTPTimeExptExit.value = IIf(IsNull(rs("TimeExptExit").value), "", rs("TimeExptExit").value) 'rs("TimeExptExit").value
             
   '       If rs("persons").value = True Then
   '     Me.RdPerson.value = True
   '     Else
   '     Me.RdPerson.value = False
   '      End If
   
   '   If rs("Companies").value = True Then
   '     Me.RdCompany.value = True
   '     Else
   '      Me.RdCompany.value = False
   '      End If
   
   'If rs("Granty").value = True Then
   'Me.ComGranty.ListIndex = IIf(IsNull(rs("Granty").value), 0, rs("Granty").value)
 '  Me.frmgranty.Visible = True
   'Else
   'Me.ComGranty.ListIndex = 1
'   Me.frmgranty.Visible = False
   'End If
   'If rs("Month_Day").value = True Then
   'Me.ComMD.ListIndex = 0
   'Else
   'Me.ComMD.ListIndex = 1
   'End If
   ' If rs("Accept").value = True Then
    ' Me.ChAccept.value = vbChecked
    ' Me.DcbOrderStatus.ListIndex = 1
    ' Else
    '  Me.ChAccept.value = vbUnchecked
    '  End If
   '   If rs("subcar1").value = True Then
   '       Me.imag1.Picture = Me.img.Picture
'Else
' Me.imag1.Picture = Me.imgnul.Picture
'
'           End If
'            If rs("subcar2").value = True Then
'           Me.imag2.Picture = Me.img.Picture
'Else
' Me.imag2.Picture = Me.imgnul.Picture
'           End If
'            If rs("subcar3").value = True Then
'        Me.imag3.Picture = Me.img.Picture
'Else
' Me.imag3.Picture = Me.imgnul.Picture
'           End If
'            If rs("subcar4").value = True Then
'
'Me.imag4.Picture = Me.img.Picture
'Else
' Me.imag4.Picture = Me.imgnul.Picture
'           End If
''            If rs("subcar5").value = True Then
 '         Me.imag5.Picture = Me.img.Picture
'E ''lse
' Me.imag5.Picture = Me.imgnul.Picture
'           End If
'            If rs("subcar6").value = True Then
'       Me.img6.Picture = Me.img.Picture
'Else
' Me.img6.Picture = Me.imgnul.Picture
'           End If
''            If rs("subcar7").value = True Then
 '          Me.img7.Picture = Me.img.Picture
'E 'lse
' Me.img7.Picture = Me.imgnul.Picture
'           End If
'            If rs("subcar8").value = True Then
'          Me.img8.Picture = Me.img.Picture
'Else
' Me.img8.Picture = Me.imgnul.Picture
'           End If
'            If rs("subcar9").value = True Then
'          Me.img9.Picture = Me.img.Picture
'Else
' Me.img9.Picture = Me.imgnul.Picture
'           End If
'            If rs("subcar10").value = True Then
'           Me.img10.Picture = Me.img.Picture
'Else
' Me.img10.Picture = Me.imgnul.Picture
'           End If
'   ' TxtAdvanceValue.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
'  '  Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
'   ' Me.TxtPaymentCounts.text = IIf(IsNull(rs("PaymentCounts").value), "", rs("PaymentCounts").value)
'
'    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
'      If IsNull(rs("posted").value) Then
'                                                 If SystemOptions.UserInterface = ArabicInterface Then
'                                                  Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
'                                                 Else
'                                                    Accredit.Caption = " send to Approval   "
'                                               End If
'                                              Accredit.Enabled = True
'  Else
'                                                  If SystemOptions.UserInterface = ArabicInterface Then
'                                                   Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
'                                                 Else
'                                                  Accredit.Caption = " sent to Approval   "
'                                              End If
'                                              Accredit.Enabled = False
'  End If
'
'
'  '  Set RsDetails = New ADODB.Recordset
''StrSQL = " SELECT     TOP 100 PERCENT dbo.TblCardAuthorizationReformDetails.ID,dbo.TblCardAuthorizationReformDetails.count,dbo.TblCardAuthorizationReformDetails.Type, dbo.TblCardAuthorizationReformDetails.[Value],"
''            StrSQL = StrSQL & "          dbo.TblMaintenanceWork.name , dbo.TblMaintenanceWork.namee, dbo.TblCardAuthorizationReformDetails.Mainte"
'          StrSQL = StrSQL & "   FROM         dbo.TblCardAuthorizationReformDetails INNER JOIN"
'       StrSQL = StrSQL & "               dbo.TblMaintenanceWork ON dbo.TblCardAuthorizationReformDetails.Mainte = dbo.TblMaintenanceWork.Id"
'' StrSQL = StrSQL & "  Where (dbo.TblCardAuthorizationReformDetails.id =" & val(XPTxtID.text) & ") And (dbo.TblCardAuthorizationReformDetails.Type = 0)"
''StrSQL = StrSQL & "   ORDER BY dbo.TblCardAuthorizationReformDetails.Mainte"
''    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
''
''
'
'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
'       With Me.Fg
'      '  RsDetails.MoveFirst
'        .Rows = .FixedRows + RsDetails.RecordCount
''
''        For i = .FixedRows To .Rows - 1
''
'            .TextMatrix(i, .ColIndex("serial")) = i
'            .TextMatrix(i, .ColIndex("value")) = RsDetails("Value").value
 '            .TextMatrix(i, .ColIndex("cod")) = RsDetails("id").value
''          .TextMatrix(i, .ColIndex("count")) = RsDetails("count").value
 '
 '              If SystemOptions.UserInterface = ArabicInterface Then
 '                   .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("name").value), "", RsDetails("name").value)
 '               Else
 '                  .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("namee").value), "", RsDetails("namee").value)
 '            End If
 '           RsDetails.MoveNext
 '
 '       Next i
'End With
'    End If
'      RsDetails.Close
'    Set RsDetails = Nothing
'
'    '//////////////////////////////////////////
'    Set RsDetails1 = New ADODB.Recordset
' StrSQL = " SELECT     TOP 100 PERCENT dbo.TblCardAuthorizationReformDetails.ID, dbo.TblCardAuthorizationReformDetails.comp, dbo.TblCardAuthorizationReformDetails.Type,"
' StrSQL = StrSQL & "                     dbo.TblCardAuthorizationReformDetails.[Value], dbo.TblExtraExpeneses.name, dbo.TblExtraExpeneses.namee, dbo.TblCardAuthorizationReformDetails.Mainte,"
'StrSQL = StrSQL & "                       dbo.TblCardAuthorizationReformDetails.[count] , dbo.TblCardAuthorizationReformDetails.bill, dbo.TblExtraExpeneses.id"
'StrSQL = StrSQL & "  FROM         dbo.TblCardAuthorizationReformDetails INNER JOIN"
'StrSQL = StrSQL & "                       dbo.TblExtraExpeneses ON dbo.TblCardAuthorizationReformDetails.Mainte = dbo.TblExtraExpeneses.Id"
''StrSQL = StrSQL & "   Where (dbo.TblCardAuthorizationReformDetails.Type = 1) And (dbo.TblCardAuthorizationReformDetails.id =1)"
'StrSQL = StrSQL & "  Where (dbo.TblCardAuthorizationReformDetails.id =" & val(XPTxtID.text) & ") And (dbo.TblCardAuthorizationReformDetails.Type = 1)"
'
'    RsDetails1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'     If Not (RsDetails1.BOF Or RsDetails1.EOF) Then
'       With Me.fg2
'      '  RsDetails.MoveFirst
'        .Rows = .FixedRows + RsDetails1.RecordCount
'
'        For i = .FixedRows To .Rows - 1
'
'            .TextMatrix(i, .ColIndex("serial")) = i
''            .TextMatrix(i, .ColIndex("value")) = RsDetails1("Value").value
'             .TextMatrix(i, .ColIndex("cod")) = RsDetails1("id").value
'            .TextMatrix(i, .ColIndex("count")) = RsDetails1("count").value
'
''           .TextMatrix(i, .ColIndex("comp")) = RsDetails1("comp").value
 ''          .TextMatrix(i, .ColIndex("bill")) = RsDetails1("bill").value
  '             If SystemOptions.UserInterface = ArabicInterface Then
  '                  .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails1("name").value), "", RsDetails1("name").value)
  '              Else
  '                 .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails1("namee").value), "", RsDetails1("namee").value)
  '           End If
  '          RsDetails1.MoveNext
         
  '      Next i
'End With
  '  End If
   

   '
    
'RsDetails1.Close
' Set RsDetails1 = Nothing

    'fillapprovData
    'ReLineGrid
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
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

 '   On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        
    If Not SystemOptions.IsMaintItemMode Then
        
        If Me.DcbCarType.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄  «·„⁄œÂ/«·”Ì«—…!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.DcbCarType.SetFocus
       '     SendKeys "{F4}"
            Exit Sub
        End If
        End If
    If Me.TxtCliientName.text = "" Then
            Msg = "ÌÃ» «œŒ«· «”„ «·⁄„Ì·!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.TxtCliientName.SetFocus
           ' SendKeys "{F4}"
            Exit Sub
        End If
   
   Dim s As String
   s = "Select CusName,CusNamee from TblCustemers Where CusName = '" & Trim(TxtCliientName) & "' Or CusNamee = '" & Trim(TxtCliientName) & "'  "
   Dim rsDummyCus As New ADODB.Recordset
   rsDummyCus.Open s, Cn, adOpenStatic, adLockReadOnly
   If rsDummyCus.EOF Then
            Msg = "ÌÃ» «œŒ«· «”„ ⁄„Ì· „”Ã· „‰ ﬁ»·!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.TxtCliientName.SetFocus
           ' SendKeys "{F4}"
            Exit Sub
   End If
                                                                                     
                                                                                     
'''''''''''//////////


If opt(2).value = True Then
    ShowPriceOrder = val(TxtShowPriceOrder.text)
  If Me.Checked(, ShowPriceOrder, 0) = True Then
  Else
    ShowPriceOrder = 1
     maxx , ShowPriceOrder
     TxtShowPriceOrder.text = ShowPriceOrder
     End If
 ElseIf opt(1).value = True Then
        AuthoOrder = val(TxtAuthoOrder.text)
  If Me.Checked(, , AuthoOrder) = True Then
   Else
     AuthoOrder = 1
     maxx , , AuthoOrder
     TxtAuthoOrder.text = AuthoOrder
  End If
 
 Else
             WorkOrder = val(TxtWorkOrder.text)
  If Me.Checked(WorkOrder, 0, 0) = True Then
   Else
     WorkOrder = 1
     maxx WorkOrder
     TxtWorkOrder.text = WorkOrder
  End If
 
 End If
'''''''''//////

        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then

            XPTxtID.text = CStr(new_id("TblCardAuthorizationReform", "ID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
'                Dim CUSTID As Double
'createCustomer TxtCliientName, TxtCliientName, 1, CUSTID
'            TxtClientCode.text = CUSTID
            
            rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
            StrSQL = "Delete From TblCardAuthorizationReformDetails Where ID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblCarOrderVouchers Where ORderID =" & val(Me.TxtWorkOrder.text)
            Cn.Execute StrSQL, , adExecuteNoRecords

        End If
'rs.Open
      rs("BranchID").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
 
        rs("ID").value = val(XPTxtID.text)
rs("TypeCustomer").value = val(Me.TxtTypeCustomer.text)
rs("OverKM").value = Me.txtKM.text
            
         rs("ClientCode").value = TxtClientCode.text
         rs("RecordDate").value = XPDtbTrans.value
         rs("RecordeTime").value = FormatDateTime(Me.DTPicker1.value, vbShortTime)
         rs("CarMetarOut").value = TxtCarMetarOut.text
         'rs("EndDate").value = Me.TxtEndDate.value
        rs("ClientName").value = Me.TxtCliientName.text
        rs("PrivateCop").value = Me.txtprivate.text
        rs("ReComentClient").value = Me.txtrecomment.text
        rs("Telephone").value = Me.TxtClientPhone.text
        rs("CarTypeID").value = val(Me.DcbCarType.BoundText)
        rs("CarModelID").value = val(Me.DcbCarModel.BoundText)
        rs("PlateNo").value = Me.TxtPlatNo.text
        rs("OrderStatus").value = Me.DcbOrderStatus.ListIndex
        rs("ColorID").value = val(Me.DcbColor.BoundText)
        rs("YearFact").value = val(Me.DcbyearFactor.text)
        
                rs("ItemID33").value = val(Me.cmbItems.BoundText)
        rs("SalesInvoiceOrder").value = Trim(Me.txtSalesInvoiceOrder.text)
    

        rs("LongGranty").value = Me.TxtLongGranty.text
        rs("CarMeter").value = Me.TXtCarMeter.text
        rs("DateStartG").value = Me.DateStartG.value
        rs("DateEndG").value = Me.DateEndg.value
        rs("PayFirst").value = val(Me.TxtFirstPrice.text)
        rs("Noteinitial").value = Me.TxtNoteIntial1.text
        rs("Complaint").value = Me.TxtComplaint.text
        rs("Shaseh").value = Me.TXtShaseh.text
        rs("NotAccept").value = txtnotacept.text
        rs("SparePart").value = TxtSparePart.text
        rs("CodeComputer").value = Me.TxtCodeComputer.text
        rs("AmountAccept").value = val(Me.TxtAmoutAccept.text)
        rs("LastWorOrder").value = val(Me.TxtLastWorOrder.text)
        rs("CarID").value = val(Me.DcbCar.BoundText)
      ' /////////////////////////////////////CodeComputer
        rs("ResonUnderWait").value = Me.txtresonwait.text
      '  rs("typerequest").value = val(Me.combtypereq)
        If Me.RdCash.value = True Then
        rs("Cash").value = 1
        Else
         rs("Cash").value = 0
         End If
         If Me.Rdacco.value = True Then
        rs("Accoun").value = 1
        Else
         rs("Accoun").value = 0
         End If
         If Me.rdCredit.value = True Then
        rs("credit").value = 1
        Else
         rs("credit").value = 0
         End If
        rs("FitterID").value = IIf(Me.DcboFitter.BoundText = "", Null, Me.DcboFitter.BoundText)
         rs("CusID").value = val(TxtCusID.text)
        
        rs("mobile").value = Me.TxtMobile.text
        rs("box").value = Me.TxtBox.text
        rs("fax").value = Me.TxtFax.text
         rs("email").value = Me.TxtEmail.text
          rs("address").value = Me.TxtAddres.text
        rs("boxzip").value = Me.txtboxzip.text
        rs("codereg").value = Me.txtCodeReg.text
        rs("typereg").value = Me.TxtTtpeReg.text
        rs("codedoor").value = Me.TxtCodeDoor.text
        rs("driver").value = Me.TxtDriver.text
         rs("EndDate").value = Me.DTPEnterDate.value
         rs("Remarkcar").value = Me.TxtRemarkCar.text
         If DcbOrderStatus.ListIndex = 2 Then
          DTPTimeAcutExite.value = Time
          
           Me.DTPDateAcutExite.value = Date
             
          '  MsgBox " „  ÕœÌÀ  «—Œ «·Œ—ÊÃ «·›⁄·Ì"
            End If
            
              rs("TimeAcutExite").value = DTPTimeAcutExite.value
          rs("DateAcutExite").value = Me.DTPDateAcutExite.value
          
             rs("TimeExptExit").value = DTPTimeExptExit.value
              rs("DateExptExit").value = DTPDateExptExit.value
             
          If Me.RdPerson.value = True Then
        rs("persons").value = 1
        Else
         rs("persons").value = 0
         End If
            If Me.RdCompany.value = True Then
        rs("Companies").value = 1
        Else
         rs("Companies").value = 0
         End If
        If Me.ChAccept.value = xtpChecked Then
     
           rs("Accept").value = 1
           Else
           rs("Accept").value = 0
        End If
     If Me.CheckBox1.value = xtpChecked Then
     
           rs("wait").value = 1
           Else
           rs("wait").value = 0
        End If
          If Me.CheckBox2.value = xtpChecked Then
     
           rs("notAcepted").value = 1
           Else
           rs("notAcepted").value = 0
        End If
      '    If Me.ComGranty.ListIndex = 0 Then
          rs("Granty").value = val(Me.ComGranty.ListIndex)
          
       '   Else
       '   rs("Granty").value = 0
       '
        '  End If
      ''///////////16 11 2015
      rs("WorkOrder").value = val(Me.TxtWorkOrder.text)
      rs("ShowPriceOrder").value = val(Me.TxtShowPriceOrder.text)
      rs("AuthoOrder").value = val(Me.TxtAuthoOrder.text)
      
      rs("DiscValue").value = val(Me.txtDiscValue.text)
      rs("DiscPercent").value = val(Me.txtDiscPercent.text)
      rs("TotalAfterDiscount").value = val(Me.txtTotalAfterDiscount.text)
      rs("Vatyo").value = val(Me.txtVatyo.text)
      rs("Vat2").value = val(Me.txtVat2.text)
      

                        
                         
      If opt(0).value = True Then
      rs("TypeOrder").value = 0
      ElseIf opt(1).value = True Then
      rs("TypeOrder").value = 1
      ElseIf opt(2).value = True Then
      rs("TypeOrder").value = 2
      Else
      rs("notAcepted").value = Null
      End If
      ''''/////////
         If Me.ComMD.ListIndex = 0 Then
         rs("Month_Day").value = 1
         Else
          rs("Month_Day").value = 0
         End If
           If Me.imag1.Picture = Me.Img.Picture Then
           rs("subcar1").value = 1
           Else
           rs("subcar1").value = 0
           End If
           If Me.imag2.Picture = Me.Img.Picture Then
           rs("subcar2").value = 1
           Else
           rs("subcar2").value = 0
           End If
           If Me.imag3.Picture = Me.Img.Picture Then
           rs("subcar3").value = 1
           Else
          rs("subcar3").value = 0
           End If
           If Me.imag4.Picture = Me.Img.Picture Then
           rs("subcar4").value = 1
           Else
           rs("subcar4").value = 0
           End If
           If Me.imag5.Picture = Me.Img.Picture Then
           rs("subcar5").value = 1
           Else
           rs("subcar5").value = 0
           End If
           If Me.img6.Picture = Me.Img.Picture Then
           rs("subcar6").value = 1
        Else
           rs("subcar6").value = 0
           End If
           If Me.img7.Picture = Me.Img.Picture Then
           rs("subcar7").value = 1
           Else
           rs("subcar7").value = 0
           End If
           If Me.img8.Picture = Me.Img.Picture Then
           rs("subcar8").value = 1
           Else
           rs("subcar8").value = 0
           End If
           If Me.img9.Picture = Me.Img.Picture Then
           rs("subcar9").value = 1
           Else
           rs("subcar9").value = 0
           End If
           If Me.img10.Picture = Me.Img.Picture Then
           rs("subcar10").value = 1
           Else
           rs("subcar10").value = 0
           End If
           ''/////////
            If Me.img11.Picture = Me.Img.Picture Then
           rs("subcar11").value = 1
           Else
           rs("subcar11").value = 0
           End If
            If Me.img12.Picture = Me.Img.Picture Then
           rs("subcar12").value = 1
           Else
           rs("subcar12").value = 0
           End If
            If Me.img13.Picture = Me.Img.Picture Then
           rs("subcar13").value = 1
           Else
           rs("subcar13").value = 0
           End If
            If Me.img14.Picture = Me.Img.Picture Then
           rs("subcar14").value = 1
           Else
           rs("subcar14").value = 0
           End If
        rs("OrderStatus").value = Me.DcbOrderStatus.ListIndex
        rs("UserID").value = IIf(Me.DCboUserName.BoundText = "", Null, Me.DCboUserName.BoundText)

        rs.update
        '''''''''/////////////////////////////////
            Set RsDetails1 = New ADODB.Recordset
    StrSQL = "SELECT     *  from dbo.TblCarOrderVouchers Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
With vchrgrid
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("ID"))) <> 0 Then
RsDetails1.AddNew
RsDetails1("ORderID").value = val(TxtWorkOrder.text)
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
        ''/////////////////////////
        
      Set RsDetails = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblCardAuthorizationReformDetails Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     '  RsDetails.Open "TblCardAuthorizationReformDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
If Fg.Rows > 1 Then
          
       For i = Me.Fg.FixedRows To Fg.Rows - 1
       If Fg.TextMatrix(i, Fg.ColIndex("name")) <> "" Then
        If Fg.TextMatrix(i, Fg.ColIndex("workshop")) = "" Then
            Msg = "ÌÃ» «Œ Ì«— «·ﬁ”„!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
           ' Me.TxtCliientName.SetFocus
           ' SendKeys "{F4}"
           Cn.CommitTrans
            Exit Sub
        End If
       '  If FG.TextMatrix(i, FG.ColIndex("supervisor")) = "" Then
       '     Msg = "ÌÃ» «Œ Ì«— «·„‘—›!! "
       '     MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       '    ' Me.TxtCliientName.SetFocus
       '    ' SendKeys "{F4}"
       '    Cn.CommitTrans
       '     Exit Sub
       ' End If
       '          If FG.TextMatrix(i, FG.ColIndex("fitter")) = "" Then
       '     Msg = "ÌÃ» «Œ Ì«— «·›‰Ì!! "
       '     MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
           ' Me.TxtCliientName.SetFocus
           ' SendKeys "{F4}"
        '   Cn.CommitTrans
        '    Exit Sub
       ' End If
        If val(Fg.TextMatrix(i, Fg.ColIndex("Deptid"))) <> 0 Then
           RsDetails.AddNew
          RsDetails("ID").value = val(XPTxtID.text)
        RsDetails("Value").value = val(Fg.TextMatrix(i, Fg.ColIndex("Value")))
        RsDetails("nohours").value = Fg.TextMatrix(i, Fg.ColIndex("nohours"))
        
        'MsgBox val(Fg.TextMatrix(i, Fg.ColIndex("finish")))
      ' If val(Fg.TextMatrix(i, Fg.ColIndex("finish"))) <> 0 Or Fg.TextMatrix(i, Fg.ColIndex("finish")) = True Then 'Or fg.TextMatrix(i, fg.ColIndex("finish")) <> "" Then
       If Fg.Cell(flexcpChecked, i, Fg.ColIndex("finish")) = flexChecked Then
       RsDetails("finish").value = -1
        Fg.TextMatrix(i, Fg.ColIndex("dateout")) = Date
         Fg.TextMatrix(i, Fg.ColIndex("TimOut")) = Time
       Else
       RsDetails("finish").value = 0
       
       End If
     '  MsgBox fg.TextMatrix(i, fg.ColIndex("EmpID"))
     ' RsDetails("EmpID").value = IIf((fg.TextMatrix(i, fg.ColIndex("EmpID"))), fg.TextMatrix(i, fg.ColIndex("EmpID")), Null)
       RsDetails("Deptid").value = val(IIf(IsNull((Fg.TextMatrix(i, Fg.ColIndex("Deptid")))), 0, Fg.TextMatrix(i, Fg.ColIndex("Deptid"))))
       RsDetails("empsuper").value = val(IIf(IsNull((Fg.TextMatrix(i, Fg.ColIndex("empsuper")))), 0, Fg.TextMatrix(i, Fg.ColIndex("empsuper"))))
      RsDetails("EmpID").value = val(IIf(IsNull((Fg.TextMatrix(i, Fg.ColIndex("EmpID")))), 0, Fg.TextMatrix(i, Fg.ColIndex("EmpID"))))
       RsDetails("Dpeterial").value = val(IIf(IsNull((Fg.TextMatrix(i, Fg.ColIndex("Dpeterial")))), 0, Fg.TextMatrix(i, Fg.ColIndex("Dpeterial"))))
        RsDetails("DeptColor").value = IIf(IsNull((Fg.TextMatrix(i, Fg.ColIndex("DeptColor")))), "", Fg.TextMatrix(i, Fg.ColIndex("DeptColor")))
' RsDetails("EmpID").value = IIf(IsNull(fg.TextMatrix(i, fg.ColIndex("EmpID"))), "", fg.TextMatrix(i, fg.ColIndex("EmpID")))
       ' RsDetails("fitter").value = Fg.TextMatrix(i, Fg.ColIndex("fitter"))
       ' RsDetails("supervisor").value = Fg.TextMatrix(i, Fg.ColIndex("supervisor"))
       'RsDetails("workshop").value = Fg.TextMatrix(i, Fg.ColIndex("workshop"))
        RsDetails("DateEnter").value = IIf(IsDate(Fg.TextMatrix(i, Fg.ColIndex("dateenter"))), Fg.TextMatrix(i, Fg.ColIndex("dateenter")), Null)
   
       RsDetails("DateExit").value = IIf(IsDate(Fg.TextMatrix(i, Fg.ColIndex("dateout"))), Fg.TextMatrix(i, Fg.ColIndex("dateout")), Null)
       RsDetails("TimOut").value = Fg.TextMatrix(i, Fg.ColIndex("TimOut"))
       
        RsDetails("TimeEnter").value = Fg.TextMatrix(i, Fg.ColIndex("timEnter")) ' IIf((Fg.TextMatrix(i, Fg.ColIndex("timEnter"))), Fg.TextMatrix(i, Fg.ColIndex("timEnter")), Null)
        RsDetails("PriceFitter").value = val(IIf(IsNull((Fg.TextMatrix(i, Fg.ColIndex("PriceFitter")))), 0, Fg.TextMatrix(i, Fg.ColIndex("PriceFitter"))))
            RsDetails("Type").value = 0
            
           RsDetails("Mainte").value = val(Fg.TextMatrix(i, Fg.ColIndex("cod")))
          ' RsDetails("allocation").value = 0
             RsDetails("payed").value = 0
           If val(Fg.TextMatrix(i, Fg.ColIndex("count"))) <> 0 Then
           RsDetails("count").value = val(Fg.TextMatrix(i, Fg.ColIndex("count")))
           Else
           RsDetails("count").value = 1
           End If
         RsDetails.update
    End If
        End If
        Next i
        End If
        '''''''''''''''//////////////////////////
        
      Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from dbo.TblCardAuthorizationReformDetails Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      ' RsDetails1.Open "TblCardAuthorizationReformDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
If fg2.Rows > 1 Then
                ' fg2.Rows = fg2.Rows - 1
     
       For i = Me.fg2.FixedRows To fg2.Rows - 1
       If val(fg2.TextMatrix(i, fg2.ColIndex("cod"))) <> 0 Then
              If fg2.TextMatrix(i, fg2.ColIndex("name")) = "" Then
            Msg = "ÌÃ» «Œ Ì«—  «·«”„!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
           ' Me.TxtCliientName.SetFocus
           ' SendKeys "{F4}"
           Cn.CommitTrans
            Exit Sub
        End If
       ' For i = Me.fg2.FixedRows To fg2.Rows - 1
              If fg2.TextMatrix(i, fg2.ColIndex("typeexpen")) = "" Then
            Msg = "ÌÃ» «Œ Ì«— ‰Ê⁄ «·„‘ —Ì«  Ê«·«⁄„«· «·Œ«—ÃÌÂ!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
           ' Me.TxtCliientName.SetFocus
           ' SendKeys "{F4}"
           Cn.CommitTrans
            Exit Sub
        End If
       
           RsDetails1.AddNew
          RsDetails1("ID").value = val(XPTxtID.text)
        RsDetails1("Value").value = val(fg2.TextMatrix(i, fg2.ColIndex("Value")))
            RsDetails1("Type").value = 1
           RsDetails1("Mainte").value = val(fg2.TextMatrix(i, fg2.ColIndex("cod")))
           RsDetails1("Codtype").value = val(fg2.TextMatrix(i, fg2.ColIndex("Codtype")))
           RsDetails1("bill").value = fg2.TextMatrix(i, fg2.ColIndex("bill"))
           RsDetails1("comp").value = fg2.TextMatrix(i, fg2.ColIndex("comp"))
           If val(fg2.TextMatrix(i, fg2.ColIndex("count"))) <> 0 Then
           RsDetails1("count").value = val(fg2.TextMatrix(i, fg2.ColIndex("count")))
           Else
           RsDetails1("count").value = 1
           End If
         RsDetails1.update
     
       End If
           Next i
        End If
        
'Dim s As String
s = " delete TblCardAuthorizationReformItems "
s = s & "  Where (dbo.TblCardAuthorizationReformItems.id =" & val(XPTxtID.text) & ")"
Cn.Execute s
s = "Select * from TblCardAuthorizationReformItems "
s = s & "  Where (dbo.TblCardAuthorizationReformItems.id =" & val(XPTxtID.text) & ")"
saveGrid s, FG22, "ItemID", "", "id", val(Me.XPTxtID.text)
 
     
        
        '''''''''''''''//////////////////////////
        
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
        
'        End If
    
        Cn.CommitTrans
        BeginTrans = False
       RsDetails.Close
     
        Set RsDetails = Nothing
          RsDetails1.Close
        Set RsDetails1 = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
                
'SaveQRCode "TblCardAuthorizationReform", "ID", val(XPTxtBillID), TxtNoteSerial1.Text, (XPDtbBill.value), _
'        (LblFinal.Caption), Picture1, 0, (LblValueAdded.Caption), (LblFinal.Caption)



        Select Case Me.TxtModFlg.text

            Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
      Else
        Msg = "Saved  " & CHR(13)
                Msg = Msg + "Need new transaction y/n"
      
      End If
      
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                           If SystemOptions.UserInterface = ArabicInterface Then
                                         MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Else
                                         MsgBox "saed succes", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            
                            End If
                            
                 Case "p"
                '     MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                If SystemOptions.UserInterface = ArabicInterface Then
                                         MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Else
                                         MsgBox "saed succes", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            
                            End If
                            
        End Select

        TxtModFlg.text = "R"
    End If

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
   Retrive val(Me.XPTxtID.text)
End Sub

Private Sub SaveData1()
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



        Cn.BeginTrans
        BeginTrans = True


        
      Set RsDetails = New ADODB.Recordset
          StrSQL = "SELECT     *  from dbo.TblCardAuthorizationReformDetails Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     '  RsDetails.Open "TblCardAuthorizationReformDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
If Fg.Rows > 2 Then
          
       For i = Me.Fg.FixedRows To Fg.Rows - 2
         If val(Fg.TextMatrix(i, Fg.ColIndex("cod"))) <> 0 Then
           RsDetails.AddNew
          RsDetails("ID").value = val(XPTxtID.text)
        RsDetails("Value").value = val(Fg.TextMatrix(i, Fg.ColIndex("Value")))
        RsDetails("nohours").value = Fg.TextMatrix(i, Fg.ColIndex("nohours"))
        
      '  MsgBox val(fg.TextMatrix(i, fg.ColIndex("finish")))
       If val(Fg.TextMatrix(i, Fg.ColIndex("finish"))) <> 0 Then
     
       RsDetails("finish").value = -1
        Fg.TextMatrix(i, Fg.ColIndex("dateout")) = Date
         Fg.TextMatrix(i, Fg.ColIndex("TimOut")) = Time
       Else
       RsDetails("finish").value = 0
       
       End If
     '  MsgBox fg.TextMatrix(i, fg.ColIndex("EmpID"))
     ' RsDetails("EmpID").value = IIf((fg.TextMatrix(i, fg.ColIndex("EmpID"))), fg.TextMatrix(i, fg.ColIndex("EmpID")), Null)
      RsDetails("EmpID").value = val(IIf(IsNull((Fg.TextMatrix(i, Fg.ColIndex("EmpID")))), 0, Fg.TextMatrix(i, Fg.ColIndex("EmpID"))))
' RsDetails("EmpID").value = IIf(IsNull(fg.TextMatrix(i, fg.ColIndex("EmpID"))), "", fg.TextMatrix(i, fg.ColIndex("EmpID")))
        RsDetails("fitter").value = Fg.TextMatrix(i, Fg.ColIndex("fitter"))
        RsDetails("supervisor").value = Fg.TextMatrix(i, Fg.ColIndex("supervisor"))
       RsDetails("workshop").value = Fg.TextMatrix(i, Fg.ColIndex("workshop"))
        RsDetails("DateEnter").value = IIf(IsDate(Fg.TextMatrix(i, Fg.ColIndex("dateenter"))), Fg.TextMatrix(i, Fg.ColIndex("dateenter")), Null)
   
       RsDetails("DateExit").value = IIf(IsDate(Fg.TextMatrix(i, Fg.ColIndex("dateout"))), Fg.TextMatrix(i, Fg.ColIndex("dateout")), Null)
       RsDetails("TimOut").value = Fg.TextMatrix(i, Fg.ColIndex("TimOut"))
       
        RsDetails("TimeEnter").value = Fg.TextMatrix(i, Fg.ColIndex("timEnter")) ' IIf((Fg.TextMatrix(i, Fg.ColIndex("timEnter"))), Fg.TextMatrix(i, Fg.ColIndex("timEnter")), Null)
        RsDetails("PriceFitter").value = val(IIf(IsNull((Fg.TextMatrix(i, Fg.ColIndex("PriceFitter")))), 0, Fg.TextMatrix(i, Fg.ColIndex("PriceFitter"))))
            RsDetails("Type").value = 0
           RsDetails("Mainte").value = val(Fg.TextMatrix(i, Fg.ColIndex("cod")))
           RsDetails("allocation").value = 0
           If val(Fg.TextMatrix(i, Fg.ColIndex("count"))) <> 0 Then
           RsDetails("count").value = val(Fg.TextMatrix(i, Fg.ColIndex("count")))
           Else
           RsDetails("count").value = 1
           End If
         RsDetails.update
        
        End If
        Next i
        End If
     
    
        Cn.CommitTrans
        BeginTrans = False
       RsDetails.Close
     
        Set RsDetails = Nothing
          RsDetails1.Close
        Set RsDetails1 = Nothing
       


            
              
                     MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       

       

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
Dim txtID As String
txtID = Me.XPTxtID
    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            
            Me.DcbScreen.ListIndex = 1
            DcbScreen_Click
            Retrive val(txtID)
            XPBtnMove_Click (1)
            Me.TxtModFlg.text = "R"

        Case "E"
            rs.find "ID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If
Me.DcbScreen.ListIndex = 1
            DcbScreen_Click
            Retrive val(txtID)
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
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
 StrSQL1 = "Delete From TblCardAuthorizationReformDetails Where ID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL1, , adExecuteNoRecords
            StrSQL1 = "Delete From TblCarOrderVouchers where ORderID =" & val(Me.TxtWorkOrder.text)
            Cn.Execute StrSQL1, , adExecuteNoRecords
                If rs.RecordCount < 1 Then
                 Me.ChAccept.value = xtpUnchecked
                   Me.CheckBox1.value = xtpUnchecked
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



'Function fillapprovData()
'Dim Num As Integer
' Dim RsDetails As New ADODB.Recordset
' Dim StrSQL As String
'
 
' StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
'StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
'StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
'StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
''StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
'StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
'StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.name & "')"
'StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

'    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 'If Not (RsDetails.EOF Or RsDetails.BOF) Then
 '       GRID2.Rows = RsDetails.RecordCount + 1
'
'
'        For Num = 1 To RsDetails.RecordCount
'
'       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
'    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
'  GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
''   Else
 '   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
 '   End If
 '       GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
 '          If SystemOptions.UserInterface = ArabicInterface Then
 '          GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
 '         Else
 '            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
 '         End If
 '          If SystemOptions.UserInterface = ArabicInterface Then
 '           GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
 '         Else
 '          GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
 '          End If
 '          GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
 '        GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 '
 
'RsDetails.MoveNext 'If Num = RsDetails.RecordCount Then

'       If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
'                                If SystemOptions.UserInterface = ArabicInterface Then
'                                     Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
'                                 Else
'                                       Label11.Caption = "Approved"
'                                 End If
'                          Label11.backcolor = &H80FF80
'        Else
'                             If SystemOptions.UserInterface = ArabicInterface Then
'                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
'                            Else
'                                     Label11.Caption = "Currently required Approve"
'                            End If
'                 Label11.backcolor = &HFFFFC0
'        End If

'End If

'        Next Num
'Else
' GRID2.Rows = 1
'    End If
'RsDetails.Close

'End Function
 
Function cheh(Optional Ch As Boolean)
Dim i As Integer
For i = 1 To Fg.Rows - 1
If Fg.TextMatrix(i, Fg.ColIndex("finish")) <> "" Then
 If Fg.Cell(flexcpChecked, i, Fg.ColIndex("finish")) = flexChecked Then
'If fg.TextMatrix(i, fg.ColIndex("finish")) <> "" Then
'If fg.TextMatrix(i, fg.ColIndex("finish")) = True Then
Ch = True
Else
Ch = False
Exit Function
End If
End If
Next i
End Function
Sub GetInformationOfCustomerCar(Optional CarID As Double)
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = "Select * from TblCusCar where ID=" & CarID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
TxtDriver.text = IIf(IsNull(rs2("DriverName").value), "", rs2("DriverName").value)
TXtShaseh.text = IIf(IsNull(rs2("ChasisNo").value), "", rs2("ChasisNo").value)
DcbCarType.BoundText = IIf(IsNull(rs2("BrandID").value), 0, rs2("BrandID").value)
DcbyearFactor.ListIndex = IIf(IsNull(rs2("ModelID").value), -1, rs2("ModelID").value)
TxtPlatNo.text = DcbCar.text
DcbCarModel.BoundText = IIf(IsNull(rs2("CarModelID").value), 0, rs2("CarModelID").value)
DcbColor.BoundText = IIf(IsNull(rs2("ColorID").value), 0, rs2("ColorID").value)
Else
TXtShaseh.text = ""
TxtDriver.text = ""
DcbCarType.BoundText = 0
DcbColor.BoundText = 0
DcbCarModel.BoundText = 0
DcbyearFactor.ListIndex = -1
End If
End Sub
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
        .Create Me.hWnd, " »ÿ«ﬁ… ≈–‰ ≈’·«Õ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
     With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘… «·«Ê«„— «·„› ÊÕ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(7), " ..." & Wrap & "   " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
     With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘… «· ”·Ì„ ··⁄„Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(3), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
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
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   «·⁄„Ê·«  «·„” Õﬁ… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(9), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
           With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    „·› «·⁄„·«¡   ", 1, 15204351, -2147483630
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
    With TTP
        .Create Me.hWnd, "»ÿ«ﬁ… ≈–‰ ≈’·«Õ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»ÿ«ﬁ… ≈–‰ ≈’·«Õ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»ÿ«ﬁ… ≈–‰ ≈’·«Õ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " »ÿ«ﬁ… ≈–‰ ≈’·«Õ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " »ÿ«ﬁ… ≈’·«Õ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "»ÿ«ﬁ… ≈–‰ ≈’·«Õ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»ÿ«ﬁ… ≈–‰ ≈’·«Õ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»ÿ«ﬁ… ≈–‰ ≈’·«Õ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»ÿ«ﬁ… ≈–‰ ≈’·«Õ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»ÿ«ﬁ… ≈–‰ ≈’·«Õ ", 1, 15204351, -2147483630
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

 



Private Sub AddNewFgRow()

    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long

    If val(Me.DcboItems.BoundText) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ «Ú”„ «·’‰› ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.DcboItems.SetFocus
        Exit Sub
    End If

    If Me.TxtModFlg.text = "E" Then
        If val(Me.DcboItems.BoundText) = val(Me.XPTxtID.text) Then
            Msg = "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ «·’‰› Ã“¡ „‰ ‰›”Â....!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.DcboItems.SetFocus
            Exit Sub
        End If
    End If

 


    With Me.Fg
'        LngFindRow = .FindRow(val(Me.DCboItemS.BoundText), .FixedRows, .ColIndex("ItemID"), False, True)
'
'        If LngFindRow <> -1 Then
'            Msg = "Â–« «·’‰› „ÊÃÊœ ›⁄·« ...!!!"
'            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            .SetFocus
'            Exit Sub
'        End If

    End With

    LngNewRow = ModFgLib.SetFgForNewRow(FG22, FG22.ColIndex("ItemID"))
    With Me.FG22
        .TextMatrix(LngNewRow, .ColIndex("ItemID")) = Me.DcboItems.BoundText
    
        .TextMatrix(LngNewRow, .ColIndex("ItemCode")) = Trim$(Me.TxtItemCode.text)
        .TextMatrix(LngNewRow, .ColIndex("ItemName")) = Me.DcboItems.text
    
       ' .TextMatrix(LngNewRow, .ColIndex("UnitId")) = Me.dcItemunit.BoundText
       ' .TextMatrix(LngNewRow, .ColIndex("UnitName")) = Me.dcItemunit.Text
        
        .TextMatrix(LngNewRow, .ColIndex("Qty")) = val(Me.Txtqty.text)
        .TextMatrix(LngNewRow, .ColIndex("Price")) = val(Me.TxtItemPrice.text)
        
        .TextMatrix(LngNewRow, .ColIndex("BeforeVat")) = val(txtTotal)
        
        
        .AutoSize 0, .Cols - 1, False
    End With

    

    Me.TxtItemCode.text = ""
    Me.DcboItems.BoundText = ""
    TxtItemPrice.text = ""
    Txtqty.text = ""
    txtTotal.text = ""
    
    
    Me.TxtItemCode.SetFocus
End Sub


   
Public Sub CalculteValueAdded(LongRow As Long, Optional TransType As Integer, Optional flg As Integer = 0, Optional AllItems As Integer = 0, Optional posDelete As Boolean = False)
TransType = 21
'If SystemOptions.PriceWithVAT = True And (TransType = 21 Or TransType = 9) Then Exit Sub
'If (TxtModFlg.Text = "R" Or TxtModFlg.Text = "" Or val(Me.FG22.TextMatrix(LongRow, FG22.ColIndex("ItemID"))) = 0) And posDelete = False Then Exit Sub
 Dim Percentg As Double
Dim LngItemID As Double
Dim cCompanyInfo As New ClsCompanyInfo
Dim AccountVATCreit As String
If True = True Then
'If TransType = 9 And ReturnSales = True Then

  
    'LngItemID = val(Me.FG22.TextMatrix(LongRow, FG22.ColIndex("ItemID")))
    If SystemOptions.AllItemInVAT = True Then
        Percentg = val(cCompanyInfo.VATItems)
    Else
      PercentgValueAddedAccount_Transec XPDtbTrans.value, 21, 0, AccountVATCreit, Percentg
        
    End If
'
'    If Percentg = -1 Then
'        Percentg = 0
'        If SystemOptions.UserInterface = ArabicInterface Then
'            FG22.TextMatrix(LongRow, FG22.ColIndex("TypeVAT")) = "„⁄›Ì"
'        Else
'            FG22.TextMatrix(LongRow, FG22.ColIndex("TypeVAT")) = "Exempt"
'        End If
'    Else
'        If FG22.ColIndex("TypeVAT") <> -1 Then 'salim1503
'            FG22.TextMatrix(LongRow, FG22.ColIndex("TypeVAT")) = Percentg
'        End If
'
'    End If
    txtVatyo = Percentg
'    If FG22.ColIndex("Vatyo") <> -1 Then
'        FG22.TextMatrix(LongRow, FG22.ColIndex("Vatyo")) = Percentg
'    End If

    ' FG22.TextMatrix(LongRow, FG22.ColIndex("beforeVat")) = val(FG22.TextMatrix(LongRow, FG22.ColIndex("Qty"))) * val(FG22.TextMatrix(LongRow, FG22.ColIndex("Price")))
     txtVat2 = val(txtTotalAfterDiscount) * Percentg / 100
     lbl(23) = val(txtTotalAfterDiscount) + val(txtVat2)
   

End If

End Sub
  



Private Sub DeleteFgRow()

    With Me.FG22

        If .Row = -1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        .RemoveItem .Row
        .AutoSize 0, .Cols - 1, False
        
    End With

End Sub

