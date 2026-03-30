VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form frmdriveassestMove 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰ﬁ· Ê ”·Ì„ ⁄Âœ «·„ÊŸ›"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   Icon            =   "frmdriveassestMove.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   11280
   Begin VB.Frame gimage 
      BackColor       =   &H80000005&
      Height          =   6615
      Left            =   630
      TabIndex        =   99
      Top             =   840
      Visible         =   0   'False
      Width           =   9855
      Begin VB.CommandButton bClose 
         BackColor       =   &H000000FF&
         Caption         =   "X"
         Height          =   375
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imag4 
         Height          =   612
         Left            =   5640
         Picture         =   "frmdriveassestMove.frx":038A
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   708
      End
      Begin VB.Image img14 
         Height          =   612
         Left            =   720
         Picture         =   "frmdriveassestMove.frx":0938
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   708
      End
      Begin VB.Image img12 
         Height          =   612
         Left            =   3720
         Picture         =   "frmdriveassestMove.frx":0EE6
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   708
      End
      Begin VB.Image img11 
         Height          =   612
         Left            =   5280
         Picture         =   "frmdriveassestMove.frx":1494
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   708
      End
      Begin VB.Image img13 
         Height          =   612
         Left            =   2280
         Picture         =   "frmdriveassestMove.frx":1A42
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   708
      End
      Begin VB.Image img8 
         Height          =   612
         Left            =   8880
         Picture         =   "frmdriveassestMove.frx":1FF0
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   708
      End
      Begin VB.Image img10 
         Height          =   612
         Left            =   7080
         Picture         =   "frmdriveassestMove.frx":259E
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   708
      End
      Begin VB.Image img9 
         Height          =   612
         Left            =   8040
         Picture         =   "frmdriveassestMove.frx":2B4C
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   708
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   612
         Left            =   8880
         Top             =   4440
         Width           =   732
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   612
         Left            =   8040
         Top             =   4200
         Width           =   732
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   612
         Left            =   7080
         Top             =   4440
         Width           =   732
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   612
         Left            =   720
         Top             =   4320
         Width           =   732
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   612
         Left            =   2280
         Top             =   4200
         Width           =   732
      End
      Begin VB.Shape Shape13 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   612
         Left            =   5280
         Top             =   4320
         Width           =   732
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   612
         Left            =   3720
         Top             =   4320
         Width           =   732
      End
      Begin VB.Image img7 
         Height          =   612
         Left            =   720
         Picture         =   "frmdriveassestMove.frx":30FA
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   708
      End
      Begin VB.Image img6 
         Height          =   612
         Left            =   2520
         Picture         =   "frmdriveassestMove.frx":36A8
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   708
      End
      Begin VB.Image imag5 
         Height          =   612
         Left            =   4200
         Picture         =   "frmdriveassestMove.frx":3C56
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   708
      End
      Begin VB.Image imag3 
         Height          =   612
         Left            =   7080
         Picture         =   "frmdriveassestMove.frx":4204
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   708
      End
      Begin VB.Image imag2 
         Height          =   612
         Left            =   7920
         Picture         =   "frmdriveassestMove.frx":47B2
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   708
      End
      Begin VB.Image imag1 
         Height          =   612
         Left            =   8760
         Picture         =   "frmdriveassestMove.frx":4D60
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   708
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   612
         Left            =   8760
         Top             =   1800
         Width           =   732
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   612
         Index           =   1
         Left            =   7920
         Top             =   1560
         Width           =   732
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   612
         Index           =   1
         Left            =   7080
         Top             =   1800
         Width           =   732
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   612
         Left            =   5640
         Top             =   1320
         Width           =   732
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   612
         Left            =   4200
         Top             =   1560
         Width           =   732
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   612
         Left            =   2520
         Top             =   1680
         Width           =   732
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         Height          =   612
         Left            =   720
         Top             =   1560
         Width           =   732
      End
      Begin VB.Image Image6 
         Height          =   5772
         Left            =   0
         Picture         =   "frmdriveassestMove.frx":530E
         Stretch         =   -1  'True
         Top             =   480
         Width           =   9732
      End
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   13920
      TabIndex        =   39
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   13440
      TabIndex        =   38
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   13920
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   14040
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   13140
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   960
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7380
      Width           =   9585
      _cx             =   16907
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
         Left            =   8400
         TabIndex        =   16
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   7575
         TabIndex        =   17
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
         Left            =   6735
         TabIndex        =   18
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
         Left            =   5880
         TabIndex        =   19
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
         Left            =   5025
         TabIndex        =   20
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
         Height          =   372
         Index           =   6
         Left            =   360
         TabIndex        =   21
         Top             =   60
         Width           =   768
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
         Left            =   2295
         TabIndex        =   22
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
         Left            =   4200
         TabIndex        =   32
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
         Left            =   3360
         TabIndex        =   40
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
      Begin ImpulseButton.ISButton ISButton1 
         Height          =   375
         Left            =   1320
         TabIndex        =   113
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "«·„—›ﬁ« "
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
      Left            =   5220
      TabIndex        =   23
      Top             =   7020
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   13200
      TabIndex        =   24
      Top             =   3570
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
      Left            =   13560
      TabIndex        =   34
      Top             =   1920
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
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   5775
      Left            =   -120
      TabIndex        =   41
      Top             =   1080
      Width           =   11400
      _cx             =   20108
      _cy             =   10186
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
      Caption         =   "«·»Ì«‰« |Õ«·Â «·«⁄ „«œ|»Ì«‰«  «·„⁄œÂ/«·”Ì«—…"
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
      Picture(0)      =   "frmdriveassestMove.frx":22A5E
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   5310
         Left            =   12345
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   45
         Width           =   11310
         _cx             =   19950
         _cy             =   9366
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic11 
            Height          =   2052
            Left            =   0
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   0
            Width           =   11292
            _cx             =   19923
            _cy             =   3625
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
            Caption         =   "„” ‰œ«  «·„—ﬂ»…"
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   6
            ChildSpacing    =   4
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
            Begin VB.TextBox authorizeExamination 
               Alignment       =   1  'Right Justify
               Height          =   372
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   1440
               Width           =   3372
            End
            Begin VB.TextBox authorizeLicense 
               Alignment       =   1  'Right Justify
               Height          =   372
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   960
               Width           =   3372
            End
            Begin VB.TextBox FormOrignal 
               Alignment       =   1  'Right Justify
               Height          =   372
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   480
               Width           =   3372
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›Õ’ «·œÊ—Ï"
               Height          =   375
               Left            =   9480
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   1440
               Width           =   1695
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " ›ÊÌ÷ «·ﬁÌ«œ…"
               Height          =   375
               Left            =   9480
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«” „«—… «·«’·Ì…"
               Height          =   375
               Left            =   9480
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   480
               Width           =   1695
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic10 
            Height          =   2652
            Left            =   0
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   2160
            Width           =   11292
            _cx             =   19923
            _cy             =   4683
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
            Caption         =   "„—›ﬁ«  «·„—ﬂ»…"
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   6
            ChildSpacing    =   4
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
            Begin VB.CheckBox Stickers 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«” Ìﬂ— „”«⁄œ… ⁄·Ï «·ÿ—Ìﬁ"
               Height          =   432
               Left            =   2160
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   1800
               Width           =   2412
            End
            Begin VB.CheckBox Guarantee 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â«œ… ÷„«‰"
               Height          =   432
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   1800
               Width           =   1812
            End
            Begin VB.CheckBox CoverKey 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„› «Õ ⁄Ã·"
               Height          =   432
               Left            =   6960
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   1800
               Width           =   1812
            End
            Begin VB.CheckBox Crane 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—«›⁄…"
               Height          =   432
               Left            =   9120
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   1800
               Width           =   1812
            End
            Begin VB.CheckBox SpareTyre 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈ÿ«— «Õ Ì«ÿÏ"
               Height          =   432
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   1200
               Width           =   1812
            End
            Begin VB.CheckBox Battery 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·»ÿ«—Ì« "
               Height          =   432
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   1200
               Width           =   1812
            End
            Begin VB.CheckBox Anntena 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÂÊ«∆Ï"
               Height          =   432
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   1200
               Width           =   1812
            End
            Begin VB.CheckBox Recorder 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—«œÌÊ Ê«·„”Ã·"
               Height          =   432
               Left            =   6960
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   1200
               Width           =   1812
            End
            Begin VB.CheckBox SunScreens 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ê«ﬁÌ«  «·‘„” "
               Height          =   432
               Left            =   9120
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   1200
               Width           =   1812
            End
            Begin VB.CheckBox Pedals 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›—‘ «·œÊ”« "
               Height          =   432
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   600
               Width           =   1812
            End
            Begin VB.CheckBox InnerLights 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·‰Ê— «·œ«Œ·Ï"
               Height          =   432
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   600
               Width           =   1812
            End
            Begin VB.CheckBox driverMirror 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„—¬… «·”«∆ﬁ"
               Height          =   432
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   600
               Width           =   1812
            End
            Begin VB.CheckBox sideMirror 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„—«Ì« «·Ã«‰»Ì…"
               Height          =   432
               Left            =   6960
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   600
               Width           =   1812
            End
            Begin VB.CheckBox cleaner 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„”«Õ«  Ê√“—⁄ Â« "
               Height          =   432
               Left            =   8880
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   600
               Width           =   2055
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5310
         Left            =   12045
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   45
         Width           =   11310
         _cx             =   19950
         _cy             =   9366
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
            TabIndex        =   43
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
            FormatString    =   $"frmdriveassestMove.frx":22DF8
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
            TabIndex        =   46
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
            TabIndex        =   44
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   5310
         Index           =   15
         Left            =   45
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   45
         Width           =   11310
         _cx             =   19950
         _cy             =   9366
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
         _GridInfo       =   $"frmdriveassestMove.frx":22F3B
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5280
            Index           =   16
            Left            =   15
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   15
            Width           =   11280
            _cx             =   19897
            _cy             =   9313
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
            Begin VB.CheckBox chkAll 
               Alignment       =   1  'Right Justify
               Caption         =   "«ŸÂ«— ﬂ· «·⁄Âœ"
               Height          =   195
               Left            =   6420
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   1470
               Width           =   1425
            End
            Begin VB.TextBox txtPrice 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   1290
               Width           =   1335
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   8880
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   2040
               Width           =   1215
            End
            Begin VB.TextBox TxtOperatorN 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2880
               TabIndex        =   107
               Top             =   1680
               Width           =   1335
            End
            Begin VB.TextBox TxtBoardNO 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2880
               TabIndex        =   105
               Top             =   2040
               Width           =   1335
            End
            Begin VB.CheckBox chkCar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”Ì«—…"
               Height          =   312
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   120
               Width           =   1092
            End
            Begin VB.CommandButton BtImage 
               Caption         =   " ÕœÌœ «·„·«ÕŸ«  "
               Height          =   375
               Left            =   240
               Picture         =   "frmdriveassestMove.frx":22F6E
               TabIndex        =   101
               Top             =   2640
               Width           =   1335
            End
            Begin VB.TextBox IdAsest 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   840
               Visible         =   0   'False
               Width           =   1335
            End
            Begin XtremeSuiteControls.RadioButton RdType 
               Height          =   255
               Left            =   9480
               TabIndex        =   71
               Top             =   120
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ”·Ì„ ⁄ÂœÂ"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.TextBox Txtamount2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox TxtSearchCode1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   8880
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   975
               Width           =   1215
            End
            Begin VB.TextBox Txtamount 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox TxtSearchCode 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   8880
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox txtreson 
               Alignment       =   1  'Right Justify
               Height          =   585
               Left            =   5520
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   48
               Top             =   2400
               Width           =   4560
            End
            Begin ImpulseButton.ISButton xxx 
               Height          =   510
               Left            =   0
               TabIndex        =   49
               Top             =   5265
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   900
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "«—”«· ··«⁄ „«œ"
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
               ColorShadow     =   4210752
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               ColorTextShadow =   4210752
            End
            Begin MSDataListLib.DataCombo DcboEmpName 
               Height          =   315
               Left            =   5520
               TabIndex        =   53
               Top             =   600
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton Accredit 
               Height          =   510
               Left            =   0
               TabIndex        =   60
               Top             =   5280
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   900
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "«—”«· ··«⁄ „«œ"
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
               ColorShadow     =   4210752
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               ColorTextShadow =   4210752
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   2070
               Left            =   120
               TabIndex        =   61
               Tag             =   "1"
               Top             =   3120
               Width           =   11055
               _cx             =   19500
               _cy             =   3651
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
               Cols            =   44
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmdriveassestMove.frx":26EB5
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   14
               Left            =   2985
               TabIndex        =   62
               Top             =   2520
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
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
               ButtonImage     =   "frmdriveassestMove.frx":2752E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   15
               Left            =   2280
               TabIndex        =   63
               Top             =   2520
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   688
               ButtonStyle     =   1
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
               ButtonImage     =   "frmdriveassestMove.frx":278C8
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSComCtl2.DTPicker DBIssueDate 
               Height          =   375
               Left            =   240
               TabIndex        =   65
               Top             =   120
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   393216
               Format          =   118554625
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DcboEmpNameTo 
               Height          =   315
               Left            =   5520
               TabIndex        =   68
               Top             =   960
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdTypeMov 
               Height          =   255
               Left            =   7800
               TabIndex        =   72
               Top             =   120
               Width           =   1455
               _Version        =   786432
               _ExtentX        =   2566
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "‰ﬁ· ⁄ÂœÂ"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker DriveDate 
               Height          =   375
               Left            =   2880
               TabIndex        =   102
               Top             =   120
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   393216
               Format          =   118554625
               CurrentDate     =   38784
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   435
               Left            =   480
               TabIndex        =   106
               TabStop         =   0   'False
               Top             =   2040
               Width           =   2325
               _cx             =   4101
               _cy             =   767
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               Begin VB.TextBox txtNum4 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   0
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   0
                  Width           =   300
               End
               Begin VB.TextBox txtLetter4 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1155
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   4
                  Top             =   0
                  Width           =   360
               End
               Begin VB.TextBox txtNum3 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   270
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   0
                  Width           =   300
               End
               Begin VB.TextBox txtNum2 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   480
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   0
                  Width           =   330
               End
               Begin VB.TextBox txtNum1 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   795
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   5
                  Top             =   0
                  Width           =   360
               End
               Begin VB.TextBox txtLetter3 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1440
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   3
                  Top             =   0
                  Width           =   315
               End
               Begin VB.TextBox txtLetter2 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1710
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   2
                  Top             =   0
                  Width           =   240
               End
               Begin VB.TextBox txtLetter1 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1935
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   1
                  Top             =   0
                  Width           =   285
               End
            End
            Begin MSDataListLib.DataCombo dcmboassest 
               Height          =   315
               Left            =   5520
               TabIndex        =   111
               Top             =   2040
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”⁄—"
               Height          =   180
               Index           =   17
               Left            =   4410
               TabIndex        =   115
               Top             =   1290
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ÊŸ›"
               Height          =   285
               Index           =   15
               Left            =   10200
               TabIndex        =   112
               Top             =   600
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «··ÊÕ…"
               Height          =   285
               Index           =   18
               Left            =   4440
               TabIndex        =   109
               Top             =   2040
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—ﬁ„ «· ‘€Ì·Ì"
               Height          =   285
               Index           =   66
               Left            =   4440
               TabIndex        =   108
               Top             =   1680
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " √—ÌŒ «· ”·Ì„"
               Height          =   300
               Index           =   16
               Left            =   4440
               TabIndex        =   103
               Top             =   120
               Width           =   975
            End
            Begin VB.Shape Shape2 
               BorderWidth     =   2
               Height          =   975
               Index           =   0
               Left            =   0
               Top             =   600
               Width           =   2775
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "‘«‘…  ‰ﬁ· Ê ”·Ì„ «·⁄Âœ ··„ÊŸ›Ì‰   ”«⁄œ Â–… «·‘«‘… ›Ì  ‰ﬁ· Ê ”·Ì„ «·⁄Âœ  „‰ „ÊŸ› ·«Œ— „⁄ „—«Ã⁄Â «·ﬂ„Ì… «·›⁄·Ì… «·„ÊÃÊœ…"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   945
               Index           =   14
               Left            =   -120
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   600
               Width           =   2775
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì… «·„‰ﬁÊ·Â"
               Height          =   300
               Index           =   12
               Left            =   4440
               TabIndex        =   73
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï «·„ÊŸ›"
               Height          =   285
               Index           =   11
               Left            =   10230
               TabIndex        =   69
               Top             =   975
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì…"
               Height          =   180
               Index           =   10
               Left            =   4440
               TabIndex        =   64
               Top             =   600
               Width           =   972
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄Âœ"
               Height          =   285
               Index           =   2
               Left            =   10230
               TabIndex        =   55
               Top             =   2040
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ «·„ÊŸ›"
               Height          =   285
               Index           =   3
               Left            =   10230
               TabIndex        =   54
               Top             =   600
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ…"
               Height          =   255
               Index           =   9
               Left            =   10545
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   2520
               Width           =   690
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " √—ÌŒ «· ”·Ì„"
               Height          =   300
               Index           =   13
               Left            =   1800
               TabIndex        =   50
               Top             =   120
               Width           =   975
            End
         End
      End
   End
   Begin MSDataListLib.DataCombo Dcbranch 
      Height          =   315
      Left            =   120
      TabIndex        =   56
      Top             =   720
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   5400
      TabIndex        =   59
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   118554625
      CurrentDate     =   38784
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   11205
      _cx             =   19764
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
      Caption         =   " ‰ﬁ· Ê ”·Ì„ ⁄Âœ «·„ÊŸ› "
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
         TabIndex        =   11
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
         ButtonImage     =   "frmdriveassestMove.frx":27E62
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
         TabIndex        =   12
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
         ButtonImage     =   "frmdriveassestMove.frx":281FC
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
         TabIndex        =   13
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
         ButtonImage     =   "frmdriveassestMove.frx":28596
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
         TabIndex        =   14
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
         ButtonImage     =   "frmdriveassestMove.frx":28930
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
         Left            =   3600
         Picture         =   "frmdriveassestMove.frx":28CCA
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
         TabIndex        =   37
         Top             =   480
         Width           =   2205
      End
   End
   Begin VB.Image imgnul 
      Height          =   1092
      Left            =   12840
      Top             =   6720
      Width           =   732
   End
   Begin VB.Image img 
      Height          =   852
      Left            =   12120
      Picture         =   "frmdriveassestMove.frx":2C932
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   720
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   975
      Index           =   0
      Left            =   120
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· √—ÌŒ"
      Height          =   300
      Index           =   1
      Left            =   6480
      TabIndex        =   58
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·›—⁄     "
      Height          =   300
      Index           =   5
      Left            =   3720
      TabIndex        =   57
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   12810
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   285
      Index           =   4
      Left            =   10110
      TabIndex        =   31
      Top             =   750
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   7965
      TabIndex        =   30
      Top             =   7035
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2550
      TabIndex        =   29
      Top             =   7020
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   840
      TabIndex        =   28
      Top             =   7020
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   210
      TabIndex        =   27
      Top             =   7020
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1860
      TabIndex        =   26
      Top             =   7020
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   13350
      TabIndex        =   25
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "frmdriveassestMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim Employee_account As String
Public bo As Boolean

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.VSFlexGrid1

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("EmpName")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("ser")) = IntCounter
  
            End If

        Next i
   
    End With


 

End Sub
Function addrow2()
 Dim i As Integer
 If Txtamount2.Text = "" Then
Txtamount2.Text = 1
End If
 If Txtamount.Text = "" Then
Txtamount.Text = 1
End If
 
      If VSFlexGrid1.Rows = 1 Then VSFlexGrid1.Rows = 2
         With VSFlexGrid1
         If Me.RdTypeMov.value = True Then
    If val(Me.Txtamount2.Text) > val(Me.Txtamount.Text) Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "«·ﬂ„ÌÂ «ﬂ»— „‰ «·ﬂ„ÌÂ «·„ Ê›—…"
    Else
    MsgBox "Quantity greater than the quantity available"
    End If
    Txtamount2.Text = ""
    Txtamount.Text = ""
    TxtPrice = ""
    
    'Me.dcmboassest.text = ""
    Exit Function
    End If
    End If
  i = .Rows
 
        .TextMatrix(i - 1, .ColIndex("FormOrignal")) = IIf(FormOrignal.Text = "", "", FormOrignal.Text)
        .TextMatrix(i - 1, .ColIndex("authorizeLicense")) = IIf(authorizeLicense.Text = "", "", authorizeLicense.Text)
        .TextMatrix(i - 1, .ColIndex("authorizeExamination")) = IIf(authorizeExamination.Text = "", "", authorizeExamination.Text)
             
        .TextMatrix(i - 1, .ColIndex("cleaner")) = cleaner.value
        .TextMatrix(i - 1, .ColIndex("sideMirror")) = sideMirror.value
       .TextMatrix(i - 1, .ColIndex("driverMirror")) = driverMirror.value
        .TextMatrix(i - 1, .ColIndex("InnerLights")) = InnerLights.value
        .TextMatrix(i - 1, .ColIndex("Pedals")) = Pedals.value
       .TextMatrix(i - 1, .ColIndex("SunScreens")) = SunScreens.value
      .TextMatrix(i - 1, .ColIndex("Recorder")) = Recorder.value
        .TextMatrix(i - 1, .ColIndex("Anntena")) = Anntena.value
       .TextMatrix(i - 1, .ColIndex("Battery")) = Battery.value
        .TextMatrix(i - 1, .ColIndex("SpareTyre")) = SpareTyre.value
       .TextMatrix(i - 1, .ColIndex("Crane")) = Crane.value
        .TextMatrix(i - 1, .ColIndex("CoverKey")) = CoverKey.value
        .TextMatrix(i - 1, .ColIndex("Guarantee")) = Guarantee.value
        .TextMatrix(i - 1, .ColIndex("Stickers")) = Stickers.value
    .TextMatrix(i - 1, .ColIndex("Price")) = (Me.TxtPrice.Text)
          .TextMatrix(i - 1, .ColIndex("Total")) = val(Me.TxtPrice.Text) * val(Txtamount2.Text)
        
        
        
           If Me.imag1.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar1")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar1")) = 0
           End If
           
           
           If Me.imag2.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar2")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar2")) = 0
           End If
            
           If Me.imag3.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar3")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar3")) = 0
           End If
           
           
           If Me.imag4.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar4")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar4")) = 0
           End If
           
                      If Me.imag5.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar5")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar5")) = 0
           End If
           
           
           If Me.img6.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar6")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar6")) = 0
           End If
           
                      If Me.img7.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar7")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar7")) = 0
           End If
           
           
           If Me.img8.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar8")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar8")) = 0
           End If
           
          If Me.img9.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar9")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar9")) = 0
           End If
           
           
           If Me.img10.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar10")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar10")) = 0
           End If
           
                      If Me.img11.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar11")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar11")) = 0
           End If
           
           
           If Me.img12.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar12")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar12")) = 0
           End If
        
                   If Me.img12.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar12")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar12")) = 0
           End If
        
                   If Me.img13.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar13")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar13")) = 0
           End If
        
        
        
                   If Me.img14.Picture <> 0 Then
                    .TextMatrix(i - 1, .ColIndex("subcar14")) = 1
           Else
                    .TextMatrix(i - 1, .ColIndex("subcar14")) = 0
           End If
          
        
        
        
 
        Clear_CarData
 
 
               .TextMatrix(i - 1, .ColIndex("EmpName")) = (dcmboassest.Text)
      .TextMatrix(i - 1, .ColIndex("id")) = (dcmboassest.BoundText)
                    
                If Me.RdTypeMov.value = True Then
                .TextMatrix(i - 1, .ColIndex("idas")) = Me.IdAsest.Text
                .TextMatrix(i - 1, .ColIndex("ApprovDate")) = (Me.Txtamount2.Text)
             
                .TextMatrix(i - 1, .ColIndex("Total")) = val(Me.TxtPrice.Text) * val(Txtamount2.Text)
                
                   .TextMatrix(i - 1, .ColIndex("diff")) = val(Txtamount.Text) - val(Txtamount2.Text)
                   Else
                   .TextMatrix(i - 1, .ColIndex("ApprovDate")) = (Me.Txtamount.Text)
                   End If
                .TextMatrix(i - 1, .ColIndex("Remarks")) = (txtreson.Text)
             
              '  .TextMatrix(i - 1, .ColIndex("workto")) = (workto.value)
              '  .TextMatrix(i - 1, .ColIndex("worktoH")) = (worktoH.value)
              '  .TextMatrix(i - 1, .ColIndex("des")) = (Text27.text)
                  
                  '.TextMatrix(i - 1, .ColIndex("des")) = (Txtdes1.text)
                  
                'Me.dcmboassest.Text = ""
                  .Rows = .Rows + 1
                  txtreson.Text = ""
                  Txtamount.Text = ""
                  Txtamount2.Text = ""
                  TxtPrice.Text = ""
              '    TxtWorkEntity.text = ""
             
                 
      '       .AutoSize 0, .Cols - 1, False
   
    End With
 
    
    ReLineGrid

End Function
'Private Sub Accredit_Click()
'    Dim BeginTrans As Boolean
'
''    Cn.BeginTrans
 '   BeginTrans = True
'
'    If IsNull(rs("Posted")) Then
'        rs("Posted") = user_id
'        rs("PostedDate") = Time
'    Else
'        rs("Posted") = Null
'       rs("PostedDate") = Time
''    End If
   
Private Sub Clear_CarData()
       
    FormOrignal.Text = ""
    authorizeLicense.Text = ""
    authorizeExamination.Text = ""
    
    cleaner.value = False
    sideMirror.value = False
    driverMirror.value = False
    InnerLights.value = False
    Pedals.value = False
    SunScreens.value = False
    Recorder.value = False
    Anntena.value = False
    
    Battery.value = False
    SpareTyre.value = False
    Crane.value = False
    CoverKey.value = False
    Guarantee.value = False
    Stickers.value = False
      

End Sub

   
Private Sub Check_Car()
If Me.TxtModFlg.Text <> "R" Then
loadcombo
If chkCar.value = False Then
lbl(66).Visible = False
lbl(16).Visible = False
lbl(15).Visible = False
TxtOperatorN.Visible = False
TxtBoardNO.Visible = False
C1Elastic7.Visible = False
        Exit Sub
  Else
  C1Elastic7.Visible = True
TxtOperatorN.Visible = True
TxtBoardNO.Visible = True
lbl(66).Visible = True
lbl(15).Visible = True
lbl(16).Visible = True
End If

Dim str As String

str = "  SELECT dbo.TblCarsData.id, dbo.TblCarsData.fixedAssetid, dbo.TblAssestes.AsFixedID, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Stickers, dbo.TblCarsData.Guarantee, "
str = str & "  dbo.TblCarsData.CoverKey, dbo.TblCarsData.Crane, dbo.TblCarsData.SpareTyre, dbo.TblCarsData.Battery, dbo.TblCarsData.Anntena, dbo.TblCarsData.Recorder,"
str = str & "  dbo.TblCarsData.SunScreens, dbo.TblCarsData.Pedals, dbo.TblCarsData.InnerLights, dbo.TblCarsData.driverMirror, dbo.TblCarsData.sideMirror, dbo.TblCarsData.cleaner,"
str = str & "  dbo.TblCarsData.authorizeLicense , dbo.TblCarsData.authorizeExamination, dbo.TblCarsData.FormOrignal"
str = str & "  FROM     dbo.FixedAssets INNER JOIN"
str = str & "  dbo.TblCarsData ON dbo.FixedAssets.id = dbo.TblCarsData.fixedAssetid INNER JOIN"
str = str & "  dbo.TblAssestes ON dbo.FixedAssets.id = dbo.TblAssestes.AsFixedID  "
str = str & "  where TblAssestes.Asid = " & val(dcmboassest.BoundText)
Set Rs_Temp = New ADODB.Recordset
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText

If Rs_Temp.RecordCount > 0 Then
                    FormOrignal.Text = IIf(IsNull(Rs_Temp("FormOrignal").value), "", Rs_Temp("FormOrignal").value)
                    authorizeLicense.Text = IIf(IsNull(Rs_Temp("authorizeLicense").value), "", Rs_Temp("authorizeLicense").value)
                    authorizeExamination.Text = IIf(IsNull(Rs_Temp("authorizeExamination").value), "", Rs_Temp("authorizeExamination").value)
                    cleaner.value = IIf(IsNull(Rs_Temp("cleaner").value), False, Rs_Temp("cleaner").value)
                    SunScreens.value = IIf(IsNull(Rs_Temp("SunScreens").value), False, Rs_Temp("SunScreens").value)
                    Crane.value = IIf(IsNull(Rs_Temp("Crane").value), False, Rs_Temp("Crane").value)
                    sideMirror.value = IIf(IsNull(Rs_Temp("sideMirror").value), False, Rs_Temp("sideMirror").value)
                    Recorder.value = IIf(IsNull(Rs_Temp("Recorder").value), False, Rs_Temp("Recorder").value)
                    CoverKey.value = IIf(IsNull(Rs_Temp("CoverKey").value), False, Rs_Temp("CoverKey").value)
                    driverMirror.value = IIf(IsNull(Rs_Temp("driverMirror").value), False, Rs_Temp("driverMirror").value)
                    Anntena.value = IIf(IsNull(Rs_Temp("Anntena").value), False, Rs_Temp("Anntena").value)
                    Guarantee.value = IIf(IsNull(Rs_Temp("Guarantee").value), False, Rs_Temp("Guarantee").value)
                    InnerLights.value = IIf(IsNull(Rs_Temp("InnerLights").value), False, Rs_Temp("InnerLights").value)
                    Battery.value = IIf(IsNull(Rs_Temp("Battery").value), False, Rs_Temp("Battery").value)
                    Stickers.value = IIf(IsNull(Rs_Temp("Stickers").value), False, Rs_Temp("Stickers").value)
                    Pedals.value = IIf(IsNull(Rs_Temp("Pedals").value), False, Rs_Temp("Pedals").value)
                    SpareTyre.value = IIf(IsNull(Rs_Temp("SpareTyre").value), False, Rs_Temp("SpareTyre").value)
End If
End If
End Sub


Private Sub bClose_Click()

BtImage.Visible = True
gimage.Visible = False

End Sub

Private Sub BtImage_Click()
If chkCar.value = 1 Then
        BtImage.Visible = False
        gimage.Visible = True
End If '

End Sub

Private Sub chkAll_Click()
  Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    If Me.RdTypeMov.value = True Then
        
        If val(Me.DcboEmpName.BoundText) <> 0 Then
            If chkAll = vbChecked Then
                Dcombos.GetAssests Me.dcmboassest
            Else
                Dcombos.GetAssestsOfEmp Me.dcmboassest, val(Me.DcboEmpName.BoundText)
            End If
        End If
    Else
        Dcombos.GetAssests Me.dcmboassest
    End If
End Sub

Private Sub chkCar_Click()
Check_Car
End Sub

 '   rs.update
 'If SystemOptions.UserInterface = ArabicInterface Then
 '   Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
'Else
'Accredit.Caption = "Sent To approval "
'End If

'    Cn.CommitTrans
'    BeginTrans = False
'FillApprovedTable
'    Retrive (val(Me.XPTxtID.text))
'End Sub

Private Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            Check_Car
            Me.RdType.value = True
      ''      lbl(20).Caption = "0"
       ''     lbl(21).Caption = "0"
       '     lbl(22).Caption = "0"
       '     lbl(23).Caption = "0"
            
'              GRID2.Clear flexClearScrollable, flexClearEverything
'    GRID2.Rows = 1
                Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 1
    VSFlexGrid1.Enabled = True
   ' VSFlexGrid1.Editable = flexEDKbd
            Me.DCboUserName.BoundText = user_id
           ' TxtPaymentCounts.text = 1
Dcbranch.BoundText = Current_branch
            'XPDtbTrans.SetFocus
            
'            Accredit.Enabled = True
'                If SystemOptions.UserInterface = ArabicInterface Then
'                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
'                                                  Else
''                                                    Accredit.Caption = " send to Approval   "
 '                                              End If
                                               
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
Me.VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
    
            Dim Msg As String

            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Dcbranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.Dcbranch.BoundText

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5
        bo = True

             Load FrmAssestSearch
             FrmAssestSearch.show

        Case 6
            Unload Me

 '       Case 7
 '           ShowGL_cc Me.TxtNoteSerial.text, , 200

 '       Case 8
 '           CalCulateParts
            
            
                 Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.Text) <> 0 Then
                If chkCar.value = 0 Then
                        print_report val(Me.XPTxtID.Text)
                ElseIf chkCar.value = 1 Then
                        print_report2
                End If
        
        
            End If
          Case 14
              
            addrow2
            BtImage.Visible = True
gimage.Visible = False
            Case 15
            RemoveGridRow2
    End Select

    Exit Sub
ErrTrap:
End Sub
Private Sub RemoveGridRow2()
      If Me.VSFlexGrid1.Rows = 1 Then Exit Sub
    With Me.VSFlexGrid1

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub
Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String


    
    
MySQL = " SELECT dbo.TblAssestes.AsName, dbo.TblAssestes.AsID, dbo.TblAssestes.AsDes, dbo.TblEmpAsestDetails.Remarks, dbo.TblEmpAsestDetails.Qunt, dbo.TblEmpAsest.EmpAsID,"
  MySQL = MySQL & "                 dbo.TblEmpAsestDetails.IDAseset, dbo.TblEmpAsest.EmpAsestID, dbo.TblEmpAsest.PostedDate, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code,"
   MySQL = MySQL & "                dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality,"
   MySQL = MySQL & "                dbo.TblEmployee.dean, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3,"
   MySQL = MySQL & "                TblEmpAsestDetails.Price,TblEmpAsestDetails.Price * TblEmpAsestDetails.Qunt Total,"
     MySQL = MySQL & "              dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Fullcode, dbo.TblEmpAsest.ToEmId, TblEmployee_1.Emp_Name AS Emp_NameTo,"
     MySQL = MySQL & "              TblEmployee_1.Emp_Name2 AS Emp_Name2To, TblEmployee_1.Emp_Name3 AS Emp_Name3To, TblEmployee_1.Emp_Name1 AS Emp_Name1To,"
     MySQL = MySQL & "              TblEmployee_1.Emp_Name4 AS Emp_Name4To, TblEmployee_1.Nationality AS NationalityTo, TblEmployee_1.Emp_Namee AS Emp_NameeTo,"
     MySQL = MySQL & "              TblEmployee_1.Emp_Namee1 AS Emp_Namee1To, TblEmployee_1.Emp_Namee2 AS Emp_Namee2To, TblEmployee_1.Emp_Namee3 AS Emp_Namee3To,"
    MySQL = MySQL & "               TblEmployee_1.Fullcode AS FullcodeTo, TblEmployee_1.Emp_Namee4 AS Emp_Namee4To, dbo.TblEmpAsestDetails.Remark2, dbo.TblEmployee.NumEkama,"
   MySQL = MySQL & "                dbo.TblEmpJobsTypes.JobTypeName , dbo.TblEmpJobsTypes.JobTypeNamee , dbo.TblEmployee.placeEkama"
MySQL = MySQL & " FROM     dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
   MySQL = MySQL & "                dbo.TblAssestes INNER JOIN"
    MySQL = MySQL & "               dbo.TblEmpAsestDetails ON dbo.TblAssestes.AsID = dbo.TblEmpAsestDetails.AsID INNER JOIN"
     MySQL = MySQL & "              dbo.TblEmpAsest ON dbo.TblEmpAsestDetails.IDAseset = dbo.TblEmpAsest.EmpAsID INNER JOIN"
      MySQL = MySQL & "             dbo.TblEmployee ON dbo.TblEmpAsest.EmpAsestID = dbo.TblEmployee.Emp_ID ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.BlnceVocat LEFT OUTER JOIN"
      MySQL = MySQL & "             dbo.TblEmployee AS TblEmployee_1 ON dbo.TblEmpAsest.ToEmId = TblEmployee_1.Emp_ID"
        
    MySQL = MySQL & "  Where (dbo.TblEmpAsestDetails.FlagAs Is Null) and (dbo.TblEmpAsestDetails.IDAseset = " & val(Me.XPTxtID.Text) & ")"

 
        If SystemOptions.UserInterface = ArabicInterface Then
        If Me.RdTypeMov.value = False Then
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpDresAseste.rpt"
              '  StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_DriveAssetMove.rpt"
            Else
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpDresAsesteMove.rpt"
              '    StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_DriveAssetMove.rpt"
             End If
        Else
          If Me.RdTypeMov.value = False Then
          StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpDresAseste.rpt"
           '   StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_DriveAssetMove.rpt"
            Else
           StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpDresAseste.rpt"
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
        End If
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
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
  '      xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
         xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value

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
Private Sub DcboEmpName_Change()
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    If Me.RdTypeMov.value = True Then
        
        If val(Me.DcboEmpName.BoundText) <> 0 Then
            If chkAll = vbChecked Then
                Dcombos.GetAssests Me.dcmboassest
            Else
                Dcombos.GetAssestsOfEmp Me.dcmboassest, val(Me.DcboEmpName.BoundText)
            End If
        End If
    Else
        Dcombos.GetAssests Me.dcmboassest
    End If
    DcboEmpName_Click (0)
End Sub



 

Sub QuntAsset(Optional AsID As Integer, Optional ByRef count As String, Optional EmpID As Integer, Optional IDAseset As Integer, Optional IsPrice As Boolean = False)
Dim str As String
Dim Rs1 As New ADODB.Recordset
              
If Not IsPrice Then
    str = "SELECT     dbo.TblEmpAsestDetails.Qunt,IsNull(dbo.TblEmpAsestDetails.Price,TblAssestes.Price) Price, dbo.TblEmpAsestDetails.IDAseset  "
    str = str & " FROM         dbo.TblAssestes INNER JOIN"
     str = str & "                     dbo.TblEmpAsestDetails ON dbo.TblAssestes.AsID = dbo.TblEmpAsestDetails.AsID INNER JOIN"
     str = str & "                     dbo.TblEmpAsest ON dbo.TblEmpAsestDetails.IDAseset = dbo.TblEmpAsest.EmpAsID"
    'str = str & " Where (dbo.TblEmpAsestDetails.FlagAs Is Null) And (dbo.TblAssestes.AsID = " & AsID & ") and (dbo.TblAssestes.EmpID= " & EmpId & " )"
    str = str & " Where (dbo.TblEmpAsestDetails.FlagAs Is Null) And (dbo.TblEmpAsestDetails.EmpID = " & EmpID & ") And (dbo.TblEmpAsestDetails.AsID = " & AsID & ")"
Else
    str = "SELECT     dbo.TblAssestes.Price from TblAssestes Where AsID = " & val(AsID)
    Rs1.Open str, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Rs1.RecordCount > 0 Then
        TxtPrice = Rs1!Price & ""
    End If
    Exit Sub
End If
Rs1.Open str, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs1.RecordCount > 0 Then
If IsNull(Rs1("Qunt").value) Or val(Rs1("Qunt").value) = 0 Then
TxtPrice = Rs1!Price & ""
count = 1
Else
count = Rs1("Qunt").value
TxtPrice = Rs1!Price & ""
IDAseset = Rs1("IDAseset").value
End If
End If
End Sub

Private Sub DcboEmpNameTo_Change()
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    
    If val(Me.DcboEmpName.BoundText) <> 0 Then
        'Dcombos.GetAssestsOfEmp Me.dcmboassest, val(Me.DcboEmpName.BoundText)
    End If
    
    If val(DcboEmpNameTo.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , Me.DcboEmpNameTo.BoundText, EmpCode
    TxtSearchCode1.Text = EmpCode

End Sub

Public Sub dcmboassest_Change()
Dim cont As String
Dim FixedID As Double
Dim AsID As Double
Dim IdAsest As Integer
If Me.RdTypeMov.value = True Then
If val(dcmboassest.BoundText) <> 0 Then
QuntAsset Me.dcmboassest.BoundText, cont, val(Me.DcboEmpName.BoundText), IdAsest
Txtamount.Text = cont
Me.IdAsest.Text = IdAsest
End If
Else
    QuntAsset val(Me.dcmboassest.BoundText), , , , True
End If
'Check_Car
If chkCar.value = vbChecked Then
GetCardID FixedID, val(Me.dcmboassest.BoundText), 0
RetriveCarsInfo FixedID, , , 0
End If
End Sub
Public Sub GetCardID(Optional ByRef AsFixedID As Double, Optional ByRef AsID As Double, Optional Typ As Integer = 0)
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = "  SELECT     AsFixedID, AsID"
sql = sql & " FROM         dbo.TblAssestes"
If Typ = 0 Then
sql = sql & "  where AsID =" & AsID & ""
Else
sql = sql & "  where AsFixedID =" & AsFixedID & ""
End If
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
AsFixedID = IIf(IsNull(Rs4("AsFixedID").value), 0, Rs4("AsFixedID").value)
AsID = IIf(IsNull(Rs4("AsID").value), 0, Rs4("AsID").value)
Else
AsFixedID = 0
AsID = 0
End If
End Sub
Private Sub dcmboassest_KeyUp(KeyCode As Integer, Shift As Integer)
If chkCar.value = vbChecked Then
   If KeyCode = vbKeyF3 Then
         Load FrmCasrShearches
        FrmCasrShearches.SendForm = "frmdriveassestMove"
        FrmCasrShearches.show vbModal
    End If
  Else
If KeyCode = vbKeyF3 Then
bo = False
 Load FixedAssetsSearch
    FixedAssetsSearch.RetrunType = 17
         FixedAssetsSearch.show vbModal
  
End If
End If
End Sub

Private Sub ISButton1_Click()
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments XPTxtID.Text, "220820170111"
ErrTrap:
End Sub

Private Sub TxtBoardNO_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
RetriveCarsInfo , , TxtBoardNO.Text, 2
End If
End Sub

Private Sub txtLetter1_KeyPress(KeyAscii As Integer)
txtLetter1.Text = ""
If Len(txtLetter1.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case 8
        Exit Sub
    Case Else
        txtLetter2.SetFocus
End Select

End Sub
Sub RetriveCarsInfo(Optional CarID As Double = 0, Optional OperNo As String, Optional BoardNO As String, Optional Typ As Integer = 0)
If chkCar.value = vbChecked Then
If Me.TxtModFlg <> "R" Then
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "select * from TblCarsData"
If Typ = 0 Then
sql = sql & "  Where FixedassetId = " & CarID & ""
ElseIf Typ = 1 Then
sql = sql & " where OperatorN='" & OperNo & "'"
ElseIf Typ = 2 Then
sql = sql & " where BoardNO='" & BoardNO & "'"
End If
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
If Typ <> 1 Then
Me.TxtOperatorN.Text = IIf(IsNull(Rs3("OperatorN").value), "", Rs3("OperatorN").value)
End If
If Typ <> 2 Then
TxtBoardNO.Text = IIf(IsNull(Rs3("BoardNO").value), "", Rs3("BoardNO").value)
End If
If Typ <> 0 Then
Dim AsID As Double
GetCardID IIf(IsNull(Rs3("FixedassetId").value), 0, Rs3("FixedassetId").value), AsID, 1
dcmboassest.BoundText = AsID
End If
DcboEmpName.BoundText = IIf(IsNull(Rs3("Emp_id").value), 0, Rs3("Emp_id").value)
Else
DcboEmpName.BoundText = 0
If Typ <> 1 Then
TxtOperatorN.Text = ""
End If
If Typ <> 2 Then
TxtBoardNO.Text = ""
End If
If Typ <> 0 Then
dcmboassest.BoundText = 0
End If
End If
End If
End If
End Sub
Private Sub Cal_Board()
    TxtBoardNO.Text = txtLetter1.Text & " " & txtLetter2.Text & " " & txtLetter3.Text & " " & txtLetter4.Text & " " & txtNum1.Text & " " & txtNum2.Text & " " & txtNum3.Text & " " & txtNum4.Text
    RetriveCarsInfo , , TxtBoardNO.Text, 2
End Sub



Private Sub txtNum3_KeyPress(KeyAscii As Integer)
txtNum3.Text = ""
If Len(txtNum3.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum4.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtNum1_KeyPress(KeyAscii As Integer)
txtNum1.Text = ""
If Len(txtNum1.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum2.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtLetter4_KeyPress(KeyAscii As Integer)
txtLetter4.Text = ""
If Len(txtLetter4.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtNum1.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtLetter1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtLetter2_KeyPress(KeyAscii As Integer)
txtLetter2.Text = ""
If Len(txtLetter2.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter3.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
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



Private Sub RdType_Click()
If RdType.value = True Then
lbl(11).Visible = False
lbl(12).Visible = False
Me.TxtSearchCode1.Visible = False
Me.DcboEmpNameTo.Visible = False
Me.Txtamount2.Visible = False
lbl(3).Visible = False
lbl(15).Visible = True
TxtSearchCode.Enabled = True
DcboEmpName.Enabled = True
End If

End Sub

Private Sub RdTypeMov_Click()
If RdTypeMov.value = True Then
lbl(11).Visible = True
lbl(12).Visible = True
TxtPrice.Enabled = False
TxtSearchCode.Enabled = True
DcboEmpName.Enabled = True
Me.TxtSearchCode1.Visible = True
Me.DcboEmpNameTo.Visible = True
Me.Txtamount2.Visible = True
lbl(3).Visible = True

lbl(15).Visible = False
Else
TxtPrice.Enabled = True
End If
End Sub

Private Sub Txtamount_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False

Load FrmAssestSearch
            FrmAssestSearch.show
            
End If
End Sub



Private Sub txtNum4_KeyPress(KeyAscii As Integer)
txtNum4.Text = ""
If Len(txtNum4.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If
Cal_Board

End Sub
Private Sub txtNum4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtNum2_KeyPress(KeyAscii As Integer)
txtNum2.Text = ""
If Len(txtNum2.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum3.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub TxtOperatorN_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
RetriveCarsInfo , TxtOperatorN.Text, , 1
End If
End Sub

Private Sub txtreson_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False
'
Load FrmAssestSearch
            FrmAssestSearch.show
            
End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If

End Sub

 

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
       
        FrmEmployeeSearch.lbltype = 15
'        Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If

End Sub

Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
    
   If Me.TxtModFlg = "R" Then Exit Sub
   
   
    Dim StrSQL As String
      
        
        Dim IssueDate As Date
        Dim DepID As Double
        Dim specid As Double
        Dim JobTypeID As Double
        Dim gradeID As Double
        Dim Account_code2 As String
           Dim Account_code  As String
        Dim Balance As String
        Dim endContractPerMonth As Double
        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_code, endContractPerMonth
        
          WriteCustomerBalPublic Account_code2, Balance
            WriteCustomerBalPublic Account_code, Balance
    '    DBIssueDate.value = issuedate
    '    DcmbManagerID.BoundText = depid
    '     DcboJobsType.BoundText = JobTypeID
    ' lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "")
        
    'End If

End Sub

Private Sub TxtSearchCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
bo = False
Load FrmAssestSearch
            FrmAssestSearch.show
            
End If
End Sub

'Private Sub XPDtbTrans_Change()
'
''    If Trim(TxtNoteSerial1.text) <> "" Then
 '       oldtxtNoteSerial1.text = TxtNoteSerial1.text
 '   End If
'
'    TxtNoteSerial.text = ""
'    TxtNoteSerial1.text = ""

'End Sub

'Private Sub dcBranch_Click(Area As Integer)
'
'    TxtNoteSerial.text = ""
'    TxtNoteSerial1.text = ""
'End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

   

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
     Dcombos.GetUsers Me.DCboUserName
  
    Dcombos.GetBranches Me.Dcbranch
    ' Dcombos.GetAssestsOfEmp Me.dcmboassest
    ' Dcombos.GetEmpDepartments Me.DcmbFromDepart
   '  Dcombos.GetEmpDepartments Me.DcmbToDepart
'     Dcombos.GetEmployees Me.DcmbManagerID
    ' Dcombos.GetEmpJobsTypes Me.DcboJobsType
'   Dcombos.GetEmpJobsTypes Me.DcmbToJob
   ' Dcombos.GetEmpLocations Me.dcmbFromProject ' location
 '  Dcombos.GetEmpLocations Me.dcmbToProject ' location
   
    Dcombos.GetAssests Me.dcmboassest
    loadcombo
    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If
     SetDtpickerDate Me.XPDtbTrans
    SetDtpickerDate Me.DriveDate
    
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblEmpAsest    "
    StrSQL = StrSQL & "  where  (BranchID=0 or BranchID is null or         BranchID in(" & Current_branchSql & "))"
    StrSQL = StrSQL & " and FlgCar is null"
    StrSQL = StrSQL & " Order By EmpAsID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
        Me.TxtModFlg.Text = "R"
    Retrive
Check_Car
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
VSFlexGrid1.Enabled = True
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    
    Exit Sub

ErrTrap:
End Sub
Sub loadcombo()
    Dim str As String
    Dim Dcombos As New ClsDataCombos
    Dim RsOptions As ADODB.Recordset
    Set RsOptions = New ADODB.Recordset
    Dim optionStr As String
    Set Dcombos = New ClsDataCombos
    
    optionStr = "select * from TblOptions"
    RsOptions.Open optionStr, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If chkCar.value = vbChecked Then
        If RsOptions("ShowDriverOnly") = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
                str = str & "                   dbo.TblEmployee.Emp_Namee"
            Else
                str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
                str = str & "                   dbo.TblEmployee.Emp_Name"
            End If
            str = str & "    FROM         dbo.TblEmployee LEFT OUTER JOIN"
            str = str & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
            str = str & "     where  ( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1) "
   
            fill_combo DcboEmpNameTo, str
            fill_combo DcboEmpName, str
        Else
            Dcombos.GetEmployees Me.DcboEmpNameTo
            Dcombos.GetEmployees Me.DcboEmpName
        End If
    Else
        Dcombos.GetEmployees Me.DcboEmpNameTo
        Dcombos.GetEmployees Me.DcboEmpName
    End If
End Sub
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    'Label1.Visible = False
    lbl(15).Caption = "Emp"
    lbl(14).Caption = "Screen transfer and delivery Of the covenant Staff help this screen from the era of employee transfer to another review with the actual quantity of existing"
    Me.RdType.Caption = "Drive Assest"
    Me.RdTypeMov.Caption = "Move Assest"
    Me.RdType.RightToLeft = False
    Me.RdTypeMov.RightToLeft = False
    lbl(66).Caption = "Working No."
    Accredit.Caption = "Send Approve"
    BtImage.Caption = "Select"
    Label5.Caption = "Car Form"
    Crane.Caption = "Crane"
    C1Elastic10.Caption = "Car Attachments "
    Stickers.Caption = "Stickers"
    InnerLights.Caption = "Inner Light"
    Pedals.Caption = "Pedals"
    SpareTyre.Caption = "Spare Tyre"
    Battery.Caption = "Battery"
    CoverKey.Caption = "Tires key"
    Anntena.Caption = "Airy"
    Guarantee.Caption = "Guarantee"
    driverMirror.Caption = "Driver Mirror"
    C1Elastic11.Caption = "Cars Data"
    Label6.Caption = "Delegating Lead"
    SunScreens.Caption = "Sun Shades"
    sideMirror.Caption = "Side Mirrors"
    Recorder.Caption = "Recorder"
lbl(10).Caption = "Quantity"
cleaner.Caption = "Windshield Wiper"
Label7.Caption = "Periodic Inspection"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(14).Caption = "Add"
 Cmd(15).Caption = "Delete"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
    Cmd(9).Caption = "Prient"

    Me.Caption = "Assest Transfer"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lbl(3).Caption = "Employee"
    lbl(2).Caption = "Assest"
    lbl(5).Caption = "Branch"
    chkCar.RightToLeft = False
    chkCar.Caption = "Cars"
    lbl(18).Caption = "Plate No. "
    lbl(16).Caption = "Delivery Date"
    'Fra(0).Caption = "payments Method"
    lbl(13).Caption = "Date Drive"
    lbl(9).Caption = "Remarks"
XPTab301.TabCaption(0) = "Data"
XPTab301.TabCaption(2) = "Cars Data"
 '   Cmd(8).Caption = "Calc Dates"
   ' ChkSaleryDis.Caption = "Auto Discount"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"

    With Me.VSFlexGrid1
      .TextMatrix(0, .ColIndex("EmpName")) = "Name"
      
                    
                
                .TextMatrix(0, .ColIndex("ApprovDate")) = "Quantity"
                
                .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
        .TextMatrix(0, .ColIndex("ser")) = "NO"
        

    End With




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

'Private Sub TxtAdvanceValue_LostFocus()
  '  Dim StrSQL As String
  '  Dim Mytot As String
  '  Dim MySal As String
  '  Exit Sub
 '   Dim Myrs As New ADODB.Recordset
    'StrSQL =
   ' Myrs.Open "SELECT * From TblEmployee  where EmpID=" & val(DcboEmpName.BoundText), Cn, adOpenStatic, adLockReadOnly

   ' If Not Myrs.EOF And Not IsNull(Myrs!Emp_Salary) Then
    '    MySal = Myrs!Emp_Salary
     '   Mytot = val(MySal) * 5
'
      '  If val(TxtAdvanceValue.text) >= Mytot Then
        '    MsgBox "⁄›Ê« «·”·›…  ⁄œ  «·Õœ  «·„”„ÊÕ »Â ÊÂÊ 5 «÷⁄«› ﬁÌ„Â «·—« »  " & Chr(13) & "   —« » «·„ÊŸ›    " & MySal, vbOKOnly, App.Title
        '    Exit Sub
   
  '      End If
  '
  '  End If
   
'End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰"
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
          '  TxtAdvanceValue.Locked = True
            Me.DcboBox.locked = True
            XPDtbTrans.Enabled = False
            Cmd(15).Enabled = False
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

        Case "N"
         Cmd(15).Enabled = True
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '      Me.XPBtnMove(0).Enabled = False
            '      Me.XPBtnMove(1).Enabled = False
            '      Me.XPBtnMove(2).Enabled = False
            '      Me.XPBtnMove(3).Enabled = False
          '  TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
         Cmd(15).Enabled = True
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
         '   TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

  
Private Sub TxtSearchCode1_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode1.Text, EmpID
        Me.DcboEmpNameTo.BoundText = EmpID
    End If
End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With Me.VSFlexGrid1

      
        Select Case .ColKey(Col)
            
            
            Case "Price", "ApprovDate"
                .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("Price"))) * val(.TextMatrix(Row, .ColIndex("ApprovDate")))
       End Select

    End With

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With Me.VSFlexGrid1

      
        Select Case .ColKey(Col)
            
            
            Case "ApprovDate"
            
               VSFlexGrid1.ComboList = ""
             Case "EmpName"
            
               Cancel = True
              Case "Remarks"
            
               Cancel = True
                 Case "ser"
                
               Cancel = True
               Case "Price"
                If RdTypeMov Then
                    Cancel = True
                Else
                    .EditMaxLength = 10
                End If
        End Select

    End With
End Sub

Private Sub VSFlexGrid1_Click()
Clear_CarData

With VSFlexGrid1
        If .TextMatrix(.Row, .ColIndex("id")) <> "" Then
        
        FormOrignal.Text = .TextMatrix(.Row, .ColIndex("FormOrignal"))
        authorizeLicense.Text = .TextMatrix(.Row, .ColIndex("authorizeLicense"))
        authorizeExamination.Text = .TextMatrix(.Row, .ColIndex("authorizeExamination"))
        
        cleaner.value = val(.TextMatrix(.Row, .ColIndex("cleaner")))
        sideMirror.value = val(.TextMatrix(.Row, .ColIndex("sideMirror")))
        driverMirror.value = val(.TextMatrix(.Row, .ColIndex("driverMirror")))
        InnerLights.value = val(.TextMatrix(.Row, .ColIndex("InnerLights")))
        Pedals.value = val(.TextMatrix(.Row, .ColIndex("Pedals")))
        SunScreens.value = val(.TextMatrix(.Row, .ColIndex("SunScreens")))
        Recorder.value = val(.TextMatrix(.Row, .ColIndex("Recorder")))
        Anntena.value = val(.TextMatrix(.Row, .ColIndex("Anntena")))
        
        Battery.value = val(.TextMatrix(.Row, .ColIndex("Battery")))
        SpareTyre.value = val(.TextMatrix(.Row, .ColIndex("SpareTyre")))
        Crane.value = val(.TextMatrix(.Row, .ColIndex("Crane")))
        CoverKey.value = val(.TextMatrix(.Row, .ColIndex("CoverKey")))
        Guarantee.value = val(.TextMatrix(.Row, .ColIndex("Guarantee")))
        Stickers.value = val(.TextMatrix(.Row, .ColIndex("Stickers")))
 
        End If
End With

End Sub

'End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
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
   ' Dim RsDetails As ADODB.Recordset
                   Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 1
    Dim i As Integer
    Dim StrSQL As String
  Dim RsDetails As New ADODB.Recordset
    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.Find "EmpAsID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
 If rs("TypeAsset").value = True Then
 Me.RdTypeMov.value = True
 
        Else
        Me.RdType.value = True
        End If
        Me.DcboEmpNameTo.BoundText = val(IIf(IsNull(rs("ToEmId").value), "", rs("ToEmId").value))
    XPTxtID.Text = IIf(IsNull(rs("EmpAsID").value), "", val(rs("EmpAsID").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    Dcbranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    DcboEmpName.BoundText = val(IIf(IsNull(rs("EmpAsestID").value), "", rs("EmpAsestID").value))
    DriveDate.value = IIf(IsNull(rs("DriveDate").value), Date, rs("DriveDate").value)
      
      chkCar.value = IIf(IsNull(rs("ISCar").value), False, rs("ISCar").value)
      TxtOperatorN.Text = IIf(IsNull(rs("OperatorN").value), "", rs("OperatorN").value)
      TxtBoardNO.Text = IIf(IsNull(rs("BoardNO").value), "", rs("BoardNO").value)
      
      
  '  dcmboassest.BoundText = val(IIf(IsNull(rs("AsID").value), "", rs("AsID").value))
  '  txtreson.text = IIf(IsNull(rs("remark").value), "", rs("remark").value)

'''''''''''''''''''''''''''''''''''
  

 StrSQL = " SELECT     dbo.TblAssestes.AsName,dbo.TblAssestes.Price Price2, dbo.TblAssestes.AsID, dbo.TblAssestes.AsDes, dbo.TblEmpAsestDetails.Remarks, dbo.TblEmpAsestDetails.Qunt, "
 StrSQL = StrSQL & "                     dbo.TblEmpAsest.EmpAsID , dbo.TblEmpAsestDetails.IDAseset, dbo.TblEmpAsestDetails.diff, dbo.TblEmpAsestDetails.FlagAs"
 
 StrSQL = StrSQL & "  ,TblEmpAsestDetails.formorignal ,TblEmpAsestDetails.Price, TblEmpAsestDetails.authorizeLicense , TblEmpAsestDetails.authorizeExamination , TblEmpAsestDetails.cleaner , TblEmpAsestDetails.sideMirror ,"
 StrSQL = StrSQL & "  TblEmpAsestDetails.driverMirror , TblEmpAsestDetails.InnerLights , TblEmpAsestDetails.Pedals , TblEmpAsestDetails.SunScreens , TblEmpAsestDetails.Recorder , TblEmpAsestDetails.Anntena ,"
 StrSQL = StrSQL & "  TblEmpAsestDetails.Battery , TblEmpAsestDetails.SpareTyre , TblEmpAsestDetails.Crane , TblEmpAsestDetails.CoverKey, TblEmpAsestDetails.Guarantee , TblEmpAsestDetails.Stickers   "
 StrSQL = StrSQL & "  ,TblEmpAsestDetails.subcar1 ,TblEmpAsestDetails.subcar2,TblEmpAsestDetails.subcar3,TblEmpAsestDetails.subcar4,TblEmpAsestDetails.subcar5,TblEmpAsestDetails.subcar6,"
 StrSQL = StrSQL & "              TblEmpAsestDetails.subcar7 , TblEmpAsestDetails.subcar8, TblEmpAsestDetails.subcar9, TblEmpAsestDetails.subcar10, TblEmpAsestDetails.subcar11, TblEmpAsestDetails.subcar12, TblEmpAsestDetails.subcar13, TblEmpAsestDetails.subcar14"
 StrSQL = StrSQL & " FROM         dbo.TblAssestes INNER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmpAsestDetails ON dbo.TblAssestes.AsID = dbo.TblEmpAsestDetails.AsID INNER JOIN"
 StrSQL = StrSQL & "                   dbo.TblEmpAsest ON dbo.TblEmpAsestDetails.IDAseset = dbo.TblEmpAsest.EmpAsID"
 
 '   If Me.RdTypeMov.value = True Then
'StrSQL = StrSQL & " Where (dbo.TblEmpAsestDetails.FlagAs Is Null) And (dbo.TblEmpAsestDetails.EmpID = " & val(Me.DcboEmpNameTo.BoundText) & ")"
'Else
'StrSQL = StrSQL & " Where (dbo.TblEmpAsestDetails.FlagAs Is Null) And (dbo.TblEmpAsestDetails.EmpID = " & val(Me.DcboEmpName.BoundText) & ")"
'End If
StrSQL = StrSQL & "  Where (dbo.TblEmpAsestDetails.FlagAs Is Null) and (dbo.TblEmpAsestDetails.IDAseset = " & val(Me.XPTxtID.Text) & ")"
    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  With Me.VSFlexGrid1
     .Clear flexClearScrollable, flexClearEverything
     .Rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
         .Rows = .FixedRows + RsDetails.RecordCount

          
          
        For i = .FixedRows To .Rows - 1
             .TextMatrix(i, .ColIndex("ser")) = i
             .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(RsDetails("AsName").value), "", RsDetails("AsName").value)
               .TextMatrix(i, .ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("Qunt").value), "", RsDetails("Qunt").value)
                If val(RsDetails!Price & "") = 0 Then
                    .TextMatrix(i, .ColIndex("Price")) = val(RsDetails!Price2 & "")
                Else
                    .TextMatrix(i, .ColIndex("Price")) = val(RsDetails!Price & "")
                End If
                
                .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("Price"))) * val(.TextMatrix(i, .ColIndex("ApprovDate")))
               .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDetails("AsID").value), "", RsDetails("AsID").value)
                .TextMatrix(i, .ColIndex("diff")) = IIf(IsNull(RsDetails("diff").value), "", RsDetails("diff").value)
            .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks").value), "", RsDetails("Remarks").value)
            
            
            .TextMatrix(i, .ColIndex("FormOrignal")) = IIf(IsNull(RsDetails("FormOrignal").value), "", RsDetails("FormOrignal").value)
            .TextMatrix(i, .ColIndex("authorizeLicense")) = IIf(IsNull(RsDetails("authorizeLicense").value), "", RsDetails("authorizeLicense").value)
            .TextMatrix(i, .ColIndex("authorizeExamination")) = IIf(IsNull(RsDetails("authorizeExamination").value), "", RsDetails("authorizeExamination").value)
            
            .TextMatrix(i, .ColIndex("cleaner")) = IIf(IsNull(RsDetails("cleaner").value), False, RsDetails("cleaner").value)
            .TextMatrix(i, .ColIndex("sideMirror")) = IIf(IsNull(RsDetails("sideMirror").value), False, RsDetails("sideMirror").value)
            .TextMatrix(i, .ColIndex("driverMirror")) = IIf(IsNull(RsDetails("driverMirror").value), False, RsDetails("driverMirror").value)
            .TextMatrix(i, .ColIndex("InnerLights")) = IIf(IsNull(RsDetails("InnerLights").value), False, RsDetails("InnerLights").value)
            .TextMatrix(i, .ColIndex("Pedals")) = IIf(IsNull(RsDetails("Pedals").value), False, RsDetails("Pedals").value)
            .TextMatrix(i, .ColIndex("SunScreens")) = IIf(IsNull(RsDetails("SunScreens").value), False, RsDetails("SunScreens").value)
            .TextMatrix(i, .ColIndex("Recorder")) = IIf(IsNull(RsDetails("Recorder").value), False, RsDetails("Recorder").value)
            .TextMatrix(i, .ColIndex("Anntena")) = IIf(IsNull(RsDetails("Anntena").value), False, RsDetails("Anntena").value)
            
            .TextMatrix(i, .ColIndex("Battery")) = IIf(IsNull(RsDetails("Battery").value), False, RsDetails("Battery").value)
            .TextMatrix(i, .ColIndex("SpareTyre")) = IIf(IsNull(RsDetails("SpareTyre").value), False, RsDetails("SpareTyre").value)
            .TextMatrix(i, .ColIndex("Crane")) = IIf(IsNull(RsDetails("Crane").value), False, RsDetails("Crane").value)
            .TextMatrix(i, .ColIndex("CoverKey")) = IIf(IsNull(RsDetails("CoverKey").value), False, RsDetails("CoverKey").value)
            .TextMatrix(i, .ColIndex("Guarantee")) = IIf(IsNull(RsDetails("Guarantee").value), False, RsDetails("Guarantee").value)
            .TextMatrix(i, .ColIndex("Stickers")) = IIf(IsNull(RsDetails("Stickers").value), False, RsDetails("Stickers").value)


.TextMatrix(i, .ColIndex("subcar1")) = IIf(IsNull(RsDetails("subcar1").value), 0, RsDetails("subcar1").value)
.TextMatrix(i, .ColIndex("subcar2")) = IIf(IsNull(RsDetails("subcar2").value), 0, RsDetails("subcar2").value)
.TextMatrix(i, .ColIndex("subcar3")) = IIf(IsNull(RsDetails("subcar3").value), 0, RsDetails("subcar3").value)
.TextMatrix(i, .ColIndex("subcar4")) = IIf(IsNull(RsDetails("subcar4").value), 0, RsDetails("subcar4").value)
.TextMatrix(i, .ColIndex("subcar5")) = IIf(IsNull(RsDetails("subcar5").value), 0, RsDetails("subcar5").value)
.TextMatrix(i, .ColIndex("subcar6")) = IIf(IsNull(RsDetails("subcar6").value), 0, RsDetails("subcar6").value)
.TextMatrix(i, .ColIndex("subcar7")) = IIf(IsNull(RsDetails("subcar7").value), 0, RsDetails("subcar7").value)
.TextMatrix(i, .ColIndex("subcar8")) = IIf(IsNull(RsDetails("subcar8").value), 0, RsDetails("subcar8").value)
.TextMatrix(i, .ColIndex("subcar9")) = IIf(IsNull(RsDetails("subcar9").value), 0, RsDetails("subcar9").value)
.TextMatrix(i, .ColIndex("subcar10")) = IIf(IsNull(RsDetails("subcar10").value), 0, RsDetails("subcar10").value)
.TextMatrix(i, .ColIndex("subcar11")) = IIf(IsNull(RsDetails("subcar11").value), 0, RsDetails("subcar11").value)
.TextMatrix(i, .ColIndex("subcar12")) = IIf(IsNull(RsDetails("subcar12").value), 0, RsDetails("subcar12").value)
.TextMatrix(i, .ColIndex("subcar13")) = IIf(IsNull(RsDetails("subcar13").value), 0, RsDetails("subcar13").value)
.TextMatrix(i, .ColIndex("subcar14")) = IIf(IsNull(RsDetails("subcar14").value), 0, RsDetails("subcar14").value)

            
           
            
            RsDetails.MoveNext
        Next i

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
  
'''''''''''''''''''''''
   
 
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
   '    If IsNull(rs("posted").value) Then
   '                                                If SystemOptions.UserInterface = ArabicInterface Then
   ''                                                 Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
    ''                                              Else
  '                                                  Accredit.Caption = " send to Approval   "
  '                                             End If
  '                                             Accredit.Enabled = True
  'Else
  '                                                 If SystemOptions.UserInterface = ArabicInterface Then
  '                                                  Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
  '                                                Else
  '                                                  Accredit.Caption = " sent to Approval   "
  '                                             End If
  '                                             Accredit.Enabled = False
  ' End If
   
   
      
  '  fillapprovData
    
       
      
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Public Sub Retrive2(Optional Lngid As Long = 0)
                Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 1
   ' Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
  Dim RsDetails As New ADODB.Recordset
    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.Find "EmpAsID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    XPTxtID.Text = IIf(IsNull(rs("EmpAsID").value), "", val(rs("EmpAsID").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    Dcbranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
      DcboEmpName.BoundText = val(IIf(IsNull(rs("EmpAsestID").value), "", rs("EmpAsestID").value))
  '  dcmboassest.BoundText = val(IIf(IsNull(rs("AsID").value), "", rs("AsID").value))
  '  txtreson.text = IIf(IsNull(rs("remark").value), "", rs("remark").value)

'''''''''''''''''''''''''''''''''''
'
'
'StrSQL = "   SELECT     dbo.TblAssestes.AsName, dbo.TblAssestes.AsID, dbo.TblAssestes.AsDes, dbo.TblEmpAsestDetails.Remarks, dbo.TblEmpAsestDetails.Qunt,"
' StrSQL = StrSQL & "                     dbo.TblEmpAsest.EmpAsID , dbo.TblEmpAsestDetails.IDAseset"
' StrSQL = StrSQL & "   FROM         dbo.TblAssestes INNER JOIN"
' StrSQL = StrSQL & "                        dbo.TblEmpAsestDetails ON dbo.TblAssestes.AsID = dbo.TblEmpAsestDetails.AsID INNER JOIN"
'   StrSQL = StrSQL & "                      dbo.TblEmpAsest ON dbo.TblEmpAsestDetails.IDAseset = dbo.TblEmpAsest.EmpAsID"
'StrSQL = StrSQL & "  Where (dbo.TblEmpAsestDetails.IDAseset = " & val(Me.XPTxtID.text) & ")"
    
 '   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 ' With Me.VSFlexGrid1
 '    .Clear flexClearScrollable, flexClearEverything
 '    .Rows = .FixedRows

 '   If Not (RsDetails.BOF Or RsDetails.EOF) Then
''        RsDetails.MoveFirst
'         .Rows = .FixedRows + RsDetails.RecordCount
'
  '      For i = .FixedRows To .Rows - 1
  '           .TextMatrix(i, .ColIndex("ser")) = i
  '           .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(RsDetails("AsName").value), "", RsDetails("AsName").value)
  '             .TextMatrix(i, .ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("Qunt").value), "", RsDetails("Qunt").value)
  '
 ''           .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks").value), "", RsDetails("Remarks").value)
 '
 '           RsDetails.MoveNext
''        Next i
'
'    End If
'End With
'    RsDetails.Close
'    Set RsDetails = Nothing
  
'''''''''''''''''''''''
   
 
'    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
'       If IsNull(rs("posted").value) Then
'                                                   If SystemOptions.UserInterface = ArabicInterface Then
'                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
'                                                  Else
'                                                    Accredit.Caption = " send to Approval   "
'                                               End If
'                                               Accredit.Enabled = True
'  Else
''                                                   If SystemOptions.UserInterface = ArabicInterface Then
 '                                                   Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
 '                                                 Else
 '                                                   Accredit.Caption = " sent to Approval   "
 '                                              End If
 '                                              Accredit.Enabled = False
 '  End If
 '
   
      
'    fillapprovData
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Public Sub retrive1(Optional Lngid As Long = 0)
                Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 1
   ' Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
  Dim RsDetails As New ADODB.Recordset
    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.Find "EmpAsID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

 If rs("TypeAsset").value = True Then
 Me.RdTypeMov.value = True
 
        Else
        Me.RdType.value = True
        End If
        Me.DcboEmpNameTo.BoundText = val(IIf(IsNull(rs("ToEmId").value), "", rs("ToEmId").value))
    XPTxtID.Text = IIf(IsNull(rs("EmpAsID").value), "", val(rs("EmpAsID").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    Dcbranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
      DcboEmpName.BoundText = val(IIf(IsNull(rs("EmpAsestID").value), "", rs("EmpAsestID").value))
  '  dcmboassest.BoundText = val(IIf(IsNull(rs("AsID").value), "", rs("AsID").value))
   ' dcmboassest.BoundText = val(IIf(IsNull(rs("AsID").value), "", rs("AsID").value))
   ' txtreson.text = IIf(IsNull(rs("remark").value), "", rs("remark").value)

'''''''''''''''''''''''''''''''''''
   

StrSQL = " SELECT     dbo.TblAssestes.AsName, dbo.TblAssestes.AsID, dbo.TblAssestes.AsDes, dbo.TblEmpAsestDetails.Remarks, dbo.TblEmpAsestDetails.Qunt, "
 StrSQL = StrSQL & "                     dbo.TblEmpAsest.EmpAsID , dbo.TblEmpAsestDetails.IDAseset, dbo.TblEmpAsestDetails.diff, dbo.TblEmpAsestDetails.FlagAs"
StrSQL = StrSQL & " FROM         dbo.TblAssestes INNER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmpAsestDetails ON dbo.TblAssestes.AsID = dbo.TblEmpAsestDetails.AsID INNER JOIN"
   StrSQL = StrSQL & "                   dbo.TblEmpAsest ON dbo.TblEmpAsestDetails.IDAseset = dbo.TblEmpAsest.EmpAsID"
   ' If Me.RdTypeMov.value = True Then
'StrSQL = StrSQL & " Where (dbo.TblEmpAsestDetails.FlagAs Is Null) And (dbo.TblEmpAsestDetails.EmpID = " & val(Me.DcboEmpNameTo.BoundText) & ")"
'Else
'StrSQL = StrSQL & " Where (dbo.TblEmpAsestDetails.FlagAs Is Null) And (dbo.TblEmpAsestDetails.EmpID = " & val(Me.DcboEmpName.BoundText) & ")"
'End If
StrSQL = StrSQL & "  Where (dbo.TblEmpAsestDetails.FlagAs Is Null) and (dbo.TblEmpAsestDetails.IDAseset = " & val(Me.XPTxtID.Text) & ")"

    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  With Me.VSFlexGrid1
     .Clear flexClearScrollable, flexClearEverything
     .Rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
         .Rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .Rows - 1
             .TextMatrix(i, .ColIndex("ser")) = i
             .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(RsDetails("AsName").value), "", RsDetails("AsName").value)
               .TextMatrix(i, .ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("Qunt").value), "", RsDetails("Qunt").value)
                
            .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks").value), "", RsDetails("Remarks").value)
            
            RsDetails.MoveNext
        Next i

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
  
'''''''''''''''''''''''
   
 
 '   Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
 '      If IsNull(rs("posted").value) Then
 ''                                                  If SystemOptions.UserInterface = ArabicInterface Then
  '                                                  Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
  '                                                Else
  '                                                  Accredit.Caption = " send to Approval   "
  '                                             End If
  '                                             Accredit.Enabled = True
  'Else
  '                                                 If SystemOptions.UserInterface = ArabicInterface Then
  '                                                  Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
  '                                                Else
  '                                                  Accredit.Caption = " sent to Approval   "
  '                                             End If
  '                                             Accredit.Enabled = False
  ' End If
  '
  '
  '
  '  fillapprovData
  '
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Private Sub SaveData()
    Dim EmpID As Double
    Dim FixedassetId As Double
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String

   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
        If Me.DcboEmpName.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ› «·„‰ﬁÊ· ⁄Âœ Â..!! "
         Else
         Msg = "Please Select Employee"
         End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            DcboEmpName.SetFocus
             SendKeys "{F4}"
            Exit Sub
        End If

    If Me.RdTypeMov.value = True Then
        If Me.DcboEmpNameTo.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ› «·„‰ﬁÊ· «·ÌÂ «·⁄Âœ… ..!! "
         Else
         Msg = "Please Select Employee"
        End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboEmpName.SetFocus
            'SendKeys "{F4}"
            Exit Sub
        End If
End If
       

      ''  If CheckDate = False Then
        '    Exit Sub
      '  End If
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then

            XPTxtID.Text = CStr(new_id("TblEmpAsest", "EmpAsID", "", True))
         
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
      
            StrSQL = "Delete From TblEmpAsestDetails Where IDAseset=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords

        End If
        If Me.RdTypeMov.value = True Then
        rs("TypeAsset").value = 1
        rs("DeliverDate").value = DBIssueDate.value
        End If
 rs("EmpAsID").value = val(XPTxtID.Text)
      rs("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
       rs("ToEmId").value = IIf(Me.DcboEmpNameTo.BoundText = "", Null, Me.DcboEmpNameTo.BoundText)
      '  rs("ToDepart").value = IIf(Me.DcmbToDepart.BoundText = "", Null, Me.DcmbToDepart.BoundText)
        
        rs("RecordDate").value = XPDtbTrans.value
        rs("EmpAsestID").value = IIf(Me.DcboEmpName.BoundText = "", Null, Me.DcboEmpName.BoundText)
'        rs("FullCodAse").value = TxtSearchCode.text
     '   rs("ManagerID").value = val(Me.DcmbManagerID.BoundText)
      '  rs("JobID").value = val(Me.DcboJobsType.BoundText)
      rs("OperatorN").value = TxtOperatorN.Text
      rs("BoardNO").value = TxtBoardNO.Text
      
        rs("DriveDate").value = Me.DriveDate.value
        rs("PostedDate").value = Me.DBIssueDate.value
      '  rs("JobTo").value = val(Me.DcmbToJob.BoundText)
      '  rs("ProjectTo").value = val(Me.dcmbToProject.BoundText)
      '  rs("ProjectFrom").value = val(Me.dcmbFromProject.BoundText)
        rs("remark").value = Me.txtreson.Text
        rs("AsID").value = val(Me.dcmboassest.BoundText)
        rs("ISCar").value = chkCar.value
        
      rs.update
        Cn.CommitTrans
        BeginTrans = False
          Set RsDetails = New ADODB.Recordset
        RsDetails.Open "TblEmpAsestDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
 If RdType.value = True Then
EmpID = val(DcboEmpName.BoundText)
ElseIf RdTypeMov.value = True Then
EmpID = val(DcboEmpNameTo.BoundText)
End If
With Me.VSFlexGrid1
                    For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, .ColIndex("EmpName")) <> "" Then
                        RsDetails.AddNew
                             RsDetails("IDAseset").value = val(XPTxtID.Text)
                             RsDetails("AsID").value = val(.TextMatrix(i, .ColIndex("id")))
                             RsDetails("diff").value = val(.TextMatrix(i, .ColIndex("diff")))
                     RsDetails("Qunt").value = .TextMatrix(i, .ColIndex("ApprovDate"))
                     If Me.RdTypeMov.value = True Then
                     RsDetails("EmpID").value = val(Me.DcboEmpNameTo.BoundText)
                     Else
                     RsDetails("EmpID").value = val(Me.DcboEmpName.BoundText)
                     End If
                       RsDetails("Remarks").value = .TextMatrix(i, .ColIndex("Remarks"))
                        
                    RsDetails("FormOrignal").value = .TextMatrix(i, .ColIndex("FormOrignal"))
                    RsDetails("authorizeLicense").value = .TextMatrix(i, .ColIndex("authorizeLicense"))
                    RsDetails("authorizeExamination").value = .TextMatrix(i, .ColIndex("authorizeExamination"))
                    
                    RsDetails("cleaner").value = val(.TextMatrix(i, .ColIndex("cleaner")))
                    RsDetails("sideMirror").value = val(.TextMatrix(i, .ColIndex("sideMirror")))
                    RsDetails("driverMirror").value = val(.TextMatrix(i, .ColIndex("driverMirror")))
                    RsDetails("InnerLights").value = val(.TextMatrix(i, .ColIndex("InnerLights")))
                    RsDetails("Pedals").value = val(.TextMatrix(i, .ColIndex("Pedals")))
                    RsDetails("SunScreens").value = val(.TextMatrix(i, .ColIndex("SunScreens")))
                    RsDetails("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
                    RsDetails("Recorder").value = val(.TextMatrix(i, .ColIndex("Recorder")))
                    RsDetails("Anntena").value = val(.TextMatrix(i, .ColIndex("Anntena")))
                    RsDetails("Battery").value = val(.TextMatrix(i, .ColIndex("Battery")))
                    RsDetails("SpareTyre").value = val(.TextMatrix(i, .ColIndex("SpareTyre")))
                    RsDetails("Crane").value = val(.TextMatrix(i, .ColIndex("Crane")))
                    RsDetails("CoverKey").value = val(.TextMatrix(i, .ColIndex("CoverKey")))
                    RsDetails("Guarantee").value = val(.TextMatrix(i, .ColIndex("Guarantee")))
                    RsDetails("Stickers").value = val(.TextMatrix(i, .ColIndex("Stickers")))
                    RsDetails("subcar1").value = val(.TextMatrix(i, .ColIndex("subcar1")))
                    RsDetails("subcar2").value = val(.TextMatrix(i, .ColIndex("subcar2")))
                    RsDetails("subcar3").value = val(.TextMatrix(i, .ColIndex("subcar3")))
                    RsDetails("subcar4").value = val(.TextMatrix(i, .ColIndex("subcar4")))
                    RsDetails("subcar5").value = val(.TextMatrix(i, .ColIndex("subcar5")))
                    RsDetails("subcar6").value = val(.TextMatrix(i, .ColIndex("subcar6")))
                    RsDetails("subcar7").value = val(.TextMatrix(i, .ColIndex("subcar7")))
                    RsDetails("subcar8").value = val(.TextMatrix(i, .ColIndex("subcar8")))
                    RsDetails("subcar9").value = val(.TextMatrix(i, .ColIndex("subcar9")))
                    RsDetails("subcar10").value = val(.TextMatrix(i, .ColIndex("subcar10")))
                    RsDetails("subcar11").value = val(.TextMatrix(i, .ColIndex("subcar11")))
                    RsDetails("subcar12").value = val(.TextMatrix(i, .ColIndex("subcar12")))
                    RsDetails("subcar13").value = val(.TextMatrix(i, .ColIndex("subcar13")))
                    RsDetails("subcar14").value = val(.TextMatrix(i, .ColIndex("subcar14")))
                        RsDetails.update
                    If chkCar.value = vbChecked Then
                                      GetCardID FixedassetId, val(.TextMatrix(i, .ColIndex("id"))), 0
                   Cn.Execute "Update TblCarsData set FlagFixedasset=1,Emp_id =" & EmpID & " where fixedAssetid=" & FixedassetId & ""
                   End If
 
                        End If
                    Next i
 End With
 If Me.RdTypeMov.value = True Then
 Dim sql As String
RsDetails.Close

Set RsDetails = Nothing
Set RsDetails = New ADODB.Recordset
          Set RsDetails = New ADODB.Recordset
        RsDetails.Open "TblEmpAsestDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
      Dim str As String
  If SystemOptions.UserInterface = ArabicInterface Then
 str = " „ ‰ﬁ· Â–Â «·⁄ÂœÂ «·Ï «·„ÊŸ› "
 str = str & Me.DcboEmpNameTo.Text
 str = str & "» √—ÌŒ"
 str = str & DBIssueDate.value
 Else
  str = "Move Assest to Employee "
 str = str & Me.DcboEmpNameTo.Text
 str = str & "Date"
 str = str & DBIssueDate.value

 End If
With Me.VSFlexGrid1
                    For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, .ColIndex("EmpName")) <> "" Then
                    If val(.TextMatrix(i, .ColIndex("diff"))) = 0 Then
                     sql = "update TblEmpAsestDetails set   Remark2 ='" & str & "'  WHERE     (dbo.TblEmpAsestDetails.FlagAs IS NULL) AND (dbo.TblEmpAsestDetails.EmpID = " & val(Me.DcboEmpName.BoundText) & ") AND (dbo.TblEmpAsestDetails.AsID = " & val(.TextMatrix(i, .ColIndex("id"))) & ")"
                                    Cn.Execute sql
     sql = "update TblEmpAsestDetails set   FlagAs =1  WHERE     (dbo.TblEmpAsestDetails.FlagAs IS NULL) AND (dbo.TblEmpAsestDetails.EmpID = " & val(Me.DcboEmpName.BoundText) & ") AND (dbo.TblEmpAsestDetails.AsID = " & val(.TextMatrix(i, .ColIndex("id"))) & ")"
                                    Cn.Execute sql
                              
                     End If
                                          If val(.TextMatrix(i, .ColIndex("diff"))) > 0 Then
                                                 sql = "update TblEmpAsestDetails set   Qunt =" & val(.TextMatrix(i, .ColIndex("ApprovDate"))) & "  WHERE     (dbo.TblEmpAsestDetails.FlagAs IS NULL) AND (dbo.TblEmpAsestDetails.EmpID = " & val(Me.DcboEmpName.BoundText) & ") AND (dbo.TblEmpAsestDetails.AsID = " & val(.TextMatrix(i, .ColIndex("id"))) & ")"
                                    Cn.Execute sql
                             sql = "update TblEmpAsestDetails set   Remark2 ='" & str & "'  WHERE     (dbo.TblEmpAsestDetails.FlagAs IS NULL) AND (dbo.TblEmpAsestDetails.EmpID = " & val(Me.DcboEmpName.BoundText) & ") AND (dbo.TblEmpAsestDetails.AsID = " & val(.TextMatrix(i, .ColIndex("id"))) & ")"
                                    Cn.Execute sql
                        sql = "update TblEmpAsestDetails set   FlagAs =1  WHERE     (dbo.TblEmpAsestDetails.FlagAs IS NULL) AND (dbo.TblEmpAsestDetails.EmpID = " & val(Me.DcboEmpName.BoundText) & ") AND (dbo.TblEmpAsestDetails.AsID = " & val(.TextMatrix(i, .ColIndex("id"))) & ")"
                                    Cn.Execute sql
                  
                        RsDetails.AddNew
                             RsDetails("IDAseset").value = val(.TextMatrix(i, .ColIndex("idas")))
                             RsDetails("AsID").value = val(.TextMatrix(i, .ColIndex("id")))
                     RsDetails("Qunt").value = .TextMatrix(i, .ColIndex("diff"))
                     RsDetails("EmpID").value = val(Me.DcboEmpName.BoundText)
                       RsDetails("Remarks").value = .TextMatrix(i, .ColIndex("Remarks"))
                        RsDetails.update
                
                        End If
                        End If
                    Next i
 End With
RsDetails.Close
Set RsDetails = Nothing

End If


'        RsDetails.Close
        Set RsDetails = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.Text

            Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
              Else
              Msg = "This is record already saved " & CHR(13)
              Msg = Msg & "You want eneter another record "
              End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
              Else
              MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
              End If
        End Select

        TxtModFlg.Text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
Else
    If Err.Number = -2147217900 Then
        Msg = "Can Not Save This  Data" & CHR(13)
        Msg = Msg + " have been insert false data " & CHR(13)
        Msg = Msg + "Make sure of the validity of the data and try again"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
 Else
 Msg = "Sorry ...error douring save"
 End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "EmpAsID='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String

    On Error GoTo ErrTrap

     If XPTxtID.Text <> "" Then
     If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
     Else
     Msg = "Confirm Delete"
     End If
Dim i As Integer
Dim FixedassetId As Double
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
               
                rs.MoveFirst
                    If chkCar.value = vbChecked Then
                    With VSFlexGrid1
                    For i = .FixedRows To .Rows - 1
                        GetCardID FixedassetId, val(.TextMatrix(i, .ColIndex("id"))), 0
                   Cn.Execute "Update TblCarsData set FlagFixedasset=null    where fixedAssetid=" & FixedassetId & ""
                   Next i
                   End With
                   End If
 StrSQL = "Delete From TblEmpAsestDetails Where IDAseset=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
                If rs.RecordCount < 1 Then
                
            Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
            End If
                If rs.RecordCount < 1 Then
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        Msg = "This process is not available "
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Sorry error douring delete"
    End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub



'Function FillApprovedTable()
' Dim RSApproval  As New ADODB.Recordset
''   Set RSApproval = New ADODB.Recordset
 '  Dim currentdate As Date
 ''  RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable


' Dim sql As String
'  Dim rs1 As New ADODB.Recordset
' Dim i As Integer
'    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
''  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
 ' sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
 '' sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
  'sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
'sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.name & "')"
'sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "
'
'    rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
''
 '   If rs1.RecordCount > 0 Then
 '           currentdate = Now
 '           For i = 1 To rs1.RecordCount
 ''             RSApproval.AddNew
  '              RSApproval("ScreenName").value = Me.name
  '              RSApproval("levelo").value = IIf(IsNull(rs1("levelo").value), Null, rs1("levelo").value)
  ''             RSApproval("EmpID").value = IIf(IsNull(rs1("EmpID").value), Null, rs1("EmpID").value)
   '             RSApproval("levelorder").value = IIf(IsNull(rs1("levelorder").value), Null, rs1("levelorder").value)
   '              RSApproval("currorder").value = IIf(IsNull(rs1("currorder").value), Null, rs1("currorder").value)
   '               RSApproval("Transaction_ID").value = val(Me.XPTxtID.text)
   '                RSApproval("NoteSerial").value = val(Me.XPTxtID.text)
   ''             RSApproval("Transaction_Date").value = Date
                
    '              RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.name), currentdate)
    '           RSApproval("SendTime").value = currentdate

    '             If i = 1 Then
    '                    RSApproval("Currcursor").value = 1
    '                     RSApproval("FromUser").value = user_name
    ''            End If
     '
     '           RSApproval.update
     '           rs1.MoveNext
     ''       Next i
'
'    End If
''
    

'End Function



'Function fillapprovData()
''Dim Num As Integer
 'Dim RsDetails As New ADODB.Recordset
 'Dim StrSQL As String
 ''
 
' StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
''StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
'StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
'StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
'StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
'StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
'StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.name & "')"
'StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"
'
'    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

' If Not (RsDetails.EOF Or RsDetails.BOF) Then
'        GRID2.Rows = RsDetails.RecordCount + 1
'

'        For Num = 1 To RsDetails.RecordCount
'
'       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
'    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
''   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
 '  Else
 ''   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
  '  End If
  '
  ''      GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
   ''        If SystemOptions.UserInterface = ArabicInterface Then
    '        GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
    ''      Else
     '        GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
     ''     End If
      '      If SystemOptions.UserInterface = ArabicInterface Then
      ''      GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
       '     Else
       '     GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
       ''     End If
        '    GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
        '  GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 ''
 '
'rsDetails.MoveNext
'if Num = RsDetails.RecordCount Then
'
'        If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
''                                If SystemOptions.UserInterface = ArabicInterface Then
'                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
'                                 Else
'                                       Label11.Caption = "Approved"
'                                 End If
'                            Label11.BackColor = &H80FF80
'        Else
'                             If SystemOptions.UserInterface = ArabicInterface Then
''                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
 '                           Else
 '                                    Label11.Caption = "Currently required Approve"
 '                           End If
 '                Label11.BackColor = &HFFFFC0
 '       End If

'End If

'        Next Num
'Else
' GRID2.Rows = 1
'    End If
'RsDetails.Close

'End Function

Private Sub txtLetter3_KeyPress(KeyAscii As Integer)
txtLetter3.Text = ""
If Len(txtLetter3.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter4.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.Text = "R" Then
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
        .Create Me.hWnd, "  ”·Ì„ ⁄Âœ «·„ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”·Ì„ ⁄Âœ «·„ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”·Ì„ ⁄Âœ «·„ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”·Ì„ ⁄Âœ «·„ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”·Ì„ ⁄Âœ «·„ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”·Ì„ ⁄Âœ «·„ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, " ”·Ì„ ⁄Âœ «·„ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”·Ì„ ⁄Âœ «·„ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”·Ì„ ⁄Âœ «·„ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”·Ì„ ⁄Âœ «·„ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”·Ì„ ⁄Âœ «·„ÊŸ›", 1, 15204351, -2147483630
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

    If Me.TxtModFlg.Text <> "R" Then

        Select Case Me.TxtModFlg.Text

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

 
 
Function print_report2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
'MySQL = MySQL & "     SELECT dbo.TblEmpAsest.EmpAsID, dbo.TblCarsData.CarsTypeId, dbo.TblBranchesData.branch_id, dbo.TblCarsData.Emp_id, dbo.TblCarsData.VColor, dbo.TblCarsData.LocationID,"
'MySQL = MySQL & "     dbo.TblCarsData.VModel, dbo.TblCarsData.id, dbo.TblCarsData.Branch_NO, dbo.TblCarsData.code, dbo.TblCarsData.Fullcode, dbo.TblCarsData.prifix,"
'MySQL = MySQL & "     dbo.TblCarsData.LicenseNO, dbo.TblCarsData.Name, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Model, dbo.TblCarsData.PurchaseDate, dbo.TblCarsData.LastKMCounter,"
'MySQL = MySQL & "     dbo.TblCarsData.LicenseExpireDate, dbo.TblCarsData.InsuranceCompanyId, dbo.TblCarsData.InsuranceExpireDate, dbo.TblCarsData.TestExpireDate,"
'MySQL = MySQL & "     dbo.TblCarsData.Notes, dbo.TblCarsData.LicenseExpireDateH, dbo.TblCarsData.InsuranceExpireDateH, dbo.TblCarsData.TestExpireDateH, dbo.TblCarsData.fixedAssetid,"
'MySQL = MySQL & "      dbo.TblCarsData.VehicleLong, dbo.TblCarsData.EquQty, dbo.TblCarsData.Capacity, dbo.TblCarsData.ContractID, dbo.TblCarsData.EndContractDate,"
'MySQL = MySQL & "      dbo.TblCarsData.SetCount, dbo.TblCarsData.Rate, dbo.TblCarsData.EndContractDateH, dbo.TblCarsData.Rep, dbo.TblCarsData.EndAllocationDate,"
'MySQL = MySQL & "      dbo.TblCarsData.MaxCap, dbo.TblCarsData.OperatorN, dbo.TblCarsData.EqupName, dbo.TblCarsData.TypeCar, dbo.TblCarsData.Gearno, dbo.TblCarsData.Gearno1,"
'MySQL = MySQL & "      dbo.TblCarsData.Machineno, dbo.TblCarsData.Machineno1, dbo.TblCarsData.VType, dbo.TblCarsData.Chesis, dbo.TblCarsData.Total, dbo.TblCarsData.LetterCount,"
'MySQL = MySQL & "      dbo.TblColor.name AS ColorName, dbo.TblCarModels.Model AS ModelName, dbo.EmpGroupDep.GroupName, dbo.TblBranchesData.branch_name,"
'MySQL = MySQL & "      dbo.TBLCarTypes.name AS TypeName, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS Emp_FullCode, dbo.TblEmpAsest.AsID, dbo.TblCarsData.LetterPrice,"
'MySQL = MySQL & "      dbo.TblEmpAsestDetails.CarsDataID, dbo.TblEmpAsestDetails.AsestName, dbo.TblEmpAsestDetails.TypeEqup, dbo.TblEmpAsestDetails.FlgCarNotFixed,"
'MySQL = MySQL & "      dbo.TblEmpAsestDetails.Stickers, dbo.TblEmpAsestDetails.Guarantee, dbo.TblEmpAsestDetails.CoverKey, dbo.TblEmpAsestDetails.Crane,"
'MySQL = MySQL & "      dbo.TblEmpAsestDetails.SpareTyre, dbo.TblEmpAsestDetails.Battery, dbo.TblEmpAsestDetails.Anntena, dbo.TblEmpAsestDetails.Recorder,"
'MySQL = MySQL & "      dbo.TblEmpAsestDetails.SunScreens, dbo.TblEmpAsestDetails.Pedals, dbo.TblEmpAsestDetails.InnerLights, dbo.TblEmpAsestDetails.driverMirror,"
'MySQL = MySQL & "      dbo.TblEmpAsestDetails.sideMirror, dbo.TblEmpAsestDetails.cleaner, dbo.TblEmpAsestDetails.FormOrignal, dbo.TblEmpAsestDetails.authorizeLicense,"
'MySQL = MySQL & "      dbo.TblEmpAsestDetails.authorizeExamination"
'
'MySQL = MySQL & " ,TblEmpAsestDetails.subcar1 ,TblEmpAsestDetails.subcar2,TblEmpAsestDetails.subcar3,TblEmpAsestDetails.subcar4,TblEmpAsestDetails.subcar5,TblEmpAsestDetails.subcar6,"
' MySQL = MySQL & "              TblEmpAsestDetails.subcar7 , TblEmpAsestDetails.subcar8, TblEmpAsestDetails.subcar9, TblEmpAsestDetails.subcar10, TblEmpAsestDetails.subcar11, TblEmpAsestDetails.subcar12, TblEmpAsestDetails.subcar13, TblEmpAsestDetails.subcar14"
'
'MySQL = MySQL & "      FROM     dbo.TblCarModels RIGHT OUTER JOIN"
'MySQL = MySQL & "      dbo.TblCarsData RIGHT OUTER JOIN"
'MySQL = MySQL & "      dbo.TblEmpAsest LEFT OUTER JOIN"
'MySQL = MySQL & "      dbo.TblEmpAsestDetails ON dbo.TblEmpAsest.EmpAsID = dbo.TblEmpAsestDetails.IDAseset LEFT OUTER JOIN"
'MySQL = MySQL & "      dbo.TblAssestes ON dbo.TblEmpAsestDetails.AsID = dbo.TblAssestes.AsID ON dbo.TblCarsData.id = dbo.TblAssestes.CarsDataID LEFT OUTER JOIN"
'MySQL = MySQL & "      dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
'MySQL = MySQL & "      dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN"
'MySQL = MySQL & "      dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
''MySQL = MySQL & "      dbo.TblColor ON dbo.TblCarsData.VColor = dbo.TblColor.Id ON dbo.TblCarsData.VModel = dbo.TblCarModels.Id LEFT OUTER JOIN"
''MySQL = MySQL & "      dbo.EmpGroupDep ON dbo.TblCarsData.LocationID = dbo.EmpGroupDep.GroupID"

MySQL = MySQL & "           SELECT dbo.TblEmpAsest.EmpAsID, dbo.TblCarsData.CarsTypeId, dbo.TblBranchesData.branch_id, dbo.TblCarsData.Emp_id, dbo.TblCarsData.VColor, dbo.TblCarsData.LocationID,"
MySQL = MySQL & "           dbo.TblCarsData.VModel, dbo.TblCarsData.id, dbo.TblCarsData.Branch_NO, dbo.TblCarsData.code, dbo.TblCarsData.Fullcode, dbo.TblCarsData.prifix,"
MySQL = MySQL & "           dbo.TblCarsData.LicenseNO, dbo.TblCarsData.Name, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Model, dbo.TblCarsData.PurchaseDate, dbo.TblCarsData.LastKMCounter,"
MySQL = MySQL & "           dbo.TblCarsData.LicenseExpireDate, dbo.TblCarsData.InsuranceCompanyId, dbo.TblCarsData.InsuranceExpireDate, dbo.TblCarsData.TestExpireDate,"
MySQL = MySQL & "           dbo.TblCarsData.Notes, dbo.TblCarsData.LicenseExpireDateH, dbo.TblCarsData.InsuranceExpireDateH, dbo.TblCarsData.TestExpireDateH, dbo.TblCarsData.fixedAssetid,"
MySQL = MySQL & "           dbo.TblCarsData.VehicleLong, dbo.TblCarsData.EquQty, dbo.TblCarsData.Capacity, dbo.TblCarsData.ContractID, dbo.TblCarsData.EndContractDate,"
MySQL = MySQL & "          dbo.TblCarsData.SetCount, dbo.TblCarsData.Rate, dbo.TblCarsData.EndContractDateH, dbo.TblCarsData.Rep, dbo.TblCarsData.EndAllocationDate,"
MySQL = MySQL & "           dbo.TblCarsData.MaxCap, dbo.TblCarsData.OperatorN, dbo.TblCarsData.EqupName, dbo.TblCarsData.TypeCar, dbo.TblCarsData.Gearno, dbo.TblCarsData.Gearno1,"
MySQL = MySQL & "           dbo.TblCarsData.Machineno, dbo.TblCarsData.Machineno1, dbo.TblCarsData.VType, dbo.TblCarsData.Chesis, dbo.TblCarsData.Total, dbo.TblCarsData.LetterCount,"
MySQL = MySQL & "           dbo.TblColor.name AS ColorName, dbo.TblCarModels.Model AS ModelName, dbo.EmpGroupDep.GroupName, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "           dbo.TBLCarTypes.name AS TypeName, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS Emp_FullCode, dbo.TblEmpAsest.AsID, dbo.TblCarsData.LetterPrice,"
MySQL = MySQL & "           dbo.TblEmpAsestDetails.CarsDataID, dbo.TblEmpAsestDetails.AsestName, dbo.TblEmpAsestDetails.TypeEqup, dbo.TblEmpAsestDetails.FlgCarNotFixed,"
MySQL = MySQL & "           dbo.TblEmpAsestDetails.Stickers, dbo.TblEmpAsestDetails.Guarantee, dbo.TblEmpAsestDetails.CoverKey, dbo.TblEmpAsestDetails.Crane,"
MySQL = MySQL & "           dbo.TblEmpAsestDetails.SpareTyre, dbo.TblEmpAsestDetails.Battery, dbo.TblEmpAsestDetails.Anntena, dbo.TblEmpAsestDetails.Recorder,"
MySQL = MySQL & "           dbo.TblEmpAsestDetails.SunScreens, dbo.TblEmpAsestDetails.Pedals, dbo.TblEmpAsestDetails.InnerLights, dbo.TblEmpAsestDetails.driverMirror,"
MySQL = MySQL & "           dbo.TblEmpAsestDetails.sideMirror, dbo.TblEmpAsestDetails.cleaner, dbo.TblEmpAsestDetails.FormOrignal, dbo.TblEmpAsestDetails.authorizeLicense,"
MySQL = MySQL & "           dbo.TblEmpAsestDetails.authorizeExamination, dbo.TblEmpAsestDetails.subcar1, dbo.TblEmpAsestDetails.subcar2, dbo.TblEmpAsestDetails.subcar3,"
MySQL = MySQL & "           dbo.TblEmpAsestDetails.subcar4, dbo.TblEmpAsestDetails.subcar5, dbo.TblEmpAsestDetails.subcar6, dbo.TblEmpAsestDetails.subcar7, dbo.TblEmpAsestDetails.subcar8,"
MySQL = MySQL & "           dbo.TblEmpAsestDetails.subcar9, dbo.TblEmpAsestDetails.subcar10, dbo.TblEmpAsestDetails.subcar11, dbo.TblEmpAsestDetails.subcar12,"
MySQL = MySQL & "           dbo.TblEmpAsestDetails.subcar13, dbo.TblEmpAsestDetails.subcar14, dbo.TblEmpAsest.ISCar, dbo.TblEmpAsest.EmpAsestID, FromEmp.Emp_Name AS FromEmpName,"
MySQL = MySQL & "           FromEmp.Fullcode AS FromEmpCode, dbo.TblEmpAsest.ToEmId, ToEmp.Emp_Name AS ToEmpName, ToEmp.Fullcode AS ToEmpCode ,dbo.TblEmpAsest.DriveDate "
 MySQL = MySQL & "          FROM     dbo.TblBranchesData RIGHT OUTER JOIN"
MySQL = MySQL & "           dbo.TblEmpAsestDetails RIGHT OUTER JOIN"
MySQL = MySQL & "           dbo.TblEmpAsest LEFT OUTER JOIN"
MySQL = MySQL & "           dbo.TblEmployee AS ToEmp ON dbo.TblEmpAsest.ToEmId = ToEmp.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "           dbo.TblEmployee AS FromEmp ON dbo.TblEmpAsest.EmpAsestID = FromEmp.Emp_ID ON dbo.TblEmpAsestDetails.IDAseset = dbo.TblEmpAsest.EmpAsID LEFT OUTER JOIN"
MySQL = MySQL & "           dbo.TblAssestes ON dbo.TblEmpAsestDetails.AsID = dbo.TblAssestes.AsID LEFT OUTER JOIN"
MySQL = MySQL & "           dbo.TblCarsData ON dbo.TblAssestes.CarsDataID = dbo.TblCarsData.id LEFT OUTER JOIN"
MySQL = MySQL & "           dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "           dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id ON dbo.TblBranchesData.branch_id = dbo.TblCarsData.Branch_NO LEFT OUTER JOIN"
MySQL = MySQL & "           dbo.TblColor ON dbo.TblCarsData.VColor = dbo.TblColor.Id LEFT OUTER JOIN"
MySQL = MySQL & "           dbo.TblCarModels ON dbo.TblCarsData.VModel = dbo.TblCarModels.Id LEFT OUTER JOIN"
MySQL = MySQL & "           dbo.EmpGroupDep ON dbo.TblCarsData.LocationID = dbo.EmpGroupDep.GroupID"
        


MySQL = MySQL & "  where    TblEmpAsest.EmpAsID  =" & val(XPTxtID.Text)
MySQL = MySQL & " order by  TblEmpAsest.EmpAsID "

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_CarsAsset.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_CarsAsset.rpt"
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
      If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
       Else
       Msg = "No Data"
       End If
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
     '         Dim xLogo As CRAXDRT.OLEObject
     '    If Dir(App.path & "\" & SystemOptions.ImagesPath & "\" & val(XPTxtID.text) & ".JPG") <> "" Then
     '       Set xLogo = xReport.Areas(1).Sections(1).AddPictureObject(App.path & "\" & SystemOptions.ImagesPath & "\" & val(XPTxtID.text) & ".JPG", 250, 2400)
     '       xLogo.Width = 6100
     '       xLogo.Height = 1200
     '       xLogo.backcolor = vbWhite
     ''       xLogo.BorderColor = 255
     '       xLogo.CloseAtPageBreak = True
     '     End If
    ''///////
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
