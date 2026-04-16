VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbaListView6.ocx"
Begin VB.Form frmsalebill333 
   BorderStyle     =   0  'None
   Caption         =   "›« Ê—… «·„»Ì⁄« "
   ClientHeight    =   11520
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   27915
   HelpContextID   =   160
   Icon            =   "FrmSaleBill333.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   27915
   ShowInTaskbar   =   0   'False
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
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   11520
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   27915
      _cx             =   49239
      _cy             =   20320
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
      AutoSizeChildren=   8
      BorderWidth     =   1
      ChildSpacing    =   1
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
      GridRows        =   5
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmSaleBill333.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Height          =   11490
         Left            =   15
         RightToLeft     =   -1  'True
         TabIndex        =   169
         Top             =   15
         Width           =   27885
         Begin VB.TextBox txtcurrenttableidMain 
            Height          =   285
            Left            =   18480
            TabIndex        =   381
            Top             =   10860
            Width           =   615
         End
         Begin VB.Frame FramePay 
            BackColor       =   &H00E0E0E0&
            Caption         =   "«·„»·€ «·„œ›Ê⁄"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   6735
            Left            =   8610
            RightToLeft     =   -1  'True
            TabIndex        =   339
            Top             =   540
            Visible         =   0   'False
            Width           =   11175
            Begin VB.Frame Frame13 
               BackColor       =   &H00FFFFFF&
               Height          =   5055
               Left            =   120
               TabIndex        =   359
               Top             =   480
               Width           =   5535
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   0
                  Left            =   4320
                  TabIndex        =   360
                  Top             =   3970
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
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
                  ButtonImage     =   "FrmSaleBill333.frx":03F3
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   1
                  Left            =   2160
                  TabIndex        =   361
                  Top             =   3000
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmSaleBill333.frx":0BB3
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   2
                  Left            =   3240
                  TabIndex        =   362
                  Top             =   3000
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmSaleBill333.frx":11B5
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   3
                  Left            =   4320
                  TabIndex        =   363
                  Top             =   3000
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmSaleBill333.frx":199C
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   4
                  Left            =   2160
                  TabIndex        =   364
                  Top             =   2040
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmSaleBill333.frx":21B1
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   5
                  Left            =   3240
                  TabIndex        =   365
                  Top             =   2040
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmSaleBill333.frx":293C
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   6
                  Left            =   4320
                  TabIndex        =   366
                  Top             =   2040
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmSaleBill333.frx":30FB
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   7
                  Left            =   2160
                  TabIndex        =   367
                  Top             =   1080
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmSaleBill333.frx":3895
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   8
                  Left            =   3240
                  TabIndex        =   368
                  Top             =   1080
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmSaleBill333.frx":3F98
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   9
                  Left            =   4320
                  TabIndex        =   369
                  Top             =   1080
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmSaleBill333.frx":47B3
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   10
                  Left            =   3240
                  TabIndex        =   370
                  Top             =   3970
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmSaleBill333.frx":4F42
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   11
                  Left            =   2160
                  TabIndex        =   371
                  Top             =   3970
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmSaleBill333.frx":5A89
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   12
                  Left            =   120
                  TabIndex        =   372
                  Top             =   1080
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmSaleBill333.frx":5F7B
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   13
                  Left            =   1200
                  TabIndex        =   373
                  Top             =   1080
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmSaleBill333.frx":67E2
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   2895
                  Index           =   14
                  Left            =   120
                  TabIndex        =   374
                  Top             =   2040
                  Width           =   2055
                  _ExtentX        =   3625
                  _ExtentY        =   5106
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmSaleBill333.frx":6EF3
                  ButtonImageDisabled=   "FrmSaleBill333.frx":82A1
                  ColorButton     =   16777215
               End
               Begin VB.Image Image13 
                  Height          =   1035
                  Left            =   120
                  Picture         =   "FrmSaleBill333.frx":863C
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   5295
               End
               Begin VB.Label LBLPayVal 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   24
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   615
                  Left            =   1200
                  TabIndex        =   375
                  Top             =   240
                  Width           =   3375
               End
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "1500"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   9
               Left            =   5760
               TabIndex        =   358
               Top             =   5280
               Width           =   1215
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "2000"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   8
               Left            =   5760
               TabIndex        =   357
               Top             =   5880
               Width           =   1215
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Height          =   2055
               Left            =   7080
               RightToLeft     =   -1  'True
               TabIndex        =   350
               Top             =   4440
               Width           =   3840
               Begin VB.TextBox TxtNetValue 
                  Alignment       =   2  'Center
                  BackColor       =   &H00000000&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   18
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FF00&
                  Height          =   600
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   353
                  Top             =   240
                  Width           =   2460
               End
               Begin VB.TextBox TxtPayedValue 
                  Alignment       =   2  'Center
                  BackColor       =   &H00000000&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   18
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FF00&
                  Height          =   555
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   352
                  Top             =   840
                  Width           =   2445
               End
               Begin VB.TextBox TxtRemainValue 
                  Alignment       =   2  'Center
                  BackColor       =   &H00000000&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   18
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FF00&
                  Height          =   555
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   351
                  Top             =   1320
                  Width           =   2445
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«Ã„«·Ì"
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
                  Index           =   58
                  Left            =   2640
                  TabIndex        =   356
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„œ›Ê⁄"
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
                  Index           =   59
                  Left            =   2640
                  TabIndex        =   355
                  Top             =   840
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„ »ﬁÌ"
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
                  Index           =   60
                  Left            =   2640
                  TabIndex        =   354
                  Top             =   1440
                  Width           =   855
               End
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "1000"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   7
               Left            =   5760
               TabIndex        =   349
               Top             =   4680
               Width           =   1215
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "500"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   6
               Left            =   5760
               TabIndex        =   348
               Top             =   4080
               Width           =   1215
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "200"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   5
               Left            =   5760
               TabIndex        =   347
               Top             =   3480
               Width           =   1215
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "100"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   4
               Left            =   5760
               TabIndex        =   346
               Top             =   2880
               Width           =   1215
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "50"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   3
               Left            =   5760
               TabIndex        =   345
               Top             =   2280
               Width           =   1215
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "20"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   2
               Left            =   5760
               TabIndex        =   344
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "10"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   1
               Left            =   5760
               TabIndex        =   343
               Top             =   1080
               Width           =   1215
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   0
               Left            =   5760
               TabIndex        =   342
               Top             =   480
               Width           =   1215
            End
            Begin ImpulseButton.ISButton CMDPAy 
               Height          =   1215
               Left            =   240
               TabIndex        =   340
               Top             =   5430
               Width           =   5295
               _ExtentX        =   9340
               _ExtentY        =   2143
               Caption         =   "”œ«œ"
               ForeColor       =   16777215
               FontSize        =   24
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmSaleBill333.frx":89F2
               ColorHoverText  =   16777215
               ColorToggledText=   16777215
               ColorToggledHoverText=   16777215
               AlignmentIgnoreImage=   -1  'True
            End
            Begin MSDataListLib.DataCombo DCboUserName 
               Height          =   315
               Left            =   4200
               TabIndex        =   341
               Top             =   -960
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               Text            =   "DataCombo2"
            End
            Begin VSFlex8UCtl.VSFlexGrid Grid 
               Height          =   3885
               Left            =   7080
               TabIndex        =   376
               Top             =   600
               Width           =   3885
               _cx             =   6853
               _cy             =   6853
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
               BackColor       =   -2147483640
               ForeColor       =   65280
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483641
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483640
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
               Rows            =   6
               Cols            =   11
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   650
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSaleBill333.frx":8F6C
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
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               Height          =   315
               Left            =   0
               TabIndex        =   377
               Top             =   -600
               Visible         =   0   'False
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lblexit 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   270
               Index           =   90
               Left            =   10200
               TabIndex        =   379
               Top             =   240
               Width           =   570
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "X"
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
               Left            =   10800
               TabIndex        =   378
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.PictureBox Picture1 
            Height          =   1815
            Left            =   15
            ScaleHeight     =   1815
            ScaleWidth      =   15
            TabIndex        =   338
            Top             =   15
            Width           =   15
         End
         Begin VB.CommandButton Command7 
            Caption         =   "› Õ «·Œ’„"
            Height          =   540
            Left            =   10230
            TabIndex        =   337
            Top             =   10320
            Width           =   1365
         End
         Begin VB.Frame frmaeDiscount 
            BackColor       =   &H00E2E9E9&
            Height          =   615
            Left            =   120
            TabIndex        =   331
            Top             =   8760
            Visible         =   0   'False
            Width           =   5775
            Begin VB.ComboBox XPCboDiscountType 
               Height          =   315
               Left            =   3000
               Style           =   2  'Dropdown List
               TabIndex        =   333
               Top             =   120
               Width           =   1470
            End
            Begin VB.TextBox XPTxtDiscountVal 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   0
               TabIndex        =   332
               Top             =   120
               Width           =   1470
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "‰Ê⁄ «·Œ’„"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000001&
               Height          =   390
               Index           =   10
               Left            =   4560
               TabIndex        =   336
               Top             =   120
               Width           =   1170
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„…"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000001&
               Height          =   390
               Index           =   8
               Left            =   1605
               TabIndex        =   334
               Top             =   120
               Width           =   1170
            End
         End
         Begin VB.CommandButton Command6 
            Caption         =   "«·„‘—›"
            Height          =   540
            Left            =   11670
            TabIndex        =   330
            Top             =   10320
            Width           =   1365
         End
         Begin VB.CommandButton Command5 
            Caption         =   "«·ÿ«Ê·« "
            Height          =   540
            Left            =   13050
            TabIndex        =   329
            Top             =   10320
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   600
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   319
            Top             =   11880
            Visible         =   0   'False
            Width           =   2460
         End
         Begin VB.TextBox TxtValueAdded 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   316
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CheckBox ChecVAT 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   " ÕœÌœ «·ﬂ·"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5280
            RightToLeft     =   -1  'True
            TabIndex        =   315
            Top             =   10920
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.TextBox txtPointsOpr 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   3120
            TabIndex        =   301
            Top             =   8400
            Width           =   1470
         End
         Begin VB.TextBox TxtTotalPoints 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   120
            TabIndex        =   300
            Top             =   8400
            Width           =   1470
         End
         Begin VB.TextBox Txtcart 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   3120
            TabIndex        =   294
            Top             =   8040
            Width           =   1470
         End
         Begin VB.TextBox CashCustomerName 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   120
            TabIndex        =   293
            Top             =   8040
            Width           =   1470
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   17040
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   292
            Top             =   10320
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Timer Timer5 
            Interval        =   3000
            Left            =   6960
            Top             =   7200
         End
         Begin MSDataListLib.DataCombo DCPaymentNet 
            Height          =   315
            Left            =   20280
            TabIndex        =   235
            Top             =   1200
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.ComboBox CBOPrinter 
            Height          =   315
            Left            =   19800
            TabIndex        =   279
            Text            =   "Combo1"
            Top             =   0
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox dbname 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   23280
            RightToLeft     =   -1  'True
            TabIndex        =   277
            Top             =   12480
            Width           =   2865
         End
         Begin vbalIml6.vbalImageList GrouplImageList 
            Left            =   8760
            Top             =   120
            _ExtentX        =   953
            _ExtentY        =   953
            IconSizeX       =   32
            IconSizeY       =   32
            ColourDepth     =   32
         End
         Begin VB.Timer Timer1 
            Interval        =   250
            Left            =   17760
            Top             =   4200
         End
         Begin VB.Frame Frame9 
            Caption         =   "Frame9"
            Height          =   2055
            Left            =   -4440
            TabIndex        =   224
            Top             =   8520
            Visible         =   0   'False
            Width           =   4215
            Begin VB.CommandButton Command4 
               Caption         =   "Command4"
               Height          =   195
               Left            =   960
               TabIndex        =   234
               Top             =   120
               Width           =   135
            End
            Begin VB.CheckBox chkPayed 
               Caption         =   "„œ›Ê⁄"
               Height          =   255
               Left            =   1680
               TabIndex        =   232
               Top             =   960
               Width           =   975
            End
            Begin VB.ComboBox CboPOSBillType 
               Height          =   315
               ItemData        =   "FrmSaleBill333.frx":912B
               Left            =   2280
               List            =   "FrmSaleBill333.frx":912D
               Style           =   2  'Dropdown List
               TabIndex        =   225
               Top             =   195
               Width           =   1635
            End
            Begin VB.Label LblStableID 
               Caption         =   "-1"
               Height          =   375
               Left            =   3000
               TabIndex        =   233
               Top             =   720
               Width           =   855
            End
            Begin VB.Label LblSessionID 
               Height          =   375
               Left            =   480
               TabIndex        =   226
               Top             =   1200
               Width           =   2055
            End
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   21945
            Style           =   2  'Dropdown List
            TabIndex        =   228
            Top             =   8595
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1875
            TabIndex        =   227
            Top             =   4800
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.PictureBox imgLarge 
            BackColor       =   &H80000005&
            Height          =   480
            Left            =   -1920
            ScaleHeight     =   420
            ScaleWidth      =   1875
            TabIndex        =   211
            Top             =   120
            Visible         =   0   'False
            Width           =   1935
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
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   2160
            TabIndex        =   210
            Top             =   1800
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Frame Frame6 
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
            TabIndex        =   188
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
               TabIndex        =   189
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
                  TabIndex        =   207
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
                  TabIndex        =   206
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
                  TabIndex        =   205
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
                  TabIndex        =   204
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
                  TabIndex        =   203
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
                  TabIndex        =   202
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
                  TabIndex        =   201
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
                  TabIndex        =   200
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
                  TabIndex        =   199
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
                  TabIndex        =   198
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
                  TabIndex        =   197
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
                  TabIndex        =   196
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
                  TabIndex        =   195
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
                  TabIndex        =   194
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
                  TabIndex        =   193
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
                  TabIndex        =   192
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
                  TabIndex        =   191
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
                  TabIndex        =   190
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
                  TabIndex        =   209
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
                  TabIndex        =   208
                  Top             =   480
                  Width           =   915
               End
            End
         End
         Begin VB.Frame Frame7 
            Height          =   4695
            Left            =   22800
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   2880
            Visible         =   0   'False
            Width           =   4815
            Begin VB.Timer tmr 
               Interval        =   1000
               Left            =   3960
               Top             =   1440
            End
            Begin VB.Timer Timer4 
               Interval        =   1000
               Left            =   840
               Top             =   1320
            End
            Begin VB.Timer Timer3 
               Interval        =   100
               Left            =   4080
               Top             =   600
            End
            Begin VB.Timer Timer2 
               Enabled         =   0   'False
               Interval        =   100
               Left            =   -1320
               Top             =   480
            End
            Begin VB.Label LblSowPrice 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   735
               Index           =   1
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   238
               Top             =   840
               Width           =   4815
            End
            Begin VB.Label lblqty 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   186
               Top             =   360
               Width           =   4725
            End
            Begin VB.Label lBLclr 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   1455
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label LBLdOT 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   735
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Top             =   3960
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   9
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   178
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   8
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   179
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   7
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   180
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   6
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   181
               Top             =   2640
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   5
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   174
               Top             =   2640
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   4
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   175
               Top             =   2640
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   3
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   182
               Top             =   3360
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   2
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   183
               Top             =   3360
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   1
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   184
               Top             =   3360
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   0
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   185
               Top             =   3360
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.Image Image1 
               Height          =   4605
               Left            =   -960
               Picture         =   "FrmSaleBill333.frx":912F
               Stretch         =   -1  'True
               Top             =   0
               Visible         =   0   'False
               Width           =   4845
            End
            Begin VB.Label Label9 
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
               TabIndex        =   187
               Top             =   360
               Width           =   1965
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
               TabIndex        =   173
               Top             =   4080
               Width           =   3045
            End
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   20880
            TabIndex        =   229
            Top             =   2880
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   0
            Left            =   15210
            TabIndex        =   239
            Top             =   10920
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   953
            ButtonPositionImage=   1
            Caption         =   "ÃœÌœ"
            BackColor       =   14737632
            ForeColor       =   16711680
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14737632
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledText=   16711680
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   1
            Left            =   13770
            TabIndex        =   240
            Top             =   10950
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   953
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ·"
            BackColor       =   14737632
            ForeColor       =   16711680
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14737632
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   16711680
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   2
            Left            =   12330
            TabIndex        =   241
            Top             =   10920
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   953
            ButtonPositionImage=   1
            Caption         =   "œ›⁄"
            BackColor       =   14737632
            ForeColor       =   16711680
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14737632
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   16711680
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   3
            Left            =   10890
            TabIndex        =   242
            Top             =   10920
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   953
            ButtonPositionImage=   1
            Caption         =   " —«Ã⁄"
            BackColor       =   14737632
            ForeColor       =   16711680
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14737632
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   16711680
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   4
            Left            =   21840
            TabIndex        =   243
            Top             =   2040
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BackColor       =   0
            ForeColor       =   65280
            FontSize        =   12
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   0
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   65280
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   7
            Left            =   8490
            TabIndex        =   244
            Top             =   10920
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   953
            ButtonPositionImage=   1
            Caption         =   "œ›⁄"
            BackColor       =   14737632
            ForeColor       =   16711680
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14737632
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   16711680
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   540
            Index           =   1
            Left            =   0
            TabIndex        =   251
            TabStop         =   0   'False
            Top             =   -480
            Visible         =   0   'False
            Width           =   20280
            _cx             =   35772
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
               Height          =   540
               Index           =   5
               Left            =   6855
               TabIndex        =   252
               Top             =   0
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   953
               ButtonStyle     =   1
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
               ColorShadow     =   4210752
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               ColorTextShadow =   4210752
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   540
               Index           =   6
               Left            =   30
               TabIndex        =   253
               Top             =   0
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   953
               ButtonStyle     =   1
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
               ColorShadow     =   4210752
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               ColorTextShadow =   4210752
            End
            Begin ImpulseButton.ISButton CmdHelp 
               Height          =   540
               Left            =   2295
               TabIndex        =   254
               Top             =   0
               Width           =   2070
               _ExtentX        =   3651
               _ExtentY        =   953
               ButtonStyle     =   1
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
               ColorShadow     =   4210752
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               ColorTextShadow =   4210752
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   435
               Index           =   3
               Left            =   7080
               TabIndex        =   255
               TabStop         =   0   'False
               Top             =   -600
               Width           =   20280
               _cx             =   35772
               _cy             =   767
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
               AutoSizeChildren=   7
               BorderWidth     =   0
               ChildSpacing    =   0
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
               Begin VB.TextBox XPTxtSum 
                  Alignment       =   2  'Center
                  BackColor       =   &H000000FF&
                  Height          =   375
                  Left            =   17385
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   256
                  TabStop         =   0   'False
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   285
               End
               Begin VB.Label LblFinal 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Left            =   2685
                  TabIndex        =   268
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1830
               End
               Begin VB.Label lblInstComm 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Left            =   7005
                  TabIndex        =   267
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   585
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
                  Height          =   315
                  Index           =   63
                  Left            =   3600
                  TabIndex        =   266
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   435
               End
               Begin VB.Label LblTotalQty 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Left            =   3765
                  TabIndex        =   265
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   675
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·≈Ã„«·Ï"
                  Height          =   285
                  Index           =   3
                  Left            =   20430
                  RightToLeft     =   -1  'True
                  TabIndex        =   264
                  Top             =   75
                  Width           =   810
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "/"
                  Height          =   285
                  Index           =   0
                  Left            =   1020
                  RightToLeft     =   -1  'True
                  TabIndex        =   263
                  Top             =   75
                  Width           =   165
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·”Ã·"
                  Height          =   285
                  Index           =   2
                  Left            =   1860
                  RightToLeft     =   -1  'True
                  TabIndex        =   262
                  Top             =   75
                  Width           =   1065
               End
               Begin VB.Label XPTxtCurrent 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Height          =   285
                  Left            =   1365
                  RightToLeft     =   -1  'True
                  TabIndex        =   261
                  Top             =   75
                  Width           =   270
               End
               Begin VB.Label XPTxtCount 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Height          =   285
                  Left            =   330
                  RightToLeft     =   -1  'True
                  TabIndex        =   260
                  Top             =   75
                  Width           =   405
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·’«›Ì"
                  Height          =   285
                  Index           =   49
                  Left            =   8850
                  RightToLeft     =   -1  'True
                  TabIndex        =   259
                  Top             =   75
                  Width           =   1020
               End
               Begin VB.Label LblTotalAll 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   """#,###.##"""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Left            =   17145
                  RightToLeft     =   -1  'True
                  TabIndex        =   258
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   795
               End
               Begin VB.Label LblDiscountsTotal 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Left            =   13455
                  RightToLeft     =   -1  'True
                  TabIndex        =   257
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   1935
               End
            End
            Begin VB.Image Image10 
               Height          =   135
               Left            =   120
               Top             =   480
               Width           =   4935
            End
         End
         Begin MSComctlLib.Toolbar TBar 
            Height          =   630
            Left            =   0
            TabIndex        =   215
            Top             =   6720
            Width           =   5970
            _ExtentX        =   10530
            _ExtentY        =   1111
            ButtonWidth     =   609
            ButtonHeight    =   1005
            Appearance      =   1
            _Version        =   393216
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   8
            Left            =   7410
            TabIndex        =   274
            Top             =   10920
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   953
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14737632
            ForeColor       =   16711680
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14737632
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   16711680
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin VSFlex8UCtl.VSFlexGrid FG 
            Height          =   4710
            Left            =   -90
            TabIndex        =   275
            Top             =   2070
            Width           =   6585
            _cx             =   11615
            _cy             =   8308
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   14737632
            ForeColorFixed  =   0
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
            Cols            =   26
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   800
            RowHeightMax    =   800
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmSaleBill333.frx":111E6
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   9
            Left            =   4560
            TabIndex        =   278
            Top             =   13080
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â"
            BackColor       =   14737632
            ForeColor       =   0
            FontSize        =   12
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14737632
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton CMDADDQty 
            Height          =   840
            Left            =   6720
            TabIndex        =   280
            Top             =   4680
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1482
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmSaleBill333.frx":1164D
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   840
            Left            =   6720
            TabIndex        =   281
            Top             =   5520
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1482
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmSaleBill333.frx":1217B
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   960
            Left            =   6720
            TabIndex        =   287
            Top             =   6360
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1693
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmSaleBill333.frx":12CC4
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton btnMove 
            Height          =   840
            Index           =   0
            Left            =   6720
            TabIndex        =   288
            Top             =   1320
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1482
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmSaleBill333.frx":137BB
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton btnMove 
            Height          =   840
            Index           =   1
            Left            =   6720
            TabIndex        =   289
            Top             =   2160
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1482
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmSaleBill333.frx":141F5
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton btnMove 
            Height          =   840
            Index           =   2
            Left            =   6720
            TabIndex        =   290
            Top             =   3000
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1482
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmSaleBill333.frx":14BE9
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton btnMove 
            Height          =   840
            Index           =   3
            Left            =   6720
            TabIndex        =   291
            Top             =   3840
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1482
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmSaleBill333.frx":1559C
            ColorButton     =   16777215
         End
         Begin ImpulseButton.ISButton SearchCashCustomer 
            Height          =   345
            Index           =   0
            Left            =   0
            TabIndex        =   295
            TabStop         =   0   'False
            Top             =   8400
            Visible         =   0   'False
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            BackStyle       =   0
            ButtonImage     =   "FrmSaleBill333.frx":15F7B
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8UCtl.VSFlexGrid VatGrid 
            Height          =   1725
            Left            =   -840
            TabIndex        =   314
            Tag             =   "1"
            Top             =   11760
            Visible         =   0   'False
            Width           =   9855
            _cx             =   17383
            _cy             =   3043
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
            FormatString    =   $"FrmSaleBill333.frx":16378
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
         Begin MSDataListLib.DataCombo DcboEmp1 
            Height          =   315
            Left            =   2760
            TabIndex        =   320
            Top             =   10560
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   420
            Index           =   12
            Left            =   480
            TabIndex        =   324
            Top             =   7290
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   741
            ButtonPositionImage=   1
            Caption         =   "„ÿ»Œ"
            BackColor       =   14737632
            ForeColor       =   16711680
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14737632
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   16711680
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   13
            Left            =   9600
            TabIndex        =   325
            Top             =   10920
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   953
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â"
            BackColor       =   14737632
            ForeColor       =   16711680
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14737632
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   16711680
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin VB.Frame FrameAdmi 
            BackColor       =   &H000000FF&
            Caption         =   "AdminLogin"
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
            Height          =   2775
            Left            =   8040
            TabIndex        =   326
            Top             =   1920
            Visible         =   0   'False
            Width           =   5775
            Begin VB.CommandButton CMDAdminLogin 
               Caption         =   "OK"
               Height          =   855
               Left            =   3240
               TabIndex        =   328
               Top             =   1800
               Width           =   2055
            End
            Begin VB.TextBox TxtAdminLogin 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               IMEMode         =   3  'DISABLE
               Left            =   1080
               PasswordChar    =   "*"
               TabIndex        =   327
               Top             =   720
               Width           =   4215
            End
            Begin VB.Image Image11 
               Height          =   855
               Left            =   360
               Picture         =   "FrmSaleBill333.frx":1648B
               Stretch         =   -1  'True
               Top             =   720
               Width           =   615
            End
         End
         Begin vbalListViewLib6.vbalListViewCtl lvwMain 
            Height          =   4395
            Left            =   7620
            TabIndex        =   212
            Top             =   120
            Width           =   19995
            _ExtentX        =   35269
            _ExtentY        =   7752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14737632
            MultiSelect     =   -1  'True
            LabelEdit       =   0   'False
            AutoArrange     =   0   'False
            HeaderButtons   =   0   'False
            HeaderTrackSelect=   0   'False
            HideSelection   =   0   'False
            InfoTips        =   0   'False
         End
         Begin vbalListViewLib6.vbalListViewCtl lvwItems 
            Height          =   5655
            Left            =   7605
            TabIndex        =   213
            Top             =   4950
            Width           =   19995
            _ExtentX        =   35269
            _ExtentY        =   9975
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14737632
            MultiSelect     =   -1  'True
            LabelEdit       =   0   'False
            AutoArrange     =   0   'False
            HeaderButtons   =   0   'False
            HeaderTrackSelect=   0   'False
            HideSelection   =   0   'False
            InfoTips        =   0   'False
         End
         Begin VB.Frame Frame8 
            Caption         =   "Frame4"
            Height          =   2415
            Left            =   -2010
            RightToLeft     =   -1  'True
            TabIndex        =   170
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
            Begin vbalIml6.vbalImageList ilsIcons32 
               Left            =   0
               Top             =   360
               _ExtentX        =   953
               _ExtentY        =   953
               IconSizeX       =   32
               IconSizeY       =   32
               ColourDepth     =   24
               Size            =   4412
               Images          =   "FrmSaleBill333.frx":16B68
               Version         =   131072
               KeyCount        =   1
               Keys            =   ""
            End
            Begin MSComctlLib.ImageList ImageList1 
               Left            =   960
               Top             =   1560
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               MaskColor       =   12632256
               _Version        =   393216
            End
            Begin vbalIml6.vbalImageList ilsIcons16 
               Left            =   8280
               Top             =   0
               _ExtentX        =   953
               _ExtentY        =   953
               IconSizeX       =   48
               IconSizeY       =   48
               ColourDepth     =   24
            End
            Begin MSComctlLib.ImageList ImageListItems 
               Left            =   315
               Top             =   1530
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   120
               ImageHeight     =   120
               MaskColor       =   12632256
               _Version        =   393216
            End
            Begin VB.Label lblStatus 
               Alignment       =   1  'Right Justify
               Caption         =   "Label10"
               Height          =   495
               Left            =   840
               RightToLeft     =   -1  'True
               TabIndex        =   171
               Top             =   960
               Width           =   135
            End
         End
         Begin vbalListViewLib6.vbalListViewCtl lvwTables 
            Height          =   5655
            Left            =   7530
            TabIndex        =   214
            Top             =   4560
            Width           =   20190
            _ExtentX        =   35613
            _ExtentY        =   9975
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14737632
            View            =   4
            MultiSelect     =   -1  'True
            LabelEdit       =   0   'False
            AutoArrange     =   0   'False
            HeaderButtons   =   0   'False
            HeaderTrackSelect=   0   'False
            HideSelection   =   0   'False
            InfoTips        =   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   11
            Left            =   16620
            TabIndex        =   380
            Top             =   10920
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   953
            ButtonPositionImage=   1
            Caption         =   "„ÿ»Œ"
            BackColor       =   14737632
            ForeColor       =   16711680
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14737632
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   16711680
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin VB.Image Image16 
            Height          =   1665
            Left            =   0
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   1860
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰ﬁ«ÿ «·⁄„·Ì…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   390
            Index           =   89
            Left            =   4560
            TabIndex        =   335
            Top             =   8520
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·ﬂ«» ‰"
            Height          =   285
            Index           =   4
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   323
            Top             =   10680
            Width           =   810
         End
         Begin VB.Label LblTable1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Take Out"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   615
            Left            =   -360
            TabIndex        =   322
            Top             =   10680
            Width           =   4095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·ÿ«Ê·… «·„Õœœ…"
            Height          =   285
            Index           =   1
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   321
            Top             =   10560
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁ. „÷«›…"
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
            Height          =   555
            Index           =   16
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   318
            Top             =   9240
            Width           =   1125
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   555
            Left            =   1920
            TabIndex        =   317
            Top             =   9600
            Width           =   1125
         End
         Begin VB.Label Label17 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   6360
            TabIndex        =   270
            Top             =   12120
            Visible         =   0   'False
            Width           =   14175
         End
         Begin VB.Label LBLGross 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """#,###.##"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   0
            TabIndex        =   313
            Top             =   12120
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   312
            Top             =   960
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   311
            Top             =   960
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   10
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   310
            Top             =   960
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   309
            Top             =   960
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "F7"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   8
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   308
            Top             =   10920
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "F6"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   7
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   307
            Top             =   10920
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "F9"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   6
            Left            =   9960
            RightToLeft     =   -1  'True
            TabIndex        =   306
            Top             =   10920
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "F10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   5
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   305
            Top             =   10920
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "F11"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   3
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   304
            Top             =   10920
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "F12"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   2
            Left            =   14160
            RightToLeft     =   -1  'True
            TabIndex        =   303
            Top             =   10920
            Width           =   810
         End
         Begin VB.Label LblTable 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   6360
            TabIndex        =   302
            Top             =   10320
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—’Ìœ«·‰ﬁ«ÿ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   390
            Index           =   86
            Left            =   1605
            TabIndex        =   299
            Top             =   8520
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·ﬂ«— "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   390
            Index           =   1
            Left            =   4725
            TabIndex        =   298
            Top             =   8160
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬂ«— "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   270
            Index           =   87
            Left            =   27000
            TabIndex        =   297
            Top             =   7920
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„Ì·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   390
            Index           =   88
            Left            =   1725
            TabIndex        =   296
            Top             =   8160
            Width           =   1170
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H00E0E0E0&
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   5
            FillColor       =   &H00FFFFFF&
            Height          =   7245
            Left            =   0
            Top             =   120
            Width           =   6585
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00E0E0E0&
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   5
            FillColor       =   &H00FFFFFF&
            Height          =   4005
            Left            =   0
            Top             =   7440
            Width           =   6090
         End
         Begin VB.Image Image14 
            Height          =   945
            Left            =   6360
            Stretch         =   -1  'True
            Top             =   360
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H00E0E0E0&
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   5
            FillColor       =   &H00FFFFFF&
            Height          =   1125
            Left            =   6240
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Image Image4 
            Height          =   1275
            Left            =   1875
            Picture         =   "FrmSaleBill333.frx":17CC4
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1395
         End
         Begin VB.Image Image7 
            Height          =   1155
            Left            =   5505
            Picture         =   "FrmSaleBill333.frx":18BF2
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1275
         End
         Begin VB.Image Image8 
            Height          =   1275
            Left            =   3720
            Picture         =   "FrmSaleBill333.frx":19860
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1515
         End
         Begin VB.Image Image12 
            Height          =   1275
            Left            =   120
            Picture         =   "FrmSaleBill333.frx":1A781
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1395
         End
         Begin VB.Label xxx 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00FF0000&
            Height          =   675
            Index           =   0
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   286
            Top             =   11640
            Width           =   2925
         End
         Begin VB.Label LblSowPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   360
            Index           =   0
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   285
            Top             =   7680
            Width           =   1470
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "”⁄—"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   555
            Index           =   85
            Left            =   2040
            TabIndex        =   284
            Top             =   7680
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬂ„Ì…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   390
            Index           =   84
            Left            =   4725
            TabIndex        =   283
            Top             =   7680
            Width           =   1170
         End
         Begin VB.Label lblShowQty2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   360
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   282
            Top             =   7680
            Width           =   1470
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Ê«ﬁ⁄"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   21600
            RightToLeft     =   -1  'True
            TabIndex        =   220
            Top             =   4200
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Image Image6 
            Height          =   435
            Left            =   23640
            Stretch         =   -1  'True
            Top             =   4080
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "«·À·«À«¡"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   18
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   675
            Index           =   83
            Left            =   3480
            TabIndex        =   276
            Top             =   10080
            Width           =   2655
         End
         Begin VB.Image Image9 
            Height          =   1695
            Left            =   24120
            Stretch         =   -1  'True
            Top             =   240
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "12:30 AM"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   18
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   675
            Index           =   82
            Left            =   19440
            TabIndex        =   273
            Top             =   8880
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "12:30 AM"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   18
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   675
            Index           =   81
            Left            =   240
            TabIndex        =   272
            Top             =   10080
            Width           =   2655
         End
         Begin VB.Label lblLabel1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   19080
            TabIndex        =   271
            Top             =   10920
            Width           =   1095
         End
         Begin VB.Label lblLabel1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   0
            Left            =   19800
            TabIndex        =   269
            Top             =   10440
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·’«›Ì"
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
            Height          =   555
            Index           =   71
            Left            =   240
            TabIndex        =   250
            Top             =   9240
            Width           =   1125
         End
         Begin VB.Label LblTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   555
            Left            =   120
            TabIndex        =   249
            Top             =   9600
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Œ’Ê„« "
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
            Height          =   555
            Index           =   50
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   248
            Top             =   9240
            Width           =   1125
         End
         Begin VB.Label LblDiscountsTotalView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   555
            Left            =   3120
            TabIndex        =   247
            Top             =   9600
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·«Ã„«·Ì"
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
            Height          =   555
            Index           =   69
            Left            =   4440
            TabIndex        =   246
            Top             =   9240
            Width           =   1125
         End
         Begin VB.Label LblTotalAllView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """#,###.##"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   555
            Left            =   4440
            TabIndex        =   245
            Top             =   9600
            Width           =   1485
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Œœ„… «·”Ì«—« "
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
            Height          =   675
            Left            =   -2280
            RightToLeft     =   -1  'True
            TabIndex        =   237
            Top             =   480
            Width           =   1485
         End
         Begin VB.Image Image5 
            Height          =   315
            Left            =   21480
            Picture         =   "FrmSaleBill333.frx":1B7EA
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·—ÃÊ⁄ ··„Ã„Ê⁄« "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   270
            Index           =   57
            Left            =   30525
            TabIndex        =   236
            Top             =   8565
            Width           =   1650
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            FillStyle       =   0  'Solid
            Height          =   5775
            Left            =   22320
            Top             =   5160
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·⁄„Ì·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   7
            Left            =   21525
            TabIndex        =   231
            Top             =   2925
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·œ›⁄"
            Height          =   300
            Index           =   9
            Left            =   3300
            TabIndex        =   230
            Top             =   15720
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
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
            Height          =   675
            Left            =   -2280
            RightToLeft     =   -1  'True
            TabIndex        =   216
            Top             =   360
            Width           =   1245
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
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
            Height          =   675
            Left            =   -2400
            RightToLeft     =   -1  'True
            TabIndex        =   217
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label13 
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
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   -2760
            RightToLeft     =   -1  'True
            TabIndex        =   218
            Top             =   4440
            Visible         =   0   'False
            Width           =   3165
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
            ForeColor       =   &H00FF0000&
            Height          =   435
            Index           =   80
            Left            =   21000
            RightToLeft     =   -1  'True
            TabIndex        =   222
            Top             =   3240
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.Label LblTotalView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   6000
            TabIndex        =   223
            Top             =   9240
            Visible         =   0   'False
            Width           =   9255
         End
         Begin VB.Image Image3 
            Height          =   435
            Left            =   23160
            Stretch         =   -1  'True
            Top             =   4320
            Width           =   4275
         End
         Begin VB.Image Image2 
            Height          =   435
            Left            =   -1080
            Stretch         =   -1  'True
            Top             =   4320
            Width           =   3555
         End
         Begin VB.Label Label10 
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
            TabIndex        =   221
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label Label12 
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
            Left            =   21720
            RightToLeft     =   -1  'True
            TabIndex        =   219
            Top             =   840
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00E0E0E0&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   20760
            Top             =   4440
            Width           =   8295
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BorderColor     =   &H00E0E0E0&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   1335
            Left            =   0
            Top             =   6480
            Width           =   4815
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00E0E0E0&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   600
            Left            =   19560
            Top             =   9480
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Image Image15 
            Height          =   7515
            Left            =   0
            Picture         =   "FrmSaleBill333.frx":1BF10
            Stretch         =   -1  'True
            Top             =   3960
            Width           =   6900
         End
         Begin VB.Shape Shape8 
            BackColor       =   &H00E0E0E0&
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   5
            FillColor       =   &H00FFFFFF&
            Height          =   1125
            Left            =   7680
            Top             =   240
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1800
         Index           =   0
         Left            =   15
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   645
         Width           =   27885
         _cx             =   49186
         _cy             =   3175
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
         Begin VB.CommandButton Command3 
            Caption         =   "«” ⁄·«„ ⁄‰ ’‰›"
            Height          =   255
            Left            =   7755
            TabIndex        =   93
            Top             =   1680
            Visible         =   0   'False
            Width           =   3150
         End
         Begin VB.TextBox TxtIssueSerial 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   0
            TabIndex        =   81
            Top             =   -240
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1875
            TabIndex        =   79
            Top             =   -240
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.TextBox TXTOrDer_no 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8880
            TabIndex        =   74
            Top             =   1080
            Width           =   1650
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   21735
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   120
            Width           =   3675
         End
         Begin VB.CommandButton Command1 
            Caption         =   " ÕÊÌ· «·Ï «–‰ ’—›"
            Height          =   255
            Left            =   15180
            TabIndex        =   69
            Top             =   -120
            Visible         =   0   'False
            Width           =   3240
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   17310
            TabIndex        =   3
            Top             =   1080
            Width           =   3360
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   14130
            TabIndex        =   6
            Top             =   60
            Width           =   1545
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   23820
            TabIndex        =   4
            Top             =   1440
            Width           =   1590
         End
         Begin VB.TextBox TxtCusID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   23820
            TabIndex        =   2
            Top             =   765
            Width           =   1590
         End
         Begin VB.ComboBox CboSaleType 
            Height          =   315
            Left            =   5865
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   690
            Visible         =   0   'False
            Width           =   2985
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   22485
            TabIndex        =   0
            Top             =   -180
            Visible         =   0   'False
            Width           =   2610
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   735
            Index           =   8
            Left            =   25575
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1680
            Visible         =   0   'False
            Width           =   7305
            _cx             =   12885
            _cy             =   1296
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
            Style           =   1
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
            Begin ImpulseButton.ISButton CmdInvProfit 
               Height          =   390
               Left            =   9885
               TabIndex        =   25
               Top             =   165
               Width           =   5400
               _ExtentX        =   9525
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "..."
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
               ButtonImage     =   "FrmSaleBill333.frx":1DBF2
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰”»… «·—»Õ"
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   23
               Left            =   15060
               TabIndex        =   30
               Top             =   420
               Width           =   7920
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·—»Õ"
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   22
               Left            =   63690
               TabIndex        =   29
               Top             =   150
               Width           =   7800
            End
            Begin VB.Label lblInvPrecent 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   6495
               TabIndex        =   28
               Top             =   390
               Width           =   10005
            End
            Begin VB.Label LblInvProfit1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   6495
               TabIndex        =   27
               Top             =   105
               Width           =   10005
            End
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   17310
            TabIndex        =   5
            Top             =   1440
            Width           =   6510
            _ExtentX        =   11483
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   345
            Left            =   21795
            TabIndex        =   1
            Top             =   420
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   609
            _Version        =   393216
            Format          =   171376641
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   360
            Left            =   25575
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   750
            Visible         =   0   'False
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   635
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            BackStyle       =   0
            ButtonImage     =   "FrmSaleBill333.frx":1DF8C
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   8760
            TabIndex        =   7
            Top             =   45
            Width           =   5370
            _ExtentX        =   9472
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdCash 
            Height          =   390
            Index           =   0
            Left            =   17205
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   900
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   14871017
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   0
            ButtonImage     =   "FrmSaleBill333.frx":1E326
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdCash 
            Height          =   270
            Index           =   1
            Left            =   16980
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1140
            Visible         =   0   'False
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   14871017
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   0
            ButtonImage     =   "FrmSaleBill333.frx":1E6C0
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   17205
            TabIndex        =   75
            Top             =   120
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   8760
            TabIndex        =   77
            Top             =   375
            Width           =   6930
            _ExtentX        =   12224
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCDocTypes 
            Height          =   315
            Left            =   17310
            TabIndex        =   128
            Top             =   480
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   1815
            Left            =   75
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   0
            Width           =   4515
            _cx             =   7964
            _cy             =   3201
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
            Begin VB.TextBox TxtManualNo2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   240
               TabIndex        =   136
               Top             =   360
               Width           =   1140
            End
            Begin VB.TextBox TxtManualNo1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   240
               TabIndex        =   134
               Top             =   0
               Width           =   1140
            End
            Begin VB.TextBox txt_Currency_rate 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   130
               Text            =   "1"
               Top             =   1440
               Width           =   1005
            End
            Begin VB.Frame Frame2 
               Caption         =   " œ·«·«  «·«·Ê«‰"
               Height          =   735
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   720
               Width           =   2280
               Begin VB.Label Label5 
                  BackColor       =   &H000000FF&
                  Height          =   255
                  Left            =   1920
                  TabIndex        =   99
                  Top             =   240
                  Width           =   255
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "»Ì⁄ «ﬁ· „‰ ”⁄— «· ﬂ·›Â"
                  Height          =   255
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.Label Label7 
                  BackColor       =   &H0000FFFF&
                  Height          =   255
                  Left            =   1920
                  TabIndex        =   97
                  Top             =   480
                  Width           =   255
               End
               Begin VB.Label Label8 
                  Alignment       =   1  'Right Justify
                  Caption         =   "»Ì⁄  »”⁄— «· ﬂ·›Â"
                  Height          =   255
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   480
                  Width           =   1215
               End
            End
            Begin MSDataListLib.DataCombo DcCurrency 
               Height          =   315
               Left            =   1140
               TabIndex        =   131
               Top             =   1440
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„  »Ê·Ì’… «·‘Õ‰"
               Height          =   195
               Index           =   67
               Left            =   1440
               TabIndex        =   137
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «–‰ «· ”·Ì„"
               Height          =   195
               Index           =   66
               Left            =   1440
               TabIndex        =   135
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "««·⁄„·…"
               Height          =   285
               Index           =   65
               Left            =   2265
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   1440
               Width           =   540
            End
         End
         Begin VB.Frame Frame400 
            Height          =   495
            Left            =   12195
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   1320
            Width           =   4725
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—»Õ «·›« Ê—…"
               ForeColor       =   &H00008000&
               Height          =   195
               Index           =   68
               Left            =   1680
               TabIndex        =   141
               Top             =   240
               Width           =   960
            End
            Begin VB.Label LblPrecenValuex 
               Caption         =   "0"
               Height          =   255
               Left            =   120
               TabIndex        =   140
               Top             =   840
               Width           =   1455
            End
            Begin VB.Label LblInvProfit 
               Alignment       =   2  'Center
               Caption         =   "0"
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   120
               TabIndex        =   139
               Top             =   240
               Width           =   1575
            End
         End
         Begin MSComCtl2.DTPicker DtpDelayDate 
            Height          =   285
            Left            =   4740
            TabIndex        =   142
            Top             =   1440
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            Format          =   162660353
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «·«” Õﬁ«ﬁ"
            Height          =   270
            Index           =   21
            Left            =   7020
            TabIndex        =   143
            Top             =   1515
            Width           =   1620
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·”‰œ"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   20130
            TabIndex        =   129
            Top             =   480
            Width           =   1350
         End
         Begin VB.Label Label4 
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   255
            Left            =   1860
            TabIndex        =   80
            Top             =   480
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·Œ“‰…"
            Height          =   225
            Index           =   11
            Left            =   15615
            TabIndex        =   78
            Top             =   480
            Width           =   1365
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   20100
            TabIndex        =   76
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÿ·»Ì…"
            Height          =   240
            Index           =   56
            Left            =   10575
            TabIndex        =   73
            Top             =   1200
            Width           =   1560
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
            Left            =   8400
            TabIndex        =   70
            Top             =   840
            Width           =   555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÏ"
            Height          =   300
            Index           =   33
            Left            =   20790
            TabIndex        =   36
            Top             =   1140
            Width           =   2100
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”Ì«”… «·»Ì⁄"
            Height          =   240
            Index           =   32
            Left            =   15615
            TabIndex        =   32
            Top             =   1410
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·»«∆⁄"
            Height          =   285
            Index           =   25
            Left            =   15825
            TabIndex        =   23
            Top             =   75
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   240
            Index           =   24
            Left            =   25095
            TabIndex        =   15
            Top             =   1485
            Width           =   2775
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·›« Ê—…"
            Height          =   285
            Index           =   6
            Left            =   23775
            TabIndex        =   14
            Top             =   420
            Width           =   3900
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·›« Ê—…"
            Height          =   285
            Index           =   5
            Left            =   24825
            TabIndex        =   13
            Top             =   75
            Width           =   2850
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   8040
         Left            =   15
         TabIndex        =   10
         Top             =   2460
         Width           =   27885
         _cx             =   49186
         _cy             =   14182
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
         BackColor       =   12648447
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "«·√’‰«›|«·«ﬁ”«ÿ  Ê «·‘Ìﬂ« |„·«ÕŸ«  ⁄·Ï «·›« Ê—…|«·„—›ﬁ« "
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
         Picture(0)      =   "FrmSaleBill333.frx":1EA5A
         Picture(1)      =   "FrmSaleBill333.frx":1EDF4
         Picture(2)      =   "FrmSaleBill333.frx":1F18E
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7575
            Index           =   19
            Left            =   29130
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   45
            Width           =   27795
            _cx             =   49027
            _cy             =   13361
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
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  ﬁÌœ «·›« Ê—Â"
               Height          =   1575
               Left            =   2160
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   720
               Width           =   4335
               Begin VB.TextBox TxtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   600
                  Width           =   2505
               End
               Begin ImpulseButton.ISButton Cmd 
                  CausesValidation=   0   'False
                  Height          =   375
                  Index           =   10
                  Left            =   240
                  TabIndex        =   91
                  Top             =   600
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   661
                  ButtonStyle     =   1
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
                  ColorShadow     =   4210752
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   4210752
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ﬁÌœ ··›« Ê—Â"
                  Height          =   435
                  Index           =   62
                  Left            =   2880
                  TabIndex        =   92
                  Top             =   240
                  Width           =   1215
               End
            End
            Begin VB.OptionButton BillBasedOn 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›« Ê—… „»Ì⁄« "
               Height          =   195
               Index           =   0
               Left            =   10320
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   360
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   4785
            End
            Begin VB.OptionButton BillBasedOn 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "√Ê«„— «·»Ì⁄"
               Height          =   195
               Index           =   2
               Left            =   10680
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   3000
               Visible         =   0   'False
               Width           =   4305
            End
            Begin VB.OptionButton BillBasedOn 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰œ«  «·’—›"
               Height          =   195
               Index           =   1
               Left            =   9480
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   600
               Width           =   5625
            End
            Begin VB.TextBox TXTNoteID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   0
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VSFlex8UCtl.VSFlexGrid GRID1 
               Height          =   2085
               Left            =   6960
               TabIndex        =   82
               Tag             =   "1"
               Top             =   840
               Width           =   8415
               _cx             =   14843
               _cy             =   3678
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   0   'False
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
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmSaleBill333.frx":1F528
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
            Begin VSFlex8UCtl.VSFlexGrid GRID2 
               Height          =   1725
               Left            =   7080
               TabIndex        =   84
               Tag             =   "1"
               Top             =   3240
               Width           =   8175
               _cx             =   14420
               _cy             =   3043
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
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSaleBill333.frx":1F675
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
            Begin ImpulseButton.ISButton Cmd1 
               CausesValidation=   0   'False
               Height          =   375
               Left            =   5160
               TabIndex        =   144
               Top             =   2640
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "„—›ﬁ«  «·›« Ê—…"
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›« Ê—Â »‰«¡ ⁄·Ï"
               Height          =   300
               Index           =   61
               Left            =   12600
               TabIndex        =   86
               Top             =   120
               Width           =   2520
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7575
            Index           =   15
            Left            =   28830
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   45
            Width           =   27795
            _cx             =   49027
            _cy             =   13361
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
            GridRows        =   7
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmSaleBill333.frx":1F768
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1575
               Index           =   18
               Left            =   15
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   795
               Visible         =   0   'False
               Width           =   27765
               _cx             =   48974
               _cy             =   2778
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
               Begin VB.Frame Frame4 
                  Height          =   30
                  Left            =   285
                  TabIndex        =   164
                  Top             =   -15
                  Width           =   90
                  Begin VB.ComboBox CboPaymentType1 
                     Height          =   315
                     Left            =   0
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   166
                     Top             =   585
                     Width           =   2685
                  End
                  Begin VB.TextBox TxtAdvPaymnt 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   0
                     MaxLength       =   10
                     RightToLeft     =   -1  'True
                     TabIndex        =   165
                     Top             =   240
                     Width           =   2685
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "ÿ—Ìﬁ… «·ﬁ»÷"
                     Height          =   315
                     Index           =   79
                     Left            =   2850
                     RightToLeft     =   -1  'True
                     TabIndex        =   168
                     Top             =   585
                     Width           =   1275
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "ﬁÌ„… «·œ›⁄Â"
                     Height          =   285
                     Index           =   78
                     Left            =   2850
                     RightToLeft     =   -1  'True
                     TabIndex        =   167
                     Top             =   255
                     Width           =   1275
                  End
               End
               Begin VB.Frame FraNote 
                  BackColor       =   &H00E2E9E9&
                  Height          =   30
                  Left            =   210
                  RightToLeft     =   -1  'True
                  TabIndex        =   152
                  Top             =   -15
                  Width           =   75
                  Begin VB.TextBox TxtChequeNumber 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   30
                     RightToLeft     =   -1  'True
                     TabIndex        =   154
                     Top             =   930
                     Width           =   2685
                  End
                  Begin VB.TextBox TXTBankName 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   153
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   2685
                  End
                  Begin MSComCtl2.DTPicker DtpChequeDueDate1 
                     Height          =   315
                     Left            =   30
                     TabIndex        =   155
                     Top             =   1260
                     Width           =   2685
                     _ExtentX        =   4736
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   162660353
                     CurrentDate     =   39614
                  End
                  Begin MSDataListLib.DataCombo DcboBankName1 
                     Height          =   315
                     Left            =   30
                     TabIndex        =   156
                     Top             =   600
                     Width           =   2685
                     _ExtentX        =   4736
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcboBox1 
                     Height          =   315
                     Left            =   30
                     TabIndex        =   157
                     Top             =   270
                     Width           =   2685
                     _ExtentX        =   4736
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcChequeBox1 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   158
                     Top             =   1680
                     Width           =   2685
                     _ExtentX        =   4736
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
                     Height          =   285
                     Index           =   77
                     Left            =   2820
                     RightToLeft     =   -1  'True
                     TabIndex        =   163
                     Top             =   1260
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «·‘Ìﬂ"
                     Height          =   285
                     Index           =   76
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   162
                     Top             =   930
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·»‰ﬂ"
                     Height          =   285
                     Index           =   75
                     Left            =   2790
                     RightToLeft     =   -1  'True
                     TabIndex        =   161
                     Top             =   630
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·Œ“‰…"
                     Height          =   285
                     Index           =   74
                     Left            =   2790
                     RightToLeft     =   -1  'True
                     TabIndex        =   160
                     Top             =   300
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ«›Ÿ… «·‘Ìﬂ« "
                     Height          =   285
                     Index           =   73
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   159
                     Top             =   1560
                     Width           =   1215
                  End
               End
               Begin VB.TextBox TxtTaxServiceValue 
                  Alignment       =   1  'Right Justify
                  Height          =   15
                  Left            =   150
                  MaxLength       =   4
                  TabIndex        =   56
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   30
               End
               Begin VB.CheckBox ChkTaxSerivce 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì… Œœ„…"
                  Height          =   15
                  Left            =   210
                  TabIndex        =   51
                  Top             =   0
                  Width           =   30
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   15
                  Index           =   54
                  Left            =   105
                  TabIndex        =   68
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   30
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
                  Height          =   15
                  Index           =   47
                  Left            =   135
                  TabIndex        =   61
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   15
                  Index           =   43
                  Left            =   180
                  TabIndex        =   57
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   15
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1575
               Index           =   17
               Left            =   15
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   795
               Visible         =   0   'False
               Width           =   27765
               _cx             =   48974
               _cy             =   2778
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
               Begin VB.TextBox TxtTaxStampValue 
                  Alignment       =   1  'Right Justify
                  Height          =   15
                  Left            =   150
                  MaxLength       =   4
                  TabIndex        =   55
                  Top             =   0
                  Width           =   30
               End
               Begin VB.CheckBox ChkTaxStamp 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ„€…"
                  Height          =   0
                  Left            =   210
                  TabIndex        =   49
                  Top             =   15
                  Width           =   0
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   15
                  Index           =   53
                  Left            =   105
                  TabIndex        =   67
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   30
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "$"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   15
                  Index           =   48
                  Left            =   135
                  TabIndex        =   62
                  Top             =   0
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   15
                  Index           =   41
                  Left            =   180
                  TabIndex        =   53
                  Top             =   0
                  Width           =   15
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1575
               Index           =   16
               Left            =   15
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   795
               Visible         =   0   'False
               Width           =   27765
               _cx             =   48974
               _cy             =   2778
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
               Begin VB.TextBox TxtTaxAddValue 
                  Alignment       =   1  'Right Justify
                  Height          =   15
                  Left            =   150
                  MaxLength       =   4
                  TabIndex        =   54
                  Top             =   0
                  Width           =   30
               End
               Begin VB.CheckBox ChkTaxAdd 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… Œ’„ Ê≈÷«›… (√—»«Õ  Ã«—Ì…)"
                  Height          =   15
                  Left            =   195
                  TabIndex        =   47
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   15
                  Index           =   52
                  Left            =   120
                  TabIndex        =   66
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   15
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
                  Height          =   15
                  Index           =   46
                  Left            =   135
                  TabIndex        =   60
                  Top             =   0
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   15
                  Index           =   39
                  Left            =   180
                  TabIndex        =   52
                  Top             =   0
                  Width           =   15
               End
            End
            Begin VB.TextBox TxtBillComment 
               Alignment       =   1  'Right Justify
               Height          =   1575
               Left            =   15
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   40
               Top             =   795
               Width           =   27765
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   765
               Index           =   4
               Left            =   15
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   15
               Visible         =   0   'False
               Width           =   27765
               _cx             =   48974
               _cy             =   1349
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
               Begin VB.CheckBox XPChkTAX 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… «·„»Ì⁄« "
                  Height          =   315
                  Left            =   2415
                  TabIndex        =   44
                  Top             =   225
                  Width           =   420
               End
               Begin VB.TextBox XPTxtTaxValue 
                  Alignment       =   1  'Right Justify
                  Height          =   510
                  Left            =   1815
                  MaxLength       =   4
                  TabIndex        =   43
                  Top             =   105
                  Width           =   300
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   360
                  Index           =   51
                  Left            =   300
                  TabIndex        =   65
                  Top             =   135
                  Width           =   120
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
                  Height          =   360
                  Index           =   45
                  Left            =   1755
                  TabIndex        =   59
                  Top             =   135
                  Width           =   60
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   4
                  Left            =   1875
                  TabIndex        =   45
                  Top             =   195
                  Width           =   420
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈÷«›… √Ì… „·«ÕŸ«  ⁄·Ï «·›« Ê—…"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   1575
               Index           =   44
               Left            =   15
               TabIndex        =   58
               Top             =   795
               Width           =   27765
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7575
            Index           =   7
            Left            =   45
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   45
            Width           =   27795
            _cx             =   49027
            _cy             =   13361
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
            AutoSizeChildren=   8
            BorderWidth     =   2
            ChildSpacing    =   1
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
            GridRows        =   8
            GridCols        =   4
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmSaleBill333.frx":1F7E3
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   930
               Index           =   2
               Left            =   30
               TabIndex        =   110
               TabStop         =   0   'False
               Top             =   30
               Width           =   27735
               _cx             =   48921
               _cy             =   1640
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
               AutoSizeChildren=   7
               BorderWidth     =   1
               ChildSpacing    =   1
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
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   510
                  Left            =   1095
                  MaxLength       =   10
                  TabIndex        =   114
                  Top             =   390
                  Width           =   2610
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   510
                  Left            =   7275
                  MaxLength       =   20
                  TabIndex        =   113
                  Top             =   390
                  Width           =   2610
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   510
                  Left            =   4080
                  MaxLength       =   10
                  TabIndex        =   112
                  Top             =   390
                  Width           =   3150
               End
               Begin VB.ComboBox CboItemCase 
                  Height          =   315
                  Left            =   10170
                  Style           =   2  'Dropdown List
                  TabIndex        =   111
                  Top             =   390
                  Width           =   2025
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   12315
                  TabIndex        =   115
                  Top             =   390
                  Width           =   11625
                  _ExtentX        =   20505
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   23970
                  TabIndex        =   116
                  Top             =   390
                  Width           =   3315
                  _ExtentX        =   5847
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   510
                  Left            =   75
                  TabIndex        =   117
                  Top             =   390
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   900
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
                  Caption         =   ""
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
                  BackStyle       =   0
                  ButtonImage     =   "FrmSaleBill333.frx":1F892
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdSearch 
                  Height          =   495
                  Left            =   855
                  TabIndex        =   118
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   873
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "..."
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
                  ButtonImage     =   "FrmSaleBill333.frx":1FC2C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”⁄—"
                  Height          =   360
                  Index           =   26
                  Left            =   1590
                  TabIndex        =   124
                  Top             =   15
                  Width           =   1590
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   375
                  Index           =   27
                  Left            =   4665
                  TabIndex        =   123
                  Top             =   30
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   360
                  Index           =   28
                  Left            =   7620
                  TabIndex        =   122
                  Top             =   15
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   360
                  Index           =   29
                  Left            =   10395
                  TabIndex        =   121
                  Top             =   15
                  Width           =   1470
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   360
                  Index           =   30
                  Left            =   20760
                  TabIndex        =   120
                  Top             =   15
                  Width           =   1500
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   375
                  Index           =   31
                  Left            =   23505
                  TabIndex        =   119
                  Top             =   30
                  Width           =   3465
               End
            End
            Begin MSComctlLib.Toolbar Toolbar1 
               Height          =   630
               Left            =   30
               TabIndex        =   125
               Top             =   30
               Width           =   13785
               _ExtentX        =   24315
               _ExtentY        =   1111
               ButtonWidth     =   609
               ButtonHeight    =   1005
               Appearance      =   1
               _Version        =   393216
            End
            Begin VB.Label LblItemsCount 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               ForeColor       =   &H0000FFFF&
               Height          =   1860
               Left            =   30
               TabIndex        =   31
               Top             =   5685
               Width           =   210
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7575
            Index           =   5
            Left            =   28530
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   45
            Width           =   27795
            _cx             =   49027
            _cy             =   13361
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
            BackColor       =   255
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   0
            ChildSpacing    =   0
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   7575
               Left            =   0
               TabIndex        =   100
               TabStop         =   0   'False
               Top             =   0
               Width           =   20295
               _cx             =   35798
               _cy             =   13361
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   690
                  Index           =   11
                  Left            =   90
                  TabIndex        =   101
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   20115
                  _cx             =   35481
                  _cy             =   1217
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
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   360
                     Index           =   1
                     Left            =   15840
                     Locked          =   -1  'True
                     TabIndex        =   150
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1635
                  End
                  Begin VB.CheckBox ChkInstall 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ”Ìÿ"
                     Height          =   195
                     Left            =   3300
                     TabIndex        =   148
                     Top             =   280
                     Width           =   930
                  End
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "¬Ã· "
                     Height          =   195
                     Index           =   1
                     Left            =   7155
                     TabIndex        =   146
                     Top             =   280
                     Width           =   1215
                  End
                  Begin VB.TextBox XPTxtValue 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Index           =   1
                     Left            =   4560
                     Locked          =   -1  'True
                     MaxLength       =   10
                     TabIndex        =   145
                     Top             =   225
                     Width           =   1500
                  End
                  Begin VB.TextBox XPTxtValue 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Index           =   0
                     Left            =   8820
                     Locked          =   -1  'True
                     MaxLength       =   10
                     TabIndex        =   104
                     Top             =   225
                     Width           =   1515
                  End
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Index           =   0
                     Left            =   14430
                     Locked          =   -1  'True
                     TabIndex        =   103
                     Top             =   75
                     Visible         =   0   'False
                     Width           =   1530
                  End
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰ﬁœ«"
                     Height          =   195
                     Index           =   0
                     Left            =   11670
                     TabIndex        =   102
                     Top             =   280
                     Width           =   930
                  End
                  Begin ImpulseButton.ISButton CmdINSTALLMENT 
                     Height          =   390
                     Left            =   240
                     TabIndex        =   149
                     Top             =   240
                     Width           =   1530
                     _ExtentX        =   2699
                     _ExtentY        =   688
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "Õ”«» «·√ﬁ”«ÿ"
                     BackColor       =   14871017
                     Enabled         =   0   'False
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmSaleBill333.frx":1FFC6
                     ColorButton     =   14871017
                     ColorHighlight  =   16777215
                     ColorHoverText  =   16711680
                     ColorShadow     =   4210752
                     ColorOutline    =   0
                     DrawFocusRectangle=   0   'False
                     ColorToggledHoverText=   16711680
                     ColorTextShadow =   4210752
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„”·”·"
                     Height          =   375
                     Index           =   14
                     Left            =   15495
                     TabIndex        =   151
                     Top             =   315
                     Visible         =   0   'False
                     Width           =   630
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   195
                     Index           =   15
                     Left            =   6330
                     TabIndex        =   147
                     Top             =   280
                     Width           =   420
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
                     Height          =   225
                     Index           =   20
                     Left            =   12780
                     TabIndex        =   107
                     Top             =   250
                     Width           =   1410
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   225
                     Index           =   13
                     Left            =   10815
                     TabIndex        =   106
                     Top             =   285
                     Width           =   600
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„”·”·"
                     Height          =   225
                     Index           =   12
                     Left            =   15270
                     TabIndex        =   105
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   810
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   90
                  Index           =   12
                  Left            =   90
                  TabIndex        =   108
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   20115
                  _cx             =   35481
                  _cy             =   159
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
               End
               Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
                  Height          =   2010
                  Left            =   90
                  TabIndex        =   109
                  Top             =   870
                  Width           =   17385
                  _cx             =   30665
                  _cy             =   3545
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
                  Rows            =   50
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmSaleBill333.frx":20360
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
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   615
         Index           =   9
         Left            =   15
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   15
         Width           =   27885
         _cx             =   49186
         _cy             =   1085
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
         Caption         =   "›« Ê—… „»Ì⁄«  "
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   0
         ChildSpacing    =   0
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   7
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
         Begin VB.CommandButton Command2 
            Caption         =   " ÕÊÌ· «·Ï «–‰ ’—›"
            Height          =   315
            Left            =   11955
            Style           =   1  'Graphical
            TabIndex        =   133
            Top             =   240
            Width           =   7410
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   13125
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   0
            Visible         =   0   'False
            Width           =   2490
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   11580
            TabIndex        =   64
            Top             =   0
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   10575
            TabIndex        =   63
            Top             =   0
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   20745
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   120
            Visible         =   0   'False
            Width           =   3225
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   4290
            TabIndex        =   19
            Top             =   30
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   609
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
            ButtonImage     =   "FrmSaleBill333.frx":20456
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
            Height          =   345
            Index           =   3
            Left            =   2355
            TabIndex        =   20
            Top             =   30
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   609
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
            ButtonImage     =   "FrmSaleBill333.frx":207F0
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
            Height          =   345
            Index           =   1
            Left            =   6270
            TabIndex        =   21
            Top             =   30
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   609
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
            ButtonImage     =   "FrmSaleBill333.frx":20B8A
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
            Height          =   345
            Index           =   2
            Left            =   60
            TabIndex        =   22
            Top             =   30
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
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
            ButtonImage     =   "FrmSaleBill333.frx":20F24
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin ImpulseButton.ISButton CmdNotes 
            Height          =   345
            Left            =   16755
            TabIndex        =   33
            Top             =   120
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   3
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
            ButtonImage     =   "FrmSaleBill333.frx":212BE
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdRetruns 
            Height          =   345
            Left            =   7470
            TabIndex        =   34
            Top             =   0
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   3
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
            ButtonImage     =   "FrmSaleBill333.frx":21658
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdInfo 
            Height          =   615
            Left            =   9255
            TabIndex        =   72
            Top             =   -120
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   1085
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmSaleBill333.frx":21BF2
            ButtonImageHover=   "FrmSaleBill333.frx":228CC
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   64
            Left            =   10680
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   0
            Width           =   11700
         End
         Begin VB.Label LblShortcutKeys 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "ÃœÌœ F12 Or Enter ,  ⁄œÌ· F11 , Õ›Ÿ F10 ,  —«Ã⁄ F9 ,Õ–› F8 ,»ÕÀ F3 "
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
            Height          =   195
            Left            =   255
            TabIndex        =   35
            Top             =   390
            Width           =   16485
         End
      End
   End
End
Attribute VB_Name = "frmsalebill333"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim NewGrid As New ClsGrid
Dim mIsChangeTable As Boolean
Dim mIsChangeTablePaySave As Boolean
Dim SaleReport As ClsSaleReport
Dim cSearchDcbo(4)   As clsDCboSearch
Dim Dcombos As ClsDataCombos
      Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
            Dim SAVESTATUS As Boolean
            Dim imageCounter As Integer
            Dim currenttableidMain As Long
Public BolPrint As Boolean
Public TimeOut_InSec As Long
'*****
Dim visapayed As Double
Public WithEvents m_Menu1 As Menu
Attribute m_Menu1.VB_VarHelpID = -1
Dim WithEvents m_MenuRefesh As Menu
Attribute m_MenuRefesh.VB_VarHelpID = -1
Dim WithEvents m_MenuCusBalance As Menu
Attribute m_MenuCusBalance.VB_VarHelpID = -1
Dim WithEvents m_MenuViewList As Menu
Attribute m_MenuViewList.VB_VarHelpID = -1
Dim WithEvents m_MenuViewNotes As Menu
Attribute m_MenuViewNotes.VB_VarHelpID = -1
Dim WithEvents m_MenuScreenPremission As Menu
Attribute m_MenuScreenPremission.VB_VarHelpID = -1
Dim WithEvents StrCashCustomerPhone As TextBox
Attribute StrCashCustomerPhone.VB_VarHelpID = -1
Dim WithEvents StrCashCustomerMobile As TextBox
Attribute StrCashCustomerMobile.VB_VarHelpID = -1
Dim WithEvents StrCashCustomerAddress As TextBox
Attribute StrCashCustomerAddress.VB_VarHelpID = -1
Dim WithEvents m_FrmSearch As Form
Attribute m_FrmSearch.VB_VarHelpID = -1
Dim first_run As Boolean
Dim bank_account As String
Dim general_noteid As Long
Dim RsNotesGeneral As ADODB.Recordset
Dim CurrentVoucherNo As String
Dim CurrentVoucherSerialNo As String

Dim DateChanged As Boolean
 Private Declare Function GetQueueStatus Lib "user32" _
   (ByVal fuFlags As Long) As Long

Private Const QS_KEY = &H1
Private Const QS_MOUSEMOVE = &H2
Private Const QS_MOUSEBUTTON = &H4
Private Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT = (QS_MOUSE Or QS_KEY)
Public bCancel As Boolean
Private Type PLASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type

Private Declare Function GetLastInputInfo Lib "user32.dll" (ByRef plii As PLASTINPUTINFO) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Sub printtomanyprinter2()
Dim VarSet As Variant
Dim a As String
'Exit Sub

Open App.path & "\printersGroup.txt" For Input As #1
    dbname.Clear

    Do Until EOF(1)
        Line Input #1, a
        'subsequent lines
 
        If a <> "" Then
            VarSet = Split(a, "*", , vbTextCompare)

            If VarSet(0) <> Empty Or VarSet(0) <> "" Then
            
                CBOPrinter.AddItem a
             SetPrinter2 (a)
          printtoAnotherprinter2 a
            DoEvents
            End If
        End If
    
    Loop

    Close #1
    


'Exit Sub

End Sub

Sub printtoAnotherprinter2(GroupPrinterName As String)
'print by Group
'-----------------------------------------------------------------------------
    
    Dim intLineCtr          As Integer
    Dim intPageCtr          As Integer
    Dim intX                As Integer
    Dim strCustFileName     As String
    Dim strBackSlash        As String
    Dim intCustFileNbr      As Integer
    
    
    Const intLINE_START_POS As Integer = 0
    Const intLINES_PER_PAGE As Integer = 60
    
    ' Have the user make sure his/her printer is ready ...
 
    
    ' Set the printer font to Courier, if available (otherwise, we would be
    ' relying on the default font for the Windows printer, which may or
    ' may not be set to an appropriate font) ...
 
    For intX = 0 To Printer.FontCount - 1
        If Printer.Fonts(intX) Like "Arabic*" Then
            Printer.FontName = Printer.Fonts(intX)
            Exit For
        End If
    Next
    
    Printer.fontsize = 10
    
    ' initialize report variables ...
    intPageCtr = 0
    intLineCtr = 99 ' initialize line counter to an arbitrarily high number
                    ' to force the first page break
                    
    Dim openingdate As Date
    Dim StrSQL  As String
    Dim rs As New ADODB.Recordset
    StrSQL = " SELECT  dbo.Transaction_Details.Remarks,   dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.printed, dbo.TblItems.ItemName,dbo.TblItems.ItemNamee, "
StrSQL = StrSQL & "                      dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice AS value, dbo.Transaction_Details.Transaction_ID, dbo.Groups.GroupPrinterName,"
StrSQL = StrSQL & "                      dbo.Transaction_Details.ID"
StrSQL = StrSQL & " FROM         dbo.Transaction_Details INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
StrSQL = StrSQL & "                      dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID"
                      
StrSQL = StrSQL & " WHERE     (dbo.Transaction_Details.printedGroup IS NULL or dbo.Transaction_Details.printedGroup =0) AND (dbo.Transaction_Details.Transaction_ID = " & val(XPTxtBillID.text) & ")"
StrSQL = StrSQL & " and  (dbo.Groups.GroupPrinterName = N'" & GroupPrinterName & "') "
 StrSQL = StrSQL & " ORDER BY dbo.Transaction_Details.ID "
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
     Exit Sub
    End If
 
 
 
    Dim RowNum As Integer
     For RowNum = 1 To rs.RecordCount
         If 1 = 1 Then
        
        If intLineCtr > intLINES_PER_PAGE Then
            GoSub PrintHeadings
        End If
        ' print a line of data
     '   Printer.Print Tab(intLINE_START_POS); _
     '                 IIf(IsNull(rs("VALUE").value), "", rs("VALUE").value); _
     '                 Tab(7 + intLINE_START_POS); _
     '                 IIf(IsNull(rs("showPrice").value), "", rs("showPrice").value); _
     '                 Tab(14 + intLINE_START_POS); _
     '                 IIf(IsNull(rs("ShowQty").value), "", rs("ShowQty").value); _
     '                 Tab(21 + intLINE_START_POS); _
     '                 IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value);
     '
                      
                              Printer.Print Tab(intLINE_START_POS); _
                      IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value); _
                      Tab(7 + intLINE_START_POS); _
                    IIf(IsNull(rs("ItemNamee").value), "", rs("ItemNamee").value); _
                      Tab(14 + intLINE_START_POS); _
                    IIf(IsNull(rs("ShowQty").value), "", rs("ShowQty").value); _
                      Tab(21 + intLINE_START_POS); _
                    IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value);


'                      Fg.TextMatrix(RowNum, Fg.ColIndex("Name"));
'Fg.TextMatrix(RowNum, Fg.ColIndex("Name"));
        ' increment the line count
        intLineCtr = intLineCtr + 1
        If intLineCtr = 1 Then Exit Sub
  '  Loop

    ' close the input file
 
 End If
 rs.MoveNext




Next RowNum
     Printer.EndDoc
    
 
 
    Dim sql As String
       sql = "update Transaction_Details set printedGroup=1   where  Transaction_ID=" & val(XPTxtBillID.text)
               
sql = sql & " and  Item_ID in ("
sql = sql & "  SELECT DISTINCT dbo.Transaction_Details.Item_ID"
sql = sql & "  FROM         dbo.Transaction_Details INNER JOIN"
sql = sql & "                       dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
sql = sql & "                       dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID"
sql = sql & "  WHERE     (dbo.Groups.GroupPrinterName = N'" & GroupPrinterName & "') ) "

            Cn.Execute sql
            Debug.Print sql
            
  Exit Sub

PrintHeadings:
'------------
     If intPageCtr > 0 Then
        Printer.NewPage
    End If
    ' increment the page counter
    intPageCtr = intPageCtr + 1
    
     Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    
    ' Print the main headings
    Printer.Print Tab(intLINE_START_POS); _
                  "Print Date: "; _
                  Format$(Date, "mm/dd/yy"); _
                  Tab(intLINE_START_POS + 31); _
                    "NO:"; Me.TxtNoteSerial1.text; _
                  Tab(intLINE_START_POS + 73); _
                  ""; _
                  'Format$(intPageCtr, "@@@")
    Printer.Print Tab(intLINE_START_POS); _
                  "Print Time: "; _
                  Format$(Time, "hh:nn:ss"); _
                  Tab(intLINE_START_POS + 33); _
                  "Table:" & " " & LBLTable1.Caption
                  '"Table:" & GroupPrinterName & LblTable1.Caption
    Printer.Print
    ' Print the column headings
    Printer.Print Tab(intLINE_START_POS); _
                  "item"; _
                  Tab(7 + intLINE_START_POS); _
                  "QTY"; _
                  Tab(14 + intLINE_START_POS); _
                  "Remarks";
                   
       
    Printer.Print Tab(intLINE_START_POS); _
                  "------"; _
                  Tab(7 + intLINE_START_POS); _
                  "------"; _
                  Tab(14 + intLINE_START_POS); _
                  "------"; _
                  Tab(21 + intLINE_START_POS); _
                  "------";
    Printer.Print
     intLineCtr = 6
    Return
            
            
End Sub



Public Sub CheckInputIdle(ByVal TimeOut_InSec As Long)
Dim t As Long
t = Timer
Do While bCancel = False
If GetQueueStatus(QS_INPUT) Then
t = Timer
DoEvents
End If
If Timer - t >= TimeOut_InSec Then Exit Do
Loop
'If bCancel = False Then SFrmScreenSaver.show
End Sub

Function addrow(ItemID As Integer, _
                ItemName As String, _
                ITEMPRICE As Double, _
                ItemType As Integer)
    lblqty.Caption = ""
    Dim Msg        As String
    Dim LngRow     As Long
    Dim LngFindRow As Long
    Dim des        As String
    On Error Resume Next
    
    'Me.DCboItemsName.Text = itemname
    Me.DCboItemsName.BoundText = ItemID
    TxtQuantity.text = 1
    NewGrid.CmdAddData_Click
    
    With FG
        .Row = .Rows - 1
    End With

    Image16.Visible = False
    Dim StrSQL As String
    Dim rs     As ADODB.Recordset
  
    StrSQL = " Select * from TblItems where ItemID=" & ItemID

    Set rs = New ADODB.Recordset
    'rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdText
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.BOF Or rs.EOF Then
       
        Exit Function
    End If

    If Not (IsNull(rs("ItemPhoto").value)) Then
        Image16.Visible = True
        LoadPictureFromDB Image16, rs, "ItemPhoto"
        rs.Close
        Set rs = Nothing
    Else
        Image16.Visible = False
    End If
   
    Exit Function
    
    '    Me.Grid.Rows = Me.Grid.Rows + 1
    '    LngRow = Me.Grid.Rows - 1
    ' With Me.Grid
    '     .TextMatrix(LngRow, .ColIndex("Code")) = ITEMID
    '     .TextMatrix(LngRow, .ColIndex("Name")) = itemname
    '      .TextMatrix(LngRow, .ColIndex("Count")) = 1
    '      .TextMatrix(LngRow, .ColIndex("Price")) = ITEMPRICE
    '       .TextMatrix(LngRow, .ColIndex("Totals")) = ITEMPRICE
    '      .TextMatrix(LngRow, .ColIndex("ItemType")) = ItemType
    '      .AutoSize 0, .Cols - 1, False
    '
    '      .Row = .Rows - 1
    'End With
    '
 
    ' ReLineGrid

End Function

Private Sub RemoveGridRow()
    'With Me.Grid
    '    If .Row <= 0 Then Exit Sub
    '      .RemoveItem .Row
    'End With
    'ReLineGrid
End Sub

Private Sub ReLineGrid()
    On Error Resume Next
    On Error Resume Next
    Dim i As Integer
    Dim IntCounter As Integer
    Dim totalPayed As Double
 totalPayed = 0
 visapayed = 0
  With Grid

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("value")) <> "" Then
               ' IntCounter = IntCounter + 1
                totalPayed = totalPayed + .TextMatrix(i, .ColIndex("value"))
               If i > 1 Then
                     visapayed = visapayed + .TextMatrix(i, .ColIndex("value"))
               End If
               
               ' .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            End If

        Next i

    End With
  TxtPayedValue = totalPayed


End Sub

Private Sub btnMove_Click(Index As Integer)
'FG.SetFocus
    With Me.FG
Select Case Index
Case 0
.Row = 1

Case 1
If .Row >= 1 Then
.Row = .Row - 1
End If


Case 2
If .Row < .Rows - 1 Then
.Row = .Row + 1
End If

Case 3
.Row = .Rows - 1


End Select
End With
End Sub

Private Sub CMDADDQty_Click()
    'If val(lblqty.Caption) = 0 Then Exit Sub

    With Me.FG
        If .TextMatrix(.Row, .ColIndex("printed")) <> "" Then
        
       MsgBox "·« Ì„ﬂ‰  ⁄œÌ· ﬂ„Ì… Â–… «·’‰› ·«‰Â «—”· «·Ï «·„ÿ»Œ", vbCritical
        Exit Sub
        End If
        .TextMatrix(.Row, .ColIndex("Count")) = .TextMatrix(.Row, .ColIndex("Count")) + 1
     If val(.TextMatrix(.Row, .ColIndex("Count"))) < 0 Then .TextMatrix(.Row, .ColIndex("Count")) = 0
        NewGrid.Grid_AfterEdit .Row, .ColIndex("Count")
    
    
    
    If lblqty.Caption <> "0" Then
    lblShowQty2.Caption = "" & .TextMatrix(.Row, .ColIndex("Count"))
   Else
  lblShowQty2.Caption = ""
  End If
  End With
End Sub
Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "SELECT     dbo.TblPaymentType.PaymentID, dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId, dbo.TblPaymentType.Accountsus, "
My_SQL = My_SQL & "  dbo.TblPaymentType.Accountcom, dbo.TblPaymentType.commision, dbo.TblPaymentType.PaymentNamee, dbo.BanksData.Account_Code AS bankAccount_Code"
My_SQL = My_SQL & " FROM         dbo.TblPaymentType LEFT OUTER JOIN"
My_SQL = My_SQL & " dbo.BanksData ON dbo.TblPaymentType.BankId = dbo.BanksData.BankID order by PaymentID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 2
            rs.MoveFirst
      If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(1, .ColIndex("PaymentName")) = " ‰ﬁœÌ"
               Else
               .TextMatrix(1, .ColIndex("PaymentName")) = " Cash"
               End If
               
                .TextMatrix(1, .ColIndex("PaymentID")) = 0
           
           
            For i = 2 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentName").value), "", rs.Fields("PaymentName").value)
               Else
               .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentNamee").value), "", rs.Fields("PaymentNamee").value)
               End If
               
                .TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(rs.Fields("PaymentID").value), "", rs.Fields("PaymentID").value)
           
                .TextMatrix(i, .ColIndex("BankId")) = IIf(IsNull(rs.Fields("BankId").value), "", rs.Fields("BankId").value)
            
            .TextMatrix(i, .ColIndex("Accountsus")) = IIf(IsNull(rs.Fields("Accountsus").value), "", rs.Fields("Accountsus").value)
            .TextMatrix(i, .ColIndex("Accountcom")) = IIf(IsNull(rs.Fields("Accountcom").value), "", rs.Fields("Accountcom").value)
            .TextMatrix(i, .ColIndex("commision")) = IIf(IsNull(rs.Fields("commision").value), "", rs.Fields("commision").value)
           .TextMatrix(i, .ColIndex("bankAccount_Code")) = IIf(IsNull(rs.Fields("bankAccount_Code").value), "", rs.Fields("bankAccount_Code").value)
            
                rs.MoveNext
            Next

            rs.Close
        End If

  '      .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

Private Sub CMDAdminLogin_Click()
FrameAdmi.Visible = False
                      
End Sub

Private Sub CMDPAy_Click()
If SystemOptions.AllowPOSPAy = False Then

If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " Ì„ﬂ‰ﬂ « „«„ ⁄„·Ì… «·œ›⁄ ·Ì” ·œÌﬂ ’·«ÕÌ…   ", vbCritical
Else
MsgBox " Can't pay not alllowed", vbCritical
End If

Exit Sub


End If

SAVESTATUS = True
Dim AskOption As Boolean
Dim Msg As String
If 1 = 1 Then 'return
'TxtPayedValue.Text = TxtNetValue.Text
End If



'************************************************************************************
         Dim RowNum As Integer
    For RowNum = 1 To Grid.Rows - 1
            
                       If val(Grid.TextMatrix(RowNum, Grid.ColIndex("Value"))) < 0 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                             MsgBox "Œÿ√ ·« Ì„ﬂ‰ «œŒ«· ﬁÌ„… ”«·»…" & CHR(13) & Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName"))
                             
                        Else
                                                     MsgBox "ERROR nEGATIVE VALUE IN" & CHR(13) & Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName"))
                        End If
                            Exit Sub
                    End If
   Next RowNum
'***************************************************************************************


          If CboPayMentType.ListIndex = 0 Then

                If val(TxtRemainValue.text) < 0 Then
                    If SystemOptions.UserInterface = EnglishInterface Then
                        Msg = "Enter Correct Payed Value"
                    Else
                        Msg = "  ﬁÌ„Â «·„œ›Ê⁄ €Ì— ’ÕÌÕÂ "
                    End If
             
                   MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
  
                  Exit Sub
                End If
            End If
            
If CboPOSBillType.ListIndex = 4 Then

        '    If DoPremis(Do_Print, Me.Name, True) = False Then
        '        Exit Sub
        '    End If

            If Me.XPTxtBillID.text = "" Then
                Msg = "·« ÊÃœ ›Ê« Ì— ·Ì „ ÿ»«⁄ Â«"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If
            mIsChangeTablePaySave = True
            
            
        Cmd_Click (2)
        mIsChangeTablePaySave = False
            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

            If AskOption = False Then
                FrmSallReportOptions.show vbModal

                If FrmSallReportOptions.UserCanceled = True Then
                    Unload FrmSallReportOptions
                    Exit Sub
                End If

                Unload FrmSallReportOptions
            End If

        '    PrintReport , 1, LBLTable.Caption
        
            'PrintReport , 1, LblTable.Caption, 0
            PrintReport2
         
Else
 Cmd_Click (2)

End If
'btnNew_Click
LBLPayVal.Caption = 0

FramePay.Visible = False
'XPTxtDiscountVal.Visible = False
'TxtItemCodeB.SetFocus
Me.LBLTable1.Caption = ""
Me.LblStableID.Caption = ""
 
SAVESTATUS = False
End Sub

'Private Sub CmdValue_Click(Index As Integer)
'TxtPayedValue.text = CmdValue(Index).Caption
'LBLPayVal.Caption = TxtPayedValue.text
'End Sub

Private Sub Command4_Click()
    FillTables
End Sub

Private Sub Command5_Click()
 lvwItems.Visible = False
lvwTables.Visible = True
mIsChangeTable = True
Cmd_Click 2
mIsChangeTable = False
'  If SystemOptions.AllowPOSPAy = True Then
  Me.TxtModFlg = "R"
loadInvoices
'End If

End Sub

Private Sub Command6_Click()

FrameAdmi.Visible = True
TxtAdminLogin.SetFocus
End Sub

Private Sub Command7_Click()
If TxtAdminLogin.text = SystemOptions.BigUserPw Then
frmaeDiscount.Visible = True
Else
frmaeDiscount.Visible = False
End If

End Sub

Private Sub fg_Click()
    lblqty.Caption = ""
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
ReLineGrid
End Sub

 

Private Sub Image11_Click()
Call Shell("OSK.exe")
End Sub

Private Sub Image12_Click()
   If Me.TxtModFlg.text = "N" Then
        LBLTable.Caption = ""
 LBLTable1.Caption = ""
  
        LblStableID.Caption = -1
        CboPOSBillType.ListIndex = -1
    End If
End Sub

Private Sub Image16_Click()
    Load FrmViewPic
    Set FrmViewPic.MainView.Picture = Image16.Picture
    FrmViewPic.show vbModal

End Sub

Private Sub Image4_Click()
    If Me.TxtModFlg.text = "N" Then
                 If SystemOptions.UserInterface = ArabicInterface Then
        LBLTable1.Caption = " Ê’Ì·"
  Else
            LBLTable1.Caption = "Delivery"
  End If
  
        LblStableID.Caption = -1
        CboPOSBillType.ListIndex = 1
    End If

End Sub

Private Sub Image7_Click()
   If Me.TxtModFlg.text = "N" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        LBLTable1.Caption = "„Õ·Ì"
  Else
            LBLTable1.Caption = "Dine In"
  End If
  
        LblStableID.Caption = -1
        CboPOSBillType.ListIndex = 4
    End If
End Sub

Private Sub Image8_Click()
 If Me.TxtModFlg.text = "N" Then
         If SystemOptions.UserInterface = ArabicInterface Then
        LBLTable1.Caption = "Œ«—ÃÌ"
  Else
            LBLTable1.Caption = "Take out"
  End If
  
        LblStableID.Caption = -1
   
        CboPOSBillType.ListIndex = 2
    End If
End Sub

Private Sub ISButton1_Click()
'    If val(lblqty.Caption) = 0 Then Exit Sub

    With Me.FG
        If .TextMatrix(.Row, .ColIndex("printed")) <> "" Then
          If TxtAdminLogin.text = SystemOptions.BigUserPw Then
                     TxtAdminLogin.text = ""
             Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                  MsgBox "·« Ì„ﬂ‰  ⁄œÌ· ﬂ„Ì… Â–… «·’‰› ·«‰Â «—”· «·Ï «·„ÿ»Œ", vbCritical
                            Else
                                MsgBox "Can't Delete this Items is Sended already to the kitchen", vbCritical
                            End If
                Exit Sub
         End If
        End If
        .TextMatrix(.Row, .ColIndex("Count")) = .TextMatrix(.Row, .ColIndex("Count")) - 1
     If val(.TextMatrix(.Row, .ColIndex("Count"))) < 0 Then .TextMatrix(.Row, .ColIndex("Count")) = 0
        NewGrid.Grid_AfterEdit .Row, .ColIndex("Count")
    
  
    If lblqty.Caption <> "0" Then
    lblShowQty2.Caption = " " & .TextMatrix(.Row, .ColIndex("Count"))
   Else
  lblShowQty2.Caption = "1"
  End If
    
    End With
End Sub

Private Sub ISButton2_Click()
    With Me.FG
        If .TextMatrix(.Row, .ColIndex("printed")) <> "" Then
                    If TxtAdminLogin.text = SystemOptions.BigUserPw Then
                                  TxtAdminLogin.text = ""
                          Else
                                          If SystemOptions.UserInterface = ArabicInterface Then
                                               MsgBox "·« Ì„ﬂ‰  ⁄œÌ· ﬂ„Ì… Â–… «·’‰› ·«‰Â «—”· «·Ï «·„ÿ»Œ", vbCritical
                                         Else
                                             MsgBox "Can't Delete this Items is Sended already to the kitchen", vbCritical
                                         End If
                             Exit Sub
                      End If
        End If
            
       End With
            
            
            If FG.Rows > 1 Then
                If FG.Rows = 2 Then
                    Me.FG.Clear flexClearScrollable, flexClearEverything
                    NewGrid.CalculteValueAdded Me.FG.Row, , , , True
                  Cala
                Else

                    If Me.FG.Rows > 1 Then
                        If Me.FG.Row <> Me.FG.FixedRows - 1 Then
                            Me.FG.RemoveItem (Me.FG.Row)
                            NewGrid.CalculteValueAdded Me.FG.Row, , , , True
                        End If
                    End If

                 Cala
                End If
            End If
    If Me.TxtModFlg.text = "E" And LblStableID.Caption <> "-1" Then

'        Cmd_Click (2)

    End If
    
End Sub

Private Sub Label1_Click(Index As Integer)

   If Me.TxtModFlg.text = "N" Then
   
   If Index = 9 Then ' „Õ·Ì
   
            If SystemOptions.UserInterface = ArabicInterface Then
                 LBLTable1.Caption = "„Õ·Ì"
             Else
                 LBLTable1.Caption = "Dine In"
             End If
    
        
        LblStableID.Caption = -1
        CboPOSBillType.ListIndex = 4
  
  ElseIf Index = 10 Then ' Œ«—ÃÌ
  
                              If SystemOptions.UserInterface = ArabicInterface Then
                            LBLTable1.Caption = "'ÿ·» Œ«—ÃÌ"
                        Else
                            LBLTable1.Caption = "Take Out"
                        End If
        LblStableID.Caption = -1
        CboPOSBillType.ListIndex = 1
        
       ElseIf Index = 11 Then '  Ê’Ì·
  
                  If SystemOptions.UserInterface = ArabicInterface Then
                            LBLTable1.Caption = "' Ê’Ì·"
                        Else
                            LBLTable1.Caption = "Delivery"
                        End If
                        
        LblStableID.Caption = -1
        CboPOSBillType.ListIndex = 2
         ElseIf Index = 12 Then ' ”Ì«—…
  
                     If SystemOptions.UserInterface = ArabicInterface Then
                            LBLTable1.Caption = "'”Ì«—…"
                        Else
                            LBLTable1.Caption = "Car"
                        End If
                        
        LblStableID.Caption = -1
        CboPOSBillType.ListIndex = 3
   
  End If
        
        
    End If
    
End Sub

Private Sub Label14_Click()

    If Me.TxtModFlg.text = "N" Then
        LBLTable.Caption = Label14.Caption
        LblStableID.Caption = -1
        CboPOSBillType.ListIndex = 1
    End If

End Sub
Sub SetText(StrText As String)
    lblLabel1(0) = StrText & Space(10)
    lblLabel1(1) = lblLabel1(0)
    lblLabel1(0).left = 0
    lblLabel1(1).left = lblLabel1(0).Width
End Sub
Public Function showmessage(Optional speed1 As Double = 0, Optional fontcolor1 As Double = 0 _
, Optional fontsize1 As Double = 0, Optional backcolor1 As Double = 0)
Dim Message As String
Dim speed As Double
Dim fontsize As Double
Dim fontcolor As Double
Dim backcolor As Double
Dim show As Boolean
On Error Resume Next
 getInfoMessage 0, Message, speed, fontsize, fontcolor, backcolor, show
    If show = True Then
    Timer2.Enabled = True
        SetText Message
 'lblLabel1(1).Caption = Message
 If speed1 > 0 Then
 
  Timer2.interval = speed1
  
  Else
 
 Timer2.interval = speed
  End If
  If fontsize1 > 0 Then
  fontsize = fontsize1
  End If
  
  If fontcolor1 > 0 Then
  fontcolor = fontcolor1
  End If
  
  If backcolor1 > 0 Then
  backcolor = backcolor1
  End If
  
lblLabel1(1).fontsize = fontsize
lblLabel1(1).ForeColor = fontcolor
 lblLabel1(1).backcolor = backcolor
  If backcolor = 0 Then
    lblLabel1(1).BackStyle = 0
  Else
    lblLabel1(1).BackStyle = 1
  End If
    Else
    Timer2.Enabled = False
    End If
End Function
'Here is where we do all the work
Public Sub ScrollText()
 Static i As Integer
 Dim k As Integer
 k = i Xor 1 'other label
 'move the label left by one pixel
 lblLabel1(i).left = lblLabel1(i).left - 30
 'other label follows like a train
 lblLabel1(k).left = lblLabel1(i).left + lblLabel1(i).Width
 'if engine is off screen, then make it caboose
 If lblLabel1(k).left = 0 Then i = k: lblLabel1(k).left = Me.Width
 
End Sub

Private Sub Label16_Click()
   If Me.TxtModFlg.text = "N" Then
        LBLTable.Caption = Label16.Caption
        LblStableID.Caption = -1
        CboPOSBillType.ListIndex = 3
    End If
    
End Sub

Private Sub Label18_Click(Index As Integer)
LBLPayVal.Caption = LBLPayVal.Caption & Index

TxtPayedValue.text = LBLPayVal.Caption
End Sub

Private Sub Label19_Click()
FramePay.Visible = False
End Sub

Private Sub lblclaer2_Click()
 
 LBLPayVal.Caption = ""

TxtPayedValue.text = ""

End Sub
Private Sub ChecVAT_Click()
  Dim i As Integer
If Me.TxtModFlg.text <> "R" Then
    If ChecVAT.value = vbChecked Then

        With Me.VatGrid
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = True
            Next i

        End With

    Else

        With Me.VatGrid

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = False
            Next i

        End With

    End If
    RelinVatGrid
    End If
End Sub
Private Sub lbldot1_Click()
LBLPayVal.Caption = lblqty.Caption & "."

TxtPayedValue.text = LBLPayVal.Caption

 

End Sub


Private Sub lbl_Click(Index As Integer)
'lvwMain.Visible = True
'lvwItems.Visible = False

FrameAdmi.Visible = True
TxtAdminLogin.SetFocus
End Sub

Private Sub lblexit_Click(Index As Integer)
FramePay.Visible = False
End Sub

Private Sub LBLPayVal_Change()
TxtPayedValue = val(LBLPayVal)
 TxtRemainValue.text = val(Me.TxtPayedValue.text) - val(Me.TxtNetValue.text)
End Sub

Private Sub Timer2_Timer()
ScrollText
If lblLabel1(0).left + lblLabel1(0).Width <= 0 Then
lblLabel1(0).left = Me.Width
End If
lblLabel1(0).left = lblLabel1(0).left - 100

'    If lblView.backcolor = &HC0E0FF Then
'        lblView.backcolor = &HC0FFFF
'    Else
'        lblView.backcolor = &HC0E0FF
'    End If
    
End Sub


Private Sub Label15_Click()
 
    If Me.TxtModFlg.text = "N" Then
        LBLTable.Caption = Label15.Caption
        LblStableID.Caption = -1
   
        CboPOSBillType.ListIndex = 2
    End If

 
End Sub

Private Sub lBLclr_Click()
    If Me.TxtModFlg.text = "R" Then
'If Me.TxtModFlg.text = "R" And LblStableID.Caption <> "-1" Then
        Cmd_Click (1)

    End If

    lblShowQty2.Caption = "0"
   lblqty.Caption = "0"
End Sub

Private Sub LBLdOT_Click()
    lblqty.Caption = lblqty.Caption & "."

End Sub

Private Sub lBLnO_Click(Index As Integer)

    If Me.TxtModFlg.text = "R" And LblStableID.Caption <> "-1" Then

        Cmd_Click (1)

    End If

    With Me.FG

        If .Row = 0 Then Exit Sub
    End With
 

    lblqty.Caption = lblqty.Caption & Index
  
End Sub

Private Sub lblqty_Change()

    If val(lblqty.Caption) = 0 Then Exit Sub

    With Me.FG
        If .TextMatrix(.Row, .ColIndex("printed")) <> "" Then
        
       MsgBox "·« Ì„ﬂ‰  ⁄œÌ· ﬂ„Ì… Â–… «·’‰› ·«‰Â «—”· «·Ï «·„ÿ»Œ", vbCritical
        Exit Sub
        End If
        .TextMatrix(.Row, .ColIndex("Count")) = val(lblqty.Caption)
        ' .TextMatrix(.Row, .ColIndex("Valu")) = Val(lblqty.Caption) * _
          Val(.TextMatrix(.Row, .ColIndex("Price")))
        'ReLineGrid
        NewGrid.Grid_AfterEdit .Row, .ColIndex("Count")
    
    
    End With
    If lblqty.Caption <> "0" Then
    lblShowQty2.Caption = "«·ﬂ„Ì… " & lblqty.Caption
   Else
  lblShowQty2.Caption = "«·ﬂ„Ì… : 1 "
  End If
  
End Sub

Private Sub lvwItems_ItemClick(Item As vbalListViewLib6.cListItem)
    lvwTables.Visible = False

    If Me.TxtModFlg.text = "R" And LblStableID.Caption <> "-1" Then

        Cmd_Click (1)

    End If

    addrow val(Item.SubItems(2).Caption), Item.text, val(Item.SubItems(1).Caption), val(Item.SubItems(3).Caption)
    If SystemOptions.UserInterface = ArabicInterface Then
        LblSowPrice(0).Caption = " " & val(Item.SubItems(1).Caption)
        lblqty.Caption = ""
        lblShowQty2.Caption = " 1 "
    Else

        LblSowPrice(0).Caption = " Price " & val(Item.SubItems(1).Caption)
        lblqty.Caption = ""
        lblShowQty2.Caption = "1"
    End If

End Sub

Private Sub lvwMain_ItemClick(Item As cListItem)
lvwTables.Visible = False
'lvwMain.Visible = False
'lvwItems.Visible = True

    lblqty.Caption = ""
    lblStatus.Caption = "Clicked Item " & Item.text
    On Error GoTo ErrorHandler
    Dim sInfo As String

    If Not lvwMain.SelectedItem Is Nothing Then

        With lvwMain.SelectedItem
       
            '    sInfo = "Key = " & Item.key & Item.text
            Label4.Caption = "«·«’‰«› «·Œ«’… » " & Item.text
            FillItems (Item.Key)
            
            
        End With
 
    End If

    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description & " [" & Err.Number & "]", vbInformation
    Exit Sub

End Sub
 
Function FillGroups()
    On Error Resume Next
    Dim colX    As cColumn
    Dim itmX    As cListItem
    Dim i       As Long
    Dim j       As Long
    Dim sql     As String
    Dim rs      As New ADODB.Recordset
    Dim Balance As Double
 
    '    sql = " SELECT * from  Groups where GroupID>1  and LastGroup=1"
    sql = " SELECT     dbo.Groups.*"
    sql = sql & " From dbo.Groups"
    sql = sql & " WHERE  POSGroup=1 and    (GroupID IN"
    sql = sql & "                          (SELECT DISTINCT GroupID"
    sql = sql & "                             FROM         dbo.TblItems))"
                             
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        GoTo XGroups
    End If
    Dim xi As Integer
    With lvwMain
        .Visible = False
        .CustomDraw = True
            
        .AutoArrange = True
        '************************
        ImageList1.ImageHeight = 120
        ImageList1.ImageWidth = 120
        '**********************
        ' Set up image lists:
        'Image17.Picture = LoadPicture(App.path & "\images\pos\" & IIf(IsNull(rs("GroupID").value), 0, rs("GroupID").value) & ".JPG")
        'GrouplImageListListImages.Add , "", Image1.Picture
        'ilsIcons16.AddFromFile App.path & "images\pos\" & IIf(IsNull(rs("GroupID").value), 0, rs("GroupID").value) & ".JPG", IMAGE_BITMAP, 0#
       
        'Picture1.Picture = LoadPicture(App.path & "images\pos\" & IIf(IsNull(rs("GroupID").value), 0, rs("GroupID").value) & ".JPG")
        '       ImageList1.ListImages.Add 0, 0, Picture1.Picture

        '.ImageList(eLVSmallIcon) = GrouplImageList ' ilsIcons16
        '.ImageList(eLVTileImages) = GrouplImageList ' ilsIcons48
        '.ImageList(eLVHeaderImages) = GrouplImageList ' ilsIcons16
      
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
        Dim CURRENTPATH As String
        With .Listitems
            rs.MoveFirst
            For i = 0 To rs.RecordCount - 1
                CURRENTPATH = App.path
                If mId(CURRENTPATH, Len(App.path), 1) = "/" Or mId(CURRENTPATH, Len(App.path), 1) = "\" Then
                    CURRENTPATH = mId(CURRENTPATH, 1, Len(CURRENTPATH) - 1)

                End If
                If Dir(App.path & "\images\pos\" & IIf(IsNull(rs("GroupID").value), 0, rs("GroupID").value) & ".JPG") = "" Then
         
                    ImageList1.ListImages.Add , "x" & i, LoadPicture(App.path & "\images\pos\blue.JPG")
                Else
                    ImageList1.ListImages.Add , "x" & i, LoadPicture(App.path & "\images\pos\" & IIf(IsNull(rs("GroupID").value), 0, rs("GroupID").value) & ".JPG")
                End If

                '  lvwMain.ImageList(eLVLargeIcon) = ImageList1  ' ilsIcons32
                lvwMain.ImageList(eLVLargeIcon) = ImageList1
           
                If SystemOptions.UserInterface = ArabicInterface Then
                
                    Set itmX = .Add(, rs("GroupID").value, rs("GroupName").value, i)
                Else
                    Set itmX = .Add(, rs("GroupID").value, rs("GroupNamee").value, i)
           
                End If
      
                '      Set itmX = .Add(, "I" & i, "Test Item " & i, 0, 1)
                If (i Mod 2) = 0 Then
                    itmX.ToolTipText = "This is a test tool tip for item " & i
                End If

                With itmX.SubItems(1)
                    .Caption = DateSerial(year(Now), Rnd * Month(Now) + 1, Rnd * day(Now) + 1)
                    .ShowInTile = ((i Mod 2) = 0)
                    '     .IconIndex = itmX.IconIndex
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

Function FillItems(GroupID As Integer)
    Dim colX    As cColumn
    Dim itmX    As cListItem
    Dim i       As Long
    Dim j       As Long
    Dim sql     As String
    Dim rs      As New ADODB.Recordset
    Dim Balance As Double
    
    sql = " SELECT * from  TblItems where GroupID=" & GroupID
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then

        With lvwItems
            lvwItems.Listitems.Clear
        End With
   
        GoTo XGroups
    End If
   
    With lvwItems
        lvwItems.Listitems.Clear
        .Visible = False
        .CustomDraw = True
            
        .AutoArrange = True
      
        ' Set up image lists:
        
        'ilsIcons32 ' ilsIcons32
        '  ilsIcons32.Clear
         
        '**********************
        Image16.Visible = False
        ImageListItems.ListImages.Clear
        ImageListItems.ImageHeight = 120
        ImageListItems.ImageWidth = 120
        rs.MoveFirst
        'Load Item Image
        Dim Row As Integer
        Row = 1
        Do While Not rs.EOF
            LoadImageToList rs, Row
            Row = Row + 1
            rs.MoveNext
        Loop

        .ImageList(eLVLargeIcon) = ImageListItems
        '**************
        rs.MoveFirst
        '****************

        '***********************
        Dim LngUnitID As Long
        Dim LngItemID As Long

        With .Listitems

            For i = 0 To rs.RecordCount - 1
                If SystemOptions.UserInterface = ArabicInterface Then
                    Set itmX = .Add(, val(rs!ItemID) & "X", rs("ItemName").value, i + 1, ImageListItems.ListImages((i + 1)))
                Else
                    Set itmX = .Add(, val(rs!ItemID) & "X", rs("ItemNamee").value, i + 1, ImageListItems.ListImages((i + 1)))
                End If

                '      Set itmX = .Add(, "I" & i, "Test Item " & i, 0, 1)
                If (i Mod 2) = 0 Then
                    itmX.ToolTipText = "This is a test tool tip for item " & i
                End If

                With itmX.SubItems(1)
                    LngItemID = IIf(IsNull(rs("ItemID").value), 0, rs("ItemID").value)
                    '       .Caption = rs("SallingPrice").value    '  DateSerial(year(Now), Rnd * Month(Now) + 1, Rnd * Day(Now) + 1)
                    
                    GetDefaultItemUnit LngItemID, LngUnitID
                        
                    .Caption = GetItemPrice(LngItemID, 1, LngUnitID)
                        
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
    Dim j As Long

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
 
    sql = " SELECT * from  Stables "
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
   
    If rs.RecordCount = 0 Then
Label1(1).Visible = False

        With lvwTables
            lvwTables.Listitems.Clear
        End With
   
        GoTo XTable
    End If

    With lvwTables
        lvwItems.Listitems.Clear
        .Visible = False
        .CustomDraw = True
            
        .AutoArrange = True
     .BorderStyle = eLVThin
    .ImageList(eLVLargeIcon) = ilsIcons32
        ' Set up image lists:
   '     .ImageList(eLVLargeIcon) = ilsIcons32
'       .ImageList(eLVSmallIcon) = ilsIcons16
        '.ImageList(eLVTileImages) = ilsIcons48
       .ImageList(eLVSmallIcon) = ilsIcons16
 
        '      .Visible = False
        '      .CustomDraw = True
            
        '      .AutoArrange = True
      
        ' Set up image lists:
      
        ' Add column headers
        '      Set colX = .Columns.Add(, "NAME", "Name")
        '      colX.Tag = "Stores the name of the item"
        '      colX.IconIndex = 0
        '      Set colX = .Columns.Add(, "DATE", "Date")
        '      colX.Tag = "Stores the date of the item"
        '      colX.IconIndex = 1
        '      Set colX = .Columns.Add(, "SIZE", "Size")
        '      colX.Tag = "Stores the size of the item"
        '      colX.Alignment = eLVColumnAlignRight
            
        'For i = 1 To 3
        '    .Columns(i).ItemData = i * 100
        ' Next i
  
        With .Listitems
            .Clear

            For i = 1 To rs.RecordCount
If SystemOptions.UserInterface = ArabicInterface Then
                If IsNull(rs("Status").value) Then
                    Set itmX = .Add(, rs("id").value, rs("name").value, 0, 0)
                Else
                    Set itmX = .Add(, rs("id").value, rs("name").value, 0, 0)
                End If
  Else
  
                  If IsNull(rs("Status").value) Then
                    Set itmX = .Add(, rs("id").value, rs("namee").value, 0, 0)
                Else
                    Set itmX = .Add(, rs("id").value, rs("namee").value, 0, 0)
                End If
  End If
                
          
                If (i Mod 2) = 0 Then
                    itmX.ToolTipText = "This is a test tool tip for item " & i
                End If

                With itmX.SubItems(1)
                    .Caption = IIf(IsNull(rs("Status").value), 0, (rs("Status").value))
                    .ShowInTile = ((i Mod 2) = 0)
                    
                    '.IconIndex = itmX.IconIndex
                End With

                With itmX.SubItems(2)
                    .Caption = CLng(Rnd * 1024 * 1024)
                    .ShowInTile = True
                End With

                If (Not IsNull(rs("Status").value)) Then
                    ' test font/colours:
                   
                 itmX.backcolor = vbRed 'RGB(98, 176, 255)
                    itmX.ForeColor = RGB(240, 248, 255)
            
'                      Dim sFnt As New StdFont
'                           sFnt.Name = "Arial"
'                         sFnt.size = 18
                       '  sFnt.Bold = True
       
                   '      itmX.Font = sFnt
                Else
                  itmX.ForeColor = vbBlack
                  
                    itmX.backcolor = vbGreen
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

Private Sub lvwMain_OLEStartDrag(Data As DataObject, _
                                 AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
End Sub

Function CuurentLogdata(Optional Currentmode As String)
   
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·›« Ê—…   " & TxtNoteSerial1.text & CHR(13) & " «· «—ÌŒ " & XPDtbBill.value & CHR(13) & " «·Œ“Ì‰… " & DcboBox.text & CHR(13) & " «·„Œ“‰  " & DCboStoreName.text & CHR(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.text & CHR(13) & "‰Ê⁄ «·”‰œ " & DCDocTypes & CHR(13) & "ÿ—Ìﬁ… «·œ›⁄ " & CboPayMentType & CHR(13) & "‰Ê⁄ «·Œ’„ " & XPCboDiscountType & CHR(13) & "ﬁÌ„… «·Œ’„ " & XPTxtDiscountVal & CHR(13) & "  «·«” Õﬁ«ﬁ " & DtpDelayDate & CHR(13) & " «·⁄„·Â " & DcCurrency & CHR(13) & "—ﬁ„ «·ﬁÌœ " & TxtNoteSerial & CHR(13) & "—ﬁ„ «·ÿ·»Ì… " & TXTOrDer_no
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Bill No " & TxtNoteSerial1.text & CHR(13) & " Date " & XPDtbBill.value & CHR(13) & " Box " & DcboBox.text & CHR(13) & " Store  " & DCboStoreName.text & CHR(13) & " Supplier/Cuxtomer" & DBCboClientName.text & CHR(13) & "Doc Type" & DCDocTypes & CHR(13) & "Payment Type" & CboPayMentType & CHR(13) & "Discount Type  " & XPCboDiscountType & CHR(13) & " Discount Vaalue   " & XPTxtDiscountVal & CHR(13) & "Due Date " & DtpDelayDate & CHR(13) & " Currency " & DcCurrency & CHR(13) & " GE NO" & TxtNoteSerial & CHR(13) & "Order No " & TXTOrDer_no
                           
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 170, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , TxtNoteSerial, TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 170, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , TxtNoteSerial, TxtNoteSerial1
    End If
    
End Function

Function CheckBillType() As Integer
    ' ›Ê„ » ŒœÌœ Â· «·ﬁ« Ê—… «’‰«› «„ Œœ„«  «„ „Ã„⁄ «’‰«ﬁ ÊŒœ„« 
    Dim DblTempItemsGoodType As Double
    Dim DblTempItemsServiceType As Double

    DblTempItemsGoodType = NewGrid.GetItemsTotal(ItemsGoodType)
    DblTempItemsServiceType = NewGrid.GetItemsTotal(ItemsServiceType)

    If DblTempItemsGoodType = 0 And DblTempItemsServiceType > 0 Then  'Œœ„« 
        CheckBillType = 0
    ElseIf DblTempItemsServiceType > 0 And DblTempItemsGoodType > 0 Then ' Ê ·’‰«›   'Œœ„« 
        CheckBillType = 1
    ElseIf DblTempItemsServiceType = 0 And DblTempItemsGoodType > 0 Then 'Ê ·’‰«›   '
        CheckBillType = 2
      
    End If

End Function

Function CheckAccounts() As Boolean
    Dim vchrcode As String
    Dim StrTempAccountCode As String
    Dim usedaccount As Integer

    If BillBasedOn(0).value = True Then
        vchrcode = Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19)

        If vchrcode = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ’—› ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": GoTo ErrTrap
        ElseIf vchrcode = "" Then
            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": GoTo ErrTrap
                       
        End If
                       
    End If
                       
    Dim Account_Code_dynamic As String
 
    If val(Me.LblDiscountsTotal.Caption) > 0 Then
        Account_Code_dynamic = get_account_code_branch(12, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "Branch Not Created ", vbCritical
            End If

            GoTo ErrTrap
        ElseIf Account_Code_dynamic = "NO account" Then

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»    «·Œ’„ «·„”„ÊÕ »Â   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            Else
                MsgBox "Allowance Discount Not Deined in this Branch", vbCritical
            End If

            GoTo ErrTrap
         
        End If
            
    End If

    If val(DCDocTypes.BoundText) > 0 Then
        getDocAccounts val(DCDocTypes.BoundText), , , StrTempAccountCode, , , , , usedaccount

        If StrTempAccountCode = "" And usedaccount = 1 Then
            MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·Œ«’ »«·Œ’„ «·„”„ÊÕ »Â ", vbCritical
            GoTo ErrTrap
        End If
               
    End If

    If val(DCDocTypes.BoundText) > 0 Then
        getDocAccounts val(DCDocTypes.BoundText), StrTempAccountCode, , , , , usedaccount

        If StrTempAccountCode = "" And usedaccount = 1 Then
            MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«»  «·„œÌ‰ «·Œ«’ »«·›« Ê—…  ", vbCritical
            GoTo ErrTrap
        End If
               
    End If

    If ChkInstall.value = vbChecked Then
        
        Account_Code_dynamic = get_account_code_branch(63, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "Branch Not Created ", vbCritical
            End If

            GoTo ErrTrap
        ElseIf Account_Code_dynamic = "NO account" Then

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«» «Ì—«œ«  «· ﬁ”Ìÿ     ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            Else
                MsgBox "   Insatllemts Revenu Not Deined in this Branch", vbCritical
            End If

            GoTo ErrTrap
         
        End If
       
    End If

    '«· √ﬂœ „‰ «Ì—«œ«  «·Œœ„« 
    Dim SngTemp As Double

    SngTemp = NewGrid.GetItemsTotal(ItemsServiceType)

    If SngTemp > 0 Then
        Account_Code_dynamic = get_account_code_branch(23, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox " Branch Not Created", vbCritical
            End If

            GoTo ErrTrap
        Else

            If Account_Code_dynamic = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «Ì—«œ«  «·Œœ„«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "Service Revenue Account not defined in this Branch", vbCritical
                End If

                GoTo ErrTrap
         
            End If
        End If
        
    End If

    Account_Code_dynamic = get_account_code_branch(2, my_branch)
        
    If Account_Code_dynamic = "NO branch" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Else
            MsgBox "Branch Not Created", vbCritical
        End If

        GoTo ErrTrap
    ElseIf Account_Code_dynamic = "NO account" Then

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
        Else
            MsgBox "Sales Account Not Defined in this Branch", vbCritical
        End If

        GoTo ErrTrap
         
    End If
   
    If val(DCDocTypes.BoundText) > 0 Then
        getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount

        If StrTempAccountCode = "" And usedaccount = 1 Then
            MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ·›« Ê—… «·„»Ì⁄« ", vbCritical
            GoTo ErrTrap
        End If
 
    End If

    If val(DCDocTypes.BoundText) > 0 Then
        getDocAccounts val(DCDocTypes.BoundText), StrTempAccountCode, , , , , usedaccount

        If StrTempAccountCode = "" And usedaccount = 1 Then
            MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰  ·›« Ê—… «·„»Ì⁄« ", vbCritical
            GoTo ErrTrap
        End If
 
    End If

    If detect_inventory_work_type = 2 Then
        '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
        Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

        If Account_Code_dynamic = "" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
            GoTo ErrTrap
        End If
    
        If val(DCDocTypes.BoundText) > 0 Then
            getDocAccounts val(DCDocTypes.BoundText), , , , , StrTempAccountCode, , , , , usedaccount

            If StrTempAccountCode = "" And usedaccount = 1 Then
                MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ·”‰œ «·’—› ", vbCritical
                GoTo ErrTrap
            End If
        End If

    End If

    If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then

        Account_Code_dynamic = get_account_code_branch(1, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
        ElseIf Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ  ﬂ·›… «·„»Ì⁄«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
                
        End If
     
        If val(DCDocTypes.BoundText) > 0 Then
            getDocAccounts val(DCDocTypes.BoundText), , , , StrTempAccountCode, , , , , usedaccount

            If StrTempAccountCode = "" And usedaccount = 1 Then
                MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ «·Œ«’ »”‰œ «·’—›", vbCritical
                GoTo ErrTrap
            End If
        End If

    End If

    CheckAccounts = True
    Exit Function
ErrTrap:
    CheckAccounts = False
End Function

Private Sub BillBasedOn_Click(Index As Integer)

    Select Case Index

        Case 1

            If BillBasedOn(1).value = True Then
                
                FillVoucherGrid
                GRID1.Enabled = True
            End If

        Case 2

            If BillBasedOn(2).value = True Then
                
                FillOrderGrid
                GRID2.Enabled = True
            End If

    End Select

End Sub

Private Sub CboPayMentType_Change()
    On Error GoTo ErrTrap

    'If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    If CboPayMentType.ListIndex = 0 Then '‰ﬁœÌ
        XPChkPayType(0).Enabled = False
        XPChkPayType(1).Enabled = False
        XPChkPayType(2).Enabled = False
        XPChkPayType(0).value = Checked
        XPChkPayType(1).value = Unchecked
        XPChkPayType(2).value = Unchecked
        XPTxtValue(0).text = XPTxtSum.text
        XPTxtValue(1).text = ""
        DcboBox.Enabled = True
        Frame1.Visible = True
        DCPaymentNet.Enabled = True
    Else
        XPChkPayType(0).Enabled = True
        XPChkPayType(1).Enabled = True
        XPChkPayType(2).Enabled = True
        XPChkPayType(0).value = Unchecked
        XPChkPayType(1).value = Checked
        XPChkPayType(2).value = Unchecked
        XPTxtValue(1).text = XPTxtSum.text
        XPTxtValue(0).text = ""
        DcboBox.BoundText = ""
        DcboBox.Enabled = False
        Frame1.Visible = False
        DCPaymentNet.Enabled = False
    End If

    'End If
    Exit Sub
ErrTrap:
End Sub

Private Sub CboPayMentType_Click()

  '  If CboPayMentType.ListIndex = 0 Then
  '      DCPaymentNet.BoundText = 1
 '  Else
 '      DCPaymentNet.text = ""
 '   End If

 '   CboPayMentType_Change
 
End Sub

Private Sub ChkInstall_Click()

    If ChkInstall.value = vbChecked Then
        Me.CmdINSTALLMENT.Enabled = True
        XPTxtValue(1).text = LblTotal.Caption
    Else
        Me.CmdINSTALLMENT.Enabled = False

        With Me.FgInstallments
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            'LblPrecenType.Caption = ""
            'LblPrecenValue.Caption = ""
            'LblInstallTotal.Caption = ""
            'LblInstallCount.Caption = ""
            'LblFirstInstallDate.Caption = ""
            'LblInstallmentType.Caption = ""
        End With

    End If

End Sub

Private Sub ChkTaxAdd_Click()

    If ChkTaxAdd.value = Checked Then
        TxtTaxAddValue.Enabled = True
        lbl(39).Enabled = True
        lbl(46).Enabled = True
    Else
        TxtTaxAddValue.text = ""
        TxtTaxAddValue.Enabled = False
        lbl(39).Enabled = False
        lbl(46).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChkTaxSerivce_Click()
    On Error GoTo ErrTrap

    If ChkTaxSerivce.value = Checked Then
        TxtTaxServiceValue.Enabled = True
        lbl(43).Enabled = True
        lbl(47).Enabled = True
    Else
        TxtTaxServiceValue.text = ""
        TxtTaxServiceValue.Enabled = False
        lbl(43).Enabled = False
        lbl(47).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChkTaxStamp_Click()

    If ChkTaxStamp.value = Checked Then
        TxtTaxStampValue.Enabled = True
        lbl(41).Enabled = True
        lbl(48).Enabled = True
    Else
        TxtTaxStampValue.text = ""
        TxtTaxStampValue.Enabled = False
        lbl(41).Enabled = False
        lbl(48).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Function CloseIssueVoucher()
    On Error Resume Next
    Dim i As Integer
    Dim sql As String
 
    If BillBasedOn(1).value = False Then Exit Function

    With GRID1

        For i = 1 To .Rows - 1
     
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
                sql = "update transactions set closed=1" & ",nots=" & val(Me.XPTxtBillID.text) & ",nots2=" & Me.TxtNoteSerial1.text & " where  Transaction_ID= " & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
            Else
                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
               
            End If
       
            Cn.Execute sql
 
        Next
       
    End With
       
End Function

Function DeleteLinkTOIssueVoucher()
    On Error Resume Next
    Dim i As Integer
    Dim sql As String
 
    If BillBasedOn(1).value = False Then Exit Function

    With GRID1

        For i = 1 To .Rows - 1
     
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then

                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID"))) ' & "nots=" & "" & "nots2=" & ""
               
            End If
       
            Cn.Execute sql
 
        Next
       
    End With
       
End Function
Sub printtomanyprinter()
Dim VarSet As Variant
Dim a As String
'Exit Sub
Open App.path & "\printers.txt" For Input As #1
    dbname.Clear

    Do Until EOF(1)
        Line Input #1, a
        'subsequent lines
 
        If a <> "" Then
            VarSet = Split(a, "*", , vbTextCompare)

            If VarSet(0) <> Empty Or VarSet(0) <> "" Then
            
                CBOPrinter.AddItem a
             SetPrinter2 (a)
          printtoAnotherprinter
            DoEvents
            End If
        End If
    
    Loop

    Close #1
    

Dim sql As String
       sql = "update Transaction_Details set printed=1   where  Transaction_ID=" & val(XPTxtBillID.text)
               
       
            Cn.Execute sql
'Exit Sub

End Sub

Sub printtoAnotherprinter()

'-----------------------------------------------------------------------------
    
    Dim intLineCtr          As Integer
    Dim intPageCtr          As Integer
    Dim intX                As Integer
    Dim strCustFileName     As String
    Dim strBackSlash        As String
    Dim intCustFileNbr      As Integer
    
    
    Const intLINE_START_POS As Integer = 0
    Const intLINES_PER_PAGE As Integer = 60
    
    ' Have the user make sure his/her printer is ready ...
 
    
    ' Set the printer font to Courier, if available (otherwise, we would be
    ' relying on the default font for the Windows printer, which may or
    ' may not be set to an appropriate font) ...
 
    For intX = 0 To Printer.FontCount - 1
        If Printer.Fonts(intX) Like "Arabic*" Then
            Printer.FontName = Printer.Fonts(intX)
            Exit For
        End If
    Next
    
    Printer.fontsize = 10
    
    ' initialize report variables ...
    intPageCtr = 0
    intLineCtr = 99 ' initialize line counter to an arbitrarily high number
                    ' to force the first page break
                    
    Dim openingdate As Date
    Dim StrSQL  As String
    Dim rs As New ADODB.Recordset
StrSQL = " SELECT        TOP (100) PERCENT dbo.Transaction_Details.Remarks, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.printed, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, "
StrSQL = StrSQL & "                         dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice AS value, dbo.Transaction_Details.Transaction_ID, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"
StrSQL = StrSQL & " FROM            dbo.Transaction_Details INNER JOIN"
StrSQL = StrSQL & "                         dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"
StrSQL = StrSQL & " WHERE     (dbo.Transaction_Details.printed IS NULL) AND (dbo.Transaction_Details.Transaction_ID = " & val(XPTxtBillID.text) & ")"
StrSQL = StrSQL & " ORDER BY dbo.Transaction_Details.ID "
 
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
     Exit Sub
    End If
 
 
 
    Dim RowNum As Integer
     For RowNum = 1 To rs.RecordCount
         If 1 = 1 Then
        
        If intLineCtr > intLINES_PER_PAGE Then
            GoSub PrintHeadings
        End If
        ' print a line of data
                If SystemOptions.UserInterface = ArabicInterface Then
        Printer.Print Tab(intLINE_START_POS); _
                      IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value); _
                       Tab(20 + intLINE_START_POS); _
                       IIf(IsNull(rs("ShowQty").value), "", rs("ShowQty").value); _
                      Tab(27 + intLINE_START_POS); _
                      IIf(IsNull(rs("UnitName").value), "", rs("UnitName").value);
Else
  Printer.Print Tab(intLINE_START_POS); _
                      IIf(IsNull(rs("ItemNameE").value), "", rs("ItemNameE").value); _
                       Tab(20 + intLINE_START_POS); _
                       IIf(IsNull(rs("ShowQty").value), "", rs("ShowQty").value); _
                      Tab(27 + intLINE_START_POS); _
                      IIf(IsNull(rs("UnitNamee").value), "", rs("UnitNamee").value);
End If
'                      Fg.TextMatrix(RowNum, Fg.ColIndex("Name"));
'Fg.TextMatrix(RowNum, Fg.ColIndex("Name"));
        ' increment the line count
        intLineCtr = intLineCtr + 1
        If intLineCtr = 1 Then Exit Sub
  '  Loop

    ' close the input file
 
 End If
 rs.MoveNext
Next RowNum
     Printer.EndDoc
    
 
    
    Exit Sub


PrintHeadings:
'------------
     If intPageCtr > 0 Then
        Printer.NewPage
    End If
    ' increment the page counter
    intPageCtr = intPageCtr + 1
    
     Printer.Print
'    Printer.Print
'    Printer.Print
'    Printer.Print
    
    ' Print the main headings
    Printer.Print Tab(intLINE_START_POS); _
                  "Print Date: "; _
                  Format$(Date, "mm/dd/yy"); _
                  Tab(intLINE_START_POS + 31); _
                  "NO:"; Me.TxtNoteSerial1.text; _
                  Tab(intLINE_START_POS + 73); _
                  ""; _
                  'Format$(intPageCtr, "@@@")
    Printer.Print Tab(intLINE_START_POS); _
                  "Print Time: "; _
                  Format$(Time, "hh:nn:ss"); _
                  Tab(intLINE_START_POS + 33); _
                  "**Table:" & LBLTable1.Caption
    Printer.Print
    ' Print the column headings
    Printer.Print Tab(intLINE_START_POS); _
                  "item"; _
                  Tab(20 + intLINE_START_POS); _
                  "QTY"; _
                  Tab(27 + intLINE_START_POS); _
                  "Unit";
                   
       
    Printer.Print Tab(intLINE_START_POS); _
                  "------"; _
                  Tab(20 + intLINE_START_POS); _
                  "------"; _
                  Tab(27 + intLINE_START_POS); _
                  "------"; _
                  Tab(34 + intLINE_START_POS); _
                  "------";
    Printer.Print
     intLineCtr = 6
    Return


End Sub

Private Sub Cmd_Click(Index As Integer)
    frmaeDiscount.Visible = False
    Dim AskOption As Boolean
    Dim intDef    As Integer
    Dim Msg       As String
    Dim StrSQL    As String
    Dim RsTest    As ADODB.Recordset
    Dim RsOptions As ADODB.Recordset
    BolPrint = True
    'On Error GoTo ErrTrap
    Timer1.Enabled = False

    lblqty.Caption = ""
    lblShowQty2.Caption = ""
    Select Case Index
        Case 11
            printtomanyprinter
        Case 12
            printtomanyprinter2
        Case 13
            CustomerPrintReport , 1, LBLTable.Caption
 
        Case 9
            PrintReport , 1, LBLTable.Caption, 1
        Case 0
 
            loadInvoices
            lvwItems.Visible = False
            lvwTables.Visible = True
            'End If

            '            If DoPremis(Do_New, Me.Name, True) = False Then
            '                Exit Sub
            '            End If
            LBLTable1.Caption = ""
        
            clear_all Me

            With lvwItems
                lvwItems.Listitems.Clear
            End With
            BillBasedOn(1).Enabled = True
            '       DCboItemsCode.SetFocus
            CboPOSBillType.ListIndex = 4
            LblStableID.Caption = -1
            LBLTable.Caption = ""
            
            ClearNotes
            TxtModFlg.text = "N"
            'XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            SetDefaults
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            VatGrid.Clear flexClearScrollable, flexClearEverything
            VatGrid.Rows = 1
      
            XPTab301.CurrTab = 0
        
            DcCurrency.BoundText = 1
        
            Me.dcBranch.BoundText = Current_branch
      
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
     
            If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
                DcboBox.Enabled = False
                DCboStoreName.Enabled = True
                DcboEmp.Enabled = True
          
                Me.dcBranch.BoundText = userbranchid
                Me.DCboStoreName.BoundText = dstore
                Me.DcboBox.BoundText = dBox
                Me.DcboEmp.BoundText = EmpID
            Else
                dcBranch.Enabled = True
                DcboBox.Enabled = True
                DCboStoreName.Enabled = True
                DcboEmp.Enabled = True
                Me.dcBranch.BoundText = ""
                Me.DCboStoreName.BoundText = dstore
                Me.DcboBox.BoundText = dBox
                Me.DcboEmp.BoundText = EmpID

            End If
          
            BillBasedOn(0).value = True
 
            If Current_branch = 0 Then
                'branch_id = my_branch
                Me.dcBranch.BoundText = Current_branch
            End If
 
            CboPOSBillType.ListIndex = 1
            Cmd(7).Enabled = True
            'Cmd(7).Visible = False  '«Œ›«¡ «·œ›⁄
      
        Case 1

            '            If DoPremis(Do_Edit, Me.Name, True) = False Then
            '                Exit Sub
            '            End If
            '

            '           If SystemOptions.usertype = UserNormal Then
            
            '    Msg = "·Ì” ·ﬂ Õﬁ  ⁄œÌ· ›Ï «·›Ê« Ì—"
            '    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
            '    Exit Sub
            'End If
        
            'If AvailableDeal = True Then
            '«·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
            If XPTxtValue(1).Tag <> "" Then
                StrSQL = "select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
                Set RsTest = New ADODB.Recordset
                RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsTest.EOF Or RsTest.BOF) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "·ﬁœ  „  ﬁ”Ìÿ «·ﬁÌ„ «·¬Ã·… ⁄·Ï Â–Â «·›« Ê—…" & CHR(13)
                        Msg = Msg + " ⁄œÌ· «·›« Ê—… ”ÌƒœÌ ≈·Ï Õ–› Â–Â «·√ﬁ”«ÿ" & CHR(13)
                        Msg = Msg + "Â·  —€» ›Ì  ⁄œÌ· Â–Â «·›« Ê—…ø"
                    Else
                
                        Msg = "this bill was linked With Installment and edit will Delete this Installment Confirm Edit?" & CHR(13)
                    End If

                    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbNo Then
                        Exit Sub
                    End If
                End If
            End If

            '«·√ﬁ”«ÿ «·„”œœ… ⁄·Ï «·›« Ê—…
            If XPTxtValue(1).Tag <> "" Then
                StrSQL = "select * From ReceiptQestForBill where NoteID=" & XPTxtValue(1).Tag
                Set RsTest = New ADODB.Recordset
                RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsTest.EOF Or RsTest.BOF) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "·ﬁœ  „  Õ’Ì· »⁄÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï Â–Â «·›« Ê—…" & CHR(13)
                        Msg = Msg + "Ê·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰« Â«" & CHR(13)
                        Msg = Msg + "≈–« ﬂ‰   —€» ›Ì  ⁄œÌ· »Ì«‰«  Â–Â «·›« Ê—…" & CHR(13)
                        Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «· Õ’Ì· «·Œ«’… »Â«"
                    Else
                        Msg = "Some premiums were collected  on this bill You Must delete Collected  premiums according to this bill First" & CHR(13)
                    End If

                    MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            End If

            '⁄„·Ì«  «·’Ì«‰… «·„— »ÿ… »«·›« Ê—…
            '            StrSQL = "select * From MaintenanceJuncTransaction where Transaction_ID=" & Trim(XPTxtBillID.text)
            '            Set RsTest = New ADODB.Recordset
            '            RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            '            If Not (RsTest.EOF Or RsTest.BOF) Then
            '                If SystemOptions.UserInterface = ArabicInterface Then
            '                    Msg = "·ﬁœ  „ ≈Ã—«¡ »⁄÷ ⁄„·Ì«  «·’Ì«‰… ⁄·Ï Â–Â «·›« Ê—… Ê·« Ì„ﬂ‰  ⁄œÌ·Â«"
            '                    Msg = Msg + "≈–« ﬂ‰   —€» ›Ì  ⁄œÌ· »Ì«‰«  Â–Â «·›« Ê—…" & Chr(13)
            '                    Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «·’Ì«‰… «·Œ«’… »Â«"
            '                Else
            '                    Msg = "this Bill Linked with Maintenance Operation You must Delete This Operation First"
            '
            '                End If

            '                MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '                Exit Sub
            '            End If

            '         Me.Retrive Val(Me.XPTxtBillID.text)
             
            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id
            CuurentLogdata

            '    txtorder_no_Change
            'End If
        Case 2

            If CboPOSBillType.ListIndex <> 4 And SAVESTATUS = False Then
                ' Cmd_Click 7
                FramePay.Visible = True
 
                FillGridWithData
 
                ReLineGrid
                FrmCustomerDisplay.LblInformation.Caption = " Total " & vbNewLine & TxtNetValue.text

                If 1 = 1 Then

                    LBLPayVal.Caption = TxtNetValue.text
 
                    With Grid
                        .TextMatrix(1, .ColIndex("Value")) = LBLPayVal.Caption
                    End With
                    ReLineGrid
   
                End If
     
                FramePay.Visible = True
                Exit Sub
            End If
 
            CboPayMentType.ListIndex = 0

            'FramePay
            If CboPayMentType.ListIndex = 0 Then

                If val(TxtRemainValue.text) < 0 Then
                    If SystemOptions.UserInterface = EnglishInterface Then
                        Msg = "Enter Correct Payed Value"
                    Else
                        Msg = "  ﬁÌ„Â «·„œ›Ê⁄ €Ì— ’ÕÌÕÂ "
                    End If
             
                    'MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
  
                    ' Exit Sub
                End If
            End If

            ' If CboPayMentType.ListIndex = 1 And XPChkPayType(0).value = Unchecked And XPChkPayType(2).value = Unchecked Then
            '     XPTxtValue(1).Text = LblTotal.Caption
            ' End If
 
            Set RsNotesGeneral = New ADODB.Recordset
            '   RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
         
            StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
            RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
       
            '    my_branch = Me.Dcbranch.BoundText
      
            'If Me.TxtModFlg.text = "N" Then
             
            ' End If

            'xxxxxxxxx
            If Trim(LblStableID.Caption) = -1 And CboPOSBillType.ListIndex = 4 Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "SpecifY Table  "
                Else
                    Msg = "Õœœ „Êﬁ⁄     «Ê·« "
                End If
             
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Timer1.Enabled = True
                '  DCPaymentNet.SetFocus
                '  SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
    
            my_branch = Me.dcBranch.BoundText
            SaveData
            
            ' Unload customer_screen
            '  Load customer_screen
            '  customer_screen.Show
        
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            If SystemOptions.usertype = UserNormal Then
                Msg = "·Ì” ·ﬂ Õﬁ Õ–› ›Ï «·›Ê« Ì—"
                MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If
   
            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            If m_FrmSearch Is Nothing Then
                Set m_FrmSearch = New FrmBuySearch
                m_FrmSearch.DealingForm = InvoiceTransaction
                m_FrmSearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄"
                Set m_FrmSearch.RetrunFrm = Me
                m_FrmSearch.show vbModal
            Else
                Msg = "Â‰«ﬂ ‘«‘… »ÕÀ Œ«’… »‘«‘… ›« Ê—… «·»Ì⁄ «·Õ«·Ì…"
                Msg = Msg & CHR(13) & "Ÿ«Â—… «„«„ﬂ ›⁄·«...·«Ì„ﬂ‰ ⁄—÷ «ﬂÀ— „‰ ‘«‘… »ÕÀ ·ﬂ· ‘«‘… ›« Ê—…"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                m_FrmSearch.ZOrder 0
                'm_FrmSearch.SetFocus
            End If

        Case 7
        
            FillGridWithData

            RelinVatGrid
            ReLineGrid
            FrmCustomerDisplay.LblInformation.Caption = " Total " & vbNewLine & TxtNetValue.text

            If 1 = 1 Then

                LBLPayVal.Caption = TxtNetValue.text
 
                With Grid
                    '.TextMatrix(1, .ColIndex("Value")) = LBLPayVal.Caption
                End With
                ReLineGrid
   
            End If
     
            FramePay.Visible = True

            'LBLPayVal.Caption = ""
        Case 6
            Unload Me

        Case 10
            ShowGL_cc TxtNoteSerial.text, , 200, val(Me.TXTNoteID.text)
            'ShowGL_cc TxtNoteSerial.text, , 200
        Case 8
            
            ' BtnUndo_Click
            CashierLogout.show
            Unload Me
            'End
    End Select

    Exit Sub
ErrTrap:
End Sub
Function loadLogo()
    Dim rs As ADODB.Recordset
    Dim BolShowLogo As Boolean
    Dim xLogo As CRAXDRT.OLEObject
    Dim StrFileName As String
    Dim MsgErr As String

     

    Set rs = New ADODB.Recordset
    rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

    If rs.BOF Or rs.EOF Then
       
        Exit Function
    End If

'    If SystemOptions.WorkWithBranchLogo = False Then
'        LoadPictureFromDB Nothing, rs, "CompanyLogo", StrFileName
'      Else
'         LoadPictureFromDB Nothing, rs, "branchLogo", StrFileName
'
'      End If

   If Not (IsNull(rs("CompanyLogo").value)) Then
        'LoadPictureFromDB ImgPic, rs, "CompanyLogo"
        LoadPictureFromDB Image16, rs, "CompanyLogo"
        
    End If
    
End Function

Function Retrive_orders_data(Transaction_ID As Integer)
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Integer
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & Transaction_ID

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        row_count = FG.Rows
    
        If FG.TextMatrix(row_count - 1, FG.ColIndex("Code")) = "" Then
            row_count = row_count - 1
        End If
     
        FG.Rows = RsDetails.RecordCount + row_count

        For Num = row_count To FG.Rows - 1 'RsDetails.RecordCount
    
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no")), "", (RsDetails("order_no").value))
            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate")), "", (RsDetails("OrderArrivalDate").value))
            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
        
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
        
            '   FG.TextMatrix(Num, FG.ColIndex("Count")) = items_qty_not_recieved_in_order(FG.TextMatrix(Num, FG.ColIndex("Code")), FG.TextMatrix(Num, FG.ColIndex("order_no")))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
        
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("SallingPrice")), "", (RsDetails("SallingPrice").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext
            ' Debug.Print Num
            ' If FG.Rows > 10 Then
            '     If Num = 8 Then FG.Refresh
            ' End If
        Next Num

    End If

End Function
 
Private Sub Cmd1_Click()
    On Error Resume Next

    If TxtNoteSerial1.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
         
            MsgBox "·«»œ „‰ «Õ Ì«—  ”‰œ  «Ê·«": Exit Sub
        Else
            MsgBox "Select Voucher Firstly": Exit Sub
        End If
 
    End If

    Unload imaged
    imaged.show

    If SystemOptions.UserInterface = EnglishInterface Then

        imaged.Label9.Caption = "Sales Invoice  #"
        imaged.Caption = "Sales Invoice  Attachment"
        imaged.txtopeation_type = "1001"
        imaged.SUBJECT_NO = TxtNoteSerial1.text
        imaged.Label6.Caption = "Sales Invoice  #"
    Else

        imaged.Label9.Caption = "„—›ﬁ«  ›« Ê—… «·»Ì⁄ —ﬁ„"
        imaged.Caption = "„—›ﬁ«  ›« Ê—… «·»Ì⁄ —ﬁ„    "
        imaged.txtopeation_type = "1001"
        imaged.SUBJECT_NO = TxtNoteSerial1.text
        imaged.Label6.Caption = "„—›ﬁ«  ›« Ê—… «·»Ì⁄ —ﬁ„"

    End If

    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type ='1001'  and subject_no='" & TxtNoteSerial1.text & "'"
    imaged.Adodc1.Refresh

    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If

End Sub

Private Sub CmdCash_Click(Index As Integer)

    Select Case Index

        Case 0

        Case 1
    End Select

End Sub

Private Sub cmdCommand1_Click()
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdInfo_Click()
    Dim xPoint As POINTAPI
    
    mdifrmmain.MnuInvInsertTemp.Visible = True
    
    'mdifrmmain.MnuInvSales_Mnu4.Enabled = Me.CmdNotes.Visible
    

    'ClientToScreen Me.CmdInfo.hwnd, xPoint
    'Me.PopupMenu MDIFrmMain.MnuInvSales, , (xPoint.X * Screen.TwipsPerPixelX), (xPoint.Y * Screen.TwipsPerPixelY)
    'Me.PopupMenu MDIFrmMain.MnuInvSales, vbPopupMenuRightAlign + vbPopupMenuRightButton, (xPoint.X * Screen.TwipsPerPixelX), (xPoint.Y * Screen.TwipsPerPixelY)

End Sub

Private Sub CmdINSTALLMENT_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim i As Integer
    XPTxtValue(1).text = LblTotal.Caption
    'If Me.TxtModFlg = "R" Then Exit Sub

    If XPTxtValue(1).text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬁÌ„… «·¬Ã·… ﬁ»·  ”ÃÌ· «·√ﬁ”«ÿ"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

        If XPTxtValue(1).Enabled = True Then
            XPTxtValue(1).SetFocus
        End If

        Exit Sub
    End If

    Load FrmInstallMent
    Set FrmInstallMent.Frm = Me

    With FrmInstallMent

        If Me.TxtModFlg.text = "E" Then
            .Tag = "E"
        
            .Retrive val(XPTxtValue(1).Tag)
            .Txt(1).text = XPTxtValue(1).text
        ElseIf Me.TxtModFlg.text = "R" Then
  
            .Tag = "R"
            .Retrive val(XPTxtValue(1).Tag)
            '              .OptInt(1).value = True
            '.Txt(7).text = 1
            '.Txt(5).text = 12
        Else
            .Tag = "N"
            .Txt(1).text = XPTxtValue(1).text
            Me.CmdINSTALLMENT.Enabled = True
    
            .LblNoteID.Caption = XPTxtSerial(1).text
            '.CboPrecenType.ListIndex = val(Me.LblPrecenType.Tag)
            '.Txt(3).Text = val(LblPrecenValue.Caption)
            '.Txt(5).Text = val(LblInstallCount.Caption)
            .OptInt(1).value = True
            .Txt(7).text = 1
            .Txt(5).text = 12

           ' If IsDate(Me.LblFirstInstallDate.Caption) Then
           '     .Dtp_First.value = Me.LblFirstInstallDate.Caption
           ' End If

            '        .Txt(7).text = Val(LblInstallSeprator.Caption)
           ' If val(LblInstallmentType.Tag) = 0 Then
                '        .OptInt(0).value = True
           ' ElseIf val(LblInstallmentType.Tag) = 1 Then
                '        .OptInt(1).value = True
           ' ElseIf val(LblInstallmentType.Tag) = 2 Then
           '     '        .OptInt(2).value = True
'            End If
'
            With .FG
                .Rows = Me.FgInstallments.Rows

                For i = 1 To Me.FgInstallments.Rows - 1
                    .TextMatrix(i, .ColIndex("Serial")) = i
                    .TextMatrix(i, .ColIndex("Value")) = Me.FgInstallments.TextMatrix(i, Me.FgInstallments.ColIndex("Value"))
                    .TextMatrix(i, .ColIndex("Due_Date")) = Me.FgInstallments.TextMatrix(i, Me.FgInstallments.ColIndex("Due_Date"))
                Next i

                .AutoSize 0, .Cols - 1, False
            End With

        End If

        .show vbModal
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdInvProfit_Click()

    If SystemOptions.SysMainStockCostMethod = LastPurPriceType Or SystemOptions.SysMainStockCostMethod = ModernWeightAverage Then
        NewGrid.ShowInvProfDialog
    End If

    'If Me.TxtModFlg.Text = "R" Then
    '
    'Else
    '    NewGrid.ShowInvProfDialog
    'End If
End Sub

Private Sub CmdNotes_Click()
    ShowRelatedNotes val(Me.XPTxtBillID.text), 1
End Sub

Private Sub CmdNotes_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    Dim StrTemp As String

    If val(Me.CmdNotes.Tag) = 0 Then
        Me.CmdNotes.ToolTipText = ""
    Else
        StrTemp = " ÊÃœ ⁄·Ï Â–Â «·Õ—ﬂ… ⁄„·Ì«  „«·Ì… „ﬁœ«—Â« : " & val(Me.CmdNotes.Tag)
        Me.CmdNotes.ToolTipText = StrTemp
    End If

End Sub

Private Sub CmdRetruns_Click()
    ShowRelatedTransactions val(Me.XPTxtBillID.text), 1
End Sub

Private Sub CmdRetruns_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    Dim StrTemp As String

    If val(Me.CmdRetruns.Tag) = 0 Then
        Me.CmdRetruns.ToolTipText = ""
    Else
        StrTemp = " ÊÃœ ⁄·Ï Â–Â «·Õ—ﬂ… Õ—ﬂ«   Ã«—Ì… √Œ—Ï ·Â« ⁄·«ﬁ… »Â« ≈Ã„«·ÌÂ«: " & val(Me.CmdRetruns.Tag)
        Me.CmdRetruns.ToolTipText = StrTemp
    End If

End Sub

Private Sub CmdSearch_Click()
    'Dim LngItemID As Long
    'Dim LngStoreID As Long
    'LngItemID = Val(Me.DCboItemsName.BoundText)
    'LngStoreID = Val(Me.DCboStoreName.BoundText)
    'If LngItemID = 0 Or LngStoreID = 0 Then
    '    Exit Sub
    'End If
    'Load FrmSerialList
    'FrmSerialList.RetrunType = 1
    'Set FrmSerialList.m_TextBox = Me.TxtSerial
    'FrmSerialList.GetData LngItemID, LngStoreID
    'FrmSerialList.Show vbModal
End Sub

Private Sub Command1_Click()
    Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    On Error GoTo ErrTrap

    If Text1.text <> "" Then
        Msg = " „  ÕÊÌ· Â–… «·›« Ê—… „‰ ﬁ»· Ê·« Ì„ﬂ‰  ÕÊÌ·Â« „—… «Œ—Ï  …  "
        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set Frm = New FrmOut

    With Frm

        .Convert
        '    .XPTxtBillID.Text = XPTxtBillID.Text
        .XPDtbBill.value = XPDtbBill.value
        .DBCboClientName.BoundText = DBCboClientName.BoundText
        .DCboStoreName.BoundText = DCboStoreName.BoundText
        .Text2.text = TxtTransSerial.text
        .CboPayMentType.ListIndex = CboPayMentType.ListIndex

        For RowNum = 1 To FG.Rows - 1

            If .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) <> "" Then
                .FG.Rows = .FG.Rows + 1
            End If

            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Name")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Name")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Code")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("ItemCase")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("HaveSerial")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Count")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Count")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Count")))
            ' .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Price")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Price")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(.FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod)
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("DiscountType")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")))
            Dim StrSQL As String
            Dim RsUnit As New ADODB.Recordset
            StrSQL = "SELECT TOP 100 PERCENT dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial,dbo.Transactions.Transaction_Type FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites INNER JOIN dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID AND dbo.Transaction_Details.Item_ID = dbo.TblItemsUnits.ItemID WHERE (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "') AND (dbo.Transactions.Transaction_Type = 21) AND (dbo.TblItemsUnits.ItemID = " & FG.TextMatrix(RowNum, FG.ColIndex("Code")) & ") ORDER BY dbo.TblItemsUnits.SecOrder"
            Set RsUnit = New ADODB.Recordset
            RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
            .FG.Cell(flexcpData, RowNum, .FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").value))
            .FG.TextMatrix(RowNum, .FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").value))

            '        FG.Cell(flexcpData, I, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").Value))
            '        FG.TextMatrix(I, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").Value))
            '           StrSQL = "SELECT dbo.Transactions.Transaction_Type, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID WHERE (dbo.Transactions.Transaction_Type = 19) AND (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "')"
            '        .FG.Cell(flexcpData, .FG.Rows - 1, FG.ColIndex("UnitID")) = 1 'FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) ' IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").Value))
            '        .FG.TextMatrix(.FG.Rows - 1, FG.ColIndex("UnitID")) = "Ã—«„" 'FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) ' IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").Value))

        Next RowNum

        .Cala
    End With

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Function CREATE_VOUCHER_GE(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchID As Integer)
    Dim usedaccount As Integer
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim TOTAL_COST As Double
    Dim LngCurItemID As Integer
    Dim LngUnitID As Long
    Dim UnitFactor As Double

    With FG

        For i = 1 To FG.Rows - 1

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("ItemType"))) <> 1 Then
                LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
                LngUnitID = val(FG.Cell(flexcpData, i, FG.ColIndex("UnitID")))
            
                GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
                    
                TOTAL_COST = TOTAL_COST + (FG.TextMatrix(i, FG.ColIndex("Count")) * ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , LngUnitID))
            End If

        Next i

    End With

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·œ«∆‰
    SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)
    my_branch = BranchID

    If TOTAL_COST > 0 Then
   
        If detect_inventory_work_type = 1 Then

            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            If val(DCDocTypes.BoundText) > 0 Then
                getDocAccounts val(DCDocTypes.BoundText), , , , , StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ·”‰œ «·’—›", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                ElseIf usedaccount = 0 Then
                    StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
                End If

            Else
                StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            End If
            
            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            StrTempDes = "”‰œ ’—› —ﬁ„ " & Me.TxtTransSerial.text
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            If val(DCDocTypes.BoundText) > 0 Then
                getDocAccounts val(DCDocTypes.BoundText), , , , , StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ·”‰œ «·’—›", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                ElseIf usedaccount = 0 Then
                    StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
                End If

            Else
                StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            End If

            '            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰
            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "”‰œ    ’—› —ﬁ„ " & TxtNoteSerial1V
            Else
                StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With FG

                For i = 1 To FG.Rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "”‰œ    ’—› —ﬁ„ " & TxtNoteSerial1V
                        Else
                            StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                        End If
            
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·„œÌ‰
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)

        If TOTAL_COST > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then

                Account_Code_dynamic = get_account_code_branch(1, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ  ﬂ·›… «·„»Ì⁄«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If

                If val(DCDocTypes.BoundText) > 0 Then
                    getDocAccounts val(DCDocTypes.BoundText), , , , StrTempAccountCode, , , , , usedaccount

                    If StrTempAccountCode = "" And usedaccount = 1 Then
                        MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ «·Œ«’ »”‰œ ’—› «·„Ê«œ", vbCritical
                        GoTo ErrTrap
                    ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                    ElseIf usedaccount = 0 Then
                        StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄« 
        
                    End If

                Else
                    StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄« 
                End If
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ    ’—› —ﬁ„ " & TxtNoteSerial1V
                Else
                    StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
         
            ElseIf detect_inventory_work_type = 3 Then

                With FG

                    For i = 1 To FG.Rows - 1

                        If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                            groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 1)

                            '  groupAccount = get_item_group_account_inventory(FG.TextMatrix(I, FG.ColIndex("Code")), DCboStoreName.BoundText, 4)
                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»    ﬂ·›… «·„»Ì⁄«    ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "”‰œ    ’—› —ﬁ„ " & TxtNoteSerial1V
                            Else
                                StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                            End If
            
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If

                    Next i

                End With

            End If
        End If
    End If

    Dim StrSQL  As String
    StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
    Cn.Execute StrSQL
ErrTrap:
End Function

Private Sub CreateIssueVoucher()
    'On Error GoTo errortrap
    'DeleteTransactiomsVoucher Val(Text1.text)

    If BillBasedOn(1).value = True Then Exit Sub

    If CheckBillType = 0 Then ' Œœ„« 
        Exit Sub
    ElseIf CheckBillType = 1 Then ' Ê«’‰«›  ' Œœ„« 

    ElseIf CheckBillType = 2 Then ' «’‰«›

    End If

    Dim i As Long
    Dim LngCurItemID As Integer
    Dim LngUnitID As Long
    Dim UnitFactor As Double

    '›Ì Õ«·… «·«‰ «Ã «·‰„ÿÌ
    If SystemOptions.TypicalProduction = True Then
        GoTo ll
    End If

    With FG

        For i = 1 To FG.Rows - 1

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("ItemType"))) <> 1 Then
                                      
                LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
                LngUnitID = val(FG.Cell(flexcpData, i, FG.ColIndex("UnitID")))
            
                GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
                    
                'TOTAL_COST = TOTAL_COST + (FG.TextMatrix(i, FG.ColIndex("Count")) * UnitFactor * ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod))
                    
                If ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , LngUnitID) = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ  ﬂ·›Â «·»Ì⁄ ·Â Ê·„ Ì „  ÕœÌœ À„‰ «·‘—«¡ Ê·Ì” ·Â ﬁÌ„Â —’Ìœ «›  «ÕÌ… ·–·ﬂ ·« Ì„ﬂ‰ «‰‘«¡ ”‰œ «·’—› "
                    Else
                        MsgBox "Item in line no " & i & "Have No Qty "
                    End If
 
                    With Me.GRID1
                        .Rows = .FixedRows
                        .ExtendLastCol = True
                        .RowHeightMin = 300
                        .Editable = flexEDKbdMouse
                        .ExplorerBar = flexExSortShowAndMove

                        '    .AutoSize 0, .Cols - 1, False
                    End With

                    Text1.text = ""
                    'Cn.Execute "UPDATE Transactions SET NOTS='" & "" & "' WHERE Transaction_ID=" & Val(Me.XPTxtBillID.text)
                    Text1_Change

                    Exit Sub
                End If
            End If

        Next i

    End With

ll:
    Dim groupAccount  As String

    If detect_inventory_work_type = 3 Then
   
        With FG

            For i = 1 To FG.Rows - 1

                If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
                
                    ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                    groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                    If groupAccount = "Error" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                        Else
                            MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                        End If

                        Exit Sub
                    End If
                End If

            Next i

        End With

    End If

    Dim MYWAER As String
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim MYinvnum As String
    Dim note_id As Long

    Dim RSTransDetails As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim RowNum As Integer
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    'Dim Note_ID As Long
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
 
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    Dim MYTEXT As String
    '>>>>>>>>>>>>>>>>>>>>>>>>>

    'rs.Close
    '19 «–‰ ’—›
    '        rs.Open "select * from Transactions where nots =' " & XPTxtBillID.text & "' and Transaction_type = 19"
    '       If rs.RecordCount > 0 Then
    '        If rs!nots <> "" Then
    '        If SystemOptions.UserInterface = ArabicInterface Then
    '             Msg = "·ﬁœ  „  ÕÊÌ· Â–… «·›« Ê—… «·Ï «–‰ ’—›    .."
    '            Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  ÕÊÌ·… „—… «Œ—Ï  ..!!"
    '        Else
    '          Msg = "This bill already converted"
    '        End If
    '          MsgBox Msg, vbOKOnly, App.Title
    '        Exit Sub
    '        End If
        
    '        End If

    '        rs.Close
    '21 ›« Ê—… „»Ì⁄« 
    '        rs.Open "select * from Transactions where Transaction_ID = " & XPTxtBillID.text & " and Transaction_type = 21"

    '        If SystemOptions.UserInterface = ArabicInterface Then
    '        Msg = "”Ê› Ì „ «‰‘«¡ «–‰ ’—› „‰ Â–… «·›« Ê—…   .."
    '        Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
    '        Else
    '        Msg = "Create ISSUE Voucher to this bill ?"
    '        End If
    '  On Error GoTo ErrTrap
    Dim xyeas As Boolean
    xyeas = True

    If xyeas = True Then
 
        MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=19"))
        'mytext = TxtTransSerial.text

        '         rs!nots = mytext
        '         rs.update

        Dim Transaction_ID As Long
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        Text1.text = Transaction_ID
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
        Dim TxtNoteSerialV As String
        Dim TxtNoteSerial1V As String
            
        my_branch = Me.dcBranch.BoundText

        If TxtNoteSerialV = "" Then
            If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                       
                If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                Else
                    TxtNoteSerialV = Notes_coding(val(my_branch), XPDtbBill.value)
                End If
            End If
        End If
        
        If TxtNoteSerial1V = "" Then
            If Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ’—› ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19) = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                    TxtNoteSerial1V = Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19)
                End If
            End If
        End If
             
        If SystemOptions.TypicalProduction = True Then
            TxtNoteSerialV = ""
 
        End If
 
        If Trim(CurrentVoucherNo) <> "" And DateChanged <> True Then
            TxtNoteSerialV = CurrentVoucherNo '—ﬁ„ «·ﬁÌœ
            TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
        End If

        Dim sql As String

        sql = "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,nots2,NoteSerial,NoteSerial1,NoteId,BranchId,Closed)SELECT " & Transaction_ID & "," & MYTEXT & ",Transaction_Date,Transaction_Type = 19,CusID,StoreID,UserID,Emp_ID,nots=" & val(XPTxtBillID.text) & ",nots2=" & TxtNoteSerial1.text & " ,NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId,1From Transactions Where  Transaction_ID =" & val(XPTxtBillID.text) & " And Transaction_Type = 21"
        Cn.Execute sql
        '
        Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,OldQty,OldCost,NewQty,NewCost)SELECT  costprice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, costprice/ QtyBySmalltUnit ,ColorID,ItemSize, UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID ,OldQty,OldCost,NewQty,NewCost  From dbo.Transaction_Details Where SavedItemType=0 and   Transaction_ID = " & XPTxtBillID.text
        Text1.text = Transaction_ID
           UpdateTransactionsCost CStr(Transaction_ID)
           
        'TxtIssueSerial.text = TxtNoteSerial1V
 
        StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
        Cn.Execute StrSQL

        If SystemOptions.TypicalProduction = True Then
            Exit Sub
        End If

        'Create big notes
        Set RsNotesGeneral = New ADODB.Recordset
      '  RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
  StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
        If Me.TxtModFlg.text = "N" Then
    
        Else
        
            general_noteid = val(TXTNoteID.text)
        End If
        
        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        general_noteid = RsNotesGeneral("NoteID").value
        TXTNoteID.text = general_noteid
        RsNotesGeneral("Transaction_ID").value = Transaction_ID
        RsNotesGeneral("NoteDate").value = XPDtbBill.value
        RsNotesGeneral("NoteType").value = 180
        RsNotesGeneral("Note_Value").value = Null
        RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
        RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(10) '«–‰ wvt
        RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
        RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
        RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        RsNotesGeneral.update
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

        CREATE_VOUCHER_GE Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, val(Me.dcBranch.BoundText)

    End If
 
    '
 
ErrTrap:

End Sub

Private Sub Command3_Click()
    FrmSearchSerial.show vbModal
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.show vbModal
        FrmCustemerSearch.SearchType = 2
    End If

End Sub

Private Sub DBCboClientName_MouseUp(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)

    If Button = vbRightButton Then
        mdifrmmain.MnuCusTools.Tag = Me.DBCboClientName.BoundText
        Me.PopupMenu mdifrmmain.MnuCusTools
    End If

End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 5
        FrmItemSearch.show vbModal
    End If

    If KeyCode = vbKeyF9 Then
                    
        FrmSearchSerial.XPTxtCode.text = DCboItemsCode.text
        FrmSearchSerial.show
        FrmSearchSerial.Cmd_Click (0)
                    
    End If

End Sub

Private Sub DCboItemsName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF9 Then
                    
        FrmSearchSerial.XPTxtCode.text = DCboItemsCode.text
        FrmSearchSerial.show
        FrmSearchSerial.Cmd_Click (0)
                    
    End If

End Sub

Private Sub Dcbranch_Change()

    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
        Dcombos.GetDocTypebyid Me.DCDocTypes, 21, val(Me.dcBranch.BoundText)
    End If

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    Dcbranch_Change
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub DcCurrency_Change()

    If Me.TxtModFlg.text = "" Or Me.TxtModFlg.text = "R" Then Exit Sub
    If Me.DcCurrency.BoundText <> "" Then
        txt_Currency_rate.text = get_currency_rate(Me.DcCurrency.BoundText)
    Else
        txt_Currency_rate.text = 1
    End If

End Sub

Private Sub DcCurrency_Click(Area As Integer)
    DcCurrency_Change
End Sub

Private Sub DCPaymentNet_Click(Area As Integer)
'frmmangerlogon.show vbModal
    If val(DCPaymentNet.BoundText) <> 1 Then
    '    DcboBox.text = ""
        
    End If

End Sub

Function FillOrderGrid()
    ' ⁄»∆… «Ê«„— «·‘—«¡ Ê «·»Ì⁄

    With Me.GRID2
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset
    My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_Type=6 and NOT(ORDER_NO IS NULL) AND CLOSED= 0 and   dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText)

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.GRID2
        .Rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("closed").value), 0, RsExp.Fields("closed").value)
         
                .TextMatrix(i, .ColIndex("order_no")) = IIf(IsNull(RsExp.Fields("order_no").value), "", RsExp.Fields("order_no").value)
               
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsExp.Fields("Transaction_Date").value), "", RsExp.Fields("Transaction_Date").value)
           
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsExp.Fields("CusName").value), "", RsExp.Fields("CusName").value)
                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsExp.Fields("Transaction_ID").value), "", RsExp.Fields("Transaction_ID").value)

                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    GRID2.Visible = True

End Function

Function FillVoucherGrid()
    ' ⁄»∆…  ”‰œ«   «·’—›
    On Error Resume Next

    With Me.GRID1
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset

    'My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_Type=19   and   dbo.TblCustemers.CusID=" & Val(DBCboClientName.BoundText)
    If BillBasedOn(0).value = True Then
        My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1,dbo.Transactions.NoteSerial, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where   ( (nots='" & Me.XPTxtBillID.text & "' and  Transaction_Type=19)   and  (dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText) & ")) "
    Else
        My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1,dbo.Transactions.NoteSerial, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where   ( (nots='" & Me.XPTxtBillID.text & "' and  Transaction_Type=19) or ( Transaction_Type=19   and  closed =0 ) and  (dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText) & ")) "
    End If
 
    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.GRID1
        .Rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
             
                .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("closed").value), 0, RsExp.Fields("closed").value)
              
                .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(RsExp.Fields("NoteSerial").value), "", RsExp.Fields("NoteSerial").value)
              
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
               
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsExp.Fields("Transaction_Date").value), "", RsExp.Fields("Transaction_Date").value)
           
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsExp.Fields("CusName").value), "", RsExp.Fields("CusName").value)
                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsExp.Fields("Transaction_ID").value), "", RsExp.Fields("Transaction_ID").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("P1")) = "⁄—÷ «·”‰œ"
                    .TextMatrix(i, .ColIndex("P2")) = "ÿ»«⁄Â  «·ﬁÌœ"
                Else
                    .TextMatrix(i, .ColIndex("P1")) = "View VCHR"
                    .TextMatrix(i, .ColIndex("P2")) = "Print GE"
                End If

                RsExp.MoveNext
            Next
       
        End If
         
        .RowHeight(-1) = 300
    End With

    RsExp.Close
    GRID1.Visible = True

End Function

Private Sub Ele_DblClick(Index As Integer)
    On Error GoTo ErrTrap

    If Index = 9 Then
        If Me.WindowState = vbNormal Then
            Me.WindowState = vbMaximized
        Else
            Me.WindowState = vbNormal
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Ele_KeyUp(Index As Integer, _
                      KeyCode As Integer, _
                      Shift As Integer)

    If Me.TxtModFlg.text = "R" And Not (Me.ActiveControl Is TxtTransSerial) Then
        '        Cmd_Click (0)
    Else
        Sendkeys "{TAB}"
    End If

End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

    If Me.TxtModFlg <> "E" Then Exit Sub
    If val(Me.TxtNoteSerial.text) = 0 Or val(Me.TxtNoteSerial1.text) = 0 Then GoTo ll

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , val(Me.TxtNoteSerial), val(Me.TxtNoteSerial1), 170
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), Me.TxtNoteSerial, Me.TxtNoteSerial1, 170

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
ll:
End Sub

Private Sub Fg_DblClick()
    'FrmItemsDetails.Show
End Sub

Private Sub Form_Activate()
    'Set m_Menu1 = mdifrmmain.MnuInvInsertTemp
    'Set m_MenuRefesh = mdifrmmain.MnuInvSales_Refresh
    'Set m_MenuCusBalance = mdifrmmain.MnuInvSales_Mnu1
    'Set m_MenuViewList = mdifrmmain.MnuInvViewList
    'Set m_MenuViewNotes = mdifrmmain.MnuInvSales_Mnu4
    'Set m_MenuScreenPremission = mdifrmmain.MnuInvSales_Mnu7

    If TxtTransSerial.Enabled = True Then
        '    TxtTransSerial.SetFocus
    End If

    'If first_run = True Then
    'Me.left = Me.left + 1420
    'Cmd_Click (0)
    'first_run = False
    'End If
    Ele(2).Enabled = True
   ' CheckInputIdle (10)
End Sub

Private Sub Grid1_Click()

    With GRID1

        Select Case .Col

            Case 2
 
                With FG
                    .Clear flexClearScrollable, flexClearEverything
                    .Rows = 1
       
                End With
 
                fillVchr

            Case 7
                FrmOut.Retrive val(.TextMatrix(.Row, 1))

            Case 8
                ShowGL_cc .TextMatrix(.Row, .ColIndex("NoteSerial")), , 200

        End Select

    End With

End Sub

Private Sub GRID2_Click()

    With FG
        .Clear flexClearScrollable, flexClearEverything
        .Rows = 1
       
    End With
 
    fillOrders

End Sub

Function fillVchr()
    Dim i As Integer
        
    With GRID1

        For i = 1 To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                Retrive_orders_data (val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
            
            End If

        Next i

    End With

End Function

Function fillOrders()
    Dim i As Integer

    With GRID2

        For i = 1 To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                Retrive_orders_data (val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
            
            End If

        Next i

    End With

End Function

Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    If val(lbl(Index).Caption) <> 0 Then
        lbl(Index).ToolTipText = WriteNo(lbl(Index).Caption, 0, True)
    End If

End Sub

Private Sub LblDiscountsTotal_Change()
    LblDiscountsTotalView.Caption = Format(val(LblDiscountsTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub LblInstallCount_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    'LblInstallCount.ToolTipText = WriteNo(LblInstallCount.Caption, 0, True)
End Sub

Private Sub LblInstallTotal_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    'LblInstallTotal.ToolTipText = WriteNo(LblInstallTotal.Caption, 0, True)
End Sub

Private Sub LblInvProfit_Change()
    CalculateInvPrecent
End Sub

Private Sub LblPrecenValue_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     X As Single, _
                                     Y As Single)
    'LblPrecenValue.ToolTipText = WriteNo(LblPrecenValue.Caption, 0, True)
End Sub

Private Sub LblTotal_Change()
    LblTotalView.Caption = Format(val(LblTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
    If SystemOptions.UserInterface = ArabicInterface Then
LblSowPrice(1).Caption = "«·«Ã„«·Ì : " & LblTotalView.Caption
Else
LblSowPrice(1).Caption = "Totals : " & LblTotalView.Caption
End If

    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Or Me.TxtModFlg = "R" Then
        TxtNetValue.text = val(LblTotal.Caption)
        'TxtPayedValue.text = TxtNetValue.text
 
     '   With Me.FgInstallments
     '       .Clear flexClearScrollable, flexClearEverything
     '       .Rows = .FixedRows
     '       LblPrecenType.Caption = ""
     '       LblPrecenValue.Caption = ""
     '       LblInstallTotal.Caption = ""
     '       LblInstallCount.Caption = ""
     '       LblFirstInstallDate.Caption = ""
     '       LblInstallmentType.Caption = ""
     '   End With

    End If
  
End Sub

Function showComm()

   ' If val(LblInstallTotal.Caption) > 0 Then
   '     lblInstComm.Caption = val(LblInstallTotal.Caption) - val(LblTotal.Caption)
  '
  '  Else
  '      lblInstComm.Caption = 0
  '      '  Me.LblFinal = 0
  '  End If

    Me.LblFinal = val(lblInstComm.Caption) + val(LblTotal.Caption)
    'Me.lblInstComm.Caption = Format(Val(lblInstComm.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
 
    Me.LblFinal.Caption = Format(val(LblFinal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

End Function

Private Sub LblTotal_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    LblTotal.ToolTipText = WriteNo(LblTotal.Caption, 0, True)

End Sub

Private Sub LblTotalAll_Change()


    LblTotalAllView.Caption = Format(val(LblTotalAll.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
txtPointsOpr.text = CheckCartDiscount(val(LblTotalAll.Caption))
End Sub

Function loadInvoices()
Dim StrSQL As String
If Me.TxtModFlg.text = "R" Then
    StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=21   and  Printed IS NULL"

            If SystemOptions.usertype <> UserAdminAll Or val(Current_branch) <> 0 Then
                StrSQL = StrSQL & "  AND   BranchId=" & Current_branch
            End If

            StrSQL = StrSQL & " Order by Transaction_ID"

            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
            If rs.RecordCount > 0 Then
            rs.MoveFirst
            End If
FillTables
End If

End Function
 Private Sub lvwTables_ItemClick(Item As vbalListViewLib6.cListItem)
    On Error GoTo ErrorHandler
    Dim sInfo As String
'rs.Resync
'XPTxtDiscountVal.Visible = False
    If Not lvwTables.SelectedItem Is Nothing Then

        With lvwTables.SelectedItem
              Cmd(7).Visible = True   '«ŸÂ«— «·œ›⁄
             If SystemOptions.UserInterface = ArabicInterface Then
      Cmd(2).Caption = "Õ›Ÿ"
      Else
      Cmd(2).Caption = "Save"
      End If
      If Me.TxtModFlg.text = "N" Then
      Cmd(7).Enabled = False
      End If


            CboPOSBillType.ListIndex = 4
            '    sInfo = "Key = " & Item.key & Item.text
            LBLTable1.Caption = Item.text
            LblStableID.Caption = Item.Key
 
 
 DcboEmp1.BoundText = GetWaiterForTable(Item.Key)
 
 If val(txtcurrenttableidMain) <> val(Item.Key) Then
 End If
 currenttableidMain = val(Item.Key)
 txtcurrenttableidMain = currenttableidMain
            If Me.TxtModFlg.text = "N" And .SubItems(1).Caption = "1" Then
            
          If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "«·„Ã·” «Ê «·ÿ«Ê·… «·„Õœœ… „‘€Ê·… Õ«·Ì« ·«»œ „‰ ”œ«œ ﬁÌ„… «·›«‰Ê—… «Ê·«", vbCritical
          Else
         MsgBox "This Location Have guest.", vbCritical
         End If
         
          
                LBLTable.Caption = ""
                LBLTable1.Caption = ""
                LblStableID.Caption = -1
                Exit Sub
            End If
 Dim currenttableid As Double
            If .SubItems(1).Caption = "1" Then
            
                                    currenttableid = getTransactionIdBytable(Item.Key)
                                    If currenttableid = -1 Then
                                         If SystemOptions.UserInterface = ArabicInterface Then
                                            MsgBox " ·« ÌÊÃœ ›Ê« Ì— ·Â–« «·„Ã·” «÷⁄ÿ ÃœÌœ «Ê·« ·«Œ Ì«— „Ã·”/ÿ«Ê·… ›«—€…", vbCritical
                                        Else
                                        MsgBox " There is no no Invoice To this Location Press New or Select Empty Location", vbCritical
                                        End If
                                        LBLTable1.Caption = ""
                                            LBLTable.Caption = ""
                                            LblStableID.Caption = -1
                        
                                            clear_all Me
                                            loadInvoices
                                            
                                            Exit Sub
                                            
                                    
                                    
                                        
                        Else
                        Retrive (getTransactionIdBytable(Item.Key))
                        End If

            Else

                If Me.TxtModFlg.text <> "N" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·« ÌÊÃœ ›Ê« Ì— ·Â–« «·„Ã·” «÷⁄ÿ ÃœÌœ «Ê·« ·«Œ Ì«— „Ã·”/ÿ«Ê·… ›«—€…", vbCritical
                Else
                MsgBox " There is no no Invoice To this Location Press New or Select Empty Location", vbCritical
                End If
                LBLTable1.Caption = ""
                    LBLTable.Caption = ""
                    LblStableID.Caption = -1

                    clear_all Me
                    Exit Sub
                End If
      
            End If
 
        End With
 
    End If

    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description & " [" & Err.Number & "]", vbInformation
    Exit Sub

End Sub

Private Sub m_FrmSearch_Unload(Cancel As Integer)
    Set m_FrmSearch = Nothing
End Sub

Private Sub m_Menu1_Click()
    On Error GoTo ErrTrap

    With FrmBuySearch
        .DealingForm = InsertTemplateToInvoice
        .Caption = "«·⁄—Ê÷ «·Ã«Â“…"
        .FG.TextMatrix(0, .FG.ColIndex("Transaction_ID")) = "ﬂÊœ «·⁄—÷"
        .FG.TextMatrix(0, .FG.ColIndex("BillDate")) = "«”„ «·⁄—÷"
        .FG.TextMatrix(0, .FG.ColIndex("ClientNmae")) = " «—ÌŒ «·⁄—÷"
        .FG.TextMatrix(0, .FG.ColIndex("StorName")) = "ﬁÌ„… «·⁄—÷"
        .XPChkSearchType.Visible = False
        .TxtVal.Visible = True
        .XPLbl(2).Visible = True
        .XPLbl(1).Visible = False
        .XPLbl(0).Visible = False
        .XPLbl(3).Visible = True
        .XPLbl(4).Visible = True
        .show vbModal
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub m_MenuCusBalance_Click()
    Dim cReport As ClsCustemerReport
    Dim LngCusID As Long

    With Me.FG

        If Me.DBCboClientName.BoundText = "" Then Exit Sub
        LngCusID = val(Me.DBCboClientName.BoundText)
        OpenScreen PopUpShowCustomerBalanceScreen, LngCusID, 0
    End With

End Sub

Private Sub m_MenuRefesh_Click()
    Dim Msg As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        Msg = " ÕœÌÀ «·»Ì«‰«  €Ì— „ «Õ ≈·« «‰  ﬂÊ‰ «·‘«‘… ›Ï Õ«·… «·⁄—÷ ›ﬁÿ..!"
        'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        'Exit Sub
    End If

    LoadCombosData
    NewGrid.FillGrid
    rs.Requery
    Exit Sub
ErrTrap:
End Sub

Private Sub m_MenuScreenPremission_Click()
    ShowScreenPermission Me.Name
End Sub

Private Sub m_MenuViewList_Click()
    Dim FrmView As FrmViewList
    Dim FG As VSFlex8UCtl.VSFlexGrid
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set FG = FrmView.vsfGroup1.VSFlexGrid

    With FG
        .Cols = 10
        .RowHeightMin = 320
        .TextMatrix(0, 0) = "—ﬁ„ «·»—‰«„Ã"
        .TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
        .TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
        .ColDataType(2) = flexDTDate
        .TextMatrix(0, 3) = "«”„ «·⁄„Ì·"
        .TextMatrix(0, 4) = "ÿ—Ìﬁ… «·œ›⁄"
        StrComboList = "#0;‰ﬁœÏ|#1;√Ã·"
        .ColComboList(4) = StrComboList
    
        .TextMatrix(0, 5) = "«”„ «·„Œ“‰"
        .TextMatrix(0, 6) = "«”„ «·„ÊŸ›"
    
        .TextMatrix(0, 7) = "‰Ê⁄ «·Œ’„"
        .TextMatrix(0, 8) = "ﬁÌ„… «·Œ’„"
        .TextMatrix(0, 9) = "≈Ã„«·Ï «·›« Ê—…"

        ',
        'QryTransactionsTotal.TransSum
        'QryTransactionsTotal.TransNet,
        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT QryTransactionsTotal.Transaction_ID, QryTransactionsTotal.Transaction_Serial," & "QryTransactionsTotal.Transaction_Date,dbo.TblCustemers.CusName, QryTransactionsTotal.PaymentType, " & "dbo.TblStore.StoreName,dbo.TblEmployee.Emp_Name ,QryTransactionsTotal.Trans_DiscountType," & "QryTransactionsTotal.Trans_Discount,QryTransactionsTotal.TotalAfterTax"
            StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal LEFT OUTER JOIN"
            StrSQL = StrSQL + " dbo.TblStore ON QryTransactionsTotal.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
            StrSQL = StrSQL + " dbo.TblEmployee ON QryTransactionsTotal.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
            StrSQL = StrSQL + " dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID"
            StrSQL = StrSQL + " WHERE QryTransactionsTotal.Transaction_Type=2 "
            StrSQL = StrSQL + " Order  By QryTransactionsTotal.Transaction_ID"
        ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "SELECT QryTransactionsTotal.Transaction_ID , QryTransactionsTotal.Transaction_Serial," & "QryTransactionsTotal.Transaction_Date,TblCustemers.CusName, QryTransactionsTotal.PaymentType," & "TblStore.StoreName,TblEmployee.Emp_Name ,QryTransactionsTotal.Trans_DiscountType," & "QryTransactionsTotal.Trans_Discount,QryTransactionsTotal.TotalAfterTax "
            StrSQL = StrSQL + "FROM (TblEmployee RIGHT JOIN (TblCustemers RIGHT JOIN QryTransactionsTotal " & "ON TblCustemers.CusID = QryTransactionsTotal.CusID) ON TblEmployee.Emp_ID = QryTransactionsTotal.Emp_ID) " & "LEFT JOIN TblStore ON QryTransactionsTotal.StoreID = TblStore.StoreID "
            StrSQL = StrSQL + " WHERE QryTransactionsTotal.Transaction_Type=2 "
            StrSQL = StrSQL + " Order  By QryTransactionsTotal.Transaction_ID"
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adAsyncExecute + adAsyncFetch
        Set cProgress = New ClsProgress
        BolFrmLoaded = True
        cProgress.ProgressType = Waiting
        cProgress.StartProgress

        Do While rs.State = adStateExecuting
            DoEvents
        Loop

        If BolFrmLoaded = True Then
            cProgress.StopProgess
            Set cProgress = Nothing
        End If

        Set .DataSource = rs
        .TextMatrix(0, 0) = "—ﬁ„ «·»—‰«„Ã"
        .TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
        .TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
        .ColDataType(2) = flexDTDate
        .TextMatrix(0, 3) = "«”„ «·⁄„Ì·"
        .TextMatrix(0, 4) = "ÿ—Ìﬁ… «·œ›⁄"
        StrComboList = "#0;‰ﬁœÏ|#1;√Ã·"
        .ColComboList(4) = StrComboList
        .TextMatrix(0, 5) = "«”„ «·„Œ“‰"
        .TextMatrix(0, 6) = "«”„ «·„ÊŸ›"
    
        .TextMatrix(0, 7) = "‰Ê⁄ «·Œ’„"
        .TextMatrix(0, 8) = "ﬁÌ„… «·Œ’„"
        .TextMatrix(0, 9) = "≈Ã„«·Ï «·›« Ê—…"
        .ColKey(9) = "TotalAfterTax"
        'Rs.Close
        'Set Rs = Nothing
    End With

    Set GrdBack = New ClsBackGroundPic
    FrmView.vsfGroup1.VSFlexGrid.WallPaper = GrdBack.Picture
    FrmView.vsfGroup1.SetRTL = True
    FrmView.vsfGroup1.TotalOnColKey = "TotalAfterTax"
    FrmView.vsfGroup1.update
    FrmView.show

End Sub

Private Sub m_MenuViewNotes_Click()
    CmdNotes_Click
End Sub

Private Sub Text1_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Command2.backcolor = vbYellow
        Command2.Enabled = False

        'Exit Sub
    End If

    If Text1.text = "" Then
        Command2.backcolor = vbGreen
        Command2.Enabled = True

        If SystemOptions.UserInterface = ArabicInterface Then
            Command2.Caption = "  ·„ Ì „ «‰‘«¡ «–‰ «·’—›- «÷€ÿ  ·«‰‘«¡ «–‰ ’—› «·Ì"
        Else
            Command2.Caption = "Create Issue Voucher"
        End If
        
    Else
        Command2.backcolor = &HC0C0C0
        Command2.Enabled = False

        If SystemOptions.UserInterface = ArabicInterface Then
            Command2.Caption = "  „ «‰‘«¡ «–‰ «·’—› "
        Else
            Command2.Caption = "Voucher Was Created"
        
        End If
    End If

    If BillBasedOn(1).value = True Then
        Command2.backcolor = &HC0C0C0
        Command2.Enabled = False

        If SystemOptions.UserInterface = ArabicInterface Then
            Command2.Caption = "·« Ì„ﬂ‰ «‰‘«¡ «·”‰œ ·«‰ «·›« Ê—Â  „ —»ÿÂ« »⁄œÂ ”‰œ«  "
        Else
            Command2.Caption = "Can't Create Voucher  "
        End If
    End If

End Sub

Private Sub Timer1_Timer()

    If Shape1.BorderColor = &H80000008 Then
        Shape1.BorderColor = &HFF0000
    Else
        Shape1.BorderColor = &H80000008
    End If

End Sub

 



Private Sub Timer4_Timer()
lbl(81).Caption = Time
End Sub

Private Sub Timer5_Timer()
On Error Resume Next
If imageCounter = 0 Then imageCounter = 1
If imageCounter = 3 Then imageCounter = 1



Image14.Picture = LoadPicture(App.path & "\Images\pos2\" & imageCounter & ".jpg")
 imageCounter = imageCounter + 1
 
End Sub

Private Sub tmr_Timer()
Exit Sub
 Dim plii As PLASTINPUTINFO
    
' Setup the size
    plii.cbSize = Len(plii)
    
' Get the time of the last user input
    GetLastInputInfo plii

' Display the idle time
' (last user input is the last ms of the input, not idle time...
' to get idle time, take the current tick count - the last input
' time)
' EX (for clarification): if last user input was at 2:00pm, and
' it's now 2:01, 60 seconds, or 60*1000 ms have elapsed
' (2:00 - 2:01 = :01 = 60sec, = 60*1000)
On Error Resume Next
Dim COUNTIDLE As Double
    COUNTIDLE = GetTickCount - plii.dwTime ' / 1000 for seconds
    Debug.Print COUNTIDLE
    If val(COUNTIDLE) >= 5000 Then
    'Unload SFrmScreenSaver
    'Load SFrmScreenSaver
    'SFrmScreenSaver.Visible = True
    COUNTIDLE = 0
'    Me.tmr.Enabled = False
    End If
End Sub

Private Sub TxtAdminLogin_GotFocus()
TxtAdminLogin.text = ""
End Sub

Private Sub Txtcart_Change()
On Error Resume Next
XPCboDiscountType.ListIndex = 0
CashCustomerName.text = ""

Dim Name As String
GetCartData Txtcart, Name
CashCustomerName.text = Name
'XPCboDiscountType.ListIndex = 1

End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.text = "F" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub

Public Sub RetriveOrder(Optional order_no As String = "")
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh

    StrSQL = "Select * from transactions where order_no='" & order_no & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Sub
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
        Me.DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        Me.DCboStoreName.BoundText = IIf(IsNull(rs("storeid").value), "", rs("storeid").value)
        Me.dcBranch.BoundText = IIf(IsNull(rs("Branchid").value), "", rs("Branchid").value)

        'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
            'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice")), 0, (RsDetails("ShowPrice").value)) ' GET_COST_PRICE_FOR_PRODUCT_ITEM(Val(FG.TextMatrix(Num, FG.ColIndex("Code"))))
      
            '  FG.TextMatrix(Num, FG.ColIndex("Expenses")) = IIf(IsNull(RsDetails("Lineexpenses")), "", (RsDetails("Lineexpenses").value))
         
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemType")) = IIf(IsNull(RsDetails("ItemType")), 0, (RsDetails("ItemType").value))
         
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If

    TxtFillData.text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub TxtNetValue_Change()
    'If Me.TxtModFlg.text <> "E" Then
    TxtRemainValue.text = val(Me.TxtPayedValue.text) - val(Me.TxtNetValue.text)
    'End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
       TxtPayedValue = val(Me.TxtNetValue.text)
    End If

End Sub

Private Sub TxtNetValue_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    TxtNetValue.ToolTipText = WriteNo(LblTotal.Caption, 0, True)
End Sub

Private Sub TXTOrDer_no_Change()

    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder Me.TXTOrDer_no
    End If

End Sub

Public Function NewBillFromOrder(orderNo As String)
 

End Function

Private Sub TXTOrDer_no_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Order_no_search.show
        Order_no_search.RetrunType = 8
        Order_no_search.DBCboClientName.BoundText = Me.DBCboClientName.BoundText
    End If

End Sub

Private Sub TxtPayedValue_Change()
    'TxtRemainValue.text = Val(Me.TxtPayedValue.text) - Val(Me.TxtNetValue.text)

    'If Me.TxtModFlg.text <> "E" Then
    TxtRemainValue.text = val(Me.TxtPayedValue.text) - val(Me.TxtNetValue.text)
    'End If

End Sub

Private Sub TxtPayedValue_GotFocus()
    TxtPayedValue.text = ""
End Sub

Private Sub TxtPayedValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPayedValue.text, 0)
End Sub

Private Sub txtPointsOpr_Change()
TxtTotalPoints.text = txtPointsOpr.text
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub

Private Sub TxtTransSerial_Change()
    FillVoucherGrid
End Sub

Private Sub TxtTransSerial_KeyDown(KeyCode As Integer, _
                                   Shift As Integer)
    Dim StrSearch As String
    Dim VarBookMark As Variant
    Dim Msg As String

    If Me.TxtModFlg.text = "R" Then
        If KeyCode = vbKeyReturn Then
            If Trim$(TxtTransSerial.text) <> "" Then
                StrSearch = Trim$(TxtTransSerial.text)

                If Not (rs.BOF Or rs.EOF) Then
                    If rs.EditMode = adEditNone Then
                        VarBookMark = rs.Bookmark
                        rs.find "Transaction_Serial='" & StrSearch & "'", , adSearchForward, adBookmarkFirst

                        If Not (rs.BOF Or rs.EOF) Then
                            Me.Retrive rs("Transaction_ID").value
                        Else
                            rs.Bookmark = VarBookMark
                            Msg = "Â–Â «·›« Ê—… €Ì— „ÊÃÊœ…...!!!"
                            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        End If
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub TxtTransSerial_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTransSerial.text, 1)
End Sub

Private Sub TxtValueAdded_Change()
RelinVatGrid
End Sub

Private Sub VatGrid_Click()
RelinVatGrid
End Sub
Sub RetriveValueAdded()
Dim sql As String
Dim i As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
    VatGrid.Clear flexClearScrollable, flexClearEverything
    VatGrid.Rows = 1
sql = " SELECT     dbo.TransactionValueAdded.Transaction_Type, dbo.TransactionValueAdded.Transaction_ID, dbo.TransactionValueAdded.Vat, dbo.TransactionValueAdded.Vatyo,"
sql = sql & " dbo.TransactionValueAdded.ItemID , dbo.TblItems.itemname, dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee ,dbo.TransactionValueAdded.selectd ,dbo.TransactionValueAdded.Valu "
sql = sql & " FROM         dbo.TransactionValueAdded LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems ON dbo.TransactionValueAdded.ItemID = dbo.TblItems.ItemID"

'salim1903
'sql = sql & " Where (dbo.TransactionValueAdded.Transaction_Type = 21) And (dbo.TransactionValueAdded.Transaction_ID = " & val(TxtInvID.Text) & ")"
sql = sql & " Where (dbo.TransactionValueAdded.Transaction_Type = 21) And (dbo.TransactionValueAdded.Transaction_ID = " & val(XPTxtBillID.text) & ")"
 
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With Me.VatGrid
rs2.MoveFirst
.Rows = .Rows + rs2.RecordCount
For i = 1 To .Rows - 1
 .TextMatrix(i, .ColIndex("index")) = i
.TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs2("ItemID").value), "", rs2("ItemID").value)
.TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(rs2("Vat").value), "", rs2("Vat").value)
.TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(rs2("Vatyo").value), "", rs2("Vatyo").value)
.TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("select")) = IIf(IsNull(rs2("selectd").value), 0, rs2("selectd").value)
.TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(rs2("Valu").value), 0, rs2("Valu").value)

If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
End If
rs2.MoveNext
Next i
End With
End If
RelinVatGrid
End Sub
Sub RelinVatGrid()
Dim i As Integer
Dim SmValu As Double
SmValu = 0
With VatGrid
For i = 1 To .Rows - 1
If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
SmValu = SmValu + val(.TextMatrix(i, .ColIndex("Vat")))
End If
Next i
End With
Label2.Caption = Format(SmValu, ".##")
TxtValueAdded.text = Format(SmValu, ".##")

showComm
If SmValu <> 0 Then
 NewGrid.Calculate 1, , , True
 End If
 
LblTotal.Caption = val(LblTotalAll.Caption) - val(LblDiscountsTotal.Caption) + IIf(SystemOptions.PriceWithVAT = False, val(TxtValueAdded.text), 0)  '- SmVal
LBLPayVal.Caption = val(TxtNetValue.text) + IIf(SystemOptions.PriceWithVAT = False, val(TxtValueAdded.text), 0)

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

'
Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    'Exit Sub
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" And Not (Me.ActiveControl Is TxtTransSerial) Then
            '        Cmd_Click (0)
        Else
            '    SendKeys "{TAB}"
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

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
    End If

    If KeyCode = vbKeyF7 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (11)
    End If
    
    If KeyCode = vbKeyF2 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            'XPBtnAdd_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            'XPBtnRemove_Click
        End If
    End If

    If KeyCode = vbKeyDelete Then
        If Me.ActiveControl Is FG Then
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
                'XPBtnRemove_Click
            End If
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            XPBtnNewClients_Click
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
                'XPFillData_Click
            End If
        End If
    End If

    If Shift = 2 Then
        XPTab301.SetFocus

        If KeyCode = vbKeyTab Then
            If XPTab301.CurrTab = 0 Then
                XPTab301.CurrTab = 1

                If XPChkPayType(0).Enabled = True Then
                    XPChkPayType(0).SetFocus
                End If

            Else
                XPTab301.CurrTab = 0
                FG.SetFocus
            End If
        End If
    End If

    If Shift = VBRUN.ShiftConstants.vbShiftMask Then

        'vbKeyX
        If KeyCode = vbKeyEscape Then
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    Dim colX         As cColumn
    Dim itmX         As cListItem
    Dim i            As Long
    Dim j            As Long
    Dim imageCounter As Integer
    
    '   lvwItems.BackgroundPicture = App.path & "\Garphics\wallpaper_Main11.jpg"
    Dim visapayed    As Double
    
    Me.show 'Force to show window
    loadLogo
  
    TimeOut_InSec = 10
    Me.Refresh
   
    With lvwItems
        lvwItems.Listitems.Clear
        .Visible = False
        .CustomDraw = True
        .AutoArrange = True
        '    .ImageList(eLVLargeIcon) = GrouplImageList ' ilsIcons32
        '    .ImageList(eLVSmallIcon) = GrouplImageList ' ilsIcons16
        '    .ImageList(eLVTileImages) = GrouplImageList ' ilsIcons48
        '    .ImageList(eLVHeaderImages) = GrouplImageList ' ilsIcons16
      
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
 
    With lvwTables
        .Visible = False
        .CustomDraw = True
            
        .AutoArrange = True
        '       .ImageList(eLVLargeIcon) = ilsIcons32
        .ImageList(eLVSmallIcon) = ilsIcons16
        '.ImageList(eLVTileImages) = ilsIcons48
        '       .ImageList(eLVHeaderImages) = ilsIcons16
      
        ' Set up image lists:
      
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
    End With
  
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
    lbl(82).Caption = Date
    lbl(83).Caption = GetWeekdayName(DatePart("w", Date) + 1)

    lblLabel1(0).Width = Me.Width

    lblLabel1(0).AutoSize = True
    ' Load lblLabel1(1)
    ' lblLabel1(1).Visible = True
    '   Load lblLabel1(1).
    lblLabel1(1).Width = Me.Width
    lblLabel1(1).left = Me.Width

    showmessage
    ' Me.left = (mdifrmmain.Width - Me.Width) / 2
    '    Me.top = (mdifrmmain.Height - Me.Height) / 2
    ScreenNameArabic = " ›« Ê—… «·„»Ì⁄«  "
    ScreenNameEnglish = " Sales Bill"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
 
    first_run = True
    Dim StrSQL  As String
    Dim Num     As Integer
    Dim StrList As String
    Dim BGround As New ClsBackGroundPic
    Dim ShowTax As Boolean

    'On Error GoTo ErrTrap
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
 
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    'Set m_menu1.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Excute").Picture

    Dim My_SQL As String
    'My_SQL = "  select branch_id,branch_name from TblBranchesData   "
    'fill_combo dcBranch, My_SQL
  
    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
        ' XPDtbBill.Enabled = False
    End If

    Set NewGrid.Grid = FG

    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
    Ele(4).Visible = ShowTax
    NewGrid.GridTrans = InvoiceTransaction
    Set NewGrid.TxtNots = Me.Text3
    Set NewGrid.VatGrid = VatGrid
    Set NewGrid.TxtInvID = Me.XPTxtBillID
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.txtTotal = XPTxtSum
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    Set NewGrid.TxtValueCash = XPTxtValue(0)
    Set NewGrid.TxtValueDelay = XPTxtValue(1)
    Set NewGrid.TxtValuechque = XPTxtValue(2)
    Set NewGrid.txt_Currency_rate = txt_Currency_rate
    Set NewGrid.Customer = Me.DBCboClientName
    Set NewGrid.LBLGross = LBLGross
    Set NewGrid.TxtValueAdded = TxtValueAdded

    '--------------------------------------
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.TxtAddTax = Me.TxtTaxAddValue
    Set NewGrid.TxtStampTax = Me.TxtTaxStampValue
    Set NewGrid.TxtServiceTax = Me.TxtTaxServiceValue
    Set NewGrid.Branch = Me.dcBranch
    
    '------------------------------------------------
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.StoreName = Me.DCboStoreName
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.CmdAddSerialLIst = Me.CmdSearch
    'Set NewGrid.CboDiscountType = CboDiscountType
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    Set NewGrid.LblInvProfit = Me.LblInvProfit
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.GrdTBar = Me.TBar
    Set NewGrid.LblTotalAll = Me.LblTotalAll
    Set NewGrid.LblDiscountsTotal = Me.LblDiscountsTotal

    Set NewGrid.LblTotalQty = Me.LblTotalQty

    Set NewGrid.LblTaxSalesValue = Me.lbl(51)
    Set NewGrid.LblTaxAddValue = Me.lbl(52)
    Set NewGrid.LblTaxStampValue = Me.lbl(53)
    Set NewGrid.LblTaxServiceValue = Me.lbl(54)

    NewGrid.FillGrid
    StrSQL = " select id,code from currency"
 
    fill_combo Me.DcCurrency, StrSQL

    FG.WallPaper = BGround.Picture
    AddTip
    XPTab301.CurrTab = 0
    XPDtbBill.value = Date

    If SystemOptions.UserInterface = ArabicInterface Then

        With XPCboDiscountType
            .Clear
            .AddItem ""
            .AddItem "  ﬁÌ„…/‰ﬁ«ÿ"
            .AddItem "‰”»…"
        End With

        With CboPayMentType
            .Clear
            .AddItem "‰ﬁœ«"
            .AddItem "¬Ã·"
        End With

        With Me.CboSaleType
            .Clear
            .AddItem "ﬁÿ«⁄Ì"
            .AddItem " Ã«—Ï"
        End With

        With CboPOSBillType
            .Clear
            .AddItem "«·ÿ«Ê·…" '0
            .AddItem "ÿ·»«  Œ«—ÃÌ…" '1
            .AddItem " Œœ„…  Ê’Ì· " '2
            .AddItem " Œœ„… ”Ì«—«  " '3
            .AddItem "«·ÿ«Ê·…" '4
        End With
    
    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With CboPOSBillType
            .Clear
            .AddItem "Table"
            .AddItem "Out Order"
            .AddItem " Delivery "
            .AddItem " Cars "
            .AddItem "Table" '4
        End With
        
        With XPCboDiscountType
            .Clear
            .AddItem "No Discount"
            .AddItem "Value Discount"
            .AddItem "Precetage Discount"
        End With

        With CboPayMentType
            .Clear
            .AddItem "Cash"
            .AddItem "Credit"
        End With

        With Me.CboSaleType
            .Clear
            .AddItem "Retail"
            .AddItem "WholeSale"
        End With

    End If

    '--------------------------------
    Set Dcombos = New ClsDataCombos
    LoadCombosData

    '--------------------------------
    If SystemOptions.UserInvoiceShowProfit = 0 Then
        '   Me.Ele(8).Visible = False
        Frame400.Visible = False
    Else
        Frame400.Visible = True
        'Me.Ele(8).Visible = True
    End If

    SetDtpickerDate Me.XPDtbBill
    '----------------------------
    SetDtpickerDate Me.DtpDelayDate
    '≈⁄œ«œ Ã—œ «·√ﬁ”«ÿ
    ChkInstall.value = Unchecked
    ChkInstall.Enabled = False

    With Me.FgInstallments
        .Rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    ' With Me.FgCheques
    '     .Rows = .FixedRows
    '     Set .WallPaper = BGround.Picture
    '     .RowHeightMin = 300
    '     .AutoSize 0, .Cols - 1, False
    ' End With

    Me.XPChkTAX.value = vbUnchecked
    XPChkTAX_Click
    Me.ChkTaxAdd.value = vbUnchecked
    ChkTaxAdd_Click
    Me.ChkTaxStamp.value = vbUnchecked
    ChkTaxStamp_Click
    Me.ChkTaxSerivce.value = vbUnchecked
    ChkTaxSerivce_Click
    '---------------------------
    'Resize_Form Me, TransactionSize
    '        Me.Height = 10000
    '        Me.Width = 17595
    '    Me.top = (mdifrmmain.ScaleHeight - Me.Height) / 2
    '    Me.left = (mdifrmmain.ScaleWidth - Me.Width) / 2

    '----------------------------
    'DB_CreateField "Transactions", "TransactionComment", adVarWChar, adColNullable, 255, , " ”ÃÌ· „·«ÕŸ«  ⁄·Ï «·›« Ê—…", False, True
    '----------------------------
    Dim rsOut As New ADODB.Recordset
    Dim Msg   As String
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!checkout = True Then
            StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=21  and  Printed IS NULL "
     
            If SystemOptions.usertype <> UserAdminAll Or val(Current_branch) <> 0 Then
                StrSQL = StrSQL & " AND   BranchId=" & Current_branch
            End If

            StrSQL = StrSQL & " Order by Transaction_ID"
                
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveLast
            End If

            Retrive
            Me.TxtModFlg.text = "R"
            InvType = 21
        Else
 
            StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=2   and  Printed IS NULL"

            If SystemOptions.usertype <> UserAdminAll Or val(Current_branch) <> 0 Then
                StrSQL = StrSQL & "  AND   BranchId=" & Current_branch
            End If

            StrSQL = StrSQL & " Order by Transaction_ID"

            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveLast
            End If

            Retrive
            Me.TxtModFlg.text = "R"
            InvType = 2
        End If
    End If

    '  If OPEN_NEW_SCREEN = True Then
    '      Cmd_Click (0)
    '  End If
    On Error Resume Next
    Image2.Picture = LoadPicture(App.path & "\Images\pos\gray.jpg")
    Image3.Picture = LoadPicture(App.path & "\Images\pos\gray.jpg")
    Image6.Picture = LoadPicture(App.path & "\Images\pos\gray.jpg")
    'Image6.Picture = LoadPicture(App.path & "\Images\pos\gray.jpg")
    'Image7.Picture = LoadPicture(App.path & "\Images\pos\gray.jpg")
    'Image5.Picture = LoadPicture(App.path & "\Images\pos\blue.jpg")
    'Image1.Picture = LoadPicture(App.path & "\Images\pos\DialPad.jpg")
    'Image4.Picture = LoadPicture(App.path & "\Images\pos\takeaway.jpg")
    'Image8.Picture = LoadPicture(App.path & "\Images\pos\phone.jpg")

    'CheckInputIdle 2
    Cmd_Click (0)
      
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set Dcombos = Nothing

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Set rs = Nothing
    Set TTP = Nothing
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Set SaleReport = Nothing

    Set m_Menu1 = Nothing
    Set m_MenuRefesh = Nothing

    If Not m_FrmSearch Is Nothing Then
        Unload m_FrmSearch
        Set m_FrmSearch = Nothing
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap
    Dim RsTest As ADODB.Recordset
    Dim StrSQL As String

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "›« Ê—…«·»Ì⁄"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Bill Invoice"
            End If

            BillBasedOn(1).Enabled = False
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
            Me.Cmd(9).Enabled = True
            Me.Cmd(11).Enabled = True
            
            Me.DcboEmp.Enabled = True
            GRID1.Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            XPBtnNewClients.Enabled = False
        
            XPCboDiscountType.locked = True
            Me.XPDtbBill.Enabled = False
            Me.DBCboClientName.locked = True
            Me.DCboStoreName.locked = True
        
            Me.XPTxtDiscountVal.locked = True
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
           ' XPChkPayType(2).Enabled = False
            XPTxtValue(0).Enabled = False
            XPTxtSerial(0).Enabled = False
            XPTxtValue(1).Enabled = False
            XPTxtSerial(1).Enabled = False
        
            FG.Editable = flexEDNone
            XPChkTAX.Enabled = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
                Me.Cmd(9).Enabled = False
                Me.Cmd(9).Enabled = False
            End If
        
            CboPayMentType.locked = True
            DtpDelayDate.Enabled = False

            If Not m_Menu1 Is Nothing Then
                m_Menu1.Enabled = False
            End If

            CmdINSTALLMENT.Enabled = False
          '  CmdCheque.Enabled = False

            '⁄—÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
            If XPTxtValue(1).Tag <> "" Then
                StrSQL = "select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
                Set RsTest = New ADODB.Recordset
                RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsTest.EOF Or RsTest.BOF) Then
                    CmdINSTALLMENT.Enabled = True

                    If SystemOptions.UserInterface = ArabicInterface Then
                        CmdINSTALLMENT.Caption = "⁄—÷ «·√ﬁ”«ÿ «·„”Ã·…"
                    Else
                        CmdINSTALLMENT.Caption = "View"
                    End If

                Else
                    CmdINSTALLMENT.Enabled = False

                    If SystemOptions.UserInterface = ArabicInterface Then
                        CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
                    Else
                        CmdINSTALLMENT.Caption = "Calc"
                    End If
                End If
            End If

            Ele(2).Enabled = False
            DcboEmp.Enabled = True
            XPChkTAX.Enabled = False
            ChkTaxAdd.Enabled = False
            ChkTaxSerivce.Enabled = False
            ChkTaxStamp.Enabled = False

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "›« Ê—…«·»Ì⁄( ÃœÌœ )"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Bill Invoice(New)"
            End If

            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
           ' Me.Cmd(7).Enabled = False
            Me.Cmd(9).Enabled = False
            Me.DcboEmp.Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
            Else
                CmdINSTALLMENT.Caption = "Calc Installments"
            End If
               
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
            FG.Enabled = True
            FG.Rows = FG.FixedRows
            FG.Rows = 2
            XPCboDiscountType.locked = False
            Me.XPDtbBill.Enabled = True
            XPDtbBill.value = Date
            Me.DBCboClientName.locked = False
            CboPayMentType.locked = False
            Me.DCboStoreName.locked = False
            Me.XPTxtDiscountVal.locked = False
        
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            'XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            'XPChkPayType(2).value = Unchecked
            FG.Editable = flexEDKbdMouse
            XPChkTAX.Enabled = True
            ChkTaxAdd.Enabled = True
            ChkTaxStamp.Enabled = True
            ChkTaxSerivce.Enabled = True
            XPTxtTaxValue.text = ""
            XPChkTAX.value = Unchecked
            XPCboDiscountType.ListIndex = 0
            CboPayMentType.ListIndex = 0
            '        XPFillData.Enabled = True
            DtpDelayDate.Enabled = True
            'm_Menu1.Enabled = True
            DtpDelayDate.value = Date
       
            CmdINSTALLMENT.Enabled = False
        '    CmdCheque.Enabled = False
            Ele(2).Enabled = True
            CboItemCase.ListIndex = 0
        
            Me.LblInvProfit.Caption = "0.0"
            Me.LblInvProfit.ForeColor = vbBlack
        
            DcboEmp.Enabled = True
            XPChkTAX.Enabled = True
            ChkTaxAdd.Enabled = True
            ChkTaxStamp.Enabled = True
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True
            ChkTaxSerivce.Enabled = True

            '        ChkTaxSerivce.Enabled = True
            '        ChkTaxStamp.Enabled = True
        Case "E"
            BillBasedOn(1).Enabled = False
    
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "›« Ê—…«·»Ì⁄(   ⁄œÌ· )"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Bill Invoice( Edit )"
            End If

            XPDtbBill.Enabled = False
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            Me.Cmd(9).Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            FG.Enabled = True
            XPCboDiscountType.locked = False
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Me.XPTxtDiscountVal.locked = False
            CboPayMentType.locked = False
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            DtpDelayDate.Enabled = True

            If XPChkPayType(0).value = Checked Then
                XPChkPayType_Click (0)
            End If

            If XPChkPayType(1).value = Checked Then
                XPChkPayType_Click (1)
            End If

            If XPChkPayType(2).value = Checked Then
                XPChkPayType_Click (2)
            End If

            If CboPayMentType.ListIndex = 0 Then
                CboPayMentType_Change
            End If

            FG.Editable = flexEDKbdMouse
            XPBtnNewClients.Enabled = True
            XPChkTAX.Enabled = True

            If Not m_Menu1 Is Nothing Then
                m_Menu1.Enabled = False
            End If

            If XPChkPayType(1).value = vbChecked Then
                If XPTxtValue(1).text <> "" Then
                    CmdINSTALLMENT.Enabled = True
                    CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
                Else
                    CmdINSTALLMENT.Enabled = False
                End If
            End If

        '    If Me.XPChkPayType(2).value = vbChecked Then
        '        CmdCheque.Enabled = True
        '    Else
        '        CmdCheque.Enabled = False
        '    End If

            DBCboClientName_Change
            Ele(2).Enabled = True
        
            DcboEmp.Enabled = True
            XPChkTAX.Enabled = True

            ChkTaxAdd.Enabled = True
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True
            '        ChkTaxSerivce.Enabled = True
            '        ChkTaxStamp.Enabled = True

    End Select

    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
        'XPDtbBill.Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0, _
                   Optional NoteSerial1 As String)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    Dim i As Long
'     LblTable1.Caption = ""
'            LblStableID.Caption = ""
 
'             clear_all Me

'    On Error GoTo ErrTrap
    '---------------------------------------------
    'Here We Reset all Setting

 '   With Me.FgInstallments
 '       .Clear flexClearScrollable, flexClearEverything
 '       .Rows = .FixedRows
 '       LblPrecenType.Caption = ""
 '       LblPrecenValue.Caption = ""
 '       LblInstallTotal.Caption = ""
 '       LblInstallCount.Caption = ""
 '       LblFirstInstallDate.Caption = ""
 '       LblInstallmentType.Caption = ""
 '   End With
    
    Me.CmdNotes.Visible = False
    Me.CmdNotes.Tag = ""
    Me.CmdRetruns.Visible = False
    Me.CmdRetruns.Tag = ""

    ChkTaxAdd.value = vbUnchecked
    Me.TxtTaxAddValue.text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.text = ""
    ChkTaxSerivce.value = vbUnchecked
    Me.TxtTaxServiceValue.text = ""

    '---------------------------------------------
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

  '  If rs.EOF And rs.BOF Then
  '      Exit Sub
  '  End If

    If Lngid <> 0 Then
        rs.find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then

            With FG
                FG.Rows = 1
   
            End With

            Exit Sub
        
        End If
    End If

    If NoteSerial1 <> "" Then

        rs.find "noteserial1='" & NoteSerial1 & "'", , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.text = "T"
    Screen.MousePointer = vbArrowHourglass
    ' »Ì«‰«  ÃœÌœ…
    Me.DCPaymentNet.BoundText = IIf(IsNull(rs("PaymentNetid").value), "", rs("PaymentNetid").value)
 TxtValueAdded.text = IIf(IsNull(rs("VAT").value), 0, (rs("VAT").value))
   
    TxtNetValue.text = IIf(IsNull(rs("NetValue").value), "", (rs("NetValue").value))
    TxtPayedValue.text = IIf(IsNull(rs("PayedValue").value), "", (rs("PayedValue").value))
    TxtRemainValue.text = IIf(IsNull(rs("RemainValue").value), "", (rs("RemainValue").value))
 
    TxtManualNo1.text = IIf(IsNull(rs("ManualNo1").value), "", (rs("ManualNo1").value))
    TxtManualNo2.text = IIf(IsNull(rs("ManualNo2").value), "", (rs("ManualNo2").value))
     
    'SessionD = IIf(IsNull(rs("SessionD").value), "", (rs("SessionD").value))
 
    '‰ﬁ«ÿ «·»Ì⁄
    If Not IsNull(rs("POSBillType").value) Then
        CboPOSBillType.ListIndex = val(rs("POSBillType").value)
        LblStableID.Caption = IIf(IsNull(rs("STableID").value), -1, (rs("STableID").value))
        If LblStableID.Caption = "-1" Then
LBLTable1.Caption = "Take Out"
End If
    Else
        CboPOSBillType.ListIndex = -1
        LblStableID.Caption = -1

    End If
 
    If CboPOSBillType.ListIndex = -1 Then
        LBLTable.Caption = ""
'        LblTable1.Caption = LblStableID.Caption
    ElseIf CboPOSBillType.ListIndex > 0 Then
        LBLTable.Caption = CboPOSBillType.List(val(CboPOSBillType.ListIndex))
    End If
         
     Dim mmm As String
    
    If Not (IsNull(rs("QrCodeImage").value)) Then
        LoadPictureFromDB Picture1, rs, "QrCodeImage", mmm
    Else
     Set Picture1.Picture = Nothing
    End If


    If Not IsNull(rs("BillBasedOn").value) Then

        If rs("BillBasedOn").value = 0 Then
            BillBasedOn(0).value = True
            '   BillBasedOn_Click (0)
        ElseIf rs("BillBasedOn").value = 1 Then
            BillBasedOn(1).value = True
            '      BillBasedOn_Click (1)
        ElseIf rs("BillBasedOn").value = 2 Then
            BillBasedOn(2).value = True
            '      BillBasedOn_Click (2)
        End If
    
    Else

        BillBasedOn(0).value = True
        '  BillBasedOn_Click (0)
    End If
'rs("empID1").value = IIf(DcboEmp1.BoundText = "", Null, DcboEmp1.BoundText)

     If Not (IsNull(rs("CashCustomerPhone").value)) Then
        Me.Txtcart.text = rs("CashCustomerPhone").value
    Else
        Me.Txtcart.text = ""
    End If


    If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.CashCustomerName.text = rs("CashCustomerName").value
    Else
        Me.CashCustomerName.text = ""
    End If
    
    
DcboEmp1.BoundText = IIf(IsNull(rs("empID1").value), "", rs("empID1").value)

    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
    DCDocTypes.BoundText = IIf(IsNull(rs("Doctype").value), "", rs("Doctype").value)
    Me.DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
    txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
 
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    'lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))
    Text1.text = IIf(IsNull(rs("NotS").value), "", (rs("NotS").value))

    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))

    TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), -1, val(rs("Trans_DiscountType").value))
    CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)

    XPTxtDiscountVal.text = IIf(IsNull(rs("Trans_Discount").value), "", (rs("Trans_Discount").value))
    Me.DBCboClientName.BoundText = 1 ' IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    FG.Clear flexClearScrollable, flexClearEverything
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    XPTxtTaxValue.text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)
    'Text1.text = IIf(IsNull(rs("nots2").value), "", (rs("nots2").value))
    Me.TXTOrDer_no.text = IIf(IsNull(rs("order_no").value), "", (rs("order_no").value))
 
    If IsNull(rs("BoxID").value) Then
        Me.DcboBox.BoundText = ""
    Else
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
    End If

    If IsNull(rs("SaleType").value) Then
        Me.CboSaleType.ListIndex = 0
    Else
        Me.CboSaleType.ListIndex = IIf(rs("SaleType").value = 0, 0, 1)
    End If

    If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.text = rs("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.text = ""
    End If

    'ChkInstall 11 10 2012
    If IsNull(rs("ChkInstall").value) Then
        Me.ChkInstall.value = vbUnchecked
    Else
        Me.ChkInstall.value = IIf(rs("ChkInstall").value = 0, vbUnchecked, vbChecked)
    End If

    '÷—»Ì… «·Œ’„ Ê«·≈÷«›…
    If Not IsNull(rs("TaxAddValue").value) Then
        If rs("TaxAddValue").value > 0 Then
            ChkTaxAdd.value = vbChecked
            Me.TxtTaxAddValue.text = rs("TaxAddValue").value
        End If
    End If

    '÷—»Ì… «·œ„€…
    If Not IsNull(rs("TaxStampValue").value) Then
        If rs("TaxStampValue").value > 0 Then
            ChkTaxStamp.value = vbChecked
            Me.TxtTaxStampValue.text = rs("TaxStampValue").value
        End If
    End If

    '÷—»Ì… «·Œœ„…
    If Not IsNull(rs("TaxServiceValue").value) Then
        If rs("TaxServiceValue").value > 0 Then
            ChkTaxSerivce.value = vbChecked
            Me.TxtTaxServiceValue.text = rs("TaxServiceValue").value
        End If
    End If

    TxtBillComment.text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))

    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
    StrSQL = StrSQL + "order by id"

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For i = 1 To RsDetails.RecordCount
            FG.TextMatrix(i, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
            FG.TextMatrix(i, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
            FG.TextMatrix(i, FG.ColIndex("TypeVAT")) = IIf(IsNull(RsDetails("TypeVAT").value), "", RsDetails("TypeVAT").value)
            FG.TextMatrix(i, FG.ColIndex("Vat")) = IIf(IsNull(RsDetails("Vat").value), "", RsDetails("Vat").value)
            FG.TextMatrix(i, FG.ColIndex("Vatyo")) = IIf(IsNull(RsDetails("Vatyo").value), "", RsDetails("Vatyo").value)
            
            FG.TextMatrix(i, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate").value), "", RsDetails("OrderArrivalDate").value)
            FG.Cell(flexcpPicture, i, FG.ColIndex("Ser")) = ""
            FG.Cell(flexcpData, i, FG.ColIndex("Ser")) = ""
            FG.TextMatrix(i, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(i, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim$(RsDetails("Item_ID").value))
            FG.TextMatrix(i, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))

            FG.TextMatrix(i, FG.ColIndex("printed")) = IIf(IsNull(RsDetails("printed")), "", Trim(RsDetails("printed").value))
            FG.TextMatrix(i, FG.ColIndex("printedGroup")) = IIf(IsNull(RsDetails("printedGroup")), "", Trim(RsDetails("printedGroup").value))
            
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(i, FG.ColIndex("HaveSerial")) = True

                '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«· «·Œ«’… »«·›« Ê—…
              '  If (RsDetails("Item_ID")) <> "" And RsDetails("ItemSerial") <> "" Then
              '      StrSQL = "select * From ReplacedItems where ReturnID=" & XPTxtBillID.Text
              '      StrSQL = StrSQL + " and ItemID=" & RsDetails("Item_ID")
              '      StrSQL = StrSQL + " and ItemSerial='" & RsDetails("ItemSerial") & "'"
              '      Set RsReplace = New ADODB.Recordset
              '      RsReplace.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                '    If Not (RsReplace.EOF Or RsReplace.BOF) Then
                '        FG.Cell(flexcpPicture, i, FG.ColIndex("Ser")) = mdifrmmain.ImgLstTree.ListImages("Request").Picture
                '        FG.Cell(flexcpData, i, FG.ColIndex("Ser")) = "X"
                '    End If
              '  End If
            End If
'
            FG.TextMatrix(i, FG.ColIndex("ItemType")) = IIf(IsNull(RsDetails("ItemType").value), "", (RsDetails("ItemType").value))

            If RsDetails("ItemType").value = 1 Then
                FG.Cell(flexcpPicture, i, FG.ColIndex("Ser")) = mdifrmmain.ImgLstTree.ListImages("Maintenance").Picture
            
            End If

            FG.TextMatrix(i, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))
            FG.TextMatrix(i, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
        
            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                FG.TextMatrix(i, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Else
                FG.TextMatrix(i, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").value))
            End If
        
            FG.TextMatrix(i, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(i, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(i, FG.ColIndex("guaranteeTime")) = IIf(IsNull(RsDetails("guaranteeTime")), "", (RsDetails("guaranteeTime").value))
        
            FG.TextMatrix(i, FG.ColIndex("ItemCostPrice")) = IIf(IsNull(RsDetails("CostPrice")), "", (RsDetails("CostPrice").value))
            FG.TextMatrix(i, FG.ColIndex("PofTransID")) = IIf(IsNull(RsDetails("CostTransID")), "", (RsDetails("CostTransID").value))
            FG.TextMatrix(i, FG.ColIndex("ItemProfit")) = IIf(IsNull(RsDetails("ItemProfit")), "", (RsDetails("ItemProfit").value))
            FG.TextMatrix(i, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
        
            FG.TextMatrix(i, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(i, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(i, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
            If val(FG.TextMatrix(i, FG.ColIndex("ItemProfit"))) = 0 Then
                Me.FG.Cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = vbYellow
            ElseIf val(FG.TextMatrix(i, FG.ColIndex("ItemProfit"))) < 0 Then
                Me.FG.Cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = vbRed
            Else
                Me.FG.Cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = 0
            End If

            FG.Cell(flexcpData, i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
        
            If SystemOptions.UserInterface = ArabicInterface Then
                FG.TextMatrix(i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            Else
                FG.TextMatrix(i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitNamee")), "", (RsDetails("UnitNamee").value))
            End If

            RsDetails.MoveNext
        
            If FG.Rows > 10 Then
                If i = 8 Then FG.Refresh
            End If

        Next i

        '----------------------------
        Me.LblInvProfit.Caption = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("ItemProfit"), FG.Rows - 1, FG.ColIndex("ItemProfit"))

        If val(Me.LblInvProfit.Caption) > 0 Then
            Me.LblInvProfit.ForeColor = &H4000&
        ElseIf val(Me.LblInvProfit.Caption) = 0 Then
            Me.LblInvProfit.ForeColor = vbBlack
        ElseIf val(Me.LblInvProfit.Caption) < 0 Then
            Me.LblInvProfit.ForeColor = vbRed
        End If

        '---------------------------
        '    Fg.AutoSize 0, Fg.Cols - 1, False
    End If
    RetriveValueAdded
RelinVatGrid
    XPChkPayType(0).value = Unchecked
    XPChkPayType(1).value = Unchecked
  '  XPChkPayType(2).value = Unchecked
    XPTxtValue(0).text = ""
    XPTxtValue(1).text = ""
    XPTxtSerial(0).text = ""
    XPTxtSerial(1).text = ""
    XPTxtValue(1).Tag = ""
    DtpDelayDate.value = Date
    '----------------------------------------------------------------------------------------
  '  StrSQL = "Select * From Notes Where Transaction_ID=" & val(rs("Transaction_ID").value)
  '  RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    'If Not (RsNotes.EOF Or RsNotes.BOF) Then

        'For i = 1 To RsNotes.RecordCount

  '          If RsNotes("NoteType").value = 170 Then
  '              XPChkPayType(0).value = Checked
  '              XPChkPayType_Click (0)
  '              XPTxtValue(0).Text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
  '              XPTxtSerial(0).Text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim$(RsNotes("NoteSerial").value))
  '              Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").value), "", RsNotes("BoxID").value)
  '          End If
'
'            If RsNotes("NoteType").value = 1 Then
'                XPChkPayType(1).value = Checked
'                XPChkPayType_Click (1)
'                XPTxtValue(1).Text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
'                XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").value), "", (RsNotes("NoteID").value))
'                XPTxtSerial(1).Text = IIf(IsNull(RsNotes("NoteSerial").value), "", (RsNotes("NoteSerial").value))
'                DtpDelayDate.value = IIf(IsNull(RsNotes("DueDate").value), "", (RsNotes("DueDate").value))
'            End If

'            If RsNotes("NoteType").value = 2 Then
'                XPChkPayType(2).value = Checked
'                XPChkPayType_Click (2)
'            End If
'
'            RsNotes.MoveNext
'        Next i

    'End If

'    Set RsNotes = New ADODB.Recordset
'    StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.NoteSerial," & "Notes.Note_Value, Notes.BankID,BanksData.BankName , Notes.ChqueNum, Notes.DueDate "
'    StrSQL = StrSQL + " FROM Notes INNER JOIN BanksData ON Notes.BankID = BanksData.BankID "
'    StrSQL = StrSQL + " Where NoteType=2 AND NOTES.Transaction_ID=" & val(rs("Transaction_ID").value)
'    StrSQL = StrSQL + " Order BY Notes.NoteID"
'    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    With Me.FgCheques
'        .Rows = .FixedRows
'
'        If Not (RsNotes.BOF Or RsNotes.EOF) Then
'            .Rows = .FixedRows + RsNotes.RecordCount
'
'            For i = .FixedRows To .Rows - 1
'                .TextMatrix(i, .ColIndex("CheckValue")) = IIf(IsNull(RsNotes("Note_Value").value), "", RsNotes("Note_Value").value)
'                .TextMatrix(i, .ColIndex("CheckNumber")) = IIf(IsNull(RsNotes("ChqueNum").value), "", RsNotes("ChqueNum").value)
'                .TextMatrix(i, .ColIndex("BankID")) = IIf(IsNull(RsNotes("BankID").value), "", RsNotes("BankID").value)
'                .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(RsNotes("BankName").value), "", RsNotes("BankName").value)
'
'                If Not IsNull(RsNotes("DueDate").value) Then
'                    .TextMatrix(i, .ColIndex("DueDate")) = DisplayDate(RsNotes("DueDate").value)
'                Else
'                    .TextMatrix(i, .ColIndex("DueDate")) = ""
'                End If
'
'                RsNotes.MoveNext
'            Next i
'
'        End If
'
'        .AutoSize 0, .Cols - 1, False
'        SumChecks
'    End With
'
    TxtFillData.text = "F"
    '-----------------------------------------------------------------------------------------------
    Dim SngRelatedNotesValues As Single
   ' Me.CmdNotes.Visible = ShowRelatedNotes(val(Me.XPTxtBillID.Text), 0, SngRelatedNotesValues)
    'Me.CmdNotes.Tag = SngRelatedNotesValues
'
'    SngRelatedNotesValues = 0
'    Me.CmdRetruns.Visible = ShowRelatedTransactions(val(Me.XPTxtBillID.Text), 0, SngRelatedNotesValues)
'    Me.CmdRetruns.Tag = SngRelatedNotesValues

    '-----------------------------------------------------------------------------------------------
           NewGrid.Calculate 1, , , True
       NewGrid.SentTypeVAT
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    showComm
'    FillVoucherGrid
'    FillOrderGrid

    '    Else
    '        CmdINSTALLMENT.Enabled = False
    '        CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
    
    '  End If
    'Else
    'FgInstallments.Clear

    '⁄—÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
  '  If XPTxtValue(1).Tag <> "" Then
  '      StrSQL = "Select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
  '      Set RsTest = New ADODB.Recordset
  '      RsTest.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'        If Not (RsTest.EOF Or RsTest.BOF) Then
'            CmdINSTALLMENT.Enabled = True
'            CmdINSTALLMENT.Caption = "⁄—÷ «·√ﬁ”«ÿ «·„”Ã·…"
'            LngPartID = RsTest("PartID").value
'          '  Me.LblPrecenType.Tag = RsTest("InterestType").value
'
'            If RsTest("InterestType").value = 0 Then
'                LblPrecenType.Caption = "‰”»… „∆ÊÌ…"
'            ElseIf RsTest("InterestType").value = 1 Then
'                LblPrecenType.Caption = "ﬁÌ„… À«» …"
'            ElseIf RsTest("InterestType").value = 2 Then
'                LblPrecenType.Caption = "·«ÌÊÃœ"
'            End If
'
'            Me.LblPrecenValue.Caption = RsTest("InterestVal").value
'            'LblDiscount.Caption = IIf(IsNull(RsTest("Discount").value), "", RsTest("Discount").value)
'            'Me.LblAdvPayment.Caption = IIf(IsNull(RsTest("AdvPayment").value), "", RsTest("AdvPayment").value)
'
'            Me.LblInstallTotal.Caption = RsTest("Total").value
'            Me.LblInstallCount.Caption = RsTest("InstallCount").value
'            Me.LblFirstInstallDate.Caption = DisplayDate(RsTest("FirstInstallDate").value)
'            Me.LblInstallmentType.Tag = RsTest("InstallmentType").value
'
''            If RsTest("InstallmentType").value = 0 Then
 '               LblInstallmentType.Caption = "ÌÊ„"
 '           ElseIf RsTest("InstallmentType").value = 1 Then
 '               LblInstallmentType.Caption = "‘Â—"
 '           ElseIf RsTest("InstallmentType").value = 2 Then
 '               LblInstallmentType.Caption = "”‰…"
 '           End If
'
'            Me.LblInstallSeprator.Caption = RsTest("InstallSeprator").value
'            Me.LblStartValue.Caption = IIf(IsNull(RsTest("StartValue").value), "", RsTest("StartValue").value)
'            LblDiscount.Caption = IIf(IsNull(RsTest("Discount").value), "", RsTest("Discount").value)
'            Me.LblAdvPayment.Caption = IIf(IsNull(RsTest("AdvPayment").value), "", RsTest("AdvPayment").value)
'
'            Set RsPartDetails = New ADODB.Recordset
'            StrSQL = "Select * From InstallMentDetails Where PartID=" & LngPartID
'            RsPartDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'            'fill Installments Grid
'            If Not (RsPartDetails.BOF Or RsPartDetails.EOF) Then
'                RsPartDetails.MoveFirst
''
 '               With Me.FgInstallments
 '                   .Rows = .FixedRows + RsPartDetails.RecordCount
'
'                    For i = .FixedRows To .Rows - 1
'                        .TextMatrix(i, .ColIndex("QestID")) = IIf(IsNull(RsPartDetails("QestID").value), "", RsPartDetails("QestID").value)
'                        .TextMatrix(i, .ColIndex("Serial")) = IIf(IsNull(RsPartDetails("QeqtNum").value), "", RsPartDetails("QeqtNum").value)
'                        .TextMatrix(i, .ColIndex("QeqtNum")) = IIf(IsNull(RsPartDetails("QeqtNum").value), "", RsPartDetails("QeqtNum").value)
'
'                        .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(RsPartDetails("Value").value), "", RsPartDetails("Value").value)
'
'                        If Not IsNull(RsPartDetails("DueDate").value) Then
'                            .TextMatrix(i, .ColIndex("Due_Date")) = DisplayDate(RsPartDetails("DueDate").value)
'                        Else
'                            .TextMatrix(i, .ColIndex("Due_Date")) = ""
'                        End If
'
'                        RsPartDetails.MoveNext
'                    Next i
'
'                End With
'
'            End If
'
      '      showComm
      '  Else
      '      CmdINSTALLMENT.Enabled = False
      '      CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
    '
    '    End If

   ' End If

    '›« Ê—… «·Œœ„« 
'    If CheckBillType = 0 Then
'        Command2.backcolor = &HC0C0C0
'        Command2.Enabled = False

'        If SystemOptions.UserInterface = ArabicInterface Then
'            Command2.Caption = "  ›« Ê—… Œœ„«  Ê·Ì” ·Â« ”‰œ ’—› "
'        Else
'            Command2.Caption = " Services Invoices"
'
'        End If

'        Exit Sub

'    End If

    DoEvents
        
    Exit Sub

ErrTrap:
    Resume
    Screen.MousePointer = vbDefault
End Sub

Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
'            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
'            Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
'
'            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.text = "R"
             '   XPBtnMove_Click (1)
                LBLTable1.Caption = ""
'            End If

        Case "E"
'            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
'            Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

'            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
If 1 = 1 Then

      LBLTable1.Caption = ""
      rs.find "Transaction_ID='" & val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Exit Sub
                End If

                If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Retrive
                End If
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_TransAction()
    Dim Msg As String
    Dim RsTest As ADODB.Recordset
    Dim StrSQL As String
    Dim IntRes As Integer
    Dim BegainTrans As Boolean
    On Error GoTo ErrTrap

    If XPTxtBillID.text = "" Then
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    If AvailableDeal = False Then
        Exit Sub
    End If

    '«·√ﬁ”«ÿ «·„”œœ… ⁄·Ï «·›« Ê—…
    If XPTxtValue(1).Tag <> "" Then
        StrSQL = "select * From ReceiptQestForBill Where NoteID=" & XPTxtValue(1).Tag
        Set RsTest = New ADODB.Recordset
        RsTest.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTest.EOF Or RsTest.BOF) Then
            Msg = "·ﬁœ  „  Õ’Ì· »⁄÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï Â–Â «·›« Ê—…" & CHR(13)
            Msg = Msg + "Ê·« Ì„ﬂ‰ Õ–› »Ì«‰« Â«" & CHR(13)
            Msg = Msg + "≈–« ﬂ‰   —€» ›Ì Õ–› »Ì«‰«  Â–Â «·›« Ê—…" & CHR(13)
            Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «· Õ’Ì· «·Œ«’… »Â«"
            MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    End If

    '⁄„·Ì«  «·’Ì«‰… «·„— »ÿ… »«·›« Ê—…
    StrSQL = "select * From MaintenanceJuncTransaction Where Transaction_ID=" & Trim(XPTxtBillID.text)
    Set RsTest = New ADODB.Recordset
    RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTest.EOF Or RsTest.BOF) Then
        Msg = "·ﬁœ  „ ≈Ã—«¡ »⁄÷ ⁄„·Ì«  «·’Ì«‰… ⁄·Ï Â–Â «·›« Ê—… Ê·« Ì„ﬂ‰ Õ–›Â«"
        Msg = Msg + "≈–« ﬂ‰   —€» ›Ì Õ–› »Ì«‰«  Â–Â «·›« Ê—…" & CHR(13)
        Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «·’Ì«‰… «·Œ«’… »Â«"
        MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If Me.CboPayMentType.ListIndex = 0 Then

        '›« Ê—… ‰ﬁœÌ…
        If CheckBoxAccount(val(Me.DcboBox.BoundText), val(Me.XPTxtValue(0).text), XPDtbBill.value, False) = False Then
            Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
            Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï Õ”«»«  «·Œ“‰…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "  √ﬂÌœ Õ–›    »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
        ' Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
    Else
        Msg = " Confirm Delete  " & CHR(13)
        '     Msg = Msg + "do you new Operation?"
       
    End If
 
    IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

    If IntRes = vbYes Then
        If Not rs.RecordCount < 1 Then
            Cn.BeginTrans
            BegainTrans = True
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & rs("Transaction_ID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
        
            '                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & _
            '         "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & get_transaction_id(rs("nots").value, 19)
            '         Cn.Execute StrSQL, , adExecuteNoRecords
                
            '         StrSQL = "Delete From Transactions  " & _
            '         "Where Transaction_ID=" & get_transaction_id(rs("nots").value, 19)
            '         Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "delete From Notes where noteid=" & val(TXTNoteID.text)
    
            Cn.Execute StrSQL, , adExecuteNoRecords
            DeleteLinkTOIssueVoucher
            DeleteTransactiomsVoucher val(Text1.text)
            CuurentLogdata ("D")

            If CboPOSBillType.ListIndex = 0 And val(LblStableID.Caption) <> -1 Then
                Cn.Execute "update Stables set Status =Null   where id=" & val(LblStableID.Caption)
                FillTables
            End If
    
            rs.delete
            Cn.CommitTrans
            BegainTrans = False
            Msg = " „  ⁄„·Ì… «·Õ–› "
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            rs.MoveFirst

            If rs.RecordCount < 1 Then
                clear_all Me
                TxtModFlg_Change
                XPTxtCurrent.Caption = 0
                XPTxtCount.Caption = 0
                  VatGrid.Clear flexClearScrollable, flexClearEverything
                VatGrid.Rows = 1
            Else
                Retrive
            End If
        End If
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title

    If BegainTrans = True Then
        rs.CancelUpdate
        Cn.RollbackTrans
        BegainTrans = False
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    Dim BolRtl As Boolean

    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… »Ì⁄ ÃœÌœ…" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F12 OR Enter", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…" & Wrap & "„›« ÌÕ «·«Œ ’«— F6", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  ⁄„·Ì… «·»Ì⁄" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F11", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… «·»Ì⁄ «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F10", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·»Ì⁄" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F9", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ⁄„·Ì… »Ì⁄" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F8", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F7", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— Ctrl + X", BolRtl
        End With
    
        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnNewClients, "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " «÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F5", BolRtl
        End With
    
        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, BolRtl
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        BolRtl = False

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New..." & Wrap & "Click here to add new Bill Invoice" & Wrap & "" & Wrap & "Shortcut (Enter Or F12)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print this Bill Invoice" & Wrap & "" & Wrap & "Shortcut (F6)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit..." & Wrap & "Edit this Bill Invoice Record" & Wrap & "  " & Wrap & "Shortcut (F11)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the New Bill Invoice Or Save the edit" & Wrap & "in the current Bill Invoice" & Wrap & "" & Wrap & "Shortcut (F10)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo..." & Wrap & "Undo in the New Bill Invoice" & Wrap & "Or Undo in the Editing" & Wrap & "" & Wrap & "Shortcut (F9)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete..." & Wrap & "Delete this current Bill Invoice" & Wrap & "" & Wrap & "Shortcut (F8)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search..." & Wrap & "Click here to display the search" & Wrap & "Screen" & Wrap & "Shortcut (F7)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit..." & Wrap & "Close this Window", BolRtl
        End With
    
        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnNewClients, "Add New Customer...." & Wrap & "To add New Customer Click here..." & Wrap & "Shortcut (F5)", BolRtl
        End With
    
        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "First..." & Wrap & "Move to first Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous..." & Wrap & "Move to Previous Record" & Wrap & " , BolRTL"
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next..." & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last..." & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help..." & Wrap & "to View Help Files" & Wrap & "click Here" & Wrap & "Shortcut(F1)" & Wrap, BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdValue_Click(Index As Integer)
LBLPayVal.Caption = 0
'TxtPayedValue.text = CmdValue(Index).Caption
LBLPayVal.Caption = CmdValue(Index).Caption
        With Grid
          .TextMatrix(.Row, .ColIndex("Value")) = LBLPayVal.Caption
          End With
     ReLineGrid
     
End Sub

 


Private Sub cleargrid()
    On Error Resume Next
    Dim i As Integer
 
  With Grid

      '  For i = .FixedRows To .Rows - 1

         .TextMatrix(.Row, .ColIndex("value")) = 0
          
      '  Next i

    End With
     TxtPayedValue = 0
    
End Sub

Private Sub CmdNos_Click(Index As Integer)
If Index <= 9 Then
LBLPayVal.Caption = LBLPayVal.Caption & Index

ElseIf Index = 10 Then
LBLPayVal.Caption = LBLPayVal.Caption & "00"

ElseIf Index = 11 Then
LBLPayVal.Caption = LBLPayVal.Caption & "."

ElseIf Index = 12 Then 'ar
If Len(LBLPayVal.Caption) > 1 Then
LBLPayVal.Caption = mId(LBLPayVal.Caption, 1, Len(LBLPayVal.Caption) - 1)
Else
LBLPayVal.Caption = ""
End If
ElseIf Index = 13 Then 'ar
 LBLPayVal.Caption = ""

TxtPayedValue.text = ""
cleargrid

ElseIf Index = 14 Then
TxtPayedValue = val(LBLPayVal)

 
        With Grid
          .TextMatrix(.Row, .ColIndex("Value")) = TxtPayedValue.text
          End With
     ReLineGrid
     
 TxtRemainValue.text = val(Me.TxtPayedValue.text) - val(Me.TxtNetValue.text)
 LBLPayVal.Caption = ""

End If

 ReLineGrid
 
End Sub
Private Sub SaveData()
    Dim Msg As String
    Dim RowNum As Integer
    Dim RSTransDetails As ADODB.Recordset
    Dim RSTransDetails1 As ADODB.Recordset
    Dim RsNotes As ADODB.Recordset
    Dim RsTemp      As New ADODB.Recordset
    Dim RsTest      As New ADODB.Recordset
    Dim RsRepeat    As ADODB.Recordset
    Dim RsDetalis   As ADODB.Recordset
    Dim StrSQL      As String
    Dim StrSqlDel   As String
    Dim note_id As Long
    Dim TransBegine As Boolean
    Dim BolTemp As Boolean
    Dim LnItemID As Long
    Dim i As Integer
    Dim DblNotesTotal As Double
    Dim SngTemp As Variant
    Dim usedaccount As Integer
    Dim TotalDiscountPerLine As Variant
    Dim TotalBillDiscount As Double
  '  On Error GoTo ErrTrap

    Me.FG.FinishEditing True

    DoEvents

    Screen.MousePointer = vbArrowHourglass
 

 

 
   
 
 

 
 

 
 

    If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
                    If XPTxtDiscountVal.text = "" Then
                                    If SystemOptions.UserInterface = ArabicInterface Then
                                        Msg = "≈–« ﬂ«‰ Â‰«ﬂ Œ’„ ⁄·Ï «·›« Ê—… " & CHR(13)
                                        Msg = Msg + "ÌÃ»  ÕœÌœ ﬁÌ„… Â–« «·Œ’„ " & CHR(13)
                                        Msg = Msg + "√Ê √Œ Ì«— ·« ÌÊÃœ Œ’„ "
                                    Else
                                        Msg = Msg + " Must Enter Discount Value " & CHR(13)
                                    End If
            
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        XPCboDiscountType.SetFocus
                        Sendkeys "{F4}"
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
    End If

 

 
    
 
    If NewGrid.CheckDataEntered = False Then
        Exit Sub
    End If

    '-------------------------------
    If NewGrid.Calculate(1, , False, True) = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '-------------------------------
    If Me.XPChkPayType(0).value = vbChecked Then
        DblNotesTotal = val(Me.XPTxtValue(0).text)
    End If

    If Me.XPChkPayType(1).value = vbChecked Then
        DblNotesTotal = DblNotesTotal + val(Me.XPTxtValue(1).text)
    End If

  '  If Me.XPChkPayType(2).value = vbChecked Then
  '      DblNotesTotal = DblNotesTotal + val(Me.lbl(18).Caption)
  '  End If

             '   If CboPayMentType.ListIndex = 1 And Me.XPChkPayType(2).value = Unchecked Then
                 '   XPChkPayType(1).value = 1
                    '  XPTxtValue(1).text = Val(LblTotalAll.Caption)
                 '   XPTxtValue(1).Text = val(LblTotal.Caption)
            
             '   Else
            
                               '     If CboPayMentType.ListIndex = 1 And Me.XPChkPayType(2).value = vbChecked Then
                               '         XPChkPayType(1).value = 0
                            
                                '    Else
                                '        XPChkPayType(0).value = 1
                                 '       '  XPTxtValue(0).text = Val(LblTotalAll.Caption)
                                 '       XPTxtValue(0).Text = val(LblTotal.Caption)
                            
                                 '   End If
           '     End If

 

    CurrentVoucherNo = ""
    CurrentVoucherSerialNo = ""

    'Create big notes
    my_branch = Current_branch 'val(Me.Dcbranch.BoundText)
    Dim mSerPos As Long
    Dim mSerPosString As String
    Dim mIsHiddenVat As Boolean
    Dim s As String
    If SystemOptions.IsByNewCoding Then
        Dim mPay As Long
        
        Dim rsMPay As ADODB.Recordset
        
        Dim mValue As Double
        
       If Not mIsChangeTable Then
        For i = 1 To Grid.Rows - 1
            mPay = Grid.TextMatrix(i, Grid.ColIndex("PaymentID"))
            mValue = val(Grid.TextMatrix(i, Grid.ColIndex("Value")))
            If mValue <> 0 Then
                s = " SELECT"
                s = s & "        IsHiddenVat, TT = (CASE"
                s = s & "              WHEN bd.BankId > 9 THEN CAST(bd.BankId AS NVARCHAR)"
                s = s & "                     Else '0' + CAST(bd.BankId AS NVARCHAR)"
                s = s & "                 END)"
                s = s & "             From TblPaymentType"
                s = s & "             INNER JOIN BanksData bd"
                s = s & "                 ON bd.BankId = TblPaymentType.BankId"
                s = s & "             Where IsNull(IsNewCode, 0) = 1"
                s = s & " and PaymentID = " & mPay
                Set rsMPay = New ADODB.Recordset
                
                rsMPay.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsMPay.EOF Then
                    mSerPos = val(rsMPay!tt & "")
                    mSerPosString = Trim(rsMPay!tt & "")
                    mIsHiddenVat = IIf(IsNull(rsMPay!IsHiddenVat & ""), False, rsMPay!IsHiddenVat & "")
                    Exit For
                End If
                rsMPay.Close
            End If
        Next
        End If
    End If
    my_branch = Current_branch

    If TxtNoteSerial1.text = "" Then
                If Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21, , , , , , , , mSerPosString) = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›…   ›« Ê—… „»Ì⁄«  ÃœÌœ… ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                               
                            If Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21) = "" Then
                                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                            Else
                                If SystemOptions.IsByNewCoding Then
                                    TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21, , , , , , , , mSerPosString)
                                Else
                                    TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21)
                                End If
                            End If
                End If
    End If
     
    Set RsNotesGeneral = New ADODB.Recordset
'    RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
  StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If Me.TxtModFlg.text = "N" Then
       'XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
    Else
    
     StrSQL = "Delete From TblTransactionPayments Where Transaction_ID=" & val(Me.XPTxtBillID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords


        StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(Me.XPTxtBillID.text)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        
        CurrentVoucherNo = GetVoucherGLNO(val(Text1.text), CurrentVoucherSerialNo)
        DeleteTransactiomsVoucher val(Text1.text)
        
        general_noteid = val(TXTNoteID.text)
    End If

 

    '---------------------------------
    Set RSTransDetails = New ADODB.Recordset
  
  
  StrSQL = "SELECT    * from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      
 

    Screen.MousePointer = vbArrowHourglass
    Cn.BeginTrans
    TransBegine = True

    If Me.TxtModFlg.text = "N" Then
        
        rs.AddNew
       XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
       
    ElseIf Me.TxtModFlg.text = "E" Then
        StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(Me.XPTxtBillID.text) 'Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
 Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.XPTxtBillID.text) & ""
    End If
   rs("Transaction_ID").value = val(XPTxtBillID.text)
   rs.update
   rs.Resync adAffectCurrent
    'rs("Transaction_ID").value = val(XPTxtBillID.Text)
    rs("BranchId").value = Current_branch '  IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
    rs("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
    rs("SerPos") = val(mSerPos)
    rs("IsHiddenVat") = mIsHiddenVat
    rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
    rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '

    If CboPayMentType.ListIndex = 0 Then
        rs("BoxID").value = dBox ' IIf(DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
    Else
        rs("BoxID").value = Null
      
    End If
      
    rs("NoteId").value = val(TXTNoteID.text)
    rs("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.text) = "", "", Trim(Me.TxtTransSerial.text))
    rs("Transaction_Date").value = XPDtbBill.value
    rs("Transaction_Type").value = 21
    rs("UserID").value = user_id
    rs("nots").value = ""

    rs("Currency_id").value = IIf(DcCurrency.BoundText = "", Null, val(DcCurrency.BoundText))
    rs("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.text), 1, txt_Currency_rate.text)

    If XPCboDiscountType.ListIndex = -1 Then
        rs("Trans_DiscountType").value = 0
    Else
        rs("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If



        If Trim$(Me.CashCustomerName.text) <> "" Then
        rs("CashCustomerName").value = Trim$(Me.CashCustomerName.text)
    Else
        rs("CashCustomerName").value = Null
    End If

    If Trim$(Me.Txtcart.text) <> "" Then
        rs("CashCustomerPhone").value = Trim$(Me.Txtcart.text)
    Else
        rs("CashCustomerPhone").value = Null
    End If
    
    
    
    rs("Trans_Discount").value = IIf(XPTxtDiscountVal.text = "", Null, val(XPTxtDiscountVal.text))
    rs("CusID").value = 2 'IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
    rs("StoreID").value = dstore 'IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
    rs("order_no") = IIf(TXTOrDer_no.text = "", Null, val(TXTOrDer_no.text))

    If CboPayMentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else
        rs("PaymentType").value = val(CboPayMentType.ListIndex)
    End If

    rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
    rs("TaxValue").value = IIf(XPTxtTaxValue.text = "", Null, val(XPTxtTaxValue.text))
    rs("Emp_ID").value = EmpID ' IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)

rs("empID1").value = IIf(DcboEmp1.BoundText = "", Null, DcboEmp1.BoundText)

    'ChkInstall 11 10 2012
    If ChkInstall.value = vbChecked Then
        rs("ChkInstall").value = 1
    Else
        rs("ChkInstall").value = 0
    End If

    If Me.CboSaleType.ListIndex = 0 Or Me.CboSaleType.ListIndex = -1 Then
        rs("SaleType").value = 0
    Else
        rs("SaleType").value = 1
    End If

    If Trim$(Me.TxtCashCustomerName.text) <> "" Then
        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.text)
    Else
        rs("CashCustomerName").value = Null
    End If

    rs("TransactionComment").value = IIf(Trim$(TxtBillComment.text) = "", Null, Trim$(TxtBillComment.text))

    '÷—»Ì… Œ’„ Ê≈÷«›…
    If ChkTaxAdd.value = vbChecked And val(Me.TxtTaxAddValue.text) > 0 Then
        rs("TaxAddValue").value = val(Me.TxtTaxAddValue.text)
    Else
        rs("TaxAddValue").value = 0
    End If

    '÷—»Ì… œ„€…
    If ChkTaxStamp.value = vbChecked And val(Me.TxtTaxStampValue.text) > 0 Then
        rs("TaxStampValue").value = val(Me.TxtTaxStampValue.text)
    Else
        rs("TaxStampValue").value = 0
    End If

    '÷—»Ì… Œœ„…
    If ChkTaxSerivce.value = vbChecked And val(Me.TxtTaxServiceValue.text) > 0 Then
        rs("TaxServiceValue").value = val(Me.TxtTaxServiceValue.text)
    Else
        rs("TaxServiceValue").value = 0
    End If

    '»Ì«‰«  ÃœÌœ…
    rs("PaymentNetid").value = IIf(DCPaymentNet.BoundText = "", Null, DCPaymentNet.BoundText)
    rs("NetValue").value = IIf(TxtNetValue.text = "", Null, val(TxtNetValue.text))
    rs("PayedValue").value = IIf(TxtPayedValue.text = "", Null, val(TxtPayedValue.text))
    rs("RemainValue").value = IIf(TxtRemainValue.text = "", Null, val(TxtRemainValue.text))
  
    rs("ManualNo1").value = IIf(TxtManualNo1.text = "", Null, val(TxtManualNo1.text))
    rs("ManualNo2").value = IIf(TxtManualNo2.text = "", Null, val(TxtManualNo2.text))
    rs("VAT").value = val(TxtValueAdded.text)
    If BillBasedOn(0).value = True Then
        rs("BillBasedOn").value = 0
    ElseIf BillBasedOn(1).value = True Then
        rs("BillBasedOn").value = 1
    ElseIf BillBasedOn(2).value = True Then
        rs("BillBasedOn").value = 2
    End If
    
    '‰ﬁ«ÿ «·»Ì⁄
'    If CboPOSBillType.ListIndex = 0 Then
'        rs("POSBillType").value = 0
'        rs("STableID").value = val(LblStableID.Caption)
'
'    Else
'        rs("POSBillType").value = val(CboPOSBillType.ListIndex)
'        rs("STableID").value = val(LblStableID.Caption)
        
''    End If
  rs("POSBillType").value = val(CboPOSBillType.ListIndex)
  
    If mIsChangeTablePaySave Then
        rs("STableID").value = Null
    Else
        rs("STableID").value = val(LblStableID.Caption)
    End If
        
        
    rs("SessionD").value = SessionD
        rs("Transaction_NetValue").value = val(lblInstComm.Caption) + val(LblTotal.Caption) '+ val(Me.TxtValueAdded.Text)

    rs.update

 

 SaveValueAdded

    For RowNum = 1 To FG.Rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then


            RSTransDetails.AddNew
            
            RSTransDetails("printed").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("printed")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("printed")))
            RSTransDetails("printedGroup").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("printedGroup")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("printedGroup")))
            
            
            RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))

            RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
            RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
            RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))

            RSTransDetails("Transaction_ID").value = val(XPTxtBillID.text)
            RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
            RSTransDetails("TypeVAT").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("TypeVAT")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("TypeVAT"))))
            RSTransDetails("Vat").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vat")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Vat"))))
            RSTransDetails("Vatyo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vatyo")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Vatyo"))))

            RSTransDetails("ShowPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
            RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
            
            RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
            
            RSTransDetails("guaranteeTime").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
            
            RSTransDetails("CostTransID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("PofTransID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("PofTransID"))))
            RSTransDetails("ItemProfit").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit"))))
        
            RSTransDetails("UnitID").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
          
            If SystemOptions.TypicalProduction = False Then
  
                RSTransDetails("CostPrice").value = 0 ' ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , RSTransDetails("UnitID").value)

                If RSTransDetails("CostPrice").value = 0 Then
                    'RSTransDetails("CostPrice").value = 0 ' ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , LastPurPriceType, , , XPDtbBill.value, , RSTransDetails("UnitID").value)
                    
                End If
                  
            Else
                RSTransDetails("CostPrice").value = 0
            
            End If
              
            RSTransDetails("SavedItemType").value = val(FG.TextMatrix(RowNum, FG.ColIndex("ItemType")))
               
            RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
            Dim cnt As Double
            cnt = FG.TextMatrix(RowNum, FG.ColIndex("Count"))

            RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
                
            RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
            '«·ÊÕœ« 
           
            Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double
        
            LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            LngUnitID = val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
            DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

       '     StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
       '     StrSQL = StrSQL + " AND UnitID=" & LngUnitID
       '     Set RsUnitData = New ADODB.Recordset
       '     RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

       '     If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = 1 'RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                'RSTransDetails("OpeningSalesQty").value = RSTransDetails("Quantity").value
                'RSTransDetails("OpeningSalesValue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Valu")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Valu"))))
                RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
            
       '     End If

            SngTemp = SngTemp + (val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))) * RSTransDetails("quantity").value)
         
            If Me.XPCboDiscountType.ListIndex = 1 Then
                TotalBillDiscount = IIf(XPTxtDiscountVal.text = "", Null, (XPTxtDiscountVal.text))
                     
            ElseIf XPCboDiscountType.ListIndex = 2 Then

                            If XPTxtDiscountVal.text <> "" Then
                                TotalBillDiscount = IIf(XPTxtDiscountVal.text = "", Null, (XPTxtDiscountVal.text)) * val(LblTotalAll.Caption) / 100
                                         
                            Else
                                TotalBillDiscount = 0
                            End If
            End If

          '  TotalDiscountPerLine = ((RSTransDetails("SHOWprice") * RSTransDetails("SHOWQTY")) / LblTotalAll.Caption) * (TotalBillDiscount)
          TotalDiscountPerLine = FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / LblTotalAll.Caption * (TotalBillDiscount)
          
            RSTransDetails("TotalDiscountPerLine") = Round(TotalDiscountPerLine, 20)
'                              Dim OldQty As Double
'             Dim OldCost As Double
'              Dim NewQty As Double
'               Dim NewCost As Double
'
'getItemCostData XPDtbBill.value, RSTransDetails("Item_ID").value, val(DCboStoreName.BoundText), val(Me.XPTxtBillID.Text), OldQty, OldCost, NewQty, NewCost,,LngUnitID
'       RSTransDetails("OldQty").value = NewQty
'       RSTransDetails("OldCost").value = NewCost
'
'      RSTransDetails("NewQty").value = RSTransDetails("OldQty").value - RSTransDetails("Quantity").value
'       RSTransDetails("NewCost").value = RSTransDetails("OldCost").value ' ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
       
       
            RSTransDetails.update
            '-------------
        
         End If

    Next RowNum

 

 

                    
  
    '‰ﬁ«ÿ «·»Ì⁄
'If CboPOSBillType.ListIndex = 4 And val(LblStableID.Caption) <> -1 Then
    If val(LblStableID.Caption) <> -1 Then
        If mIsChangeTablePaySave Then
        
            Cn.Execute "update Stables set Status =null   where id=" & val(LblStableID.Caption)
        Else
            Cn.Execute "update Stables set Status =1   where id=" & val(LblStableID.Caption)
        End If
        FillTables
      
    End If

'************************************************************************************
Dim PayDes As String
   Set RSTransDetails1 = New ADODB.Recordset
   StrSQL = "SELECT   * from dbo.TblTransactionPayments Where (1 = -1)"
   RSTransDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
            PayDes = ""
    For RowNum = 1 To Grid.Rows - 1
            
                       If Grid.TextMatrix(RowNum, Grid.ColIndex("Value")) <> "" Then
                        
                                    'Check Repeat Serial
                                     
If PayDes <> "" Then
          PayDes = PayDes & CHR(13) & Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName")) & ":" & Grid.TextMatrix(RowNum, Grid.ColIndex("value"))
 Else
           PayDes = Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName")) & ":" & Grid.TextMatrix(RowNum, Grid.ColIndex("value"))
 End If
 
                                
                                           RSTransDetails1.AddNew
                                            RSTransDetails1("boxid").value = val(Me.DcboBox.BoundText)
                                           RSTransDetails1("Recorddate").value = XPDtbBill.value
                                           
                                           RSTransDetails1("PointID").value = PPointID
                                           RSTransDetails1("CurrentCashireID").value = CurrentCashireID
                                           
                                           RSTransDetails1("Transaction_ID").value = val(XPTxtBillID.text)
                                           RSTransDetails1("PaymentID").value = IIf((Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentID")) = ""), Null, val(Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentID"))))
                                              RSTransDetails1("Value").value = IIf((Grid.TextMatrix(RowNum, Grid.ColIndex("Value")) = ""), 0, val(Grid.TextMatrix(RowNum, Grid.ColIndex("Value"))))


                                           If RSTransDetails1("PaymentID").value = 0 Then
                                                '    If RSTransDetails1("Value").value > val(TxtNetValue.text) Then
                                                    RSTransDetails1("Value").value = val(TxtNetValue.text) - visapayed
                                                     
                                                    
                                                '    End If
                                           
                                           End If
                                           
                                           RSTransDetails1("CardNo").value = IIf((Grid.TextMatrix(RowNum, Grid.ColIndex("CardNo")) = ""), "", (Grid.TextMatrix(RowNum, Grid.ColIndex("CardNo"))))
                                           
                                     
                                                
                                           RSTransDetails1.update
                                  
                             
                    End If
   Next RowNum
        
   ' For RowNum = 1 To Grid.Rows - 1
   '
   '                    If Grid.TextMatrix(RowNum, Grid.ColIndex("Value")) <> "" Then
   '
   '                                 'Check Repeat Serial
   '
   '
   '                                        RSTransDetails1.AddNew
   '
   '
   '                                        RSTransDetails1("PointID").value = PPointID
   '                                        RSTransDetails1("CurrentCashireID").value = CurrentCashireID
   '                                        RSTransDetails1("boxid").value = dBox
   '                                        RSTransDetails1("Transaction_ID").value = val(XPTxtBillID.Text)
   '                                        RSTransDetails1("PaymentID").value = IIf((Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentID")) = ""), Null, val(Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentID"))))
   '                                           RSTransDetails1("Value").value = IIf((Grid.TextMatrix(RowNum, Grid.ColIndex("Value")) = ""), 0, val(Grid.TextMatrix(RowNum, Grid.ColIndex("Value"))))
'
'
'                                           If RSTransDetails1("PaymentID").value = 0 Then
'                                                '    If RSTransDetails1("Value").value > val(TxtNetValue.text) Then
'                                                    RSTransDetails1("Value").value = val(TxtNetValue.Text) - visapayed
'
'
'                                                '    End If
'
'                                           End If
'
'                                           RSTransDetails1("CardNo").value = IIf((Grid.TextMatrix(RowNum, Grid.ColIndex("CardNo")) = ""), "", (Grid.TextMatrix(RowNum, Grid.ColIndex("CardNo"))))
'
'                                            '    If optsale(1).value = True Then   ' return sallimng
'                                                    RSTransDetails1("Effect").value = -1
'                                            '      Else
'                                            '         RSTransDetails1("Effect").value = 1
'                                            '    End If
'
'                                           RSTransDetails1.update
'
'
'                    End If
'   Next RowNum
'***************************************************************************************



    Cn.CommitTrans
LblSowPrice(0).Caption = ""
    TransBegine = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
 
    If SystemOptions.autoIssueVoucher = True Then
 '       CreateIssueVoucher
    End If
SaveQRCode "transactions", "Transaction_ID", val(XPTxtBillID), TxtNoteSerial1.text, (XPDtbBill.value), _
        (LblTotal.Caption), Picture1, 0, (TxtValueAdded.text), (LblTotal.Caption)


    'If SystemOptions.usertype = UserAdminAll Then
 '   CloseIssueVoucher
    'End If
    '----------------------------------------------------------------
    '·√‰‰« ﬁ„‰« »≈÷«›… Õ—ﬂ… „‰ ‰Ê⁄ „Œ ·›…
           Cn.Execute "update Transactions set PayDes ='" & PayDes & "'   where Transaction_ID=" & val(XPTxtBillID.text)
    
  '  StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=21  and  Printed IS NULL" ' & InvType
  '  StrSQL = StrSQL & "  AND   BranchId=" & Current_branch
'
  '  Set rs = New ADODB.Recordset
  '  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
  '  Me.Retrive val(Me.XPTxtBillID.Text)
    '----------------------------------------------------------------

    CuurentLogdata
    DoEvents
' „ «Ìﬁ«› «·ﬂÊœ Â‰« „ƒﬁ « ·ÕÌ‰ ›Â„Â „‰ «·»—‰” «·ﬂ»Ì— „ «Õ„œ ”«·„
   '  printtomanyprinter
            'DoEvents
            
           '  printtomanyprinter2
             
' „ «Ìﬁ«› «·ﬂÊœ Â‰« „ƒﬁ « ·ÕÌ‰ ›Â„Â „‰ «·»—‰” «·ﬂ»Ì— „ «Õ„œ ”«·„
 'WaelComment
DoEvents
    Select Case Me.TxtModFlg.text
    
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
'                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
'                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
'                Msg = " Data Was Saved do you want Another Entry" & Chr(13)
                
            End If
            
            XPBtnMove_Click (2)
            
            If 1 = 2 Then
           '     PrintReport

           '     DoEvents
           '     DoEvents
           '     DoEvents
        
            ElseIf CboPOSBillType.ListIndex <> 4 Then
           '  PrintReport , 1, LblTable.Caption
            PrintReport2
            Cmd_Click (11)
              Cmd_Click (0)
               Screen.MousePointer = vbDefault
              Exit Sub
            End If
            
    '        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton1, App.title) = vbYes Then
                
                Cmd_Click (0)
                Screen.MousePointer = vbDefault
                
    '        Else
    '            TxtModFlg.Text = "R"
    '        End If
'
 
 
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                '    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                '    Msg = " changes Was Saved   & Chr(13)"
    
            End If

            lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
       
            '    Me.Retrive Val(Me.XPTxtBillID.text)
            TxtModFlg.text = "R"
    End Select

    Screen.MousePointer = vbDefault

    'her
    Exit Sub
ErrTrap:

    If TransBegine = True Then
        TransBegine = False
        Cn.RollbackTrans
    End If

    'Resume
    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If Not RsNotes Is Nothing Then
        If RsNotes.EditMode <> adEditNone Then
            RsNotes.CancelUpdate
        End If
    End If

    If Not RSTransDetails Is Nothing Then
        If RSTransDetails.EditMode <> adEditNone Then
            RSTransDetails.CancelUpdate
        End If
    End If

    Screen.MousePointer = vbDefault

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
            Msg = Msg & CHR(13) & Err.Description
            Msg = Msg & CHR(13) & Err.Number
            Msg = Msg & CHR(13) & Err.Source
            Msg = Msg & CHR(13) & Err.LastDllError
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Else
            Msg = "Can't Save error in Data" & CHR(13)
        End If

        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)

        Msg = Msg & CHR(13) & Err.Description
        Msg = Msg & CHR(13) & Err.Number
        Msg = Msg & CHR(13) & Err.Source
        Msg = Msg & CHR(13) & Err.LastDllError
    Else
        Msg = "Sorry........Error During Save " & CHR(13)

    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub XPBtnNewClients_Click()
    On Error GoTo ErrTrap

    'With FrmAddNewCustemer
    '    .DealingForm = InvoiceTransaction
    '    FrmAddNewCustemer.AddType = 1
    '    .Caption = "≈÷«›… ⁄„Ì· ÃœÌœ"
    '    .lbl(1).Caption = "ﬂÊœ «·⁄„Ì·"
    '    .lbl(0).Caption = "«”„ «·⁄„Ì·"
    '    Set .DcboCustomers = DBCboClientName
    '    .show vbModal
    '    cSearchDcbo(0).Refresh
    'End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPCboDiscountType_Change()
    XPCboDiscountType_Click
End Sub

Private Sub XPCboDiscountType_Click()
    On Error GoTo ErrTrap

    If XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = 3 Or XPCboDiscountType.ListIndex = -1 Then
    
        XPTxtDiscountVal.Enabled = False
        XPTxtDiscountVal.text = ""
    Else
    
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.text = ""
    End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If FG.TextMatrix(1, FG.ColIndex("Code")) <> "" Then
            NewGrid.Calculate 1, , , True
        End If
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

    Exit Sub
ErrTrap:
End Sub

Private Sub XPChkPayType_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If XPChkPayType(0).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(0).text = ""
                    XPTxtSerial(0).text = ""
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(0).Enabled = True
                    '                XPTxtSerial(0).Enabled = True
                    XPTxtValue(0).locked = False
                    '                XPTxtSerial(0).Locked = False
                End If

            Else
                XPTxtValue(0).Enabled = False
                XPTxtValue(0).text = ""
                '            XPTxtSerial(0).Enabled = False
            End If

        Case 1

            If XPChkPayType(1).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(1).text = ""
                    XPTxtSerial(1).text = ""
                    DtpDelayDate.value = Date
                    XPTxtSerial(1).text = CStr(new_id("Notes", "NoteSerial", "", True))
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(1).Enabled = True
                    XPTxtValue(1).locked = False
                    DtpDelayDate.Enabled = True
                Else
                    DtpDelayDate.Enabled = False
                End If

                Me.ChkInstall.Enabled = True
            Else
                XPTxtValue(1).Enabled = False
                XPTxtSerial(1).Enabled = False
                XPTxtValue(1).text = ""
                Me.ChkInstall.Enabled = False
            End If

        Case 2

        '    If XPChkPayType(2).value = Checked And Me.TxtModFlg.Text <> "R" Then
        '        Me.CmdCheque.Enabled = True
        '    Else
        '        Me.CmdCheque.Enabled = False
        '        Me.lbl(18).Caption = 0
        '        Me.lbl(19).Caption = 0
        '        Me.FgCheques.Rows = Me.FgCheques.FixedRows
        '    End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub XPChkTAX_Click()
    On Error GoTo ErrTrap

    If XPChkTAX.value = Checked Then
        XPTxtTaxValue.Enabled = True
        lbl(4).Enabled = True
        lbl(45).Enabled = True
    Else
        XPTxtTaxValue.text = ""
        XPTxtTaxValue.Enabled = False
        lbl(4).Enabled = False
        lbl(45).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPDtbBill_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    CurrentVoucherNo = ""
    DateChanged = True
    'updateProfit
End Sub

Private Sub XPTxtDiscountVal_Change()
    Dim Msg As String
    On Error GoTo ErrTrap



    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        NewGrid.Calculate 1, , , True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CustomerPrintReport(Optional PrinterTarget As Boolean = False, _
                        Optional pos As Integer = 0, _
                        Optional sTitle As String, Optional View As Integer = 0, Optional printername As String = "")
    Dim RowNum As Integer
    Dim PayDes As String
    If LBLTable.Caption = "" Then
       Exit Sub
    End If

'‰ﬁ«ÿ «·»Ì⁄
    If View = 0 Then
If CboPOSBillType.ListIndex = 4 Then
'     Cmd_Click (1)
'    Cmd_Click (2)
End If
    DoEvents
    DoEvents
    DoEvents


                 
'                 Cn.Execute "update Transactions set Printed =1   where Transaction_ID=" & val(XPTxtBillID.Text)
                
                 If CboPOSBillType.ListIndex = 4 And val(LblStableID.Caption) <> -1 Then
'                     Cn.Execute "update Stables set Status =Null   where id=" & val(LblStableID.Caption)
'
'                     FillTables
'
                 End If
    TxtPayedValue = 0
 TxtRemainValue.text = 0
    End If
    Dim ShowType As Integer
    'Dim clrep As ClsReportProp
    Dim StrPath As String
    Dim Msg As String
    Dim P_Target As PrintTarget

    On Error GoTo ErrTrap

    'If MDIFrmMain.MnuInvPrintDirect.Checked = True Then
    '    P_Target = PrinterTarget

    'End If
  '  PayDes = ""
  '  For RowNum = 1 To Grid.Rows - 1
  ' If val(Grid.TextMatrix(RowNum, Grid.ColIndex("Value"))) <> 0 Then
  ' If PayDes <> "" Then
  '        PayDes = PayDes & Chr(13) & Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName")) & "  : " & Grid.TextMatrix(RowNum, Grid.ColIndex("value"))
  ' Else
  '         PayDes = Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName")) & "  : " & Grid.TextMatrix(RowNum, Grid.ColIndex("value"))
  'End If
  'If RowNum = Grid.Rows - 1 Then
  'PayDes = PayDes & Chr(13)
  'End If
  'End If
  'Next RowNum
  ' Cn.Execute "update Transactions set PayDes ='" & PayDes & "'   where Transaction_ID=" & val(XPTxtBillID.Text)
    If SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3 Then
        P_Target = PrinterTarget
    Else
        P_Target = WindowTarget
    End If

    ShowType = GetSetting(StrAppRegPath, "View_Type", "SallReportType", 1)

    If ShowType = 1 Then
        If XPTxtBillID.text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingDataDetailed XPTxtBillID.text, 5, , , LblTotal, TxtSearchCode.text, TxtBillComment.text, val(lblInstComm.Caption)
            '    SaleReport.ShowSallingData XPTxtBillID.text, 4, , val(Me.TxtPayedValue.text), val(Me.TxtRemainValue.text), pos, sTitle

            '  If MDIFrmMain.MnuInvPrintReceipt.Checked = True Then
            '      SaleReport.PrintInvoiceReceipt Val(XPTxtBillID.text), P_Target
            '  End If
        End If

    ElseIf (ShowType = 2) Or (ShowType = 4) Then
        '    P_Target = IIf(MDIFrmMain.MnuInvPrintSave.Checked = True, PrintTarget.PrinterTarget, PrintTarget.WindowTarget)

        If SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3 Then
            P_Target = PrinterTarget
        Else
            P_Target = WindowTarget
        End If

        If XPTxtBillID.text <> "" Then
            '     P_Target = WindowTarget
            Set SaleReport = New ClsSaleReport
            'SaleReport.ShowSallingDataShort XPTxtBillID.text, P_Target
            SaleReport.ShowSallingData XPTxtBillID.text, 0, , val(Me.TxtPayedValue.text), val(Me.TxtRemainValue.text), pos, sTitle, printername
        
            '      P_Target = PrinterTarget
        
            'ÿ»«⁄… ≈Ì’«· ≈” ·«„ «·‰ﬁœÌ…
     
        End If

    ElseIf ShowType = 3 Then

        If XPTxtBillID.text <> "" Then
            StrPath = GetSetting(StrAppRegPath, "PrintReport", "ReportPath", App.path & "\Bill_Template\SaleMain.drp")

            If StrPath = "" Then
                Msg = "⁄›Ê« : Â‰«ﬂ Œÿ√„« ›Ì „”«— «· ﬁ—Ì— "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            Set crep = New ClsReportProp
            crep.OpenFile = StrPath
            crep.LoadFile StrPath, FrmPreview
            crep.InvoID = XPTxtBillID.text
            crep.ShowReport
            FrmPreview.show vbModal
            Set crep = Nothing
        End If

    ElseIf ShowType = 5 Then

        If XPTxtBillID.text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingData val(XPTxtBillID.text), 1, val(Me.DBCboClientName.BoundText)

 
        End If

    ElseIf ShowType = 6 Then

        If XPTxtBillID.text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingData val(XPTxtBillID.text), 2, val(Me.DBCboClientName.BoundText)
        
            SaleReport.PrintInvoiceReceipt val(XPTxtBillID.text), P_Target
       
        End If
    End If
If View = 0 Then
  '  clear_all Me
 End If
    Exit Sub
ErrTrap:
End Sub

Private Sub PrintReport(Optional PrinterTarget As Boolean = False, _
                        Optional pos As Integer = 0, _
                        Optional sTitle As String, Optional View As Integer = 0, Optional printername As String = "")
    Dim RowNum As Integer
    Dim PayDes As String
    If LBLTable.Caption = "" Then
'        Exit Sub
    End If

'‰ﬁ«ÿ «·»Ì⁄
    If View = 0 Then
If CboPOSBillType.ListIndex = 4 Then
     Cmd_Click (1)
    Cmd_Click (2)
End If
    DoEvents
    DoEvents
    DoEvents


                 
                 Cn.Execute "update Transactions set Printed =1   where Transaction_ID=" & val(XPTxtBillID.text)
                 
                 
                
                 If CboPOSBillType.ListIndex = 4 And val(LblStableID.Caption) <> -1 Then
                  Cn.Execute "update Transactions set Printed =1   where StableID=" & val(LblStableID.Caption)
                  
                     Cn.Execute "update Stables set Status =Null   where id=" & val(LblStableID.Caption)
                 DoEvents
                     FillTables
                   
                 End If
    TxtPayedValue = val(Me.LBLPayVal)
 TxtRemainValue.text = val(Me.LBLPayVal) - val(Me.TxtNetValue.text)
    End If
    Dim ShowType As Integer
    'Dim clrep As ClsReportProp
    Dim StrPath As String
    Dim Msg As String
    Dim P_Target As PrintTarget

    On Error GoTo ErrTrap

    'If MDIFrmMain.MnuInvPrintDirect.Checked = True Then
    '    P_Target = PrinterTarget

    'End If
  '  PayDes = ""
  '  For RowNum = 1 To Grid.Rows - 1
  ' If val(Grid.TextMatrix(RowNum, Grid.ColIndex("Value"))) <> 0 Then
  ' If PayDes <> "" Then
  '        PayDes = PayDes & Chr(13) & Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName")) & "  : " & Grid.TextMatrix(RowNum, Grid.ColIndex("value"))
  ' Else
  '         PayDes = Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName")) & "  : " & Grid.TextMatrix(RowNum, Grid.ColIndex("value"))
  'End If
  'If RowNum = Grid.Rows - 1 Then
  'PayDes = PayDes & Chr(13)
  'End If
  'End If
  'Next RowNum
  ' Cn.Execute "update Transactions set PayDes ='" & PayDes & "'   where Transaction_ID=" & val(XPTxtBillID.Text)
    If SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3 Then
        P_Target = PrinterTarget
    Else
        P_Target = WindowTarget
    End If

    ShowType = GetSetting(StrAppRegPath, "View_Type", "SallReportType", 1)

    If ShowType = 1 Then
        If XPTxtBillID.text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingDataDetailed XPTxtBillID.text, 4, , , LblTotal, TxtSearchCode.text, TxtBillComment.text, val(lblInstComm.Caption)
            '    SaleReport.ShowSallingData XPTxtBillID.text, 4, , val(Me.TxtPayedValue.text), val(Me.TxtRemainValue.text), pos, sTitle

            '  If MDIFrmMain.MnuInvPrintReceipt.Checked = True Then
            '      SaleReport.PrintInvoiceReceipt Val(XPTxtBillID.text), P_Target
            '  End If
            
            
            
'            If DoPremis(Do_Print, Me.Name, True) = False Then
'                Exit Sub
'            End If
'
'            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)
'
'          If AskOption = False Then
'
'             FrmSallReportOptions.show vbModal
'
'              If FrmSallReportOptions.UserCanceled = True Then
'                   Unload FrmSallReportOptions
'
'             Exit Sub
'               End If
'
'            Unload FrmSallReportOptions
'
'            End If
'
'   updateCopyNo "Transactions", "CopyNO", "Transaction_ID", val(Me.XPTxtBillID.Text)
'           Dim tempbranch As Integer
'           'ÌﬁÊ„ Â‰« »«—”«· —ﬁ„ ««·›—⁄ «·Õ«·Ì ⁄‘«‰ ÿ»«⁄Â «··ÊÃÊ
'          tempbranch = Current_branch
'          Current_branch = val(Me.Dcbranch.BoundText)
'
'
'
'
'        PrintReport
'            Current_branch = tempbranch
'
'   rs.Resync adAffectCurrent
        End If

    ElseIf (ShowType = 2) Or (ShowType = 4) Then
        '    P_Target = IIf(MDIFrmMain.MnuInvPrintSave.Checked = True, PrintTarget.PrinterTarget, PrintTarget.WindowTarget)

        If SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3 Then
            P_Target = PrinterTarget
        Else
            P_Target = WindowTarget
        End If

        If XPTxtBillID.text <> "" Then
            '     P_Target = WindowTarget
            Set SaleReport = New ClsSaleReport
            'SaleReport.ShowSallingDataShort XPTxtBillID.text, P_Target
       '     SaleReport.ShowSallingData XPTxtBillID.Text, 0, , val(Me.TxtPayedValue.Text), val(Me.TxtRemainValue.Text), pos, sTitle, printername
            Dim X As Integer
        
              Set SaleReport = New ClsSaleReport
            'SaleReport.ShowSallingDataShort XPTxtBillID.text, P_Target
            SaleReport.ShowSallingData XPTxtBillID.text, 0, , val(Me.TxtPayedValue.text), val(Me.TxtRemainValue.text), , , , , LblFinal.Caption, , val(Me.dcBranch.BoundText), PayDes, , , X
            'ÿ»«⁄… ≈Ì’«· ≈” ·«„ «·‰ﬁœÌ…
    
    
                Set crep = New ClsReportProp
            crep.OpenFile = StrPath
            
           crep.LoadFile StrPath, FrmPreview
           
            crep.InvoID = XPTxtBillID.text
            crep.ShowReport

            FrmPreview.show vbModal
   
           
            'SaleReport.ShowSallingDataShort XPTxtBillID.text, P_Target
            
            'ÿ»«⁄… ≈Ì’«· ≈” ·«„ «·‰ﬁœÌ…
    
            '      P_Target = PrinterTarget
        
            'ÿ»«⁄… ≈Ì’«· ≈” ·«„ «·‰ﬁœÌ…
  
        End If

    ElseIf ShowType = 3 Then

        If XPTxtBillID.text <> "" Then
            StrPath = GetSetting(StrAppRegPath, "PrintReport", "ReportPath", App.path & "\Bill_Template\SaleMain.drp")

            If StrPath = "" Then
                Msg = "⁄›Ê« : Â‰«ﬂ Œÿ√„« ›Ì „”«— «· ﬁ—Ì— "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            Set crep = New ClsReportProp
            crep.OpenFile = StrPath
            crep.LoadFile StrPath, FrmPreview
            crep.InvoID = XPTxtBillID.text
            crep.ShowReport
            FrmPreview.show vbModal
            Set crep = Nothing
        End If

    ElseIf ShowType = 5 Then

        If XPTxtBillID.text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingData val(XPTxtBillID.text), 1, val(Me.DBCboClientName.BoundText)

 
        End If

    ElseIf ShowType = 6 Then

        If XPTxtBillID.text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingData val(XPTxtBillID.text), 2, val(Me.DBCboClientName.BoundText)
        
            SaleReport.PrintInvoiceReceipt val(XPTxtBillID.text), P_Target
       
        End If
    End If
If View = 0 Then
    clear_all Me
 End If
    Exit Sub
ErrTrap:
End Sub


Private Sub PrintReport2(Optional PrinterTarget As Boolean = False, _
                        Optional reportid As Integer, _
                        Optional AdvPayment As String, _
                        Optional LblInstallCount As String, _
                        Optional LblPrecenValue As String, _
                        Optional LblFirstInstallDate As String)

    Dim ShowType As Integer
    'Dim clrep As ClsReportProp
    Dim StrPath As String
    Dim Msg As String
    Dim P_Target As PrintTarget
    Dim RowNum As Integer
    Dim PayDes As String
    On Error GoTo ErrTrap
    PayDes = ""
'    For RowNum = 1 To GRID2.Rows - 1
'   If val(GRID2.TextMatrix(RowNum, GRID2.ColIndex("Value"))) <> 0 Then
'   If PayDes <> "" Then
'          PayDes = PayDes & CHR(13) & GRID2.TextMatrix(RowNum, GRID2.ColIndex("PaymentName")) & "  : " & GRID2.TextMatrix(RowNum, GRID2.ColIndex("value"))
'   Else
'           PayDes = GRID2.TextMatrix(RowNum, GRID2.ColIndex("PaymentName")) & "  : " & GRID2.TextMatrix(RowNum, GRID2.ColIndex("value"))
'  End If
'  If RowNum = GRID2.Rows - 1 Then
'  PayDes = PayDes & CHR(13)
'  End If
'  End If
'  Next RowNum
Dim Balance As String, balanceString As String
If SystemOptions.ShowBalanceCustInv Then
Dim mAccount As String

    mAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
    WriteCustomerBalPublic mAccount, Balance, balanceString, , , , , , XPDtbBill.value
    
End If
  
      
    
    


    
 
    
  
    If reportid = 15 Then
 
        Dim SaleReport2 As ClsSaleReport
   
        Set SaleReport2 = New ClsSaleReport
        SaleReport2.ShowPrice XPTxtBillID.text, 10, DcboEmp.text
Exit Sub

    
    
    End If
    
    If reportid = 16 Then
  

  Dim SaleReport3 As ClsSaleReport
Set SaleReport3 = New ClsSaleReport
        SaleReport3.ShowPrice XPTxtBillID.text, 16, DcboEmp.text
Exit Sub

    
    End If
    
    
    
    
    'If MDIFrmMain.MnuInvPrintDirect.Checked = True Then
    '    P_Target = PrinterTarget

    'End If
Dim X As Integer
    If SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3 Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                            X = MsgBox("ÿ»«⁄Â „»«‘—Â", vbInformation + vbYesNo)
                Else
                            X = MsgBox(" Direct Print  ", vbInformation + vbYesNo)
                End If
   
   If X = vbYes Then
    P_Target = PrinterTarget
   Else
   P_Target = WindowTarget
   End If
   
        
    Else
        P_Target = WindowTarget
    End If

    ShowType = GetSetting(StrAppRegPath, "View_Type", "SallReportType", 1)

     If reportid = 1 Or reportid = 7 Or reportid = 17 Or reportid = 13 Or reportid = 14 Or reportid = 16 Then
        If XPTxtBillID.text <> "" Then
            Dim SaleReportReturn As New ClsSaleReport
            
            Dim Nationality As String
            Dim GID As Double
        
            GetCustomerAllData val(DBCboClientName.BoundText), , , , , , , , , , , , , , , Nationality, , GID
                    
               
            If reportid = 7 Then
                Set SaleReport = New ClsSaleReport
                SaleReport.ShowSallingDataDetailed XPTxtBillID.text, , , , LblFinal.Caption, TxtSearchCode.text, TxtBillComment.text, val(lblInstComm.Caption), reportid, Nationality, GID, XPDtbBill.value, AdvPayment, LblInstallCount, LblPrecenValue, LblFirstInstallDate, , , , , val(DcCurrency.BoundText), , , , , Balance, 0
            Else
                 SaleReportReturn.ShowSallingDataDetailed XPTxtBillID.text, , , , LblFinal.Caption, TxtSearchCode.text, TxtBillComment.text, val(lblInstComm.Caption), reportid, Nationality, GID, XPDtbBill.value, AdvPayment, LblInstallCount, LblPrecenValue, LblFirstInstallDate, , , , , val(DcCurrency.BoundText), , , , , Balance
            End If
        End If
    
        Exit Sub
    End If

    If reportid = 2 Then
        If XPTxtBillID.text <> "" Then
            Dim SaleReport6 As New ClsSaleReport
            
 
            Dim Madyna As String
            Dim hay As String
            Dim Address As String
        
            GetCustomerAllData val(DBCboClientName.BoundText), , , , , , , , , , , , , Madyna, hay, Nationality, , GID, Address
            SaleReport6.ShowSallingDataDetailed XPTxtBillID.text, , , , LblTotal, TxtSearchCode.text, TxtBillComment.text, val(lblInstComm.Caption), 2, Nationality, GID, XPDtbBill.value, AdvPayment, LblInstallCount, LblPrecenValue, LblFirstInstallDate, Madyna, hay, Address, val(val(lblInstComm.Caption) + val(LblTotal.Caption)) - val(AdvPayment), val(DcCurrency.BoundText), , , val(dcBranch.BoundText), , Balance
 
        End If
    
        Exit Sub
    End If

    If ShowType = 1 Then
        If XPTxtBillID.text <> "" Then
            
           ' Dim SaleReport7 As New ClsSaleReport
           
 
            '  If MDIFrmMain.MnuInvPrintReceipt.Checked = True Then
            '      SaleReport.PrintInvoiceReceipt Val(XPTxtBillID.text), P_Target
            '  End If
            
            
            
            
                 '  If XPTxtBillID.Text <> "" Then
            '     P_Target = WindowTarget
            Set SaleReport = New ClsSaleReport
            'SaleReport.ShowSallingDataShort XPTxtBillID.text, P_Target
            SaleReport.ShowSallingData XPTxtBillID.text, 0, , val(Me.TxtPayedValue.text), val(Me.TxtRemainValue.text), , , , , LblFinal.Caption, , val(Me.dcBranch.BoundText), PayDes, , , X
            'ÿ»«⁄… ≈Ì’«· ≈” ·«„ «·‰ﬁœÌ…
    
       ' End If
        
        End If
    
    ElseIf ShowType = 40 Then

        If XPTxtBillID.text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingDataDetailed2 XPTxtBillID.text, , , , LblFinal.Caption, TxtSearchCode.text, TxtBillComment.text, val(lblInstComm.Caption), val(DcCurrency.BoundText), , val(dcBranch.BoundText), PayDes
        
            '  If MDIFrmMain.MnuInvPrintReceipt.Checked = True Then
            '      SaleReport.PrintInvoiceReceipt Val(XPTxtBillID.text), P_Target
            '  End If
        End If
    
    ElseIf (ShowType = 2) Or (ShowType = 4) Then
        '    P_Target = IIf(MDIFrmMain.MnuInvPrintSave.Checked = True, PrintTarget.PrinterTarget, PrintTarget.WindowTarget)

        If SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3 Then
'            P_Target = PrinterTarget
   If X = vbYes Then
    P_Target = PrinterTarget
   Else
   P_Target = WindowTarget
   End If
   
        Else
            P_Target = WindowTarget
        End If

        If XPTxtBillID.text <> "" Then
            '     P_Target = WindowTarget
            Set SaleReport = New ClsSaleReport
            'SaleReport.ShowSallingDataShort XPTxtBillID.text, P_Target
            SaleReport.ShowSallingData XPTxtBillID.text, 0, , val(Me.TxtPayedValue.text), val(Me.TxtRemainValue.text), , , , , LblFinal.Caption, , val(Me.dcBranch.BoundText), PayDes, , , X
            'ÿ»«⁄… ≈Ì’«· ≈” ·«„ «·‰ﬁœÌ…
    
        End If

    ElseIf ShowType = 3 Then

        If XPTxtBillID.text <> "" Then
            StrPath = GetSetting(StrAppRegPath, "PrintReport", "ReportPath", App.path & "\Bill_Template\SaleMain.drp")

            If StrPath = "" Then
                Msg = "⁄›Ê« : Â‰«ﬂ Œÿ√„« ›Ì „”«— «· ﬁ—Ì— "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            Set crep = New ClsReportProp
            crep.OpenFile = StrPath
            
           crep.LoadFile StrPath, FrmPreview
           
            crep.InvoID = XPTxtBillID.text
            crep.ShowReport

            FrmPreview.show vbModal
   
            Set crep = Nothing
        End If

    ElseIf ShowType = 5 Then

        If XPTxtBillID.text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingData val(XPTxtBillID.text), 1, val(Me.DBCboClientName.BoundText)

    
        End If

    ElseIf ShowType = 6 Then

        If XPTxtBillID.text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingData val(XPTxtBillID.text), 2, val(Me.DBCboClientName.BoundText)
        
            SaleReport.PrintInvoiceReceipt val(XPTxtBillID.text), P_Target
       
        End If
    End If

    Exit Sub
ErrTrap:
End Sub




Private Sub XPTxtDiscountVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtDiscountVal.text, 0)
End Sub

Private Sub XPTxtSum_Change()
    On Error GoTo ErrTrap

    If CboPayMentType.ListIndex = 0 Then
        XPChkPayType(0).value = Checked
        XPTxtValue(0).text = XPTxtSum.text
    End If

    Me.LblTotal.Caption = XPTxtSum.text
    CalculateInvPrecent
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
               ' Unload customer_screen

            Case vbCancel
                Cancel = True
             '   Unload customer_screen
        End Select

   '     Unload customer_screen
    End If

    Exit Sub
ErrTrap:
End Sub

Public Sub Convert()
    Cmd_Click (0)
End Sub

Public Sub Cala()
    NewGrid.Calculate 1, , , True
End Sub

Private Sub DBCboClientName_Change()
    Dim Msg As String
    Dim RsTemp  As ADODB.Recordset
    Dim StrSQL As String

    On Error GoTo ErrTrap
    Dim Fullcode As String
 
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 1
    TxtSearchCode.text = Fullcode

    If val(DBCboClientName.BoundText) <> 0 Then
        If (DBCboClientName.BoundText = 1 Or DBCboClientName.BoundText = 2) And Me.TxtModFlg.text <> "R" Then
            CboPayMentType.locked = True
            '        CboPaymentType.ListIndex = 0
            Me.TxtCashCustomerName.Enabled = True
            Me.CmdCash(0).Enabled = True
            Me.CmdCash(1).Enabled = True
        Else
            CboPayMentType.locked = False
            Me.TxtCashCustomerName.Enabled = False
            Me.CmdCash(0).Enabled = False
            Me.CmdCash(1).Enabled = False
        End If
    
        If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            Dim DefaultSalesPersonId As Integer
            GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId

            If Not DefaultSalesPersonId = 0 Then

                Me.DcboEmp.BoundText = DefaultSalesPersonId
            End If

            StrSQL = "Select * From TblCustemers Where CusID=" & val(DBCboClientName.BoundText)
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                If Not (IsNull(RsTemp("SaleType").value)) Then
                    If RsTemp("SaleType").value = 0 Then
                        Me.CboSaleType.ListIndex = 0
                    ElseIf RsTemp("SaleType").value = 1 Then
                        Me.CboSaleType.ListIndex = 1
                    End If

                Else
                    Me.CboSaleType.ListIndex = -1
                End If

                If Not (IsNull(RsTemp("Trans_DiscountType").value)) Then
                    If RsTemp("Trans_DiscountType").value = 0 Then
                        '                 mina   Me.XPCboDiscountType.ListIndex = 0
                        '                 mina   Me.XPTxtDiscountVal.text = 0
                    ElseIf RsTemp("Trans_DiscountType").value = 1 Then
                        Me.XPCboDiscountType.ListIndex = 1
                        Me.XPTxtDiscountVal.text = IIf(IsNull(RsTemp("Trans_Discount").value), "", RsTemp("Trans_Discount").value)
                    ElseIf RsTemp("Trans_DiscountType").value = 2 Then
                        Me.XPCboDiscountType.ListIndex = 2
                        Me.XPTxtDiscountVal.text = IIf(IsNull(RsTemp("Trans_Discount").value), "", RsTemp("Trans_Discount").value)
                    End If

                Else
                    Me.XPCboDiscountType.ListIndex = 0
                    Me.XPTxtDiscountVal.text = 0
                End If

            Else
                Me.CboSaleType.ListIndex = -1
                Me.XPCboDiscountType.ListIndex = 0
                Me.XPTxtDiscountVal.text = 0
            End If

            RsTemp.Close
            Set RsTemp = Nothing
        End If
    End If

    FillVoucherGrid
    FillOrderGrid
    Exit Sub
ErrTrap:
    Msg = Err.Description & CHR(13) & ""
    Msg = Msg & Err.Source & CHR(13) & ""
    Msg = Msg & Me.Name & " DBCboClientName_Change:" & CHR(13) & ""
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub XPTxtValue_Change(Index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If XPTxtValue(1).text <> "" Then
            If val(Me.XPTxtValue(1).text) > 0 Then
                ChkInstall.Enabled = True
            End If

        End If
    End If

    'If XPChkPayType(1).Value = 1 Then
    '            XPTxtValue(0).text = Val(LblTotal.Caption) - Val(XPTxtValue(1).text)
    'End If
    'If XPChkPayType(0).Value = 1 Then
    '    XPTxtValue(1).text = Val(LblTotal.Caption) - Val(XPTxtValue(0).text)
    'End If
    Exit Sub
ErrTrap:
End Sub

Public Sub ReplacementData()
    Dim Msg As String
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsReplace As ADODB.Recordset

    If Me.TxtModFlg.text <> "R" Then Exit Sub

    '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«· «·Œ«’… »«·›« Ê—…
    If FG.TextMatrix(FG.Row, FG.ColIndex("Code")) <> "" And FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) <> "" Then
        StrSQL = "select * From ReplacedItems where ReturnID=" & XPTxtBillID.text
        StrSQL = StrSQL + " and ItemID=" & FG.TextMatrix(FG.Row, FG.ColIndex("Code"))
        StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & "'"
        Set RsReplace = New ADODB.Recordset
        RsReplace.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsReplace.EOF Or RsReplace.BOF) Then
            Msg = "·ﬁœ  „ «” »œ«· «·ﬁÿ⁄… : " & FG.Cell(flexcpTextDisplay, FG.Row, FG.ColIndex("Name")) & CHR(13)
            Msg = Msg + "–«  «·”Ì—Ì«· : " & FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & CHR(13)
            Msg = Msg + " »«·ﬁÿ⁄… –«  «·”Ì—Ì«· : " & RsReplace("newSerial").value & CHR(13)
            Msg = Msg + "›Ì ⁄„·Ì… ’Ì«‰…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, "ﬁÿ⁄…  „ «” »œ«·Â«"
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Function AvailableDeal() As Boolean
    'On Error GoTo ErrTrap
    Dim RowNum As Integer
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RsSalle As ADODB.Recordset
    Dim LngItemID As Long

    For RowNum = 1 To FG.Rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            StrSQL = "select * From QryDelPurchase where Transaction_Date >=" & SQLDate(XPDtbBill.value, True) & ""
            StrSQL = StrSQL + " and Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))
            StrSQL = StrSQL + " and Transaction_Type=9"

            If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                End If
            End If

            Set RsSalle = New ADODB.Recordset
            RsSalle.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsSalle.EOF Or RsSalle.BOF) Then
                If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then

                    '                StrSql = "select * From QryGardComplete where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                    '                StrSql = StrSql + " AND ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                    '                StrSql = StrSql + " AND StoreID=" & DCboStoreName.BoundText
                    '                Set RsTemp = New ADODB.Recordset
                    '                RsTemp.Open StrSql, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    '                If RsTemp.EOF Or RsTemp.BOF Then
                    With FrmAlarm
                        .DealingForm = InvoiceTransaction
                        .show vbModal
                    End With

                    AvailableDeal = False
                    Exit Function
                    '                End If
                    RsTemp.Close
                Else
                    LngItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                    Set RsTemp = New ADODB.Recordset
                    Set RsTemp = GetItemQuantityStock(LngItemID, Me.DCboStoreName.BoundText, Me.XPDtbBill.value, val(Me.XPTxtBillID.text))

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        If val(RsTemp("totalqty").value) < val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) Then

                            With FrmAlarm
                                .DealingForm = InvoiceTransaction
                                .show vbModal
                            End With

                            AvailableDeal = False
                            Exit Function
                        End If
                    End If

                    RsTemp.Close
                End If
            End If

            RsSalle.Close
        End If

    Next RowNum

    AvailableDeal = True
    Exit Function
ErrTrap:
End Function

Private Sub SetDefaults()
    On Error Resume Next
    Dim StrTemp As String
    Dim RsTemp As ADODB.Recordset

    Me.CboSaleType.ListIndex = 0

    If SystemOptions.SysPurDateTakeType = InvDateFromLocalCompuer Then
        XPDtbBill.value = Date
    ElseIf SystemOptions.SysPurDateTakeType = InvDateFromServerComputer Then
        StrTemp = "select Getdate() as ServerDate"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrTemp, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If Not IsNull(RsTemp("ServerDate").value) Then
                XPDtbBill.value = Format(RsTemp("ServerDate").value, "yyyy/M/d")
            End If

            'XPDtbBill.Value = IIf(IsNull(RsTemp("ServerDate").Value), Date, (RsTemp("ServerDate").Value))
        End If


        RsTemp.Close
        Set RsTemp = Nothing
    End If

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast

        If SystemOptions.SysPurDateTakeType = InvDateFromLastInvDate Then
            XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), Date, (rs("Transaction_Date").value))
        End If

        Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)

        If Not IsNull(rs("Transaction_Serial").value) Then
            StrTemp = rs("Transaction_Serial").value
            StrTemp = val(StrTemp) + 1
            TxtTransSerial.text = StrTemp
        Else
            TxtTransSerial.text = 1
        End If

    Else
        TxtTransSerial.text = 1
    End If

    DCPaymentNet.BoundText = 1

Dim Hour As String
Hour = mId(Time$(), 1, 2)

If Hour >= 0 And Hour <= 5 Then
'MsgBox HOUR
Dim NewDate As Date
XPDtbBill.value = DateAdd("d", -1, Date)
 
End If

End Sub

Private Sub CalculateInvPrecent()
    Dim DblInvTotal As Double
    Dim DblInvProfit As Double
    Dim DblRes As Double

    DblInvProfit = val(Me.LblInvProfit.Caption)
    DblInvTotal = val(Me.XPTxtSum.text)

    If DblInvProfit = 0 Or DblInvTotal = 0 Then
        DblRes = 0
    Else
        DblRes = 100 * (DblInvProfit / DblInvTotal)
    End If

    Me.lblInvPrecent.Caption = "%" & CStr(Int(DblRes)) 'Format(DblRes, SystemOptions.SysDefCurrencyForamt)
End Sub

Private Sub LoadCombosData()
    Dim StrSQL As String
    Dcombos.GetPaymentType Me.DCPaymentNet
    Dcombos.GetSalesRepData Me.DcboEmp
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBranches Me.dcBranch
Dcombos.GetSalesRepData Me.DcboEmp1

    Dcombos.GetDocTypebyid Me.DCDocTypes, 21, val(Me.dcBranch.BoundText)

    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
    Dcombos.GetStores Me.DCboStoreName

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName
    cSearchDcbo(0).SetBuddyText Me.TxtCusID

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboStoreName
    cSearchDcbo(1).SetBuddyText Me.TxtStoreID

    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DcboEmp
    cSearchDcbo(3).SetBuddyText Me.TxtEmployeeID

   ' StrSQL = "  select  BankID,BankName  from BanksData   "
   ' fill_combo Dcbanks, StrSQL
      
End Sub
Sub SaveValueAdded()
Dim i As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

sql = "Select * from  TransactionValueAdded where 1=-1"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With Me.VatGrid
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
rs2.AddNew
rs2("Transaction_ID").value = val(Me.XPTxtBillID.text)
rs2("Transaction_Type").value = 21
rs2("ItemID").value = val(.TextMatrix(i, .ColIndex("ItemID")))
rs2("Vatyo").value = val(.TextMatrix(i, .ColIndex("Vatyo")))
rs2("Vat").value = val(.TextMatrix(i, .ColIndex("Vat")))
rs2("Valu").value = val(.TextMatrix(i, .ColIndex("Valu")))
If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
rs2("selectd").value = 1
Else
rs2("selectd").value = 1
End If

rs2.update
End If
Next i
End With
End Sub
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    lbl(1).Caption = "Card"
    lbl(89).Caption = "Bala.Points"
    lbl(86).Caption = "Oper.Points"
    lbl(88).Caption = "Customer"
  Command5.Caption = "Tables"
  Command6.Caption = "Admin Login"
  Command7.Caption = "Discount"
    Cmd(13).Caption = "Print"
 '   Label1(9).Caption = "Dine In"
 '   Label1(10).Caption = "Take OUT"
 '     Label1(11).Caption = "Delivery"
 '   Label1(12).Caption = "Car"
    
    Label1(4).Caption = "User"
    Label1(1).Caption = "TABLE"
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(57).Caption = "Cash.visa"
    Label3.Caption = "Branch"
    Frame1.Caption = "Info"
    lbl(56).Caption = "Order No."
    lbl(58).Caption = " Total"
    lbl(59).Caption = " Payed"
    lbl(60).Caption = " Changed"
    lbl(63).Caption = " Total Qty"
    Frame2.Caption = "Color Map"
    lbl(68).Caption = " Profit"
            Label11.Caption = "Tables"
            lbl(80).Caption = "Groups"
            Label13.Caption = "Items"
 
 Cmd(8).Caption = "Exit"
 
 Cmd(11).Caption = "Kitc."
 Cmd(9).Caption = "Print"
 With Me.VatGrid
 .TextMatrix(0, .ColIndex("index")) = "Serial"
.TextMatrix(0, .ColIndex("select")) = "Select"
.TextMatrix(0, .ColIndex("Code")) = "Item Code"
.TextMatrix(0, .ColIndex("Name")) = "Item Name"
.TextMatrix(0, .ColIndex("Vatyo")) = "Percentage"
.TextMatrix(0, .ColIndex("Vat")) = "Value"
.TextMatrix(0, .ColIndex("Valu")) = "Item Value"
End With
 With Me.Grid
 .TextMatrix(0, .ColIndex("PaymentName")) = "Payment Name"
.TextMatrix(0, .ColIndex("Value")) = "Value"
.TextMatrix(0, .ColIndex("CardNo")) = "Card No"
End With

    'Label1.Caption = "Doc Type"
    lbl(65).Caption = "Curr"
    lbl(66).Caption = "Rec No"
    lbl(67).Caption = "Manf No"
    Label6.Caption = "Price<cost"
    Label8.Caption = "Price=cost"
    Me.XPTab301.TabCaption(3) = "Attachments"
    Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
    'Command2.Caption = "Convert to I. Voucher"
    Me.Caption = "Sales Invoice"
    Ele(9).Caption = Me.Caption
    lbl(5).Caption = "Invoice ID"
    lbl(6).Caption = "Invoice Date"
    lbl(7).Caption = "Customer Name"
    lbl(24).Caption = "Store Name"
    lbl(25).Caption = "Employee"
    lbl(9).Caption = "Cash/Credit"
    lbl(10).Caption = "Dis. Type"
    lbl(8).Caption = "Value"
    lbl(22).Caption = "Profit Value"
    lbl(23).Caption = "Profit Perce"
lbl(84).Caption = "Qty"
lbl(85).Caption = "Price"
    lbl(3).Caption = "Total:"
    lbl(49).Caption = "Net "
    lbl(50).Caption = "Disc"
    'lbl(1).Caption = "By:"
    lbl(2).Caption = "Rec. Count:"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = "Item Case"
    lbl(28).Caption = "Item Serial"
    lbl(27).Caption = "Quantity"
    lbl(26).Caption = "Price"
    lbl(32).Caption = "Sales Type"
    lbl(33).Caption = "Customer Name"
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Pay"
    Me.CmdHelp.Caption = "Help"
    Me.XPTab301.TabCaption(0) = "Items"
    
    Me.XPTab301.TabCaption(1) = "Notes"
    lbl(20).Caption = "Payment Method"
    XPChkPayType(0).Caption = "Cahs"
    XPChkPayType(1).Caption = "Due"
    XPChkPayType(0).Caption = "Check"
    lbl(13).Caption = "Value"
    lbl(15).Caption = "Value"
    lbl(16).Caption = "Value"
    lbl(12).Caption = "Serial"
    lbl(14).Caption = "Serial"
    lbl(11).Caption = "Box Name"
    lbl(21).Caption = "Due Date"
    Label14.Caption = "Order"
    Label15.Caption = "Delivery"
    Label16.Caption = "By Cars"
    lbl(69).Caption = "Totals"
    lbl(71).Caption = "Nets"
    
    
    
    
 '   lbl(18).Caption = "Check NO."
  '  lbl(17).Caption = "Bank Name"
 '   lbl(19).Caption = "Due Date"
 '   CmdINSTALLMENT.Caption = "INSTALLMENT"
    Me.XPTab301.TabCaption(2) = "Comment On Invoice"
    Me.Ele(15).Caption = "Write any Comments about this Invoice"
    
    lbl(44).Caption = "Comment"
    XPChkPayType(0).Caption = "Cash"
    lbl(13).Caption = "Value"
    lbl(12).Caption = "ID"
    lbl(2).Caption = "Box"
    lbl(20).Caption = "Currency"
    XPChkPayType(1).Caption = "Credit"
    lbl(15).Caption = "Value"
    lbl(14).Caption = "ID"
    'Label1.Caption = "Due Date"
    ChkInstall.Caption = "Installment"
    CmdINSTALLMENT.Caption = "Calc"
    Label2.Caption = "Bank"
   ' CmdCheque.Caption = "Register"

    With FgInstallments
        .TextMatrix(0, .ColIndex("QestID")) = "ID"
        .TextMatrix(0, .ColIndex("value")) = "value"
        .TextMatrix(0, .ColIndex("Due_Date")) = "Due_Date"
 
    End With

    With FG
        .TextMatrix(0, .ColIndex("order_no")) = "ORD/INV NO."
 
    End With

  '  With FgCheques
 '
 '       .TextMatrix(0, .ColIndex("CheckValue")) = "Value"
 '       .TextMatrix(0, .ColIndex("CheckNumber")) = "Cheque Number"
 '       .TextMatrix(0, .ColIndex("BankName")) = "Bank Name"
 '       .TextMatrix(0, .ColIndex("DueDate")) = "Due Date"
 '       .TextMatrix(0, .ColIndex("ReleaseDate")) = "Release Date"
 '
 '   End With

  '  XPChkPayType(2).Caption = "Cheques"
    '«·ÊﬁÊ› ⁄‰œ «·«Ê—«ﬁ «·„«·ÌÂ
    lbl(61).Caption = "Bill type"
    BillBasedOn(0).Caption = "Direct Sales Invoices"
    BillBasedOn(1).Caption = "Based On Issue Vouchere"
    BillBasedOn(2).Caption = "Based On Purchase Orders"

    With Me.GRID1
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("noteserial1")) = "Voucher NO"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Voucher Date"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "JE Voucher NO"
    End With

    With Me.GRID2
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("order_no")) = "Order No"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Voucher Date"
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
    End With

    Frame3.Caption = "JE Voucher NO"
    lbl(62).Caption = "JE Voucher NO"
    Cmd(10).Caption = "Print JE"
FramePay.Caption = "Pay"
CMDPAy.Caption = "Pay"
lbl(16).Caption = "VAT"
End Sub

Private Sub XPTxtValue_KeyPress(Index As Integer, _
                                KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtValue(Index).text, 0)
End Sub

Private Function CheckCashCustomer() As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If Trim$(Me.TxtCashCustomerName.text) = "" Then
        CheckCashCustomer = True
    Else
        StrSQL = "Select * From Transactions Where CashCustomerName='" & Trim$(Me.TxtCashCustomerName.text) & "'"
    
    End If

End Function

Private Sub XPTxtValue_MouseMove(Index As Integer, _
                                 Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)

    If val(Me.XPTxtValue(Index).text) <> 0 Then
        Me.XPTxtValue(Index).ToolTipText = WriteNo(Me.XPTxtValue(Index).text, 1, True)
    Else
        Me.XPTxtValue(Index).ToolTipText = ""
    End If

End Sub

Private Sub SumChecks()

     
End Sub

Private Sub ClearNotes()

   ' LblPrecenType.Caption = 0
   ' LblPrecenValue.Caption = 0
   ' LblInstallTotal.Caption = 0
   ' LblInstallCount.Caption = 0
   ' LblFirstInstallDate.Caption = ""
   ' LblInstallSeprator.Caption = ""
   ' LblInstallmentType.Caption = ""
   ' LblStartValue.Caption = ""
   ' Me.LblDiscount.Caption = 0
   ' Me.LblAdvPayment.Caption = 0
   ' lbl(19).Caption = ""
   ' lbl(18).Caption = ""
End Sub

Function LoadImageToList(rs As ADODB.Recordset, Row) As Boolean
    On Error GoTo EmptyImgeLBL
    Dim TempPath    As String
    Dim TempFile    As String
    Dim slength     As Long
    Dim lastfour    As Long
    Dim tmpFileName As String
    Dim oPict       As StdPicture
    Dim sDir        As String
    Dim sTempFile   As String
    Dim iFileNum    As Integer
    Dim lFileLength As Long
    Dim abBytes()   As Byte
    Dim iCtr        As Integer
  
    '*****************************************
    If (IsNull(rs("ItemPhoto").value) Or rs("ItemPhoto").value & "" = "") Then
        'ImageListItems.ListImages.Add , val(rs!ItemID) & "X", LoadPicture(App.path & "\images\pos\blue.JPG")
        ' Exit Function
        GoTo EmptyImgeLBL
    End If

    '******************************************
  
    TempPath = Space(255)
    slength = GetTempPath(255, TempPath)
    TempPath = left(TempPath, slength)
   
    TempFile = Space(255)
    lastfour = GetTempFileName(TempPath, rs!ItemID & "X", 0, TempFile)
    tmpFileName = left(TempFile, InStr(TempFile, vbNullChar) - 1)
 
    sTempFile = tmpFileName

    If Len(Dir$(sTempFile)) > 0 Then
        On Error Resume Next
        Kill sTempFile
    End If

    iFileNum = FreeFile
    Open sTempFile For Binary As #iFileNum
    lFileLength = LenB(rs("ItemPhoto") & "")

    If lFileLength = 0 Then
        Exit Function
    End If

    abBytes = rs("ItemPhoto").GetChunk(lFileLength)
    Put #iFileNum, , abBytes()
    Close #iFileNum
    
    ImageListItems.ListImages.Add Row, val(rs!ItemID) & "X", LoadPicture(sTempFile)
    LoadImageToList = True
    Exit Function
EmptyImgeLBL:
    ImageListItems.ListImages.Add Row, val(rs!ItemID) & "X", LoadPicture(App.path & "\images\pos\blue.JPG")
End Function

