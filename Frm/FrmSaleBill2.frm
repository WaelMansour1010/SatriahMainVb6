VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbaListView6.ocx"
Begin VB.Form frmsalebill2 
   BackColor       =   &H00000003&
   BorderStyle     =   0  'None
   Caption         =   "›« Ê—… «·„»Ì⁄« "
   ClientHeight    =   13545
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   18840
   HelpContextID   =   160
   Icon            =   "FrmSaleBill2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   13545
   ScaleWidth      =   18840
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
      Height          =   13545
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   18840
      _cx             =   33232
      _cy             =   23892
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
      AutoSizeChildren=   0
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   12525
         Left            =   15
         RightToLeft     =   -1  'True
         TabIndex        =   193
         Top             =   -45
         Width           =   19125
         Begin VB.ComboBox DefaultInvoicetype 
            Height          =   315
            ItemData        =   "FrmSaleBill2.frx":038A
            Left            =   6810
            List            =   "FrmSaleBill2.frx":038C
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   427
            Top             =   120
            Width           =   1740
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
            Height          =   8655
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   383
            Top             =   4650
            Visible         =   0   'False
            Width           =   13455
            Begin VB.PictureBox Picture1 
               Height          =   3735
               Left            =   9840
               ScaleHeight     =   3675
               ScaleWidth      =   3435
               TabIndex        =   419
               Top             =   4800
               Visible         =   0   'False
               Width           =   3495
            End
            Begin VB.TextBox TXtCopon 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   418
               Top             =   6480
               Visible         =   0   'False
               Width           =   2055
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
               Left            =   15120
               TabIndex        =   417
               Top             =   480
               Visible         =   0   'False
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
               Left            =   14760
               TabIndex        =   416
               Top             =   2640
               Visible         =   0   'False
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
               Height          =   555
               Index           =   2
               Left            =   240
               TabIndex        =   415
               Top             =   6840
               Width           =   1335
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
               Height          =   555
               Index           =   3
               Left            =   1560
               TabIndex        =   414
               Top             =   6840
               Width           =   1455
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
               Height          =   555
               Index           =   4
               Left            =   3000
               TabIndex        =   413
               Top             =   6840
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
               Height          =   555
               Index           =   5
               Left            =   4230
               TabIndex        =   412
               Top             =   6840
               Width           =   1335
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
               Height          =   555
               Index           =   6
               Left            =   240
               TabIndex        =   411
               Top             =   7320
               Width           =   1335
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
               Height          =   555
               Index           =   7
               Left            =   1560
               TabIndex        =   410
               Top             =   7320
               Width           =   1455
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Height          =   1935
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   403
               Top             =   4440
               Width           =   3840
               Begin VB.TextBox TxtRemainValue 
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
                  Height          =   555
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   406
                  Top             =   1320
                  Width           =   2445
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
                  TabIndex        =   405
                  Top             =   840
                  Width           =   2445
               End
               Begin VB.TextBox TxtNetValue 
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
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   404
                  Top             =   240
                  Width           =   2460
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
                  TabIndex        =   409
                  Top             =   1440
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
                  TabIndex        =   408
                  Top             =   840
                  Width           =   855
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
                  TabIndex        =   407
                  Top             =   240
                  Width           =   855
               End
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
               Height          =   555
               Index           =   8
               Left            =   4200
               TabIndex        =   402
               Top             =   7320
               Width           =   1335
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
               Height          =   555
               Index           =   9
               Left            =   3000
               TabIndex        =   401
               Top             =   7320
               Width           =   1215
            End
            Begin VB.Frame Frame13 
               BackColor       =   &H00FFFFFF&
               Height          =   5055
               Left            =   120
               TabIndex        =   384
               Top             =   480
               Width           =   5535
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   0
                  Left            =   4320
                  TabIndex        =   385
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
                  ButtonImage     =   "FrmSaleBill2.frx":038E
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   1
                  Left            =   2160
                  TabIndex        =   386
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
                  ButtonImage     =   "FrmSaleBill2.frx":0B4E
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   2
                  Left            =   3240
                  TabIndex        =   387
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
                  ButtonImage     =   "FrmSaleBill2.frx":1150
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   3
                  Left            =   4320
                  TabIndex        =   388
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
                  ButtonImage     =   "FrmSaleBill2.frx":1937
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   4
                  Left            =   2160
                  TabIndex        =   389
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
                  ButtonImage     =   "FrmSaleBill2.frx":214C
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   5
                  Left            =   3240
                  TabIndex        =   390
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
                  ButtonImage     =   "FrmSaleBill2.frx":28D7
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   6
                  Left            =   4320
                  TabIndex        =   391
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
                  ButtonImage     =   "FrmSaleBill2.frx":3096
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   7
                  Left            =   2160
                  TabIndex        =   392
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
                  ButtonImage     =   "FrmSaleBill2.frx":3830
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   8
                  Left            =   3240
                  TabIndex        =   393
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
                  ButtonImage     =   "FrmSaleBill2.frx":3F33
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   9
                  Left            =   4320
                  TabIndex        =   394
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
                  ButtonImage     =   "FrmSaleBill2.frx":474E
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   10
                  Left            =   3240
                  TabIndex        =   395
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
                  ButtonImage     =   "FrmSaleBill2.frx":4EDD
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   11
                  Left            =   2160
                  TabIndex        =   396
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
                  ButtonImage     =   "FrmSaleBill2.frx":5A24
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   12
                  Left            =   120
                  TabIndex        =   397
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
                  ButtonImage     =   "FrmSaleBill2.frx":5F16
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   13
                  Left            =   1200
                  TabIndex        =   398
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
                  ButtonImage     =   "FrmSaleBill2.frx":677D
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   2895
                  Index           =   14
                  Left            =   120
                  TabIndex        =   399
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
                  ButtonImage     =   "FrmSaleBill2.frx":6E8E
                  ButtonImageDisabled=   "FrmSaleBill2.frx":823C
                  ColorButton     =   16777215
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
                  TabIndex        =   400
                  Top             =   240
                  Width           =   3375
               End
               Begin VB.Image Image13 
                  Height          =   1035
                  Left            =   120
                  Picture         =   "FrmSaleBill2.frx":85D7
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   5295
               End
            End
            Begin ImpulseButton.ISButton CMDPAy 
               Height          =   615
               Index           =   0
               Left            =   240
               TabIndex        =   420
               Top             =   5565
               Width           =   5295
               _ExtentX        =   9340
               _ExtentY        =   1085
               Caption         =   "”œ«œ+«·ÿ»«⁄…"
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
               ButtonImage     =   "FrmSaleBill2.frx":898D
               ColorHoverText  =   16777215
               ColorToggledText=   16777215
               ColorToggledHoverText=   16777215
               AlignmentIgnoreImage=   -1  'True
            End
            Begin VSFlex8UCtl.VSFlexGrid Grid 
               Height          =   3885
               Left            =   5820
               TabIndex        =   421
               Top             =   600
               Width           =   7485
               _cx             =   13203
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
               Cols            =   12
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   650
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSaleBill2.frx":8F07
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
            Begin VSFlex8Ctl.VSFlexGrid FgC 
               Height          =   1755
               Left            =   5760
               TabIndex        =   422
               Top             =   6840
               Width           =   3825
               _cx             =   6747
               _cy             =   3096
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   16776960
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
               Rows            =   3
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmSaleBill2.frx":90F1
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   0   'False
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
            Begin ImpulseButton.ISButton CMDPAy 
               Height          =   615
               Index           =   1
               Left            =   240
               TabIndex        =   423
               Top             =   6240
               Width           =   5295
               _ExtentX        =   9340
               _ExtentY        =   1085
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
               ButtonImage     =   "FrmSaleBill2.frx":9266
               ColorHoverText  =   16777215
               ColorToggledText=   16777215
               ColorToggledHoverText=   16777215
               AlignmentIgnoreImage=   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰›Ì– «·ﬁ”«∆„"
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
               Index           =   94
               Left            =   7920
               TabIndex        =   426
               Top             =   6360
               Width           =   1815
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
               TabIndex        =   425
               Top             =   120
               Width           =   1335
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
               Left            =   9000
               TabIndex        =   424
               Top             =   240
               Width           =   570
            End
         End
         Begin VB.TextBox TxtPhone 
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
            Index           =   3
            Left            =   12240
            TabIndex        =   377
            Top             =   360
            Width           =   2145
         End
         Begin VB.TextBox TXTPrintInvoice 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   5400
            TabIndex        =   376
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox CashCustomerName 
            Height          =   495
            Left            =   7920
            TabIndex        =   375
            Top             =   2280
            Width           =   2175
         End
         Begin VB.TextBox XPTxtDiscountVal 
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
            Height          =   465
            Left            =   5400
            TabIndex        =   373
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox TxtPhone 
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
            Index           =   2
            Left            =   8400
            TabIndex        =   372
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox TxtItemCodeB1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11880
            TabIndex        =   351
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Frame Frame8 
            Caption         =   "Frame4"
            Height          =   1695
            Left            =   21840
            RightToLeft     =   -1  'True
            TabIndex        =   269
            Top             =   -240
            Visible         =   0   'False
            Width           =   1335
            Begin vbalIml6.vbalImageList ilsIcons32 
               Left            =   120
               Top             =   240
               _ExtentX        =   953
               _ExtentY        =   953
               IconSizeX       =   32
               IconSizeY       =   32
               ColourDepth     =   24
               Size            =   8824
               Images          =   "FrmSaleBill2.frx":97E0
               Version         =   131072
               KeyCount        =   2
               Keys            =   "ˇ"
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
               Images          =   "FrmSaleBill2.frx":BA78
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
               Images          =   "FrmSaleBill2.frx":17744
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
               TabIndex        =   270
               Top             =   960
               Width           =   135
            End
         End
         Begin VB.Frame Frame7 
            Height          =   3975
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   249
            Top             =   480
            Width           =   4935
            Begin VB.Timer Timer2 
               Enabled         =   0   'False
               Interval        =   100
               Left            =   -1320
               Top             =   480
            End
            Begin VB.Timer Timer4 
               Interval        =   1000
               Left            =   840
               Top             =   1320
            End
            Begin VB.Frame Frame11 
               Height          =   1095
               Left            =   0
               TabIndex        =   250
               Top             =   120
               Width           =   4815
               Begin VB.Shape Shape4 
                  BorderColor     =   &H00400000&
                  BorderWidth     =   5
                  Height          =   975
                  Left            =   0
                  Top             =   120
                  Width           =   4815
               End
               Begin VB.Label LblSowPrice 
                  Alignment       =   2  'Center
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "Tahoma"
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
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   251
                  Top             =   240
                  Width           =   4455
               End
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
               TabIndex        =   268
               Top             =   3600
               Width           =   3045
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
               TabIndex        =   267
               Top             =   360
               Width           =   1965
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   495
               Index           =   0
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   266
               Top             =   2880
               Width           =   1935
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   1
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   265
               Top             =   2880
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   2
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   264
               Top             =   2880
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   3
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   263
               Top             =   2880
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   4
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   262
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   5
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   261
               Top             =   2160
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   6
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   260
               Top             =   2160
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   7
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   259
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   8
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   258
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lBLnO 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Index           =   9
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   257
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   735
               Index           =   96
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   256
               Top             =   3600
               Width           =   975
            End
            Begin VB.Label lBLclr 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   1455
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   255
               Top             =   1320
               Width           =   975
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
               TabIndex        =   254
               Top             =   360
               Width           =   4725
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
               Index           =   0
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   253
               Top             =   120
               Width           =   2295
            End
            Begin VB.Label lblShowQty2 
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
               Height          =   795
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   252
               Top             =   120
               Width           =   2565
            End
            Begin VB.Image Image1 
               Height          =   4245
               Left            =   0
               Picture         =   "FrmSaleBill2.frx":23410
               Stretch         =   -1  'True
               Top             =   -270
               Width           =   4845
            End
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
            TabIndex        =   227
            Top             =   -4800
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
               TabIndex        =   228
               Top             =   0
               Width           =   12150
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
                  TabIndex        =   246
                  Top             =   2280
                  Width           =   2775
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
                  TabIndex        =   245
                  Top             =   2040
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
                  TabIndex        =   244
                  Top             =   1800
                  Width           =   2235
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
                  TabIndex        =   243
                  Top             =   1560
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
                  TabIndex        =   242
                  Top             =   1560
                  Value           =   1  'Checked
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
                  TabIndex        =   241
                  Top             =   1320
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
                  TabIndex        =   240
                  Top             =   1320
                  Width           =   2295
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
                  TabIndex        =   239
                  Top             =   1080
                  Value           =   1  'Checked
                  Width           =   2235
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
                  TabIndex        =   238
                  Top             =   840
                  Width           =   2235
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
                  TabIndex        =   237
                  Top             =   1080
                  Value           =   1  'Checked
                  Width           =   2295
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
                  TabIndex        =   236
                  Top             =   840
                  Width           =   2295
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
                  TabIndex        =   235
                  Top             =   420
                  Width           =   2235
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
                  TabIndex        =   234
                  Top             =   60
                  Width           =   2235
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
                  TabIndex        =   233
                  Top             =   60
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
                  TabIndex        =   232
                  Top             =   1800
                  UseMaskColor    =   -1  'True
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
                  TabIndex        =   231
                  Top             =   300
                  Value           =   1  'Checked
                  Width           =   2295
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
                  TabIndex        =   230
                  Top             =   2040
                  Width           =   2295
               End
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
                  TabIndex        =   229
                  Top             =   2520
                  Value           =   1  'Checked
                  Width           =   2775
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
                  TabIndex        =   248
                  Top             =   480
                  Width           =   915
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
                  TabIndex        =   247
                  Top             =   120
                  Width           =   915
               End
            End
         End
         Begin VB.PictureBox imgLarge 
            BackColor       =   &H80000005&
            Height          =   480
            Left            =   -1920
            ScaleHeight     =   420
            ScaleWidth      =   1875
            TabIndex        =   226
            Top             =   120
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   -2325
            TabIndex        =   225
            Top             =   4800
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   5400
            Style           =   2  'Dropdown List
            TabIndex        =   224
            Top             =   12600
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.Frame Frame9 
            Caption         =   "Frame9"
            Height          =   2055
            Left            =   -4440
            TabIndex        =   218
            Top             =   8520
            Visible         =   0   'False
            Width           =   4215
            Begin VB.ComboBox CboPOSBillType 
               Height          =   315
               Left            =   2280
               Style           =   2  'Dropdown List
               TabIndex        =   221
               Top             =   195
               Width           =   1635
            End
            Begin VB.CheckBox chkPayed 
               Caption         =   "„œ›Ê⁄"
               Height          =   255
               Left            =   1680
               TabIndex        =   220
               Top             =   960
               Width           =   975
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Command4"
               Height          =   195
               Left            =   960
               TabIndex        =   219
               Top             =   120
               Width           =   135
            End
            Begin VB.Label LblSessionID 
               Height          =   375
               Left            =   480
               TabIndex        =   223
               Top             =   1200
               Width           =   2055
            End
            Begin VB.Label LblStableID 
               Caption         =   "-1"
               Height          =   375
               Left            =   3000
               TabIndex        =   222
               Top             =   720
               Width           =   855
            End
         End
         Begin VB.Timer Timer1 
            Interval        =   250
            Left            =   17760
            Top             =   4200
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   315
            Left            =   7920
            Style           =   2  'Dropdown List
            TabIndex        =   217
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox TxtInvSerial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8040
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   216
            Top             =   1200
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   8970
            TabIndex        =   215
            Top             =   -105
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TxtPhone 
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
            Height          =   465
            Index           =   0
            Left            =   5400
            TabIndex        =   214
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   5160
            TabIndex        =   208
            Top             =   3720
            Width           =   9855
            Begin ALLButtonS.ALLButton btnNew 
               Height          =   375
               Index           =   0
               Left            =   8655
               TabIndex        =   209
               Top             =   120
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
               BCOL            =   16777088
               BCOLO           =   16777088
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmSaleBill2.frx":2B4C7
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ALLButtonS.ALLButton btnNew 
               Height          =   375
               Index           =   1
               Left            =   7440
               TabIndex        =   210
               Top             =   120
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   " —«Ã⁄"
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
               BCOL            =   16777088
               BCOLO           =   16777088
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmSaleBill2.frx":2B4E3
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ALLButtonS.ALLButton btnpay 
               Height          =   375
               Index           =   0
               Left            =   3240
               TabIndex        =   211
               Top             =   120
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "œ›⁄ "
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
               BCOL            =   16777088
               BCOLO           =   16777088
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmSaleBill2.frx":2B4FF
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ALLButtonS.ALLButton btnExit 
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   212
               Top             =   120
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "«€·«ﬁ «·‰ﬁÿ…"
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
               BCOL            =   16777088
               BCOLO           =   16777088
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmSaleBill2.frx":2B51B
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   -1  'True
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   540
               Index           =   3
               Left            =   0
               TabIndex        =   213
               Top             =   -720
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   953
               ButtonStyle     =   1
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
               ColorShadow     =   4210752
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               ColorTextShadow =   4210752
            End
            Begin ALLButtonS.ALLButton btnExit 
               Height          =   375
               Index           =   1
               Left            =   1680
               TabIndex        =   341
               Top             =   120
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "œŒÊ· «·„‘—›"
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
               BCOL            =   16777088
               BCOLO           =   16777088
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmSaleBill2.frx":2B537
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   -1  'True
            End
            Begin ALLButtonS.ALLButton btnExit 
               Height          =   375
               Index           =   2
               Left            =   4800
               TabIndex        =   348
               Top             =   120
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "«” œ⁄«¡"
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
               BCOL            =   49344
               BCOLO           =   49344
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmSaleBill2.frx":2B553
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   -1  'True
            End
            Begin ALLButtonS.ALLButton btnExit 
               Height          =   375
               Index           =   3
               Left            =   6240
               TabIndex        =   349
               Top             =   120
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   " ⁄·Ìﬁ"
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
               BCOL            =   49344
               BCOLO           =   49344
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmSaleBill2.frx":2B56F
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   -1  'True
            End
            Begin ALLButtonS.ALLButton btnExit 
               Height          =   375
               Index           =   4
               Left            =   4200
               TabIndex        =   352
               Top             =   600
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "«·«” ⁄·«„ ⁄‰ ’‰›"
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
               BCOL            =   49344
               BCOLO           =   49344
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmSaleBill2.frx":2B58B
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   -1  'True
            End
            Begin ALLButtonS.ALLButton btnExit 
               Height          =   375
               Index           =   6
               Left            =   8040
               TabIndex        =   368
               Top             =   600
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   " «·«” ⁄·«„ ⁄‰ Õœ «·ÿ·»"
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
               BCOL            =   49344
               BCOLO           =   49344
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmSaleBill2.frx":2B5A7
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   -1  'True
            End
            Begin ALLButtonS.ALLButton btnExit 
               Height          =   375
               Index           =   7
               Left            =   6000
               TabIndex        =   369
               Top             =   600
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "«·«” ⁄·«„ ⁄‰  «—ÌŒ «·«‰ Â«¡"
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
               BCOL            =   49344
               BCOLO           =   49344
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmSaleBill2.frx":2B5C3
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   -1  'True
            End
            Begin ALLButtonS.ALLButton btnExit 
               Height          =   375
               Index           =   8
               Left            =   2160
               TabIndex        =   370
               Top             =   600
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   " ﬁ«—Ì— «·‘»ﬂ…"
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
               BCOL            =   49344
               BCOLO           =   49344
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmSaleBill2.frx":2B5DF
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   -1  'True
            End
            Begin ALLButtonS.ALLButton btnExit 
               Height          =   375
               Index           =   9
               Left            =   120
               TabIndex        =   371
               Top             =   600
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "«· ﬁ«—Ì— «·⁄«„…"
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
               BCOL            =   49344
               BCOLO           =   49344
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmSaleBill2.frx":2B5FB
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   -1  'True
            End
            Begin ALLButtonS.ALLButton btnpay 
               Height          =   375
               Index           =   1
               Left            =   3240
               TabIndex        =   374
               Top             =   120
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "œ›⁄"
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
               BCOL            =   16777088
               BCOLO           =   16777088
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmSaleBill2.frx":2B617
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
         Begin VB.Timer Timer5 
            Interval        =   1000
            Left            =   1920
            Top             =   8280
         End
         Begin VB.Frame Frame12 
            Height          =   2535
            Index           =   0
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   204
            Top             =   7080
            Width           =   4695
            Begin VB.Label LblDiscountsTotalView 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Caption         =   "VAT"
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   """#,###.##"""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
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
               Index           =   3
               Left            =   3480
               TabIndex        =   361
               Top             =   1560
               Width           =   1125
            End
            Begin VB.Label LblDiscountsTotalView 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
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
               ForeColor       =   &H0000FF00&
               Height          =   435
               Index           =   6
               Left            =   120
               TabIndex        =   360
               Top             =   1560
               Width           =   3165
            End
            Begin VB.Label LblTotal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
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
               ForeColor       =   &H0000FF00&
               Height          =   435
               Left            =   120
               TabIndex        =   357
               Top             =   2040
               Width           =   3165
            End
            Begin VB.Label LblTotalAllView 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
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
               ForeColor       =   &H0000FF00&
               Height          =   435
               Left            =   120
               TabIndex        =   356
               Top             =   120
               Width           =   3165
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·ﬁ”«∆„"
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
               Index           =   92
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   347
               Top             =   1080
               Width           =   1125
            End
            Begin VB.Label LblDiscountsTotalView 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
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
               ForeColor       =   &H0000FF00&
               Height          =   435
               Index           =   1
               Left            =   120
               TabIndex        =   346
               Top             =   1080
               Width           =   3165
            End
            Begin VB.Label LblDiscountsTotalView 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
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
               ForeColor       =   &H0000FF00&
               Height          =   435
               Index           =   0
               Left            =   120
               TabIndex        =   345
               Top             =   600
               Width           =   3165
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
               TabIndex        =   344
               Top             =   1800
               Visible         =   0   'False
               Width           =   1440
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
               Left            =   3360
               TabIndex        =   207
               Top             =   120
               Width           =   1125
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
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   206
               Top             =   600
               Width           =   1125
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
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   205
               Top             =   2040
               Width           =   1125
            End
         End
         Begin VB.OptionButton optsale 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "„»Ì⁄« "
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
            Index           =   0
            Left            =   10200
            TabIndex        =   203
            Top             =   960
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optsale 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "„—œÊœ« "
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
            Height          =   255
            Index           =   1
            Left            =   8880
            TabIndex        =   202
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox TxtInvID 
            Height          =   285
            Left            =   11520
            TabIndex        =   201
            Text            =   "Text2"
            Top             =   600
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox CboRetrunType 
            Height          =   315
            ItemData        =   "FrmSaleBill2.frx":2B633
            Left            =   5400
            List            =   "FrmSaleBill2.frx":2B635
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   200
            Top             =   1200
            Width           =   2475
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«” »œ«·"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   0
            Left            =   11880
            RightToLeft     =   -1  'True
            TabIndex        =   199
            Top             =   480
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.TextBox txtInvDate 
            Height          =   285
            Left            =   12480
            TabIndex        =   198
            Top             =   360
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txt_Currency_rate 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   13335
            RightToLeft     =   -1  'True
            TabIndex        =   197
            Text            =   "1"
            Top             =   840
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox Txtcard 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   9600
            PasswordChar    =   "*"
            TabIndex        =   196
            Top             =   1920
            Width           =   510
         End
         Begin VB.TextBox Txtcard 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   9600
            PasswordChar    =   "*"
            TabIndex        =   195
            Top             =   1560
            Width           =   510
         End
         Begin VB.TextBox TxtQuantity 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   6480
            MaxLength       =   10
            TabIndex        =   312
            Top             =   3360
            Width           =   615
         End
         Begin VB.TextBox TxtPrice 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5730
            MaxLength       =   6
            TabIndex        =   314
            Top             =   3360
            Width           =   600
         End
         Begin VB.TextBox TxtShortName 
            Height          =   375
            Left            =   10920
            TabIndex        =   194
            Top             =   2880
            Width           =   2655
         End
         Begin ALLButtonS.ALLButton ALLButton9 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   6120
            TabIndex        =   271
            Top             =   3960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "«€·«ﬁ «·‘Ì› "
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
            MICON           =   "FrmSaleBill2.frx":2B637
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin vbalIml6.vbalImageList GrouplImageList 
            Left            =   0
            Top             =   600
            _ExtentX        =   953
            _ExtentY        =   953
            IconSizeX       =   32
            IconSizeY       =   32
            ColourDepth     =   24
            Size            =   4412
            Images          =   "FrmSaleBill2.frx":2B653
            Version         =   131072
            KeyCount        =   1
            Keys            =   ""
         End
         Begin vbalListViewLib6.vbalListViewCtl lvwMain 
            Height          =   8655
            Left            =   -5400
            TabIndex        =   272
            Top             =   0
            Visible         =   0   'False
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   15266
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
         Begin vbalListViewLib6.vbalListViewCtl lvwItems 
            Height          =   8655
            Left            =   21600
            TabIndex        =   273
            Top             =   4800
            Visible         =   0   'False
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   15266
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
         Begin vbalListViewLib6.vbalListViewCtl lvwTables 
            Height          =   8655
            Left            =   20880
            TabIndex        =   274
            Top             =   4800
            Visible         =   0   'False
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   15266
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
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   -3000
            TabIndex        =   275
            Top             =   4680
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
         Begin MSDataListLib.DataCombo DCPaymentNet 
            Height          =   315
            Left            =   12240
            TabIndex        =   276
            Top             =   120
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   0
            Left            =   23280
            TabIndex        =   277
            Top             =   2760
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÃœÌœ"
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
            DisabledImageExtraction=   0
            ColorToggledText=   65280
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   1
            Left            =   22440
            TabIndex        =   278
            Top             =   2760
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ·"
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
            Index           =   2
            Left            =   21480
            TabIndex        =   279
            Top             =   2760
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ›Ÿ"
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
            Index           =   4
            Left            =   21840
            TabIndex        =   280
            Top             =   1920
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
            Left            =   21840
            TabIndex        =   281
            Top             =   3360
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "œ›⁄"
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   540
            Index           =   1
            Left            =   -720
            TabIndex        =   282
            TabStop         =   0   'False
            Top             =   -1320
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
               TabIndex        =   283
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
               TabIndex        =   284
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
               TabIndex        =   285
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
               TabIndex        =   286
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
                  TabIndex        =   287
                  TabStop         =   0   'False
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   285
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
                  TabIndex        =   299
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   1935
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
                  TabIndex        =   298
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·’«›Ì"
                  Height          =   285
                  Index           =   49
                  Left            =   8850
                  RightToLeft     =   -1  'True
                  TabIndex        =   297
                  Top             =   75
                  Width           =   1020
               End
               Begin VB.Label XPTxtCount 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Height          =   285
                  Left            =   330
                  RightToLeft     =   -1  'True
                  TabIndex        =   296
                  Top             =   75
                  Width           =   405
               End
               Begin VB.Label XPTxtCurrent 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Height          =   285
                  Left            =   1365
                  RightToLeft     =   -1  'True
                  TabIndex        =   295
                  Top             =   75
                  Width           =   270
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·”Ã·"
                  Height          =   285
                  Index           =   2
                  Left            =   1860
                  RightToLeft     =   -1  'True
                  TabIndex        =   294
                  Top             =   75
                  Width           =   1065
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "/"
                  Height          =   285
                  Index           =   0
                  Left            =   1020
                  RightToLeft     =   -1  'True
                  TabIndex        =   293
                  Top             =   75
                  Width           =   165
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·≈Ã„«·Ï"
                  Height          =   285
                  Index           =   3
                  Left            =   20430
                  RightToLeft     =   -1  'True
                  TabIndex        =   292
                  Top             =   75
                  Width           =   810
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
                  TabIndex        =   291
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   675
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
                  Height          =   315
                  Index           =   63
                  Left            =   3600
                  TabIndex        =   290
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   435
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
                  TabIndex        =   289
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   585
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
                  TabIndex        =   288
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1830
               End
            End
         End
         Begin MSComctlLib.Toolbar TBar 
            Height          =   630
            Left            =   5040
            TabIndex        =   300
            Top             =   10080
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   1111
            ButtonWidth     =   609
            ButtonHeight    =   1005
            Appearance      =   1
            _Version        =   393216
            Begin ALLButtonS.ALLButton btnExit 
               Height          =   255
               Index           =   5
               Left            =   3240
               TabIndex        =   353
               Top             =   0
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   "Õ–› «·„Õœœ"
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
               BCOL            =   128
               BCOLO           =   128
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmSaleBill2.frx":2C7AF
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   -1  'True
            End
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   8
            Left            =   20880
            TabIndex        =   301
            Top             =   3360
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
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
            ColorHoverText  =   65280
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   65280
            ColorToggledHoverText=   65280
            ColorTextShadow =   4210752
         End
         Begin VSFlex8UCtl.VSFlexGrid FG 
            Height          =   4875
            Left            =   5130
            TabIndex        =   302
            Top             =   5160
            Width           =   9885
            _cx             =   17436
            _cy             =   8599
            Appearance      =   3
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
            BackColorFixed  =   16777088
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
            Cols            =   29
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmSaleBill2.frx":2C7CB
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
            Left            =   20760
            TabIndex        =   303
            Top             =   2640
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â"
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
            Index           =   11
            Left            =   20640
            TabIndex        =   304
            Top             =   2040
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "„ÿ»Œ"
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
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   5400
            TabIndex        =   305
            Top             =   1575
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboItemsCodexxx 
            Height          =   315
            Left            =   -240
            TabIndex        =   306
            Top             =   -240
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton SearchCashCustomer 
            Height          =   345
            Index           =   0
            Left            =   5160
            TabIndex        =   307
            TabStop         =   0   'False
            Top             =   2400
            Width           =   270
            _ExtentX        =   476
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
            ButtonImage     =   "FrmSaleBill2.frx":2CC8B
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton SearchCashCustomer 
            Height          =   315
            Index           =   1
            Left            =   -10560
            TabIndex        =   308
            TabStop         =   0   'False
            Top             =   2880
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   556
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
            ButtonImage     =   "FrmSaleBill2.frx":2D088
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcCurrency 
            Height          =   315
            Left            =   12900
            TabIndex        =   309
            Top             =   840
            Visible         =   0   'False
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboItemsName 
            Height          =   315
            Left            =   7200
            TabIndex        =   311
            Top             =   3360
            Width           =   4545
            _ExtentX        =   8017
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
         Begin ImpulseButton.ISButton CmdAdd 
            Height          =   420
            Left            =   5280
            TabIndex        =   316
            Top             =   3240
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   741
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
            BackStyle       =   0
            ButtonImage     =   "FrmSaleBill2.frx":2D485
            ColorButton     =   16777215
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   16777215
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            MaskColor       =   16777215
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   16777215
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   5160
            TabIndex        =   342
            Top             =   10920
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VSFlex8UCtl.VSFlexGrid VatGrid 
            Height          =   1725
            Left            =   5160
            TabIndex        =   354
            Tag             =   "1"
            Top             =   9000
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
            FormatString    =   $"FrmSaleBill2.frx":2D81F
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
         Begin VB.CheckBox ChecVAT 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   " ÕœÌœ «·ﬂ·"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   11280
            RightToLeft     =   -1  'True
            TabIndex        =   355
            Top             =   8760
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.TextBox TxtValueAdded 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11040
            RightToLeft     =   -1  'True
            TabIndex        =   362
            Top             =   10800
            Width           =   2055
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   4
            Left            =   2655
            TabIndex        =   364
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
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
            ButtonImage     =   "FrmSaleBill2.frx":2D932
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
            Index           =   5
            Left            =   1425
            TabIndex        =   365
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
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
            ButtonImage     =   "FrmSaleBill2.frx":2DCCC
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
            Index           =   6
            Left            =   3900
            TabIndex        =   366
            Top             =   120
            Width           =   1020
            _ExtentX        =   1799
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
            ButtonImage     =   "FrmSaleBill2.frx":2E066
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
            Index           =   7
            Left            =   75
            TabIndex        =   367
            Top             =   105
            Width           =   1320
            _ExtentX        =   2328
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
            ButtonImage     =   "FrmSaleBill2.frx":2E400
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin VB.TextBox TxtPhone 
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
            Height          =   375
            Index           =   1
            Left            =   8400
            TabIndex        =   363
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox TxtItemCodeB 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5400
            TabIndex        =   378
            Top             =   2880
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo DCboItemsCode 
            Height          =   315
            Left            =   13800
            TabIndex        =   310
            Top             =   4080
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.TextBox txtItemCodeSearch2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11760
            TabIndex        =   379
            Top             =   3360
            Width           =   2175
         End
         Begin ALLButtonS.ALLButton btnNew 
            Height          =   375
            Index           =   2
            Left            =   11550
            TabIndex        =   380
            Top             =   4725
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   " Õ„Ì· „·› «ﬂ”Ì· «·Ï «·‰Ÿ«„"
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
            BCOL            =   16777088
            BCOLO           =   16777088
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmSaleBill2.frx":2E79A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ALLButtonS.ALLButton btnNew 
            Height          =   270
            Index           =   3
            Left            =   10140
            TabIndex        =   381
            Top             =   2475
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   476
            BTYPE           =   3
            TX              =   "«ŸÂ«— ‰ﬁ«ÿ «·⁄„Ì·"
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
            BCOL            =   16777088
            BCOLO           =   16777088
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmSaleBill2.frx":2E7B6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "time"
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
            Index           =   98
            Left            =   5400
            TabIndex        =   382
            Top             =   4815
            Visible         =   0   'False
            Width           =   6045
         End
         Begin VB.Label LblDiscountsTotalView 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·›« "
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """#,###.##"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   9840
            TabIndex        =   359
            Top             =   2880
            Width           =   885
         End
         Begin VB.Label LblDiscountsTotalView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "«’‰«› «·ﬁÌ„… «·„÷«›…"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """#,###.##"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   4
            Left            =   12720
            TabIndex        =   358
            Top             =   8640
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·»«—ﬂÊœ"
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
            Index           =   95
            Left            =   13440
            TabIndex        =   350
            Top             =   2400
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„” Œœ„"
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
            Index           =   93
            Left            =   8160
            TabIndex        =   343
            Top             =   10920
            Width           =   1650
         End
         Begin VB.Shape Shape2 
            FillStyle       =   0  'Solid
            Height          =   1335
            Left            =   -6000
            Top             =   4800
            Visible         =   0   'False
            Width           =   4815
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
            TabIndex        =   340
            Top             =   120
            Width           =   1245
         End
         Begin VB.Image Image2 
            Height          =   435
            Left            =   21480
            Stretch         =   -1  'True
            Top             =   4320
            Visible         =   0   'False
            Width           =   3555
         End
         Begin VB.Image Image3 
            Height          =   435
            Left            =   22080
            Stretch         =   -1  'True
            Top             =   4320
            Width           =   4275
         End
         Begin VB.Image Image4 
            Height          =   555
            Left            =   22080
            Stretch         =   -1  'True
            Top             =   2040
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.Image Image7 
            Height          =   555
            Left            =   21600
            Stretch         =   -1  'True
            Top             =   3240
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.Image Image8 
            Height          =   555
            Left            =   21120
            Stretch         =   -1  'True
            Top             =   2640
            Visible         =   0   'False
            Width           =   2235
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
            Left            =   21120
            TabIndex        =   339
            Top             =   9240
            Visible         =   0   'False
            Width           =   9255
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
            Index           =   80
            Left            =   21360
            RightToLeft     =   -1  'True
            TabIndex        =   338
            Top             =   4320
            Width           =   1965
         End
         Begin VB.Label Label14 
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
            Left            =   21480
            RightToLeft     =   -1  'True
            TabIndex        =   337
            Top             =   2400
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·œ›⁄"
            Height          =   300
            Index           =   9
            Left            =   3300
            TabIndex        =   336
            Top             =   9840
            Visible         =   0   'False
            Width           =   1035
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
            ForeColor       =   &H0000FF00&
            Height          =   270
            Index           =   7
            Left            =   -4035
            TabIndex        =   335
            Top             =   5085
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·Œ’„"
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
            Index           =   10
            Left            =   10200
            TabIndex        =   334
            Top             =   1920
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„…"
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
            Index           =   8
            Left            =   7200
            TabIndex        =   333
            Top             =   1920
            Width           =   570
         End
         Begin VB.Image Image5 
            Height          =   315
            Left            =   21240
            Picture         =   "FrmSaleBill2.frx":2E7D2
            Stretch         =   -1  'True
            Top             =   3360
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Œœ„… «·„⁄œ« /«·”Ì«—« "
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
            Left            =   21120
            RightToLeft     =   -1  'True
            TabIndex        =   332
            Top             =   3360
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.Label lblLabel1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   0
            Left            =   1560
            TabIndex        =   331
            Top             =   10800
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblLabel1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
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
            TabIndex        =   330
            Top             =   10920
            Width           =   1095
         End
         Begin VB.Image Image9 
            Height          =   1455
            Left            =   11880
            Stretch         =   -1  'True
            Top             =   840
            Width           =   3015
         End
         Begin VB.Image Image6 
            Height          =   435
            Left            =   20640
            Stretch         =   -1  'True
            Top             =   4320
            Width           =   2235
         End
         Begin VB.Label LBLTable 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   20520
            RightToLeft     =   -1  'True
            TabIndex        =   329
            Top             =   3480
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·‘—ﬂ… «·⁄—»Ì… ·· Ã«—…"
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
            Height          =   390
            Index           =   84
            Left            =   5400
            TabIndex        =   328
            Top             =   120
            Width           =   9330
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "«·À·«À«¡"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   83
            Left            =   120
            TabIndex        =   327
            Top             =   10080
            Width           =   4695
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "12:30 AM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   795
            Index           =   82
            Left            =   120
            TabIndex        =   326
            Top             =   10440
            Width           =   4695
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "12:30 AM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   555
            Index           =   81
            Left            =   120
            TabIndex        =   325
            Top             =   9600
            Width           =   4695
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "›« Ê—… "
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
            Index           =   85
            Left            =   9240
            TabIndex        =   324
            Top             =   600
            Width           =   1530
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·›« Ê—…"
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
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   86
            Left            =   10200
            TabIndex        =   323
            Top             =   1260
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·»«∆⁄"
            DataField       =   "«·»«"
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
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   1
            Left            =   10200
            RightToLeft     =   -1  'True
            TabIndex        =   322
            Top             =   1560
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÃÊ«·"
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
            Index           =   87
            Left            =   7200
            TabIndex        =   321
            Top             =   2400
            Width           =   570
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Index           =   88
            Left            =   10200
            TabIndex        =   320
            Top             =   2250
            Width           =   1410
         End
         Begin VB.Image Image10 
            Height          =   2145
            Left            =   360
            Stretch         =   -1  'True
            Top             =   4680
            Width           =   4140
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H00400000&
            BorderWidth     =   5
            Height          =   1695
            Left            =   120
            Top             =   9600
            Width           =   4695
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H000000FF&
            BorderColor     =   &H00400000&
            BorderWidth     =   5
            Height          =   11295
            Left            =   5025
            Top             =   -15
            Width           =   10110
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BorderColor     =   &H00400000&
            BorderWidth     =   5
            Height          =   2655
            Left            =   120
            Top             =   4440
            Width           =   4695
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H00E2E9E9&
            BorderColor     =   &H00E2E9E9&
            FillStyle       =   0  'Solid
            Height          =   15495
            Left            =   21240
            Top             =   360
            Width           =   30855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·ﬂÊœ «· Õ·Ì·Ì"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   89
            Left            =   6840
            TabIndex        =   319
            Top             =   3000
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„·…"
            Height          =   285
            Index           =   65
            Left            =   13065
            RightToLeft     =   -1  'True
            TabIndex        =   318
            Top             =   840
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "⁄—÷ Œ«’"
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
            Index           =   57
            Left            =   8040
            TabIndex        =   317
            Top             =   600
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬂÊœ «·’‰›"
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
            Index           =   90
            Left            =   13560
            TabIndex        =   315
            Top             =   3360
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»ÕÀ ”—Ì⁄"
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
            Height          =   390
            Index           =   91
            Left            =   13320
            TabIndex        =   313
            Top             =   3000
            Width           =   1650
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1785
         Index           =   0
         Left            =   15
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   -3435
         Visible         =   0   'False
         Width           =   19170
         _cx             =   33814
         _cy             =   3149
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
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4290
            TabIndex        =   190
            Top             =   15
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "«” ⁄·«„ ⁄‰ ’‰›"
            Height          =   255
            Left            =   5355
            TabIndex        =   90
            Top             =   1695
            Visible         =   0   'False
            Width           =   2160
         End
         Begin VB.TextBox TxtIssueSerial 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   270
            Left            =   0
            TabIndex        =   78
            Top             =   -240
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   270
            Left            =   1290
            TabIndex        =   76
            Top             =   -240
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.TextBox TXTOrDer_no 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6135
            TabIndex        =   71
            Top             =   1095
            Width           =   1140
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   14940
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   135
            Width           =   2535
         End
         Begin VB.CommandButton Command1 
            Caption         =   " ÕÊÌ· «·Ï «–‰ ’—›"
            Height          =   285
            Left            =   10395
            TabIndex        =   66
            Top             =   -135
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11865
            TabIndex        =   3
            Top             =   1095
            Width           =   2340
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   16380
            TabIndex        =   4
            Top             =   1425
            Width           =   1095
         End
         Begin VB.TextBox TxtCusID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   16380
            TabIndex        =   2
            Top             =   750
            Width           =   1095
         End
         Begin VB.ComboBox CboSaleType 
            Height          =   315
            Left            =   4065
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   690
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   15435
            TabIndex        =   0
            Top             =   -195
            Visible         =   0   'False
            Width           =   1845
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   735
            Index           =   8
            Left            =   17610
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1695
            Visible         =   0   'False
            Width           =   5040
            _cx             =   8890
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
               Height          =   360
               Left            =   5565
               TabIndex        =   22
               Top             =   180
               Width           =   3030
               _ExtentX        =   5345
               _ExtentY        =   635
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
               ButtonImage     =   "FrmSaleBill2.frx":2EEF8
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
               Left            =   8475
               TabIndex        =   27
               Top             =   435
               Width           =   4440
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·—»Õ"
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   22
               Left            =   35820
               TabIndex        =   26
               Top             =   165
               Width           =   4380
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
               Left            =   3645
               TabIndex        =   25
               Top             =   405
               Width           =   5640
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
               Height          =   315
               Left            =   3645
               TabIndex        =   24
               Top             =   90
               Width           =   5640
            End
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   11865
            TabIndex        =   5
            Top             =   1425
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   330
            Left            =   14970
            TabIndex        =   1
            Top             =   420
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   582
            _Version        =   393216
            Format          =   222494721
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   375
            Left            =   17595
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   735
            Visible         =   0   'False
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   661
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
            ButtonImage     =   "FrmSaleBill2.frx":2F292
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdCash 
            Height          =   405
            Index           =   0
            Left            =   11805
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   900
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   714
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
            ButtonImage     =   "FrmSaleBill2.frx":2F62C
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdCash 
            Height          =   255
            Index           =   1
            Left            =   11655
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1140
            Visible         =   0   'False
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   450
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
            ButtonImage     =   "FrmSaleBill2.frx":2F9C6
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   11805
            TabIndex        =   72
            Top             =   135
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   6060
            TabIndex        =   74
            Top             =   390
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCDocTypes 
            Height          =   315
            Left            =   11865
            TabIndex        =   146
            Top             =   480
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   1815
            Left            =   90
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   0
            Width           =   3090
            _cx             =   5450
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
               TabIndex        =   151
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
               TabIndex        =   149
               Top             =   0
               Width           =   1140
            End
            Begin VB.Frame Frame2 
               Caption         =   " œ·«·«  «·«·Ê«‰"
               Height          =   735
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   720
               Width           =   2280
               Begin VB.Label Label5 
                  BackColor       =   &H000000FF&
                  Height          =   255
                  Left            =   1920
                  TabIndex        =   96
                  Top             =   240
                  Width           =   255
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "»Ì⁄ «ﬁ· „‰ ”⁄— «· ﬂ·›Â"
                  Height          =   255
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.Label lbl 
                  BackColor       =   &H0000FFFF&
                  Height          =   255
                  Index           =   97
                  Left            =   1920
                  TabIndex        =   94
                  Top             =   480
                  Width           =   255
               End
               Begin VB.Label Label8 
                  Alignment       =   1  'Right Justify
                  Caption         =   "»Ì⁄  »”⁄— «· ﬂ·›Â"
                  Height          =   255
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   480
                  Width           =   1215
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„  »Ê·Ì’… «·‘Õ‰"
               Height          =   195
               Index           =   67
               Left            =   1440
               TabIndex        =   152
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
               TabIndex        =   150
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.Frame Frame400 
            Height          =   495
            Left            =   8430
            RightToLeft     =   -1  'True
            TabIndex        =   153
            Top             =   1320
            Width           =   3165
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—»Õ «·›« Ê—…"
               ForeColor       =   &H00008000&
               Height          =   195
               Index           =   68
               Left            =   1680
               TabIndex        =   156
               Top             =   240
               Width           =   960
            End
            Begin VB.Label LblPrecenValuex 
               Caption         =   "0"
               Height          =   255
               Left            =   120
               TabIndex        =   155
               Top             =   840
               Width           =   1455
            End
            Begin VB.Label LblInvProfit 
               Alignment       =   2  'Center
               Caption         =   "0"
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   120
               TabIndex        =   154
               Top             =   240
               Width           =   1575
            End
         End
         Begin MSComCtl2.DTPicker DtpDelayDate 
            Height          =   315
            Left            =   3285
            TabIndex        =   158
            Top             =   1425
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Format          =   222494721
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcboEmpxxxx 
            Height          =   315
            Left            =   720
            TabIndex        =   191
            Top             =   0
            Width           =   3570
            _ExtentX        =   6297
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·»«∆⁄"
            Height          =   270
            Index           =   25
            Left            =   5490
            TabIndex        =   192
            Top             =   30
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «·«” Õﬁ«ﬁ"
            Height          =   255
            Index           =   21
            Left            =   4830
            TabIndex        =   159
            Top             =   1530
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·”‰œ"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   13815
            TabIndex        =   147
            Top             =   480
            Width           =   945
         End
         Begin VB.Label Label4 
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   240
            Left            =   1275
            TabIndex        =   77
            Top             =   480
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·Œ“‰…"
            Height          =   210
            Index           =   11
            Left            =   10665
            TabIndex        =   75
            Top             =   480
            Width           =   990
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   13770
            TabIndex        =   73
            Top             =   135
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÿ·»Ì…"
            Height          =   240
            Index           =   56
            Left            =   7305
            TabIndex        =   70
            Top             =   1185
            Width           =   1095
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
            Height          =   240
            Index           =   55
            Left            =   5820
            TabIndex        =   67
            Top             =   855
            Width           =   375
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÏ"
            Height          =   285
            Index           =   33
            Left            =   14280
            TabIndex        =   33
            Top             =   1140
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”Ì«”… «·»Ì⁄"
            Height          =   240
            Index           =   32
            Left            =   10665
            TabIndex        =   29
            Top             =   1395
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   240
            Index           =   24
            Left            =   17265
            TabIndex        =   13
            Top             =   1500
            Width           =   1875
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·›« Ê—…"
            Height          =   270
            Index           =   6
            Left            =   16335
            TabIndex        =   12
            Top             =   420
            Width           =   2670
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·›« Ê—…"
            Height          =   315
            Index           =   5
            Left            =   17055
            TabIndex        =   11
            Top             =   60
            Width           =   1950
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   9975
         Left            =   15
         TabIndex        =   8
         Top             =   14925
         Visible         =   0   'False
         Width           =   19170
         _cx             =   33814
         _cy             =   17595
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
         Picture(0)      =   "FrmSaleBill2.frx":2FD60
         Picture(1)      =   "FrmSaleBill2.frx":300FA
         Picture(2)      =   "FrmSaleBill2.frx":30494
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   9510
            Index           =   19
            Left            =   20115
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   45
            Width           =   19080
            _cx             =   33655
            _cy             =   16775
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
               TabIndex        =   86
               Top             =   720
               Width           =   4335
               Begin VB.TextBox TxtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   600
                  Width           =   2505
               End
               Begin ImpulseButton.ISButton Cmd 
                  CausesValidation=   0   'False
                  Height          =   375
                  Index           =   10
                  Left            =   240
                  TabIndex        =   88
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
                  TabIndex        =   89
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
               TabIndex        =   85
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
               TabIndex        =   84
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
               TabIndex        =   82
               Top             =   600
               Width           =   5625
            End
            Begin VB.TextBox TXTNoteID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   0
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VSFlex8UCtl.VSFlexGrid GRID1 
               Height          =   2085
               Left            =   6960
               TabIndex        =   79
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
               FormatString    =   $"FrmSaleBill2.frx":3082E
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
               TabIndex        =   81
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
               FormatString    =   $"FrmSaleBill2.frx":3097B
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
               TabIndex        =   160
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
               TabIndex        =   83
               Top             =   120
               Width           =   2520
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   9510
            Index           =   15
            Left            =   19815
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   45
            Width           =   19080
            _cx             =   33655
            _cy             =   16775
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
            _GridInfo       =   $"FrmSaleBill2.frx":30A6E
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1560
               Index           =   18
               Left            =   15
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   795
               Visible         =   0   'False
               Width           =   19050
               _cx             =   33602
               _cy             =   2752
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
                  TabIndex        =   185
                  Top             =   -15
                  Width           =   90
                  Begin VB.ComboBox CboPaymentType1 
                     Height          =   315
                     Left            =   0
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   187
                     Top             =   585
                     Width           =   2685
                  End
                  Begin VB.TextBox TxtAdvPaymnt 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   0
                     MaxLength       =   10
                     RightToLeft     =   -1  'True
                     TabIndex        =   186
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
                     TabIndex        =   189
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
                     TabIndex        =   188
                     Top             =   255
                     Width           =   1275
                  End
               End
               Begin VB.Frame FraNote 
                  BackColor       =   &H00E2E9E9&
                  Height          =   30
                  Left            =   210
                  RightToLeft     =   -1  'True
                  TabIndex        =   173
                  Top             =   -15
                  Width           =   75
                  Begin VB.TextBox TxtChequeNumber 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   30
                     RightToLeft     =   -1  'True
                     TabIndex        =   175
                     Top             =   930
                     Width           =   2685
                  End
                  Begin VB.TextBox TXTBankName 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   174
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   2685
                  End
                  Begin MSComCtl2.DTPicker DtpChequeDueDate1 
                     Height          =   315
                     Left            =   30
                     TabIndex        =   176
                     Top             =   1260
                     Width           =   2685
                     _ExtentX        =   4736
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   222560257
                     CurrentDate     =   39614
                  End
                  Begin MSDataListLib.DataCombo DcboBankName1 
                     Height          =   315
                     Left            =   30
                     TabIndex        =   177
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
                     TabIndex        =   178
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
                     TabIndex        =   179
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
                     TabIndex        =   184
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
                     TabIndex        =   183
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
                     TabIndex        =   182
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
                     TabIndex        =   181
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
                     TabIndex        =   180
                     Top             =   1560
                     Width           =   1215
                  End
               End
               Begin VB.TextBox TxtTaxServiceValue 
                  Alignment       =   1  'Right Justify
                  Height          =   15
                  Left            =   150
                  MaxLength       =   4
                  TabIndex        =   53
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
                  TabIndex        =   48
                  Top             =   0
                  Width           =   30
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   15
                  Index           =   54
                  Left            =   105
                  TabIndex        =   65
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
                  TabIndex        =   58
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
                  TabIndex        =   54
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   15
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1560
               Index           =   17
               Left            =   15
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   795
               Visible         =   0   'False
               Width           =   19050
               _cx             =   33602
               _cy             =   2752
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
                  TabIndex        =   52
                  Top             =   0
                  Width           =   30
               End
               Begin VB.CheckBox ChkTaxStamp 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ„€…"
                  Height          =   0
                  Left            =   210
                  TabIndex        =   46
                  Top             =   15
                  Width           =   0
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   15
                  Index           =   53
                  Left            =   105
                  TabIndex        =   64
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
                  TabIndex        =   59
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
                  TabIndex        =   50
                  Top             =   0
                  Width           =   15
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1560
               Index           =   16
               Left            =   15
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   795
               Visible         =   0   'False
               Width           =   19050
               _cx             =   33602
               _cy             =   2752
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
                  TabIndex        =   51
                  Top             =   0
                  Width           =   30
               End
               Begin VB.CheckBox ChkTaxAdd 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… Œ’„ Ê≈÷«›… (√—»«Õ  Ã«—Ì…)"
                  Height          =   15
                  Left            =   195
                  TabIndex        =   44
                  Top             =   0
                  Width           =   45
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   15
                  Index           =   52
                  Left            =   120
                  TabIndex        =   63
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
                  TabIndex        =   57
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
                  TabIndex        =   49
                  Top             =   0
                  Width           =   15
               End
            End
            Begin VB.TextBox TxtBillComment 
               Alignment       =   1  'Right Justify
               Height          =   1560
               Left            =   0
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   37
               Top             =   0
               Width           =   19050
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   765
               Index           =   4
               Left            =   15
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   15
               Visible         =   0   'False
               Width           =   19050
               _cx             =   33602
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
                  TabIndex        =   41
                  Top             =   225
                  Width           =   420
               End
               Begin VB.TextBox XPTxtTaxValue 
                  Alignment       =   1  'Right Justify
                  Height          =   510
                  Left            =   1815
                  MaxLength       =   4
                  TabIndex        =   40
                  Top             =   105
                  Width           =   300
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   360
                  Index           =   51
                  Left            =   300
                  TabIndex        =   62
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
                  TabIndex        =   56
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
                  TabIndex        =   42
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
               Height          =   1560
               Index           =   44
               Left            =   15
               TabIndex        =   55
               Top             =   795
               Width           =   19050
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   9510
            Index           =   7
            Left            =   0
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   9840
            Visible         =   0   'False
            Width           =   19080
            _cx             =   33655
            _cy             =   16775
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
               Height          =   1170
               Index           =   2
               Left            =   30
               TabIndex        =   131
               TabStop         =   0   'False
               Top             =   30
               Width           =   19020
               _cx             =   33549
               _cy             =   2064
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
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   660
                  Left            =   5025
                  MaxLength       =   20
                  TabIndex        =   133
                  Top             =   495
                  Width           =   1800
               End
               Begin VB.ComboBox CboItemCase 
                  Height          =   315
                  Left            =   7020
                  Style           =   2  'Dropdown List
                  TabIndex        =   132
                  Top             =   495
                  Width           =   1410
               End
               Begin ImpulseButton.ISButton CmdSearch 
                  Height          =   600
                  Left            =   615
                  TabIndex        =   134
                  Top             =   60
                  Visible         =   0   'False
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   1058
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
                  ButtonImage     =   "FrmSaleBill2.frx":30AEA
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”⁄—"
                  Height          =   465
                  Index           =   26
                  Left            =   1110
                  TabIndex        =   140
                  Top             =   15
                  Width           =   1110
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   465
                  Index           =   27
                  Left            =   3210
                  TabIndex        =   139
                  Top             =   45
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   465
                  Index           =   28
                  Left            =   5265
                  TabIndex        =   138
                  Top             =   15
                  Width           =   1035
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   465
                  Index           =   29
                  Left            =   7155
                  TabIndex        =   137
                  Top             =   15
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   465
                  Index           =   30
                  Left            =   14235
                  TabIndex        =   136
                  Top             =   15
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   465
                  Index           =   31
                  Left            =   16080
                  TabIndex        =   135
                  Top             =   45
                  Width           =   2385
               End
            End
            Begin MSComctlLib.Toolbar Toolbar1 
               Height          =   630
               Left            =   30
               TabIndex        =   141
               Top             =   30
               Width           =   9450
               _ExtentX        =   16669
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
               Height          =   2355
               Left            =   30
               TabIndex        =   28
               Top             =   7125
               Width           =   210
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   9510
            Index           =   5
            Left            =   45
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   45
            Width           =   19080
            _cx             =   33655
            _cy             =   16775
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
               TabIndex        =   97
               TabStop         =   0   'False
               Top             =   11400
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
                  TabIndex        =   98
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
                     TabIndex        =   168
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
                     TabIndex        =   166
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
                     TabIndex        =   164
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
                     TabIndex        =   163
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
                     TabIndex        =   101
                     Top             =   225
                     Width           =   1515
                  End
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Index           =   0
                     Left            =   14430
                     Locked          =   -1  'True
                     TabIndex        =   100
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
                     TabIndex        =   99
                     Top             =   280
                     Width           =   930
                  End
                  Begin ImpulseButton.ISButton CmdINSTALLMENT 
                     Height          =   390
                     Left            =   240
                     TabIndex        =   167
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
                     ButtonImage     =   "FrmSaleBill2.frx":30E84
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
                     TabIndex        =   169
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
                     TabIndex        =   165
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
                     TabIndex        =   104
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
                     TabIndex        =   103
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
                     TabIndex        =   102
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   810
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   90
                  Index           =   12
                  Left            =   90
                  TabIndex        =   105
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
                  TabIndex        =   106
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
                  FormatString    =   $"FrmSaleBill2.frx":3121E
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   690
                  Index           =   13
                  Left            =   90
                  TabIndex        =   107
                  TabStop         =   0   'False
                  Top             =   2700
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
                  Begin VB.Label LblAdvPayment 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   10680
                     TabIndex        =   172
                     Top             =   240
                     Width           =   690
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «·œ›⁄Â «·„ﬁœ„Â"
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
                     Height          =   285
                     Index           =   72
                     Left            =   11400
                     TabIndex        =   171
                     Top             =   240
                     Width           =   1125
                  End
                  Begin VB.Label LBLaDVpAY 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   165
                     Left            =   240
                     TabIndex        =   170
                     Top             =   480
                     Width           =   720
                  End
                  Begin VB.Label LblDiscount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   165
                     Left            =   120
                     TabIndex        =   162
                     Top             =   240
                     Width           =   720
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Œ’„  ﬁ”Ìÿ"
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
                     Height          =   285
                     Index           =   70
                     Left            =   1320
                     TabIndex        =   161
                     Top             =   240
                     Width           =   990
                  End
                  Begin VB.Label LblPrecenValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   14400
                     TabIndex        =   157
                     Top             =   240
                     Width           =   405
                  End
                  Begin VB.Label LblInstallmentType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   285
                     Left            =   5025
                     TabIndex        =   122
                     Top             =   285
                     Width           =   555
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "› —… «· ﬁ”Ìÿ"
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
                     Height          =   285
                     Index           =   42
                     Left            =   5760
                     TabIndex        =   121
                     Top             =   285
                     Width           =   1170
                  End
                  Begin VB.Label LblFirstInstallDate 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   6990
                     TabIndex        =   120
                     Top             =   285
                     Width           =   870
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ê· ﬁ”ÿ"
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
                     Height          =   285
                     Index           =   40
                     Left            =   7920
                     TabIndex        =   119
                     Top             =   285
                     Width           =   885
                  End
                  Begin VB.Label LblInstallCount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   8940
                     TabIndex        =   118
                     Top             =   285
                     Width           =   375
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·√ﬁ”«ÿ"
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
                     Height          =   285
                     Index           =   38
                     Left            =   9570
                     TabIndex        =   117
                     Top             =   285
                     Width           =   960
                  End
                  Begin VB.Label LblInstallTotal 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   12585
                     TabIndex        =   116
                     Top             =   285
                     Width           =   690
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„»·€ «·ﬂ·Ï"
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
                     Height          =   285
                     Index           =   36
                     Left            =   13425
                     TabIndex        =   115
                     Top             =   285
                     Width           =   885
                  End
                  Begin VB.Label LblPrecenType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   15525
                     TabIndex        =   114
                     Top             =   285
                     Width           =   780
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰Ê⁄ «·›«∆œ…"
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
                     Height          =   285
                     Index           =   34
                     Left            =   16455
                     TabIndex        =   113
                     Top             =   285
                     Width           =   780
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰”»… «·›«∆œ…"
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
                     Height          =   405
                     Index           =   35
                     Left            =   14640
                     TabIndex        =   112
                     Top             =   165
                     Width           =   750
                  End
                  Begin VB.Label LblPrecenValue1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   13620
                     TabIndex        =   111
                     Top             =   285
                     Width           =   765
                  End
                  Begin VB.Label LblInstallSeprator 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   285
                     Left            =   5475
                     TabIndex        =   110
                     Top             =   285
                     Width           =   240
                  End
                  Begin VB.Label LblStartValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   165
                     Left            =   2550
                     TabIndex        =   109
                     Top             =   285
                     Width           =   720
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„… «·„»œ∆Ì…"
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
                     Height          =   285
                     Index           =   37
                     Left            =   3285
                     TabIndex        =   108
                     Top             =   285
                     Width           =   1110
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   675
                  Index           =   14
                  Left            =   90
                  TabIndex        =   123
                  TabStop         =   0   'False
                  Top             =   3450
                  Width           =   20115
                  _cx             =   35481
                  _cy             =   1191
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
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Ìﬂ« "
                     Height          =   495
                     Index           =   2
                     Left            =   11820
                     TabIndex        =   124
                     Top             =   0
                     Width           =   915
                  End
                  Begin ImpulseButton.ISButton CmdCheque 
                     Height          =   510
                     Left            =   3690
                     TabIndex        =   125
                     Top             =   0
                     Width           =   1485
                     _ExtentX        =   2619
                     _ExtentY        =   900
                     ButtonStyle     =   1
                     Caption         =   " ”ÃÌ· «·‘Ìﬂ« "
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
                     DrawFocusRectangle=   0   'False
                  End
                  Begin MSDataListLib.DataCombo Dcbanks 
                     Height          =   315
                     Left            =   13320
                     TabIndex        =   142
                     Top             =   0
                     Width           =   1500
                     _ExtentX        =   2646
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   "6"
                     BoundColumn     =   ""
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label2 
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·»‰ﬂ"
                     Height          =   315
                     Left            =   15060
                     TabIndex        =   143
                     Top             =   0
                     Width           =   420
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   375
                     Index           =   19
                     Left            =   8370
                     TabIndex        =   129
                     Top             =   105
                     Width           =   1275
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·‘Ìﬂ« "
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
                     Height          =   495
                     Index           =   17
                     Left            =   9765
                     TabIndex        =   128
                     Top             =   105
                     Width           =   1260
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "≈Ã„«·Ï ﬁÌ„… «·‘Ìﬂ« "
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
                     Height          =   495
                     Index           =   16
                     Left            =   6270
                     TabIndex        =   127
                     Top             =   105
                     Width           =   1860
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   495
                     Index           =   18
                     Left            =   5175
                     TabIndex        =   126
                     Top             =   105
                     Width           =   1065
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FgCheques 
                  Height          =   3300
                  Left            =   90
                  TabIndex        =   130
                  Top             =   4185
                  Width           =   17355
                  _cx             =   30612
                  _cy             =   5821
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
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmSaleBill2.frx":31314
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
         Height          =   735
         Index           =   9
         Left            =   15
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   -2385
         Visible         =   0   'False
         Width           =   19170
         _cx             =   33814
         _cy             =   1296
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
            Left            =   8280
            Style           =   1  'Graphical
            TabIndex        =   148
            Top             =   240
            Width           =   5010
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9090
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Top             =   0
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   8010
            TabIndex        =   61
            Top             =   0
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   7305
            TabIndex        =   60
            Top             =   0
            Visible         =   0   'False
            Width           =   660
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
            Left            =   14235
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   120
            Visible         =   0   'False
            Width           =   2220
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   2985
            TabIndex        =   17
            Top             =   30
            Width           =   1320
            _ExtentX        =   2328
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
            ButtonImage     =   "FrmSaleBill2.frx":31449
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
            Left            =   1620
            TabIndex        =   18
            Top             =   30
            Width           =   1305
            _ExtentX        =   2302
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
            ButtonImage     =   "FrmSaleBill2.frx":317E3
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
            Left            =   4335
            TabIndex        =   19
            Top             =   30
            Width           =   1260
            _ExtentX        =   2223
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
            ButtonImage     =   "FrmSaleBill2.frx":31B7D
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
            TabIndex        =   20
            Top             =   30
            Width           =   1455
            _ExtentX        =   2566
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
            ButtonImage     =   "FrmSaleBill2.frx":31F17
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
            Left            =   11475
            TabIndex        =   30
            Top             =   120
            Width           =   1290
            _ExtentX        =   2275
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
            ButtonImage     =   "FrmSaleBill2.frx":322B1
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdRetruns 
            Height          =   345
            Left            =   5175
            TabIndex        =   31
            Top             =   0
            Width           =   1485
            _ExtentX        =   2619
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
            ButtonImage     =   "FrmSaleBill2.frx":3264B
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdInfo 
            Height          =   615
            Left            =   6435
            TabIndex        =   69
            Top             =   -120
            Width           =   945
            _ExtentX        =   1667
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
            ButtonImage     =   "FrmSaleBill2.frx":32BE5
            ButtonImageHover=   "FrmSaleBill2.frx":338BF
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
            Left            =   7365
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   0
            Width           =   7995
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
            Left            =   195
            TabIndex        =   32
            Top             =   390
            Width           =   11265
         End
      End
   End
End
Attribute VB_Name = "frmsalebill2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim Pay_Print As Integer
Dim PayMode As Boolean
Dim IsVouc         As Boolean
     Dim isFound As Boolean
    Dim IsPayType As Boolean
Public inddx As Integer
Dim rs As ADODB.Recordset
  Dim PayDes As String
Dim TTP As clstooltip
Dim NewGrid As New ClsGrid
Dim SaleReport As ClsSaleReport
Dim cSearchDcbo(4)   As clsDCboSearch
Dim Dcombos As ClsDataCombos
Dim TxtNoteSerial1V As String
Public BolPrint As Boolean
Public TimeOut_InSec As Long
Dim zatcaStatus As Integer
Dim Export As Integer

Dim imageCounter As Integer
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
Dim SpecialOffer As Integer
Dim Sales As Double
Dim GetFree As Double
Dim discount As Double
Dim FromPrice As Double '0 min   1  max
             
 Dim isFromExcel As Boolean
Dim first_run As Boolean
Dim bank_account As String
Dim general_noteid As Long
Dim RsNotesGeneral As ADODB.Recordset
Dim CurrentVoucherNo As String
Dim CurrentVoucherSerialNo As String
Dim visapayed As Double
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
'********************
Private OldSerialNo As String
Private LoadExcelFlage As Boolean
'*********************
Private Sub btnpay_Click(Index As Integer)
    'NoOfItemsHaveOffersInBill «·Œ«’…
    Dim NoOfItemsInOffers As Double
    Dim OfferCount        As Double
    Dim RowNum            As Integer
    NoOfItemsInOffers = 0
    If XPCboDiscountType.ListIndex = 3 And val(XPTxtDiscountVal.text) <> 0 Then
        If val(XPTxtDiscountVal.text) > GetBalance() Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ «·Œ’„ «ﬂ»— „‰ —’Ìœ «·⁄„Ì· ÕÌÀ «‰ —’Ìœ «·⁄„Ì· ÂÊ " & GetBalance
            Else
                MsgBox "Discount Value larger than Balance " & GetBalance
            End If

            Screen.MousePointer = vbDefault
            XPTxtDiscountVal.SetFocus
            Exit Sub
        End If
    End If
    If SystemOptions.CashCustomerNameMustenter = True Then
        If CashCustomerName.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ «œŒ«· «”„ «·⁄„Ì·"
            Else
                MsgBox "Please Enter Customer"
            End If
            CashCustomerName.SetFocus
            Exit Sub
        End If
        If TxtPhone(0).text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ «œŒ«·  —ﬁ„ «· ·Ì›Ê‰"
            Else
                MsgBox "Please Enter No.Telephone"
            End If
            TxtPhone(0).SetFocus
            Exit Sub
        End If
    End If
           
    '«· √ﬂœ „‰ Õ’„ «·„‰œÊ» Ê«ﬁ’Ì ’·«ÕÌ…

    If SystemOptions.usertype <> UserAdminAll Then
        If SystemOptions.EmpNotExcceedDiscount = True Then
            If checkEmpDiscount(val(DcboEmp.BoundText), val(LblTotalAll.Caption), val(LblDiscountsTotal.Caption)) = False Then
                MsgBox "«·Œ’„ Ì ⁄œÌ «·„”„ÊÕ »Â ··„‰œÊ» Ê·« Ì„ﬂ‰  ﬂ„·… «·Õ›Ÿ", vbCritical
                Exit Sub
            End If
                
        End If
 
    End If
    RelinVatGrid
    
    If lbl(57).Visible = True Then

        Dim discId    As Double
        Dim discValue As Double

        With FG 'clear old
            For RowNum = 1 To FG.rows - 1
                FG.TextMatrix(RowNum, FG.ColIndex("MinID")) = ""
                .TextMatrix(RowNum, .ColIndex("DiscountType")) = 1
                .TextMatrix(RowNum, .ColIndex("DiscountVal")) = 0
                If CheckItem(val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))), Me.XPDtbBill.value, discId, discValue) = True Then
                                    
                    If .ColIndex("DiscountType") <> -1 Then
                                                                      
                        '0 value 1 percentage
                        If discId = 0 Then
                            discId = 2
                        ElseIf discId = 1 Then
                            discId = 3
                        ElseIf discId = 4 Then
                            discId = 4
                            GoTo ll
                        End If
                                                                                      
                        .cell(flexcpData, RowNum, .ColIndex("DiscountType")) = discId
                        .TextMatrix(RowNum, .ColIndex("DiscountType")) = discId
                        .TextMatrix(RowNum, .ColIndex("DiscountVal")) = discValue
                    End If
                                          
                End If
                                          
ll:
            Next RowNum
        End With
        NewGrid.Calculate 1, , , True
           
        For RowNum = 1 To FG.rows - 1
            
            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                             
                If CheckItemSpecialOffer(val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))), Me.XPDtbBill, 1, Sales, GetFree, discount, FromPrice, 4, val(Me.dcBranch.BoundText)) = True Then
                              
                    FG.TextMatrix(RowNum, FG.ColIndex("SpecialOffer")) = 1
                    NoOfItemsInOffers = NoOfItemsInOffers + 1
                                     
                Else
                             
                    FG.TextMatrix(RowNum, FG.ColIndex("SpecialOffer")) = 0
                                           
                End If
                        
            End If
        Next RowNum

    End If
    Dim j As Integer
    If NoOfItemsInOffers >= Sales Then '⁄œœ «·«’‰«› «·Œ«÷⁄Â ··⁄—÷  ›Ì «·ﬁ« Ê—… «ﬂÀ— „‰ «Ê Ì”«ÊÌ «·⁄—÷

        If (GetFree + Sales) > 0 Then
            OfferCount = NoOfItemsInOffers \ (GetFree + Sales) '«Ã„«·Ì «·⁄—Ê÷
        
            '«·Õ’Ê· ⁄·Ì «—ﬁ«„ «·«’‰«› ··Œ’„ ÿ»ﬁ« ··”Ì«”…
            If FromPrice = 0 Then 'min
                    
                If GetFree = 1 And Sales = 1 Then
                    '               " — Ì» ÿ»ﬁ« ·«⁄·Ì ”⁄—
                    FG.ExplorerBar = 2
                    FG.SelectionMode = 0
                    With FG
                        .Select 1, FG.ColIndex("Price")
                        FG.Sort = flexSortGenericDescending
                        For RowNum = 1 To FG.rows - 1
                            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                                If RowNum Mod 2 = 0 Then
                                    If discount = 100 Then '„Ã«‰Ì
                                        .cell(flexcpData, RowNum, .ColIndex("DiscountType")) = 4
                                        .TextMatrix(RowNum, .ColIndex("DiscountType")) = 4
                                        .TextMatrix(RowNum, .ColIndex("DiscountVal")) = 0
                                    Else
                                        .cell(flexcpData, RowNum, .ColIndex("DiscountType")) = 2
                                        .TextMatrix(RowNum, .ColIndex("DiscountType")) = 2
                                        .TextMatrix(RowNum, .ColIndex("DiscountVal")) = discount
                                                                     
                                    End If
                                                                    
                                End If
                            End If
                        Next RowNum
                                        
                    End With
               
                Else ' «·‘ﬂ· «·ﬁœÌ„ ›Ì Õ«·Â 2 Ê«ﬂÀ—
                    For j = 1 To OfferCount
                        Getmin (discount)
                    Next j
                End If
              
            ElseIf FromPrice = 0 Then 'min
                    
                '        For j = 1 To OfferCount
                '           Getmin
                '       Next j
                    
            End If
        
        End If

    End If
    NewGrid.Calculate 1, , , True
          
    '   Exit Sub
        
    If optsale(1).value = True Then
        If checkretutn = False Then

            Exit Sub
        End If
    End If
     
    If SystemOptions.usertype <> UserAdminAll Then
        If SystemOptions.EmpNotExcceedDiscount = True Then
            If checkEmpDiscount(val(DcboEmp.BoundText), val(LblTotalAll.Caption), val(LblDiscountsTotalView(0).Caption)) = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«·Œ’„ Ì ⁄œÌ «·„”„ÊÕ »Â ··„‰œÊ» Ê·« Ì„ﬂ‰  ﬂ„·… «·Õ›Ÿ", vbCritical
                Else
                    MsgBox "Discount not allowed", vbCritical
                End If
                                    
                Exit Sub
            End If
                
        End If
 
    End If
    
    FillGridWithData
    If optsale(1).value = True Then
        If CheckFilegrid() = False Then
     
            Exit Sub
        End If
     
        Grid.Enabled = False
        Frame13.Enabled = False
    Else
        Grid.Enabled = True
        Frame13.Enabled = True
    End If

    Cmd_Click (7)
    ReLineGrid
    If SystemOptions.UserInterface = EnglishInterface Then
        FrmCustomerDisplay.LblInformation2.Caption = " Total " & "" & TxtNetValue.text
    Else
        FrmCustomerDisplay.LblInformation2.Caption = " «·«Ã„«·Ì " & "" & TxtNetValue.text 'vbNewLine
    End If
    LBLPayVal.Caption = ""
    If optsale(1).value = True Then

        LBLPayVal.Caption = val(TxtNetValue.text) '+ val(TxtValueAdded.Text)
 
        With Grid
            .TextMatrix(1, .ColIndex("Value")) = LBLPayVal.Caption
        End With
        ReLineGrid
   
    End If
End Sub

Private Sub ChecVAT_Click()
  Dim i As Integer
If Me.TxtModFlg.text <> "R" Then
    If ChecVAT.value = vbChecked Then

        With Me.VatGrid
 
            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = True
            Next i

        End With

    Else

        With Me.VatGrid

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = False
            Next i

        End With

    End If
    RelinVatGrid
    End If
End Sub
Sub RelinVatGrid()
Dim i As Integer
Dim SmValu As Double
SmValu = 0
With VatGrid
For i = 1 To .rows - 1
If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
SmValu = SmValu + val(.TextMatrix(i, .ColIndex("Vat")))
End If
Next i
End With
LblDiscountsTotalView(6).Caption = Format(SmValu, ".##")
TxtValueAdded.text = Format(SmValu, ".##")

'showComm
If SmValu <> 0 Then
 NewGrid.Calculate 1, , , True
 End If
 showComm
 'LblTotalAll.Caption = val(LblTotalAll.Caption) ' - val(TxtValueAdded.text)
LblTotal.Caption = val(LblTotalAll.Caption) - val(LblDiscountsTotal.Caption) + IIf(SystemOptions.PriceWithVAT = True, 0, val(TxtValueAdded.text)) '- SmVal
LblTotal.Caption = Round(LblTotal.Caption, 2)
LBLPayVal.Caption = val(TxtNetValue.text) + IIf(SystemOptions.PriceWithVAT = True, 0, val(TxtValueAdded.text))

End Sub
Private Sub SaveDataPanding()
    Dim Msg As String
     Dim TotalValue As Variant
     
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
        Dim ItemsGoodsTotalsnew As Variant
        Dim ItemsServiceTotalsnew As Variant
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
' On Error GoTo ErrTrap
'
    Me.FG.FinishEditing True
ItemsGoodsTotalsnew = 0
ItemsServiceTotalsnew = 0
    DoEvents
    Screen.MousePointer = vbArrowHourglass
    my_branch = val(Me.dcBranch.BoundText)



Dim NoteID As Long
  Dim NoteDate As Date
    Dim NoteSerial As String
    Dim Notevalue As Double
    Dim des As String
       StrSqlDel = "delete From Transactions where Transaction_ID=" & val(Me.TxtPhone(3).text) 'Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        
     StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(Me.TxtPhone(3).text) 'Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
    '---------------------------------
    Set RSTransDetails = New ADODB.Recordset
 
  StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Set RsNotes = New ADODB.Recordset
 StrSQL = "SELECT    * from dbo.Transactions Where (1 = -1)"
   Rs3.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


    Screen.MousePointer = vbArrowHourglass
    Cn.BeginTrans
    TransBegine = True
        Rs3.AddNew
         TxtPhone(3).text = CStr(new_id("Transactions", "Transaction_ID", "", True))
         Rs3("Transaction_ID").value = val(TxtPhone(3).text)
    '  Rs3("ExtraAccount").value = IIf(DCExtraAccount.BoundText = "", Null, (DCExtraAccount.BoundText))

  '  If DCExtraAccount.BoundText = "" Then
   '     Rs3("ExtraValue").value = 0
     '   TxtExtraValue.Text = 0
   ' Else
   '     Rs3("ExtraValue").value = val(TxtExtraValue.Text)
   ' End If
' Rs3("AdvPay").value = IIf(txtAdvPay.Text = "", 0, val(txtAdvPay.Text))

' Rs3("CustomerlocationID").value = IIf(Me.DCCustomerLocation.BoundText = "", 0, val(DCCustomerLocation.BoundText))
    Rs3("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
    Rs3("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
     
   ' Rs3("CarId").value = IIf(Me.DCCar.BoundText = "", Null, (Me.DCCar.BoundText))
 '   Rs3("DriverId").value = IIf(Me.DCDriver.BoundText = "", Null, (Me.DCDriver.BoundText))
    
    Rs3("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    Rs3("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
 '   Rs3("Fullcode").value = IIf(DCPreFix.BoundText = "", Null, DCPreFix.Text) & IIf(Trim(TxtNoteSerial1.Text) = "", Null, TxtNoteSerial1.Text)
   ' Rs3("Prefix").value = IIf(DCPreFix.Text = "", Null, DCPreFix.Text)
    Rs3("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '

    If CboPayMentType.ListIndex = 0 Or CboPayMentType.ListIndex = 2 Then
        Rs3("BoxID").value = IIf(DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
    Else
        Rs3("BoxID").value = Null
      
    End If
    Rs3("RecTime").value = Time
    Rs3("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.text) = "", "", Trim(Me.TxtTransSerial.text))
    Rs3("Transaction_Date").value = XPDtbBill.value
    Rs3("Transaction_Type").value = 70
   ' Rs3("Fromdate").value = DpFrom.value
   ' Rs3("todate").value = DpTo.value

    Rs3("UserID").value = val(DCboUserName.BoundText)
    Rs3("nots").value = ""

   '  If CBoBasedON.ListIndex = -1 Then
  '      Rs3("CBoBasedON").value = 0
  '  Else
   '     Rs3("CBoBasedON").value = val(CBoBasedON.ListIndex)
   ' End If
    
    Rs3("Currency_id").value = IIf(DcCurrency.BoundText = "", Null, val(DcCurrency.BoundText))
    Rs3("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.text), 1, txt_Currency_rate.text)

    If XPCboDiscountType.ListIndex = -1 Then
        Rs3("Trans_DiscountType").value = 0
    Else
        Rs3("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If

   Rs3("Trans_Discount").value = IIf(XPTxtDiscountVal.text = "", Null, val(XPTxtDiscountVal.text))
    Rs3("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
    Rs3("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
    Rs3("order_no") = IIf(TXTOrDer_no.text = "", Null, (TXTOrDer_no.text))
   ' Rs3("PurchaseBill") = IIf(TxtPurchaseBill.Text = "", Null, val(TxtPurchaseBill.Text))

    If CboPayMentType.ListIndex = -1 Then
        Rs3("PaymentType").value = 0
    Else
        Rs3("PaymentType").value = val(CboPayMentType.ListIndex)
    End If
'Rs3("LawFirmValue").value = IIf(LawFirmValue.Text = "", Null, val(LawFirmValue.Text))
'Rs3("TotalQest").value = IIf(TotalQest.Text = "", Null, val(TotalQest.Text))
'Rs3("QstValue").value = IIf(QstValue.Text = "", Null, val(QstValue.Text))
'Rs3("QstNo").value = IIf(QstNo.Text = "", Null, val(QstNo.Text))
'Rs3("Sandts").value = IIf(Sandts.Text = "", Null, (Sandts.Text))
'Rs3("QestStartDate").value = QestStartDate.value
'Rs3("QestEndtDate").value = QestEndtDate.value
'Rs3("QestStartDateH").value = QestStartDateH.value
'Rs3("QestEndtDateH").value = QestEndtDateH.value
 'If OptInt(0).value = True Then
 'Rs3("YMD").value = 0
 'ElseIf OptInt(1).value = True Then
 'Rs3("YMD").value = 1
 ' ElseIf OptInt(2).value = True Then
 'Rs3("YMD").value = 2
 'End If
'''/////////
    Rs3("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
    Rs3("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)
    'Rs3("LocationID").value = IIf(DCGroupID.BoundText = "", Null, DCGroupID.BoundText)
    'ChkInstall 11 10 2012
    If ChkInstall.value = vbChecked Then
        Rs3("ChkInstall").value = 1
    Else
        Rs3("ChkInstall").value = 0
    End If

    If Me.CboSaleType.ListIndex = 0 Or Me.CboSaleType.ListIndex = -1 Then
        Rs3("SaleType").value = 0
    Else
        Rs3("SaleType").value = 1
    End If

    If Trim$(Me.TxtCashCustomerName.text) <> "" Then
        Rs3("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.text)
    Else
        Rs3("CashCustomerName").value = Null
    End If
     Rs3("VATNO").value = IIf(Trim(Me.TxtPhone(1).text) = "", Null, Trim(Me.TxtPhone(1).text))
    If Trim$(Me.TxtPhone(0).text) <> "" Then
        Rs3("CashCustomerPhone").value = Trim$(Me.TxtPhone(0).text)
    Else
        Rs3("CashCustomerPhone").value = Null
    End If
    
    Rs3("TransactionComment").value = IIf(Trim$(TxtBillComment.text) = "", Null, Trim$(TxtBillComment.text))

    '÷—»Ì… Œ’„ Ê≈÷«›…
    If ChkTaxAdd.value = vbChecked And val(Me.TxtTaxAddValue.text) > 0 Then
        Rs3("TaxAddValue").value = val(Me.TxtTaxAddValue.text)
    Else
        Rs3("TaxAddValue").value = 0
    End If

    '÷—»Ì… œ„€…
    If ChkTaxStamp.value = vbChecked And val(Me.TxtTaxStampValue.text) > 0 Then
        Rs3("TaxStampValue").value = val(Me.TxtTaxStampValue.text)
    Else
        Rs3("TaxStampValue").value = 0
    End If

    '÷—»Ì… Œœ„…
    If ChkTaxSerivce.value = vbChecked And val(Me.TxtTaxServiceValue.text) > 0 Then
        Rs3("TaxServiceValue").value = val(Me.TxtTaxServiceValue.text)
    Else
        Rs3("TaxServiceValue").value = 0
    End If

    '»Ì«‰«  ÃœÌœ…
    Rs3("PaymentNetid").value = IIf(DCPaymentNet.BoundText = "", Null, val(DCPaymentNet.BoundText))
    Rs3("NetValue").value = IIf(TxtNetValue.text = "", Null, val(TxtNetValue.text))
    Rs3("PayedValue").value = IIf(TxtPayedValue.text = "", Null, val(TxtPayedValue.text))
    Rs3("RemainValue").value = IIf(TxtRemainValue.text = "", Null, val(TxtRemainValue.text))
   ' Rs3("lotNo").value = IIf(TxtLotNo.Text = "", Null, (TxtLotNo.Text))
  
    Rs3("ManualNo1").value = IIf(TxtManualNo1.text = "", Null, val(TxtManualNo1.text))
    Rs3("ManualNo2").value = IIf(TxtManualNo2.text = "", Null, val(TxtManualNo2.text))
  
    If BillBasedOn(0).value = True Then
        Rs3("BillBasedOn").value = 0
    ElseIf BillBasedOn(1).value = True Then
        Rs3("BillBasedOn").value = 1
    ElseIf BillBasedOn(2).value = True Then
        Rs3("BillBasedOn").value = 2
    End If
    
    '‰ﬁ«ÿ «·»Ì⁄
    Rs3("Printed").value = 1
   

       
    '‰ﬁ«ÿ «·»Ì⁄
    If CboPOSBillType.ListIndex = 0 Then
        Rs3("POSBillType").value = 0
        Rs3("STableID").value = val(LblStableID.Caption)
    Else
        Rs3("POSBillType").value = val(CboPOSBillType.ListIndex)
        Rs3("STableID").value = Null
    End If

  '  Rs3("SessionD").value = lblSessionD
    ''//26 05 2015
'Rs3("ManualNO").value = IIf(Me.txtManualNO.Text = "", Null, txtManualNO.Text)
    Rs3.update

    For RowNum = 1 To FG.rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then

            'Check Repeat Serial

            RSTransDetails.AddNew
            
            RSTransDetails("printed").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("printed")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("printed")))
            RSTransDetails("PrintName").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("PrintName")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("PrintName")))
            
            
            RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))

            RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
            RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
            RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))

            RSTransDetails("Transaction_ID").value = val(TxtPhone(3).text)
            RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))

            'RSTransDetails("Quantity").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
            '            RSTransDetails("ItemName").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Name")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Name"))))
            If Not FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "" Then
                StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Name"))
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsTemp.EOF Or RsTemp.BOF) Then
                    If RsTemp("HaveSerial").value = True Then
                        RSTransDetails("ItemSerial").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("Serial")))
                    End If
                End If

                RsTemp.Close
            End If

            RSTransDetails("ShowPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
            RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
            
            RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
            
            RSTransDetails("guaranteeTime").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
          RSTransDetails("ParrtNoCode").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode"))))
  RSTransDetails("ItemDetailedCode").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode"))))
  
            
            
            RSTransDetails("CostTransID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("PofTransID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("PofTransID"))))
            RSTransDetails("ItemProfit").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit"))))
        
            RSTransDetails("UnitID").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
          
            If SystemOptions.TypicalProduction = False Then
  
                RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , RSTransDetails("UnitID").value)

                If RSTransDetails("CostPrice").value = 0 Then
                    RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , LastPurPriceType, , , XPDtbBill.value, , RSTransDetails("UnitID").value)
                    
                End If
                  
            Else
                RSTransDetails("CostPrice").value = 0
            
            End If
            
               If optsale(1).value = True Then   ' return sallimng
                    RSTransDetails("CostPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))))
           
             
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
            LngUnitID = val(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
            DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                RSTransDetails("OpeningSalesQty").value = RSTransDetails("Quantity").value
                RSTransDetails("OpeningSalesValue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Valu")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Valu"))))
                RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
            
            End If

            SngTemp = SngTemp + (val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))) * RSTransDetails("quantity").value)
         
            If Me.XPCboDiscountType.ListIndex = 1 Then
                TotalBillDiscount = IIf(XPTxtDiscountVal.text = "", Null, (XPTxtDiscountVal.text))
                     'XPTxtDiscountVal
                     
            ElseIf XPCboDiscountType.ListIndex = 2 Then

                If XPTxtDiscountVal.text <> "" Then
                '    TotalBillDiscount = IIf(XPTxtDiscountVal.Text = "", Null, (XPTxtDiscountVal.Text)) * val(LblTotalAll.Caption) / 100
                   TotalBillDiscount = IIf(XPTxtDiscountVal.text = "", Null, (XPTxtDiscountVal.text)) * val(LBLGross.Caption) / 100
                                                         
                Else
                    TotalBillDiscount = 0
                End If
            End If
  
       If LblTotalAll.Caption > 0 Then
           If val(FG.TextMatrix(RowNum, FG.ColIndex("Valu"))) > 0 Then
           
          TotalDiscountPerLine = FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / (LBLGross) * TotalBillDiscount
           
         TotalDiscountPerLine = Round(TotalDiscountPerLine, 2)
           Else
           TotalDiscountPerLine = 0
           End If
         If val(FG.TextMatrix(RowNum, FG.ColIndex("itemtype"))) = 1 Then
                                                                                
         'ItemsServiceTotalsnew = ItemsServiceTotalsnew + TotalDiscountPerLine + val(Fg.TextMatrix(RowNum, Fg.ColIndex("discountvalue")))
         Else
         'ItemsGoodsTotalsnew = ItemsGoodsTotalsnew + TotalDiscountPerLine + val(Fg.TextMatrix(RowNum, Fg.ColIndex("discountvalue")))
         End If
 Else
 TotalDiscountPerLine = 0
 End If
     
     
            RSTransDetails("TotalDiscountPerLine") = Round(TotalDiscountPerLine, 20)
       
            RSTransDetails.update
            '-------------
        
        End If

    Next RowNum


    Cn.CommitTrans

    TransBegine = False

    Screen.MousePointer = vbDefault

    Exit Sub
ErrTrap:

    If TransBegine = True Then
        TransBegine = False
        Cn.RollbackTrans
    End If

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
            Msg = "·« Ì„ﬂ‰  ⁄·Ìﬁ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
            Msg = Msg & CHR(13) & Err.Description
            Msg = Msg & CHR(13) & Err.Number
            Msg = Msg & CHR(13) & Err.Source
            Msg = Msg & CHR(13) & Err.LastDllError
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Else
            Msg = "Can't Pending error in Data" & CHR(13)
        End If
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
       Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡  ⁄·Ìﬁ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry........Error During Save " & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Function checkretutn() As Boolean
    Dim Msg As String
    checkretutn = False
  If SystemOptions.ReturnSallingOption = True Then
      If val(Me.TxtInvSerial.text) = 0 Or val(TxtInvID.text) = 0 Then
            
            If SystemOptions.UserInterface = ArabicInterface Then
              Msg = "»—Ã«¡ ﬂ «»… —ﬁ„ «·›« Ê—… ·Ì „ ⁄—÷Â«..!!"
            Else
            Msg = "Plz enter invoice number"
            End If
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If
    Else
    checkretutn = True
         Exit Function
  End If



    Dim Transaction_ID As Long
    Dim Transaction_Date  As Date
   
    If Not IsDate(txtInvDate) Then Exit Function
    '    If SystemOptions.ReturnSallingOption = True Then
    Dim NoofDays As Integer

    If Me.TxtModFlg = "R" Or Me.TxtModFlg = "" Then Exit Function
    NoofDays = DateDiff("d", IIf(IsDate(Me.txtInvDate.text), Me.txtInvDate.text, Date), Me.XPDtbBill.value)
 
    If opt(0).value = True Then
        If NoofDays > SystemOptions.ReturnSallingIntervalCount Then
If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " ·« Ì„ﬂ‰ «—Ã«⁄ Â–… «·›« Ê—… ·«‰ «·Õœ «·«ﬁ’Ï ··«—Ã«⁄ " & SystemOptions.ReturnSallingIntervalCount & "  ÌÊ„ " & CHR(13)
            Msg = Msg & " «·›« Ê—Â „‰  " & NoofDays & "  ÌÊ„ "
Else
            Msg = "Can't return : you must return in " & SystemOptions.ReturnSallingIntervalCount & "  Day " & CHR(13)
            Msg = Msg & "Invoice From " & NoofDays & " Day "
End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            checkretutn = False
            Exit Function
        
        End If
  
    Else

        If NoofDays > SystemOptions.ReturnSallingIntervalCount1 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " ·« Ì„ﬂ‰ «” »œ«·  Â–… «·›« Ê—… ·«‰ «·Õœ «·«ﬁ’Ï ··«” »œ«· " & SystemOptions.ReturnSallingIntervalCount1 & "  ÌÊ„ " & CHR(13)
            Msg = Msg & " «·›« Ê—Â „‰  " & NoofDays & "  ÌÊ„ "
       Else
               Msg = "Can't return then invoice ; Return intrtval id" & SystemOptions.ReturnSallingIntervalCount1 & "  day " & CHR(13)
            Msg = Msg & "invoice from " & NoofDays & "  day "
       End If
       
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            checkretutn = False
            Exit Function
        End If

    End If
   
    checkretutn = True
         
    'End If
End Function

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

Function addrow(ItemID As Integer, ItemName As String, ITEMPRICE As Double, ItemType As Integer)
    lblqty.Caption = ""
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
    Dim des As String
    On Error Resume Next
    
    Me.DCboItemsName.text = ItemName
    TxtQuantity.text = 1
    NewGrid.CmdAddData_Click
    
    With FG
        .Row = .rows - 1
    End With

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

Private Sub ReLineGrid()
    On Error Resume Next
    Dim i As Integer
    Dim IntCounter As Integer
    Dim totalPayed As Double
 totalPayed = 0
 visapayed = 0
  With Grid

        For i = .FixedRows To .rows - 1

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

Private Sub ALLButton8_Click()

End Sub



Private Sub ALLButton1_Click()
'FrmReturnSalling.show
optsale(1).value = True
    If SystemOptions.ReturnSallingOption = True Then

                CboRetrunType.ListIndex = 0
                CboRetrunType.Enabled = False
            End If
            
End Sub

Private Sub ALLButton9_Click()
btnNew_Click 1
 CashierLogout.show
 CashierLogout.ScreenJob = 0
Unload Me
End Sub
Private Sub BtnEdit_Click()
'Cmd_Click (1)
End Sub

Private Sub btnExit_Click(Index As Integer)
    If Index = 0 Then
        btnNew_Click 1

        CashierLogout.show
        CashierLogout.ScreenJob = 1
        Unload Me
   
    ElseIf Index = 1 Then

        FrmEditPW1.FrameAdmin.Visible = True
        FrmEditPW1.LoadAdmins (PPointID)
        FrmEditPW1.show
    ElseIf Index = 2 Then
        Unload FrmBillsPnding
        FrmBillsPnding.Indxx = 2
        Load FrmBillsPnding
        FrmBillsPnding.show vbModal
   
    ElseIf Index = 3 Then
        SaveDataPanding
        btnNew_Click 0
        'Cmd_Click (0)
    ElseIf Index = 4 Then
        '            If checkApility("FrmSearchSerial") = False Then
        '                Exit Sub
        '            End If
        mdifrmmain.Enabled = True
        FrmSearchSerial.show
        Exit Sub
    ElseIf Index = 5 Then
        Dim i As Integer
        With FG
            i = .rows - 1
            Do While i > 0
                If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                    .RemoveItem i
                End If
                i = i - 1
            Loop
        End With
        NewGrid.Calculate 1, , , True

        '        NewGrid.Calculate 1, , , True
        NewGrid.SentTypeVAT
   
    ElseIf Index = 6 Then
        Load FrmRequest
        FrmRequest.show
    ElseIf Index = 7 Then
        Load FrmRequest
        FrmGuaranteeAlram.Ind = 1
        FrmGuaranteeAlram.show
    ElseIf Index = 8 Then
        Load FrmAnalysItems
        FrmAnalysItems.show
        
        FrmAnalysItems.C1Tab1.CurrTab = 1
    ElseIf Index = 9 Then
        FrmReports.show
        FrmReports.C1TabMain.CurrTab = 0
    End If
End Sub

 




Private Sub CmdAdd_LostFocus()
showComm
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
 TxtRemainValue.text = Round(TxtRemainValue.text, 2)

End If

 ReLineGrid
 
End Sub

Private Sub btnNew_Click(Index As Integer)
    If Index = 0 Then
        Cmd_Click (0)
        'XPCboDiscountType.ListIndex = 2
        'XPTxtDiscountVal.text = 25
        'If SystemOptions.WorkWithItemsDetails = True Then
        'TxtItemCodeB.SetFocus
        'End If
   
        btnpay(0).Enabled = True
        btnpay(1).Enabled = True
        btnExit(2).Enabled = True
          If TxtItemCodeB.Enabled Then
        If SystemOptions.WorkWithItemsDetails = True Then
      
            TxtItemCodeB.SetFocus
        Else
            TxtItemCodeB1.SetFocus
        End If
        End If

        'DCboItemsCode.SetFocus
        CboPayMentType.ListIndex = 0
      
        FrmCustomerDisplay.LblInformation.Caption = ""
        FrmCustomerDisplay.LblInformation2.Caption = ""

        Shape6.BorderColor = &H400000
        optsale(0).value = True
        CboPOSBillType.ListIndex = 1

        Image1.Visible = True
        TBar.Visible = True
        SystemOptions.usertype = UserNormal
    ElseIf Index = 1 Then
        Cmd_Click (3)
    ElseIf Index = 3 Then
        If TxtPhone(0).text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ «·ÃÊ«·"
            Else
                MsgBox "Please Enter No.Mobile"
            End If
            TxtPhone(0).SetFocus
            XPTxtDiscountVal.Visible = False
            Exit Sub
        End If
        XPTxtDiscountVal.Visible = True
        XPTxtDiscountVal.text = GetBalance
    ElseIf Index = 2 Then
    LoadExcelFlage = True
        Cmd_Click 0
        
        On Error GoTo eh
        CommonDialog1.DialogTitle = "Select Upload list"
        CommonDialog1.CancelError = True
        CommonDialog1.filter = "xls Files (*.xls)|*.xls"
        CommonDialog1.Flags = cdlOFNAllowMultiselect Or cdlOFNLongNames Or cdlOFNExplorer
        CommonDialog1.ShowOpen
        Dim vFiles() As String
        CboPayMentType.ListIndex = 0
        'grdFiles.Visible = True
    
        '  grdFiles.rows = 1
        Dim Row As Integer
        Row = 1
        Dim i As Integer
        vFiles = Split(CommonDialog1.FileName, CHR(0))
        If UBound(vFiles) = 0 Then
            SaveExcelFile CommonDialog1.FileName
            '            grdFiles.AddItem row
            '            grdFiles.TextMatrix(row, grdFiles.ColIndex("File")) = row
            '            grdFiles.TextMatrix(row, grdFiles.ColIndex("Rows")) = 0
            '            grdFiles.TextMatrix(row, grdFiles.ColIndex("DRows")) = 0
            '            grdFiles.TextMatrix(row, grdFiles.ColIndex("FileName")) = CommonDialog1.FileName
        Else
        
            ' txtMainPath = vFiles(0)
            For i = 1 To UBound(vFiles)
                SaveExcelFile vFiles(0) & "\" & vFiles(i)
                '                grdFiles.AddItem row
                '                grdFiles.TextMatrix(row, grdFiles.ColIndex("File")) = row
                '                grdFiles.TextMatrix(row, grdFiles.ColIndex("Rows")) = 0
                '                grdFiles.TextMatrix(row, grdFiles.ColIndex("DRows")) = 0
                '                grdFiles.TextMatrix(row, grdFiles.ColIndex("FileName")) = vFiles(0) & "\" & vFiles(i)
                Row = Row + 1
            Next
        End If
        ' txtFile.text = CommonDialog1.FileName
   LoadExcelFlage = False
        Exit Sub
eh:
LoadExcelFlage = False
        MsgBox Err.Description
    End If
End Sub

Private Sub SaveExcelFile(FileName)
    On Error GoTo eh

    If FileName = "" Then
        MsgBox "«Œ — „·› «Ê·«"
        Exit Sub
    End If
   
    Dim i              As Long
    Dim s              As String
    Dim mPrice         As Double
    Dim RsData         As New ADODB.Recordset
    Dim AllFinshedRows As Integer
    Dim allExcelRows   As Integer
    Dim startTime      As Date
    Dim moConn         As New ADODB.Connection
    Dim mrs            As ADODB.Recordset
    Dim tblname        As String
    Dim shortFileName  As String
    ' lblTime.Visible = True
    Me.Enabled = False
    moConn.CursorLocation = adUseClient
    Dim rsCheck As New ADODB.Recordset
    Dim sfo     As New FileSystemObject

    'For i = 1 To grdFiles.rows - 1
    '        filename = grdFiles.TextMatrix(i, grdFiles.ColIndex("FileName"))
    shortFileName = sfo.GetFileName(FileName) 'grdFiles.TextMatrix(i, grdFiles.ColIndex("Name"))
    '
    moConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source='" & FileName & "'; Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
    Set mrs = moConn.OpenSchema(adSchemaTables)
    '******************
    Dim POSname   As String
    Dim ItemName  As String
    Dim Qty       As String
    Dim paymethod As String
    Dim SerialNo  As String
    Dim mDate     As String
    
    '****************

    If Not mrs.EOF Then
        tblname = mrs.Fields("table_name").value
        RsData.CursorLocation = adUseClient
        RsData.Open "Select *   from [" & tblname & "]", moConn, adOpenKeyset, adLockReadOnly
        Dim RowID          As Integer
          
        Dim isREcoredSaved As Boolean
        Dim strQuery       As String
        Dim OLDSec         As Long
        Dim Secondes       As Long

        Dim AllSec         As Long
        lbl(98).Visible = True
        RowID = 0
        AllFinshedRows = 0
        Dim currentRows As Long
        Dim AllFileRows As Long
        RsData.MoveLast
        Dim LngFindRow As Long
        AllFileRows = RsData.RecordCount
         allExcelRows = AllFileRows
        RsData.MoveFirst
        'val(grdFiles.TextMatrix(i, grdFiles.ColIndex("Rows")))
        Dim RsStore As New ADODB.Recordset
        
        Do While Not RsData.EOF
        
            POSname = RsData.Fields("‰ﬁÿ… »Ì⁄")
            Dim itemNameTmp As String
            itemNameTmp = RsData.Fields("«·Ê’›")
            Qty = Trim(Split(itemNameTmp, "x")(0))
            ItemName = Trim(Split(itemNameTmp, "x")(1))
            paymethod = RsData.Fields("ÿ—Ìﬁ… «·œ›⁄")
            SerialNo = RsData.Fields("—ﬁ„ «·«Ì’«·")
            mDate = RsData.Fields("«· «—ÌŒ")
            mPrice = val(RsData.Fields("«·»Ì⁄ «·’«›Ì")) / Qty
            RowID = RowID + 1
            AllFinshedRows = AllFinshedRows + 1
            currentRows = currentRows + 1
            ' lbl(32).Caption = "F[" & currentRows & "]>[" & AllFileRows & "] A[" & AllFinshedRows & "]>[" & AllFileRows & "]"
            DoEvents
            '            strQuery = "SELECT Count(*) cnt "
            '            strQuery = strQuery & "From notes_all "
            '            strQuery = strQuery & "WHERE ExcelFile = '" & shortFileName & "' "
            '            strQuery = strQuery & " AND NoteType = 85 "
            '            strQuery = strQuery & "  AND ExcelRow =  " & RowID & " ;"
            '            ' rsCheck.CursorLocation = adUseClient
            '            rsCheck.Open strQuery, Cn, adOpenForwardOnly, adLockReadOnly
            '            isREcoredSaved = rsCheck!cnt > 0
            '            rsCheck.Close

            '*********************
            ' If Not isREcoredSaved Then
            startTime = Now
            btnpay_Click 0
            If Trim(SerialNo) <> "" Then
                'SaveItemsExcelMeth_New RsData, RowID, shortFileName
                '*************************
                '*************************
                oldtxtNoteSerial1.text = SerialNo
                ' XPDtbBill.value = Replace(Replace(Replace(mdate, "„", "PM"), "’", ""), "˛", "")
                Dim arr() As String
                arr = Split(ItemName, " ")
               
                s = ""
                s = s & "SELECT ItemCode,ItemID "
                s = s & "FROM TblItems "
                If UBound(arr) = 0 Then
                    s = s & "WHERE ItemName LIKE N'%" & arr(0) & "%' "  'AND ItemName LIKE '%91%' "
                Else
                    s = s & "WHERE 1 = 1   "
                   
                    For i = 0 To UBound(arr)
                        If arr(i) <> "" Then
                            s = s & " And ItemName LIKE N'%" & arr(i) & "%'  "
                        End If
                    Next
            
                End If
                Dim rsITem As New ADODB.Recordset
                Set rsITem = New ADODB.Recordset
                rsITem.Open s, Cn, adOpenForwardOnly, adLockReadOnly
                If rsITem.EOF Then
                    GoTo NextRow
                Else
                    s = "Select * from TblLink_Item_To_Store_Details2 where ItemId = " & val(rsITem!ItemID & "")
                    Set RsStore = New ADODB.Recordset
                    RsStore.Open s, Cn, adOpenKeyset, adLockReadOnly
                    If Not RsStore.EOF Then
                        DCboStoreName.BoundText = val(RsStore!StoreID & "")
                    End If
                    XPDtbBill.value = mDate
                    txtItemCodeSearch2.text = rsITem!itemcode & ""
                    '   txtItemCodeSearch2_KeyPress vbKeyReturn
                    DCboItemsCode.BoundText = val(rsITem!ItemID & "")
                    TxtQuantity.text = Qty
                    TxtPrice.text = mPrice
                    NewGrid.CmdAddData_Click
                    
                    LngFindRow = Grid.FindRow(paymethod, Grid.FixedRows, Grid.ColIndex("PaymentName"), False, True)
                    Grid.TextMatrix(LngFindRow, Grid.ColIndex("Value")) = mPrice * Qty
                    ReLineGrid
                                    
                End If
                'CMDPAy_Click 2
                isFromExcel = True
                If optsale(1).value = True Then 'return
                    TxtPayedValue.text = TxtNetValue.text
                End If
                Cmd_Click (2)
                
                ' End If
                btnNew_Click 0
                LBLPayVal.Caption = 0
                FramePay.Visible = False

                'SaveData
                ' btnNew_Click 0
            End If

            OLDSec = AllSec
            Secondes = DateDiff("s", startTime, Now)

            AllSec = ((allExcelRows - AllFinshedRows) * Secondes)

            If AllSec = 0 Then
                AllSec = OLDSec
            End If

            lbl(98).Caption = StringDotFormat("{0} of {1} Recored(s)  Estimated Time : {2} ", AllFinshedRows, allExcelRows, GetTimeHour(AllSec))
            '  End If

            '*********************
NextRow:
            RsData.MoveNext
            
        Loop

        RsData.Close
    End If

    mrs.Close
    moConn.Close
    isFromExcel = False
    'Next

    Me.Enabled = True
    lbl(98).Visible = False
    'grdFiles.rows = 1
    MsgBox " „ Õ›Ÿ «·Õ—ﬂ« "
    Exit Sub
eh:
    Me.Enabled = True
    lbl(98).Visible = False
    MsgBox Err.Description
End Sub

Public Function GetTimeHour(ByVal inSec As Double) As String
    Dim ss As Boolean
      Dim aHr As Double
     Dim aMin As Double
        Dim aSec As Double
    ss = (inSec < 0)
    If (inSec <> 0) Then
        inSec = IIf(inSec < 0, -1, 1) * inSec
        aHr = Fix(inSec / 3600)
        aMin = Fix((inSec - (aHr * 3600)) / 60)
        aSec = inSec - (aHr * 3600) - (aMin * 60)
        GetTimeHour = IIf(ss, "-", "") & Format(aHr, "00000") & ":" & Format(aMin, "00") & ":" & Format(aSec, "00")
    Else
        GetTimeHour = "00000:00:00"
    End If
End Function
Private Sub CMDPAy_Click(Index As Integer)
Dim i As Integer
Dim Msg As String

    For i = 1 To Grid.rows - 1
        If val(Grid.TextMatrix(i, Grid.ColIndex("Value"))) <> 0 Then
            isFound = True
        End If
        If Trim(Grid.TextMatrix(i, Grid.ColIndex("PaymentName"))) <> "" Then
            IsPayType = True
        End If
    Next
    
    If Not isFound Then
        Msg = "ÌÃ» ≈œŒ«· ﬁÌ„… «·œ›⁄...!!"
        'MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

                           ' We pause for this number of seconds
    messageResult = MsgBoxPause(hWnd, Msg, App.Title, vbExclamation, MsssageSeconde)
        
      
      '  TxtTransSerial.SetFocus
        Screen.MousePointer = vbDefault
        FillGridWithData
        FramePay.Visible = True
        
        Exit Sub
        
    End If
 
        
        
       Pay_Print = Index
'«· √ﬂœ „‰ Õ’„ «·„‰œÊ» Ê«ﬁ’Ì ’·«ÕÌ…
        If optsale(0).value = True Then
                    If SystemOptions.usertype <> UserAdminAll Then
                                     If SystemOptions.EmpNotExcceedDiscount = True Then
                                                If checkEmpDiscount(val(DcboEmp.BoundText), val(LblTotalAll.Caption), val(LblDiscountsTotal.Caption)) = False Then
                                                 ' MsgBox "«·Œ’„ Ì ⁄œÌ «·„”„ÊÕ »Â ··„‰œÊ» Ê·« Ì„ﬂ‰  ﬂ„·… «·Õ›Ÿ", vbCritical
                                                 Msg = "«·Œ’„ Ì ⁄œÌ «·„”„ÊÕ »Â ··„‰œÊ» Ê·« Ì„ﬂ‰  ﬂ„·… «·Õ›Ÿ"
                                                   messageResult = MsgBoxPause(hWnd, Msg, App.Title, vbExclamation, MsssageSeconde)
        
       
                                                   Exit Sub
                                                End If
                                    
                                    End If
                     
                        End If
    
        End If
        
        
        If optsale(1).value = True Then
              If checkretutn = False Then

                    Exit Sub
                End If
     End If
     
If optsale(1).value = True Then
  If CheckFilegrid() = False Then
     
            Exit Sub
End If
       End If
       
  If SystemOptions.usertype <> UserAdminAll Then
                 If SystemOptions.EmpNotExcceedDiscount = True Then
                            If checkEmpDiscount(val(DcboEmp.BoundText), val(LblTotalAll.Caption), val(LblDiscountsTotalView(0).Caption)) = False Then
                                    If SystemOptions.UserInterface = ArabicInterface Then
                                  '  MsgBox "«·Œ’„ Ì ⁄œÌ «·„”„ÊÕ »Â ··„‰œÊ» Ê·« Ì„ﬂ‰  ﬂ„·… «·Õ›Ÿ", vbCritical
                                                   Msg = "«·Œ’„ Ì ⁄œÌ «·„”„ÊÕ »Â ··„‰œÊ» Ê·« Ì„ﬂ‰  ﬂ„·… «·Õ›Ÿ"
                                    Else
                                  '  MsgBox "Discount not allowed", vbCritical
                                                   Msg = "Discount not allowed"
                                    End If
                                    
                                                  
                                                   messageResult = MsgBoxPause(hWnd, Msg, App.Title, vbExclamation, MsssageSeconde)
        
                                    
                                       Exit Sub
                            End If
                
                End If
 
    End If
    
CMDPAy(0).Enabled = False
CMDPAy(1).Enabled = False
Dim AskOption As Boolean
 
If optsale(1).value = True Then 'return
TxtPayedValue.text = TxtNetValue.text
End If



'************************************************************************************
         Dim RowNum As Integer
    For RowNum = 1 To Grid.rows - 1
            
                       If val(Grid.TextMatrix(RowNum, Grid.ColIndex("Value"))) < 0 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                             'MsgBox "Œÿ√ ·« Ì„ﬂ‰ «œŒ«· ﬁÌ„… ”«·»…" & CHR(13) & Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName"))
                             Msg = "Œÿ√ ·« Ì„ﬂ‰ «œŒ«· ﬁÌ„… ”«·»…" & CHR(13) & Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName"))
                        Else
                              '                       MsgBox "ERROR Negative Value  " & CHR(13) & Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName"))
                              Msg = "ERROR Negative Value  " & CHR(13) & Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName"))
                              
                        End If
                        
                                                   messageResult = MsgBoxPause(hWnd, Msg, App.Title, vbExclamation, MsssageSeconde)
        
        
                        
                        CMDPAy(0).Enabled = True
                        CMDPAy(1).Enabled = True
                            Exit Sub
                    End If
   Next RowNum
   
   
   
 
    For RowNum = 2 To Grid.rows - 1
            
                       If val(Grid.TextMatrix(RowNum, Grid.ColIndex("Value"))) > val(TxtNetValue.text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                           '  MsgBox "Œÿ√ ·« Ì„ﬂ‰ «œŒ«· ﬁÌ„… «ﬂ»— „‰  «·ﬁÌ„Â «·«Ã„«·Ì…" & CHR(13) & Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName"))
                           Msg = "Œÿ√ ·« Ì„ﬂ‰ «œŒ«· ﬁÌ„… «ﬂ»— „‰  «·ﬁÌ„Â «·«Ã„«·Ì…" & CHR(13) & Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName"))
                        Else
                               'MsgBox "ERROR Incorrect Value" & CHR(13) & Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName"))
                                Msg = "ERROR Incorrect Value" & CHR(13) & Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName"))
                        End If
                        
                       'Msg = "«·Œ’„ Ì ⁄œÌ «·„”„ÊÕ »Â ··„‰œÊ» Ê·« Ì„ﬂ‰  ﬂ„·… «·Õ›Ÿ"
                                                   messageResult = MsgBoxPause(hWnd, Msg, App.Title, vbExclamation, MsssageSeconde)
         
         
                        CMDPAy(0).Enabled = True
                        CMDPAy(1).Enabled = True
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
             
                   'MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                       messageResult = MsgBoxPause(hWnd, Msg, App.Title, vbExclamation, MsssageSeconde)
         
  CMDPAy(0).Enabled = True
  CMDPAy(1).Enabled = True
                  Exit Sub
                End If
            End If
            
If CboPOSBillType.ListIndex = 0 Then
If Index = 0 Then
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
            If Me.XPTxtBillID.text = "" Then
                Msg = "·« ÊÃœ ›Ê« Ì— ·Ì „ ÿ»«⁄ Â«"
               ' MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                '  Msg = "«·Œ’„ Ì ⁄œÌ «·„”„ÊÕ »Â ··„‰œÊ» Ê·« Ì„ﬂ‰  ﬂ„·… «·Õ›Ÿ"
                 messageResult = MsgBoxPause(hWnd, Msg, App.Title, vbExclamation, MsssageSeconde)
                                                   
                Exit Sub
            End If

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
            PrintReport , 1, LBLTable.Caption, 1
 End If
Else
 

 Cmd_Click (2)

End If
btnNew_Click 0
LBLPayVal.Caption = 0
FramePay.Visible = False
If SystemOptions.WorkWithItemsDetails = True Then
TxtItemCodeB.SetFocus
Else
TxtItemCodeB1.SetFocus
End If

CMDPAy(0).Enabled = True
CMDPAy(1).Enabled = True

End Sub

'Private Sub cmdShowPoints_Click()
'    If TxtPhone(0).text = "" Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ «·ÃÊ«·"
'        Else
'            MsgBox "Please Enter No.Mobile"
'        End If
'        TxtPhone(0).SetFocus
'        XPTxtDiscountVal.Visible = False
'        Exit Sub
'    End If
'        XPTxtDiscountVal.Visible = True
'        XPTxtDiscountVal.text = GetBalance
'
'
'End Sub

Private Sub CmdValue_Click(Index As Integer)
LBLPayVal.Caption = 0
'TxtPayedValue.text = CmdValue(Index).Caption
LBLPayVal.Caption = CmdValue(Index).Caption
        With Grid
          .TextMatrix(.Row, .ColIndex("Value")) = LBLPayVal.Caption
          End With
     ReLineGrid
     
End Sub

Private Sub DCboItemsCodex_Change()

  
End Sub

Private Sub DCboItemsCodex_Click(Area As Integer)

End Sub

Private Sub fg_Click()
    lblqty.Caption = ""
End Sub

Private Sub FgC_AfterEdit(ByVal Row As Long, ByVal Col As Long)
ChangCel Row, Col
With FgC
Select Case .ColKey(Col)
Case "Num2"
If .TextMatrix(Row, .ColIndex("Num2")) = .TextMatrix(Row, .ColIndex("Num")) And .TextMatrix(Row, .ColIndex("Num2")) <> "" Then
.cell(flexcpChecked, Row, .ColIndex("IsRetCopon")) = 1
Else
.cell(flexcpChecked, Row, .ColIndex("IsRetCopon")) = 0
.TextMatrix(Row, .ColIndex("Num2")) = ""
End If
Relim
End Select
End With
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
 ReLineGrid

End Sub

Private Sub Grid_Click()
 LBLPayVal.Caption = ""
End Sub

Private Sub Grid_DblClick()
Dim i As Long
For i = 1 To Grid.rows - 1
    Grid.TextMatrix(i, Grid.ColIndex("Value")) = ""
Next
Grid.TextMatrix(Grid.Row, Grid.ColIndex("Value")) = TxtNetValue
ReLineGrid
End Sub

Private Sub Grid_GotFocus()
LBLPayVal.Caption = 0
End Sub


Private Sub Image9_Click()
FrmCustomerDisplay.show

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

 
Private Sub Label19_Click()
FramePay.Visible = False
End Sub

Private Sub lbl_Click(Index As Integer)
FramePay.Visible = False
End Sub

 

 

Private Sub lblexit_Click(Index As Integer)
FramePay.Visible = False
End Sub

Private Sub LblTotalAllView_Change()
RelinVatGrid
End Sub

Private Sub optsale_Click(Index As Integer)

If optsale(0).value = True Then
ClearTex
   If CheckPeriodCopon(Me.XPDtbBill.value) = True Then
   TXtCopon.Visible = True
   lbl(94).Visible = True
   Else
   TXtCopon.Visible = False
   lbl(94).Visible = False
   End If
 Else
   TXtCopon.Visible = False
   lbl(94).Visible = False
End If
   btnExit(5).Visible = False
XPCboDiscountType.ListIndex = 0
XPTxtDiscountVal.text = 0
If Index = 0 Then
With Me.FgC
.ColHidden(.ColIndex("IsRetCopon")) = True
.ColHidden(.ColIndex("Num2")) = True
.ColHidden(.ColIndex("Num")) = False
End With
Else
With Me.FgC
.ColHidden(.ColIndex("IsRetCopon")) = False
.ColHidden(.ColIndex("Num2")) = False
.ColHidden(.ColIndex("Num")) = True
End With
End If
 lbl(57).Visible = False
 TBar.Visible = True

      NewGrid.GridDefaultValue 1
     FgC.Clear flexClearScrollable, flexClearEverything
     FgC.rows = 1
     VatGrid.Clear flexClearScrollable, flexClearEverything
     VatGrid.rows = 1
     LblTotalAllView.Caption = "0.00"
     LblDiscountsTotalView(0).Caption = "0.00"
     LblDiscountsTotalView(1).Caption = "0.00"
     LblTotal.Caption = "0.00"
     LBLPayVal.Caption = "0.00"
     LblDiscountsTotalView(6).Caption = "0.00"
Select Case Index

Case 0
TxtPhone(2).Visible = True
  Shape6.BorderColor = &H400000
TxtInvSerial.locked = True
TxtInvSerial.text = ""
TxtInvSerial.Visible = False
lbl(86).Visible = True

'
'    mTransaction_Type = 21
'              mSanad = 7
'              mNoteType = 170
              NewGrid.GridTrans = InvoiceTransaction
    TxtInvSerial.Visible = True
   ' lbl(86).Visible = False
   ' lbl(153).Visible = False
    TxtInvSerial.Visible = False
   
  TxtPhone(2).Visible = True
        CboRetrunType.ListIndex = -1
        CboRetrunType.Enabled = False
    
    '   Shape6.BorderColor = &HFF&
    TxtInvSerial.locked = False
    NewGrid.GridTrans = InvoiceTransaction
    
Case 1
'TxtPhone(2).Visible = False
'    mTransaction_Type = 9
'    mSanad = 14
'    mNoteType = 220
    TxtInvSerial.Visible = True
  '  lbl(86).Visible = True
    'lbl(153).Visible = True
    TxtInvSerial.Visible = True
    If SystemOptions.ReturnSallingOption = True Then
    
        CboRetrunType.ListIndex = 0
        CboRetrunType.Enabled = False
    End If
    '   Shape6.BorderColor = &HFF&
    TxtInvSerial.locked = False
    NewGrid.GridTrans = ReturnSalling

TxtPhone(2).Visible = False
btnExit(5).Visible = True
TxtInvSerial.Visible = True
lbl(86).Visible = True
    If SystemOptions.ReturnSallingOption = True Then

                CboRetrunType.ListIndex = 0
                CboRetrunType.Enabled = False
            End If
             Shape6.BorderColor = &HFF&
             TxtInvSerial.locked = False
End Select
End Sub

Private Sub SearchCashCustomer_Click(Index As Integer)
Select Case Index


Case 0
 
frmCashCustomerSearch.RetrunType = 2
frmCashCustomerSearch.show vbModal
Case 1
        Load FrmItemSearch2
        FrmItemSearch2.RetrunType = 1
        FrmItemSearch2.show vbModal
End Select
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
        'LBLTable.Caption = Label15.Caption
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

Private Sub LBLdOT_Click(Index As Integer)
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

  '  If val(lblqty.Caption) = 0 Then Exit Sub

    With Me.FG
        If .TextMatrix(.Row, .ColIndex("printed")) <> "" Then
        
       MsgBox "·« Ì„ﬂ‰  ⁄œÌ· ﬂ„Ì… Â–… «·’‰› ·«‰Â    ÿ»⁄", vbCritical
        Exit Sub
        End If
        .TextMatrix(.Row, .ColIndex("Count")) = val(lblqty.Caption)
        ' .TextMatrix(.Row, .ColIndex("Valu")) = Val(lblqty.Caption) * _
          val (.TextMatrix(.Row, .ColIndex("Price")))
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
Exit Sub
'    If Me.TxtModFlg.Text = "R" And LblStableID.Caption <> "-1" Then

'        Cmd_Click (1)

'    End If

    addrow val(Item.SubItems(2).Caption), Item.text, val(Item.SubItems(1).Caption), val(Item.SubItems(3).Caption)
    LblSowPrice(0).Caption = " «·”⁄— " & val(Item.SubItems(1).Caption)
    lblqty.Caption = ""
      lblShowQty2.Caption = "«·ﬂ„Ì… : 1 "

End Sub

Private Sub lvwMain_ItemClick(Item As cListItem)
Exit Sub

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
    MsgBox "Error: " & Err.Description & " [" & Err.Number & "]", vbInformation
    Exit Sub

End Sub

Function FillGroups()
Exit Function
    Dim colX As cColumn
    Dim itmX As cListItem
    Dim i As Long
    Dim j As Long
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
 
    sql = " SELECT * from  Groups where GroupID>1  and LastGroup=1"
 
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
      
        With .Listitems

            For i = 0 To rs.RecordCount - 1
                Set itmX = .Add(, rs("GroupID").value, rs("GroupName").value, 0, i)

                '      Set itmX = .Add(, "I" & i, "Test Item " & i, 0, 1)
                If (i Mod 2) = 0 Then
                    itmX.ToolTipText = "This is a test tool tip for item " & i
                End If

                With itmX.SubItems(1)
                    .Caption = DateSerial(year(Now), Rnd * Month(Now) + 1, Rnd * day(Now) + 1)
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

Function FillItems(GroupID As Integer)
Exit Function
    Dim colX As cColumn
    Dim itmX As cListItem
    Dim i As Long
    Dim j As Long
    Dim sql As String
    Dim rs As New ADODB.Recordset
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
        .ImageList(eLVLargeIcon) = GrouplImageList ' ilsIcons32
        .ImageList(eLVSmallIcon) = GrouplImageList ' ilsIcons16
        .ImageList(eLVTileImages) = GrouplImageList ' ilsIcons48
        .ImageList(eLVHeaderImages) = GrouplImageList ' ilsIcons16
            
        '  For i = 1 To 3
        '     .Columns(i).ItemData = i * 100
        '  Next i
      
        With .Listitems

            For i = 0 To rs.RecordCount - 1
                Set itmX = .Add(, rs("ItemID").value & i, rs("ItemName").value, 0, i)

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
Exit Function
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
      
        ' Set up image lists:
        .ImageList(eLVLargeIcon) = ilsIcons32
        .ImageList(eLVSmallIcon) = ilsIcons16
        .ImageList(eLVTileImages) = ilsIcons48
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

                If IsNull(rs("Status").value) Then
                    Set itmX = .Add(, rs("id").value, " " & rs("name").value, 0, 0)
                Else
                    Set itmX = .Add(, rs("id").value, " " & rs("name").value, 1, 1)
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
                    '   Dim sFnt As New StdFont
                    '      sFnt.name = "Tahoma"
                    '      sFnt.Size = 20
                    '      sFnt.Bold = True
        
                    '      itmX.Font = sFnt
                Else
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
        AddToLogFile CInt(user_id), 170, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , 0, TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 170, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , 0, TxtNoteSerial1
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
CheckAccounts = True
Exit Function
    Dim vchrcode As String
    Dim StrTempAccountCode As String
    Dim usedaccount As Integer
 Dim Account_Code_dynamic As String
    If BillBasedOn(0).value = True Then
        vchrcode = Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19, , , , , , val(DCboUserName.BoundText))

        If vchrcode = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ’—› ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": GoTo ErrTrap
        ElseIf vchrcode = "" Then
            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": GoTo ErrTrap
                       
        End If
                       
    End If
                       
                       
                       
                       
         If optsale(1).value = True Then 'return sales
             Account_Code_dynamic = get_account_code_branch(3, my_branch)
        
    If Account_Code_dynamic = "NO branch" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Else
            MsgBox "Branch Not Created", vbCritical
        End If

        GoTo ErrTrap
    ElseIf Account_Code_dynamic = "NO account" Then

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»   „—œÊœ«  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
        Else
            MsgBox "Sales Account Not Defined in this Branch", vbCritical
        End If

        GoTo ErrTrap
         
    End If
        vchrcode = Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20, , , , , , val(DCboUserName.BoundText))

        If vchrcode = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  «” ·«„ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": GoTo ErrTrap
        ElseIf vchrcode = "" Then
            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": GoTo ErrTrap
                       
        End If
                       
    End If
    
    
   
 




    If val(DCDocTypes.BoundText) > 0 Then
        getDocAccounts val(DCDocTypes.BoundText), StrTempAccountCode, , , , , usedaccount

        If StrTempAccountCode = "" And usedaccount = 1 Then
            MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«»  «·„œÌ‰ «·Œ«’ »«·›« Ê—…  ", vbCritical
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

    If CboPayMentType.ListIndex = 0 Then
        DCPaymentNet.BoundText = 1
    Else
        DCPaymentNet.text = ""
    End If

    CboPayMentType_Change
 
End Sub

Private Sub ChkInstall_Click()

    If ChkInstall.value = vbChecked Then
        Me.CmdINSTALLMENT.Enabled = True
        XPTxtValue(1).text = LblTotal.Caption
    Else
        Me.CmdINSTALLMENT.Enabled = False

        With Me.FgInstallments
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            LblPrecenType.Caption = ""
            LblPrecenValue.Caption = ""
            LblInstallTotal.Caption = ""
            LblInstallCount.Caption = ""
            LblFirstInstallDate.Caption = ""
            LblInstallmentType.Caption = ""
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

        For i = 1 To .rows - 1
     
            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
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

        For i = 1 To .rows - 1
     
            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then

                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID"))) ' & "nots=" & "" & "nots2=" & ""
               
            End If
       
            Cn.Execute sql
 
        Next
       
    End With
       
End Function
Sub printtomanyprinter()
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
    StrSQL = " SELECT     dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.printed, dbo.TblItems.ItemName, "
StrSQL = StrSQL & " dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice AS value, dbo.Transaction_Details.Transaction_ID"
StrSQL = StrSQL & " FROM         dbo.Transaction_Details INNER JOIN"
StrSQL = StrSQL & " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
StrSQL = StrSQL & " WHERE     (dbo.Transaction_Details.printed IS NULL) AND (dbo.Transaction_Details.Transaction_ID = " & val(XPTxtBillID.text) & ")"

 
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
        Printer.Print Tab(intLINE_START_POS); _
                      IIf(IsNull(rs("VALUE").value), "", rs("VALUE").value); _
                      Tab(7 + intLINE_START_POS); _
                      IIf(IsNull(rs("showPrice").value), "", rs("showPrice").value); _
                      Tab(14 + intLINE_START_POS); _
                      IIf(IsNull(rs("ShowQty").value), "", rs("ShowQty").value); _
                      Tab(21 + intLINE_START_POS); _
                      IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value);
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
    Printer.Print
    Printer.Print
    Printer.Print
    
    ' Print the main headings
    Printer.Print Tab(intLINE_START_POS); _
                  "Print Date: "; _
                  Format$(Date, "mm/dd/yy"); _
                  Tab(intLINE_START_POS + 31); _
                  ""; _
                  Tab(intLINE_START_POS + 73); _
                  ""; _
                  'Format$(intPageCtr, "@@@")
    Printer.Print Tab(intLINE_START_POS); _
                  "Print Time: "; _
                  Format$(Time, "hh:nn:ss"); _
                  Tab(intLINE_START_POS + 33); _
                  LBLTable.Caption
    Printer.Print
    ' Print the column headings
    Printer.Print Tab(intLINE_START_POS); _
                  "≈Ã„«·Ì"; _
                  Tab(7 + intLINE_START_POS); _
                  "«·”⁄—"; _
                  Tab(14 + intLINE_START_POS); _
                  "«·ﬂ„Ì…"; _
                  Tab(21 + intLINE_START_POS); _
                  "«·’‰›";
       
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

Private Sub Cmd_Click(Index As Integer)
    Dim AskOption As Boolean
    Dim intDef    As Integer
    Dim Msg       As String
    Dim i         As Integer
    Dim StrSQL    As String
    Dim RsTest    As ADODB.Recordset
    Dim RsOptions As ADODB.Recordset
    BolPrint = True
    '    On Error GoTo ErrTrap
    Timer1.Enabled = False

    '    If Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, 21) = "" And val(my_branch) <> 0 Then
    '        TxtNoteSerial1.Locked = False
    '    Else
    '        TxtNoteSerial1.Locked = True
    '
    '    End If


If (Index = 1 Or Index = 4) And zatcaStatus = 1 Then
    ' If SystemOptions.IsBluee = True Then
           Msg = "·« Ì„ﬂ‰  ⁄œÌ· «Ê Õ–› «Ì „” ‰œ Ì„ﬂ‰ﬂ ⁄„· „” ‰œ ⁄ﬂ”Ì ›ﬁÿ"
                        Msg = Msg & CHR(13) & ""
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
   '     End If

End If
    
    
    lblqty.Caption = ""
    lblShowQty2.Caption = ""
    Select Case Index
        Case 11
            printtomanyprinter
        Case 9
            PrintReport , 1, LBLTable.Caption, 1
        Case 0

            '       If DoPremis(Do_New, Me.Name, True) = False Then
            '           Exit Sub
            '       End If
            clear_all Me
            LblTotalAllView.Caption = 0

            '    FrmCustomerDisplay.LblInformation.Caption = ""
            ' FrmCustomerDisplay.LblInformation2.Caption = ""
            ' FrmCustomerDisplay.Image1.Visible = False

            '   With lvwItems
            '       lvwItems.Listitems.Clear
            '   End With
            BillBasedOn(1).Enabled = True
            '           DCboItemsCode.SetFocus
            CboPOSBillType.ListIndex = 0
            LblStableID.Caption = -1
            LBLTable.Caption = ""
            
            ClearNotes
            TxtModFlg.text = "N"
            DefaultInvoicetype.ListIndex = SystemOptions.DefaultInvoicetype
            '            XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            SetDefaults
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            VatGrid.Clear flexClearScrollable, flexClearEverything
            VatGrid.rows = 1
            intDef = val(GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2))
            If intDef = 0 Then intDef = 2
            DBCboClientName.BoundText = 2 ' intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
        
            '            Set RsOptions = New ADODB.Recordset
            '            RsOptions.Open "tbloptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
            '
            '            If Not (RsOptions.BOF Or RsOptions.EOF) Then
            '                Me.DcboBox.BoundText = IIf(IsNull(RsOptions("SalesBoxID").value), "", RsOptions("SalesBoxID").value)
            '            End If

            XPTab301.CurrTab = 0
            '------------------
            '        Me.XPDtbBill.SetFocus
            '   customer_screen.Show
            '--------------------
        
            DcCurrency.BoundText = 1
        
            Me.dcBranch.BoundText = Current_branch
            Dim dstore       As Integer
            Dim dBox         As Integer
            Dim usertype     As Integer
            Dim EmpID        As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
     
            If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
                DcboBox.Enabled = False
                DCboStoreName.Enabled = True
                DcboEmp.Enabled = False
          
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
                Me.DCboStoreName.BoundText = ""
                Me.DcboBox.BoundText = ""
                Me.DcboEmp.BoundText = ""
                Me.dcBranch.BoundText = userbranchid
                Me.DCboStoreName.BoundText = dstore
                Me.DcboBox.BoundText = dBox
                Me.DcboEmp.BoundText = EmpID

            End If
            DefaultInvoicetype.ListIndex = SystemOptions.DefaultInvoicetype
            zatcaStatus = 0
            BillBasedOn(0).value = True
 
            If Current_branch = 0 Then
                'branch_id = my_branch
                Me.dcBranch.BoundText = Current_branch
            End If
     
            LblStableID.Caption = -1
            CboPOSBillType.ListIndex = 1
            optsale(0).value = True
 
            If CheckSpecialOffer(Me.XPDtbBill.value, val(Me.dcBranch.BoundText), Sales, GetFree, discount, FromPrice) = True Then
                lbl(57).Visible = True
 
            Else
                lbl(57).Visible = False
                inddx = 1
            End If
 
            If CheckPeriodCopon(Me.XPDtbBill.value) = True Then
                TXtCopon.Visible = True
                lbl(94).Visible = True
            Else
                TXtCopon.Visible = False
                lbl(94).Visible = False
            End If
   
        Case 1

            '  If DoPremis(Do_Edit, Me.Name, True) = False Then
            '      Exit Sub
            '  End If

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

                    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbNo Then
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

                    MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If

            '⁄„·Ì«  «·’Ì«‰… «·„— »ÿ… »«·›« Ê—…
            StrSQL = "select * From MaintenanceJuncTransaction where Transaction_ID=" & Trim(XPTxtBillID.text)
            Set RsTest = New ADODB.Recordset
            RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsTest.EOF Or RsTest.BOF) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·ﬁœ  „ ≈Ã—«¡ »⁄÷ ⁄„·Ì«  «·’Ì«‰… ⁄·Ï Â–Â «·›« Ê—… Ê·« Ì„ﬂ‰  ⁄œÌ·Â«"
                    Msg = Msg + "≈–« ﬂ‰   —€» ›Ì  ⁄œÌ· »Ì«‰«  Â–Â «·›« Ê—…" & CHR(13)
                    Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «·’Ì«‰… «·Œ«’… »Â«"
                Else
                    Msg = "this Bill Linked with Maintenance Operation You must Delete This Operation First"
            
                End If

                MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            '         Me.Retrive Val(Me.XPTxtBillID.text)
             
            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id
            CuurentLogdata

            '    txtorder_no_Change
            'End If
        Case 2
            With Me.FgC
                For i = 1 To .rows - 1
                    If val(.TextMatrix(i, .ColIndex("Vlue"))) <> 0 Then
                        If .TextMatrix(i, .ColIndex("Num")) = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                Msg = "Ì—ÃÏ «œŒ«· —ﬁ„ «·ﬁ”Ì„… ›Ì «·”ÿ— —ﬁ„"
                                Msg = Msg & " " & i
                            Else
                                Msg = "Please eneter .No of coupon in row No."
                                Msg = Msg & " " & i
                            End If
                            MsgBox Msg
                            Screen.MousePointer = vbDefault '
                            Exit Sub
                        End If
                    End If
                Next i
            End With
            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·« "
                End If
              
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                '          dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault '
                Exit Sub
            End If
 
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

            If CboPayMentType.ListIndex = 1 And XPChkPayType(0).value = Unchecked And XPChkPayType(2).value = Unchecked Then
                XPTxtValue(1).text = LblTotal.Caption
            End If
 
            If optsale(1).value = True Then 'return sales
                If CboRetrunType.ListIndex = 0 Then
                    Dim bill_id    As Double
                    Dim voucher_id As Double
                    bill_id = val(TxtInvID.text)
                
                    voucher_id = check_bill_voucher(bill_id, 19)  '·«ÌÃ«œ —ﬁ„ «–‰ «·’—› „‰ ﬁ«⁄œ… «·»Ì«‰« 

                    If voucher_id = 0 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            '         MsgBox "·« ÌÊÃœ ”‰œ ’—› „Œ“‰Ì ·Â–… «·›« Ê—… Õ Ï Ì„ﬂ‰ Õ”«»  ﬂ·›… «·„»Ì⁄« ", vbCritical
                        Else
                            '         MsgBox " There is no issue voucher to this bill ", vbCritical
                        End If

                        '     GoTo ErrTrap
                    End If
                   
                    If checkretutn = False Then

                        Exit Sub
                    End If

                End If

            End If
    
            my_branch = Me.dcBranch.BoundText
            If optsale(1).value = True Then
                If CheckFilegrid() = True Then
                    SaveData
                Else
                    Exit Sub
                End If
            Else
                SaveData
            End If
     
            '   MsgBox ""
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            If SystemOptions.usertype = UserNormal Then
                Msg = "·Ì” ·ﬂ Õﬁ Õ–› ›Ï «·›Ê« Ì—"
                MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
   
            Del_TransAction

        Case 5

            '      If DoPremis(Do_Search, Me.Name, True) = False Then
            '          Exit Sub
            '      End If

            If m_FrmSearch Is Nothing Then
               
                Set m_FrmSearch = New FrmBuySearch
                m_FrmSearch.DealingForm = InvoiceTransaction
                m_FrmSearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄"
                Set m_FrmSearch.RetrunFrm = Me
                m_FrmSearch.show vbModal

            Else
                Msg = "Â‰«ﬂ ‘«‘… »ÕÀ Œ«’… »‘«‘… ›« Ê—… «·»Ì⁄ «·Õ«·Ì…"
                Msg = Msg & CHR(13) & "Ÿ«Â—… «„«„ﬂ ›⁄·«...·«Ì„ﬂ‰ ⁄—÷ «ﬂÀ— „‰ ‘«‘… »ÕÀ ·ﬂ· ‘«‘… ›« Ê—…"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                m_FrmSearch.ZOrder 0
                'm_FrmSearch.SetFocus
            End If

        Case 7
            If Not LoadExcelFlage Then
            FramePay.Visible = True
            CMDPAy(0).Enabled = True
            CMDPAy(1).Enabled = True
            End If

        Case 6
            Unload Me

        Case 10
            ShowGL_cc TxtNoteSerial.text, , 200, val(Me.TXTNoteID.text)
            'ShowGL_cc TxtNoteSerial.text, , 200
        Case 8
            End
    End Select

    Exit Sub
ErrTrap:
End Sub
Function CheckFilegrid() As Boolean
Dim i As Integer
Dim j As Integer
Dim Item_ID As Double
Dim SumQty As Double
Dim ClassId As Integer
Dim itemsize As Integer
Dim ColorID As Integer
Dim UnitID As Integer
Dim total As Double
Dim Msg As String
If SystemOptions.ReturnSallingOption = False And Me.TxtInvSerial.text <> "" Then
 
  CheckFilegrid = True
  Exit Function
End If
With FG

CheckFilegrid = True
For j = .FixedRows To .rows - 1

SumQty = 0
Item_ID = val(.TextMatrix(j, .ColIndex("Code")))
ClassId = val(.TextMatrix(j, .ColIndex("ClassId")))
itemsize = val(.TextMatrix(j, .ColIndex("ItemSize")))
ColorID = val(.TextMatrix(j, .ColIndex("ColorID")))
UnitID = IIf(.cell(flexcpData, j, .ColIndex("UnitID")) = "", 0, (.cell(flexcpData, j, .ColIndex("UnitID"))))
For i = .FixedRows To .rows - 1

If Item_ID = val(.TextMatrix(i, .ColIndex("Code"))) And UnitID = IIf(.cell(flexcpData, i, .ColIndex("UnitID")) = "", 0, (.cell(flexcpData, i, .ColIndex("UnitID")))) And ClassId = val(.TextMatrix(i, .ColIndex("ClassId"))) And itemsize = val(.TextMatrix(i, .ColIndex("ItemSize"))) And ColorID = val(.TextMatrix(i, .ColIndex("ColorID"))) Then
SumQty = SumQty + val(.TextMatrix(i, .ColIndex("Count")))
End If
Next i
total = RetriveQtyItem(TxtInvSerial.text, Item_ID, ColorID, ClassId, itemsize, UnitID)
If total < SumQty Then
If total > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
Msg = .cell(flexcpTextDisplay, j, .ColIndex("Name")) & "  ·«Ì„ﬂ‰ «— Ã«⁄ ﬂ„Ì… «ﬂ»— „‰ «·ﬂ„Ì… «·«’·Ì… ··’‰› "
Msg = Msg & CHR(13)
Msg = Msg & (total) & " " & "«·ﬂ„Ì… «·„ »ﬁÌ…"
Else
Msg = .cell(flexcpTextDisplay, j, .ColIndex("Name")) & " can't return this qty for this item "
Msg = Msg & CHR(13)
Msg = Msg & " " & "Avilable Qty to Return is   " & (total)

End If
Else
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = .cell(flexcpTextDisplay, j, .ColIndex("Name")) ' & "  ·«ÌÊÃœ  ﬂ„Ì… „‰  «·’‰›   "
        Msg = Msg & CHR(13)
        Msg = Msg & " „ «—Ã«⁄ ﬂ«„· «·ﬂ„Ì… „‰ Â–… «·›« Ê—… «Ê «·’‰› €Ì— „ÊÃÊœ «’·« ›Ì «·›« Ê—…"
        Else
        
                Msg = .cell(flexcpTextDisplay, j, .ColIndex("Name")) ' & "  ·«ÌÊÃœ  ﬂ„Ì… „‰  «·’‰›   "
        Msg = Msg & CHR(13)
        Msg = Msg & "ˆAll Qty From this item Already returnd or not exist in the invoice"
        End If
End If
MsgBox Msg
GoTo l
Else
CheckFilegrid = True
End If
Next j
Exit Function
End With
l: CheckFilegrid = False


End Function
Function RetriveQtyItem(Optional NoteSerial1 As String, Optional Item_ID As Double, Optional ColorID As Integer, Optional ClassId As Integer, Optional itemsize As Integer, Optional UnitID As Integer) As Double
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Integer
 '**************************************************************************
  StrSQL = "SELECT     dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.ClassId, dbo.Transaction_Details.Item_ID, "
  StrSQL = StrSQL & "                     dbo.Transaction_Details.UnitId, SUM(dbo.Transaction_Details.ShowQty * isnull( dbo.Transaction_Details.FLgReturn,1)) AS smQty"
  StrSQL = StrSQL & "  FROM         dbo.Transaction_Details RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
  StrSQL = StrSQL & "   WHERE     (((dbo.Transactions.NoteSerial1 = N'" & NoteSerial1 & "') AND (dbo.Transactions.Transaction_Type = 21)) OR"
  StrSQL = StrSQL & "                  (   (dbo.Transactions.ReturnSerial = N'" & NoteSerial1 & "') AND (dbo.Transactions.Transaction_Type = 9)))"
  StrSQL = StrSQL & "  AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ") AND (dbo.Transaction_Details.UnitId = " & UnitID & ") and(dbo.Transaction_Details.ColorID = " & ColorID & ") and(dbo.Transaction_Details.ClassId = " & ClassId & ")and(dbo.Transaction_Details.ItemSize = " & itemsize & ")"
  StrSQL = StrSQL & "  GROUP BY dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.ClassId, dbo.Transaction_Details.Item_ID,"
  StrSQL = StrSQL & "                     dbo.Transaction_Details.unitid"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsDetails.RecordCount > 0 Then
RetriveQtyItem = IIf(IsNull(RsDetails("smQty").value), 0, RsDetails("smQty").value)
Else
RetriveQtyItem = 0
End If
End Function
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
 
   If Not (IsNull(rs("CompanyLogo").value)) Then
        'LoadPictureFromDB ImgPic, rs, "CompanyLogo"
        LoadPictureFromDB Image9, rs, "CompanyLogo"
        
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
        row_count = FG.rows
    
        If FG.TextMatrix(row_count - 1, FG.ColIndex("Code")) = "" Then
            row_count = row_count - 1
        End If
     
        FG.rows = RsDetails.RecordCount + row_count

        For Num = row_count To FG.rows - 1 'RsDetails.RecordCount
    
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
        
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
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
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

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
            .CboPrecenType.ListIndex = val(Me.LblPrecenType.Tag)
            .Txt(3).text = val(LblPrecenValue.Caption)
            .Txt(5).text = val(LblInstallCount.Caption)
            .OptInt(1).value = True
            .Txt(7).text = 1
            .Txt(5).text = 12

            If IsDate(Me.LblFirstInstallDate.Caption) Then
                .Dtp_First.value = Me.LblFirstInstallDate.Caption
            End If

            '        .Txt(7).text = Val(LblInstallSeprator.Caption)
            If val(LblInstallmentType.Tag) = 0 Then
                '        .OptInt(0).value = True
            ElseIf val(LblInstallmentType.Tag) = 1 Then
                '        .OptInt(1).value = True
            ElseIf val(LblInstallmentType.Tag) = 2 Then
                '        .OptInt(2).value = True
            End If

            With .FG
                .rows = Me.FgInstallments.rows

                For i = 1 To Me.FgInstallments.rows - 1
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
        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
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

        For RowNum = 1 To FG.rows - 1

            If .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Code")) <> "" Then
                .FG.rows = .FG.rows + 1
            End If

            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Name")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Name")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Code")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Code")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("ItemCase")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("HaveSerial")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Count")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Count")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Count")))
            ' .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Price")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Price")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(.FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod)
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("DiscountType")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")))
            Dim StrSQL As String
            Dim RsUnit As New ADODB.Recordset
            StrSQL = "SELECT TOP 100 PERCENT dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial,dbo.Transactions.Transaction_Type FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites INNER JOIN dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID AND dbo.Transaction_Details.Item_ID = dbo.TblItemsUnits.ItemID WHERE (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "') AND (dbo.Transactions.Transaction_Type = 21) AND (dbo.TblItemsUnits.ItemID = " & FG.TextMatrix(RowNum, FG.ColIndex("Code")) & ") ORDER BY dbo.TblItemsUnits.SecOrder"
            Set RsUnit = New ADODB.Recordset
            RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
            .FG.cell(flexcpData, RowNum, .FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").value))
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
    Dim LngCurItemID As Double
    Dim LngUnitID As Long
    Dim UnitFactor As Double

    With FG

        For i = 1 To FG.rows - 1

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("ItemType"))) <> 1 Then
                LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
                LngUnitID = val(FG.cell(flexcpData, i, FG.ColIndex("UnitID")))
            
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

                For i = 1 To FG.rows - 1

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

                    For i = 1 To FG.rows - 1

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
    Dim LngCurItemID As Double
    Dim LngUnitID As Long
    Dim UnitFactor As Double

    '›Ì Õ«·… «·«‰ «Ã «·‰„ÿÌ
    If SystemOptions.TypicalProduction = True Then
        GoTo ll
    End If
GoTo ll

    With FG

        For i = 1 To FG.rows - 1

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("ItemType"))) <> 1 Then
                                      
                LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
                LngUnitID = val(FG.cell(flexcpData, i, FG.ColIndex("UnitID")))
            
                GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
                    
                'TOTAL_COST = TOTAL_COST + (FG.TextMatrix(i, FG.ColIndex("Count")) * UnitFactor * ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod))
                    
                If ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , LngUnitID) = 0 Then
'                    If SystemOptions.UserInterface = ArabicInterface Then
'                        MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ  ﬂ·›Â «·»Ì⁄ ·Â Ê·„ Ì „  ÕœÌœ À„‰ «·‘—«¡ Ê·Ì” ·Â ﬁÌ„Â —’Ìœ «›  «ÕÌ… ·–·ﬂ ·« Ì„ﬂ‰ «‰‘«¡ ”‰œ «·’—› "
'                    Else
'                        MsgBox "Item in line no " & i & "Have No Qty "
'                    End If
 
                    With Me.GRID1
                        .rows = .FixedRows
                        .ExtendLastCol = True
                        .RowHeightMin = 300
                        .Editable = flexEDKbdMouse
                        .ExplorerBar = flexExSortShowAndMove

                        '    .AutoSize 0, .Cols - 1, False
                    End With

                    Text1.text = ""
                    'Cn.Execute "UPDATE Transactions SET NOTS='" & "" & "' WHERE Transaction_ID=" & Val(Me.XPTxtBillID.text)
                    Text1_Change

    '                Exit Sub
                End If
            End If

        Next i

    End With

ll:
    Dim groupAccount  As String

    If 4 = 3 Then
   
        With FG

            For i = 1 To FG.rows - 1

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
 
        'MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=19"))
        'mytext = TxtTransSerial.text

        '         rs!nots = mytext
        '         rs.update

        Dim Transaction_ID As Long

        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
        Dim TxtNoteSerialV As String
        Dim TxtNoteSerial1V As String
            
        my_branch = val(Me.dcBranch.BoundText)
'
'        If TxtNoteSerialV = "" Then
'            If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
'                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
'            Else
'
'                If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
'                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
'
'                Else
'                    TxtNoteSerialV = Notes_coding(val(my_branch), XPDtbBill.value)
'                End If
'            End If
'        End If
'
            If TxtNoteSerialV = "" Then
                TxtNoteSerialV = Notes_coding(val(my_branch), XPDtbBill.value)
           End If
           
'        If TxtNoteSerial1V = "" Then
'            If Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19, , , , , , val(DCboUserName.BoundText)) = "error" Then
'                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ’—› ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
'            Else
'
'                If Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19, , , , , , val(DCboUserName.BoundText)) = "" Then
'                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
'                Else
'                    TxtNoteSerial1V = Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19, , , , , , val(DCboUserName.BoundText))
'                End If
'            End If
'        End If
            If TxtNoteSerial1V = "" Then
                TxtNoteSerial1V = Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19, , , , , , val(DCboUserName.BoundText))
            End If
        If SystemOptions.TypicalProduction = True Then
            TxtNoteSerialV = ""
 
        End If
 
        If Trim(CurrentVoucherNo) <> "" And DateChanged <> True Then
            TxtNoteSerialV = CurrentVoucherNo '—ﬁ„ «·ﬁÌœ
            TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
        End If

        Dim sql As String

        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        MYTEXT = Transaction_ID
        Text1.text = Transaction_ID
        
        sql = "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,nots2,NoteSerial,NoteSerial1,NoteId,BranchId,Closed)SELECT " & Transaction_ID & "," & MYTEXT & ",Transaction_Date,Transaction_Type = 19,CusID,StoreID,UserID,Emp_ID,nots=" & val(XPTxtBillID.text) & ",nots2=" & TxtNoteSerial1.text & " ,NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId,1From Transactions Where  Transaction_ID =" & val(XPTxtBillID.text) & " And Transaction_Type = 21"
        Cn.Execute sql
        '
        Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID ,OldQty,OldCost,NewQty,NewCost,ProductionDate,ExpiryDate,LotNO)SELECT  costprice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, costprice/ QtyBySmalltUnit ,ColorID,ItemSize, UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID ,OldQty,OldCost,NewQty,NewCost,ProductionDate,ExpiryDate,LotNO From dbo.Transaction_Details Where SavedItemType=0 and   Transaction_ID = " & XPTxtBillID.text
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
'        RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
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
        sql = " update Transactions Set NoteID = " & general_noteid & " where Transaction_ID = " & Transaction_ID
        
        Cn.Execute sql
    End If
 
    '
 
ErrTrap:

End Sub

Private Sub Command2_Click()

    If Me.TxtModFlg = "R" Then
        Cmd_Click (1)
        Cmd_Click (2)
        CreateIssueVoucher
    End If

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
        FrmItemSearch.RetrunType = 7715
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
'    If val(DCPaymentNet.BoundText) <> 1 Then
  '      DcboBox.text = ""
        
  '  End If

End Sub

Function FillOrderGrid()
    ' ⁄»∆… «Ê«„— «·‘—«¡ Ê «·»Ì⁄

    With Me.GRID2
        .rows = .FixedRows
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
        .rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .rows - 1
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
        .rows = .FixedRows
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
        .rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .rows - 1
             
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
'RelinVatGrid

    If Me.TxtModFlg <> "E" Then Exit Sub
    
    If val(Me.TxtNoteSerial.text) = 0 Or val(Me.TxtNoteSerial1.text) = 0 Then GoTo ll

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , val(Me.TxtNoteSerial), val(Me.TxtNoteSerial1), 170
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), Me.TxtNoteSerial, Me.TxtNoteSerial1, 170

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
                    .rows = 1
       
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
        .rows = 1
       
    End With
 
    fillOrders

End Sub

Function fillVchr()
    Dim i As Integer
        
    With GRID1

        For i = 1 To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                Retrive_orders_data (val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
            
            End If

        Next i

    End With

End Function

Function fillOrders()
    Dim i As Integer

    With GRID2

        For i = 1 To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
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

 

End Sub

Private Sub LblDiscountsTotal_Change()
    LblDiscountsTotalView(0).Caption = Format(val(LblDiscountsTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub LblInstallCount_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    LblInstallCount.ToolTipText = WriteNo(LblInstallCount.Caption, 0, True)
End Sub

Private Sub LblInstallTotal_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    LblInstallTotal.ToolTipText = WriteNo(LblInstallTotal.Caption, 0, True)
End Sub

Private Sub LblInvProfit_Change()
    CalculateInvPrecent
End Sub

Private Sub LblPrecenValue_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     X As Single, _
                                     Y As Single)
    LblPrecenValue.ToolTipText = WriteNo(LblPrecenValue.Caption, 0, True)
End Sub

Private Sub LblTotal_Change()
    LblTotalView.Caption = Format(val(LblTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
If SystemOptions.UserInterface = ArabicInterface Then
LblSowPrice(1).Caption = "«·«Ã„«·Ì : " & LblTotalView.Caption
 Else
 LblSowPrice(1).Caption = "Totals : " & LblTotalView.Caption
 End If
    
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
        TxtNetValue.text = val(LblTotal.Caption)
        'TxtPayedValue.text = TxtNetValue.text
 
        With Me.FgInstallments
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            LblPrecenType.Caption = ""
            LblPrecenValue.Caption = ""
            LblInstallTotal.Caption = ""
            LblInstallCount.Caption = ""
            LblFirstInstallDate.Caption = ""
            LblInstallmentType.Caption = ""
        End With

    End If
  
End Sub

Function showComm()

    If val(LblInstallTotal.Caption) > 0 Then
        lblInstComm.Caption = val(LblInstallTotal.Caption) - val(LblTotal.Caption)
  
    Else
        lblInstComm.Caption = 0
        '  Me.LblFinal = 0
    End If
LblTotal.Caption = val(LblTotalAll.Caption) - val(LblDiscountsTotal.Caption) + IIf(SystemOptions.PriceWithVAT = True, 0, val(TxtValueAdded.text)) '- SmVal
 LblTotal.Caption = Round(LblTotal.Caption, 2)
    Me.LblFinal = val(lblInstComm.Caption) + val(LblTotal.Caption) + IIf(SystemOptions.PriceWithVAT = True, 0, val(TxtValueAdded.text))
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
    If SystemOptions.PriceWithVAT = True Then
        LblTotalAllView.Caption = Format(val(LblTotalAll.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
    Else
        LblTotalAllView.Caption = Format(val(LblTotalAll.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
    End If
End Sub

Private Sub lvwTables_ItemClick(Item As vbalListViewLib6.cListItem)
Exit Sub
    On Error GoTo ErrorHandler
    Dim sInfo As String

    If Not lvwTables.SelectedItem Is Nothing Then

        With lvwTables.SelectedItem
            CboPOSBillType.ListIndex = 0
            '    sInfo = "Key = " & Item.key & Item.text
            LBLTable.Caption = Item.text
            LblStableID.Caption = Item.key
 
            If Me.TxtModFlg.text = "N" And .SubItems(1).Caption = "1" Then
                MsgBox "«·„Ã·” «Ê «·ÿ«Ê·… «·„Õœœ… „‘€Ê·… Õ«·Ì« ·«»œ „‰ ”œ«œ ﬁÌ„… «·›«‰Ê—… «Ê·«", vbCritical
                LBLTable.Caption = ""
                LblStableID.Caption = -1
                Exit Sub
            End If
 
            If .SubItems(1).Caption = "1" Then
                Retrive (getTransactionIdBytable(Item.key))

            Else

                If Me.TxtModFlg.text <> "N" Then
                    MsgBox " ·« ÌÊÃœ ›Ê« Ì— ·Â–« «·„Ã·” «÷⁄ÿ ÃœÌœ «Ê·« ·«Œ Ì«— „Ã·”/ÿ«Ê·… ›«—€…", vbCritical
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

 



Private Sub Timer3_Timer()

End Sub

Private Sub Timer4_Timer()
lbl(81).Caption = Time
End Sub


Private Sub Timer5_Timer()
On Error Resume Next
If imageCounter = 0 Then imageCounter = 1
If imageCounter = 3 Then imageCounter = 1
On Error Resume Next
 Image10.Picture = LoadPicture(App.path & "\Images\pos2\" & imageCounter & ".jpg")
  imageCounter = imageCounter + 1
 
End Sub




Private Sub Txtcard_Click(Index As Integer)
Txtcard(Index) = ""
Select Case Index
Case 0

 CashCustomerName.text = ""
   TxtPhone(0).text = ""
        XPCboDiscountType.ListIndex = 0
         XPTxtDiscountVal.text = 0
         

Case 1

        DcboEmp.BoundText = 0
                
End Select
End Sub

 
Private Sub Txtcard_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
Dim Name As String
Dim phone As String
Dim discount As Double
If KeyAscii = vbKeyReturn Then
   
   Select Case Index
   Case 0
                     GetCashCustomernamebycard Txtcard(0).text, Name, phone, discount
                     
                     CashCustomerName.text = Name
                     TxtPhone(0).text = phone
                     If discount > 0 Then
                     XPCboDiscountType.ListIndex = 2
                     XPTxtDiscountVal.text = discount
                     Else
                            XPCboDiscountType.ListIndex = 0
                     XPTxtDiscountVal.text = 0
                     End If
                     
          
    
    Case 1
       Dim EmpID As Integer

   
                    If Txtcard(1).text <> "" Then
                        GetEmployeeIDFromCode mId(Txtcard(1).text, 2, Len(Txtcard(1).text) - 2), EmpID
                    
                        DcboEmp.BoundText = EmpID
                      End If
                      
 



    
    End Select
       End If

End Sub

Private Sub TXtCopon_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
If TXtCopon.text <> "" Then
If CheckParcode3(TXtCopon.text, Me.XPDtbBill.value) = True Then
                If ChekGrid2(TXtCopon.text) = True Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "Â–« «·ﬂÊ»Ê‰ „” Œœ„ ·«Ì„ﬂ‰ «· ﬂ—«—"
                                Else
                                     MsgBox "This is coupon were used previously"
                                End If
                    Exit Sub
                                
                End If
FillGridC
                
Else
                If SystemOptions.UserInterface = ArabicInterface Then
                                      MsgBox "Â–« «·ﬂÊ»Ê‰ „” Œœ„ „‰ ﬁ»· «Ê €Ì— „ÊÃÊœ Œ·«· Â–Â «·› —…"
                 Else
                                     MsgBox "This is coupon were used previously.or not in period"
                  End If
                                
                Exit Sub
            
End If
End If
End If
End Sub
Sub FillGridC()
Dim i As Integer
Dim k As Integer
Dim ID As Double
Dim Valu As Double
If CheckParcode4(Me.TXtCopon.text, XPDtbBill.value, ID, Valu) = True Then
With Me.FgC
k = .rows
.rows = .rows + 1
For i = k To .rows - 1
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("Vlue")) = Valu
.TextMatrix(i, .ColIndex("Num")) = TXtCopon.text
.TextMatrix(i, .ColIndex("PerioDID")) = ID
.TextMatrix(i, .ColIndex("Selcd2")) = 1
 inddx = 0
Next i
End With
End If
RelimSali
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
    FG.rows = 2
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
        FG.rows = RsDetails.RecordCount + 1

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
        
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If FG.rows > 10 Then
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

Sub ClearTex()
TxtInvSerial.text = ""
 TxtInvID.text = 0
         FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
End Sub
Public Sub RetriveReSalin()
XPCboDiscountType.ListIndex = 0
XPTxtDiscountVal.text = 0
CashCustomerName.text = ""
TxtPhone(0).text = ""
lbl(57).Visible = False
    Dim Msg As String
    Dim FrmNewSales As frmsalebill2
   FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    If val(Me.TxtInvSerial.text) = 0 Then
        FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    End If
    

    Dim Transaction_ID As Long
    Dim Transaction_Date  As Date
 SpecialOffer = 0
    GetTransIDFromNoteSerial1 Me.TxtInvSerial.text, Transaction_ID, Transaction_Date, 21, SpecialOffer
    Me.TxtInvID.text = Transaction_ID
     Me.txtInvDate.text = Transaction_Date

 If SpecialOffer = 1 Then
 lbl(57).Visible = True
 Image1.Visible = False
TBar.Visible = False
 Else
 lbl(57).Visible = False
 Image1.Visible = True
 TBar.Visible = True
 End If
 
    
    If val(Me.TxtInvID.text) = 0 Then
      
        Exit Sub
    Else
     
        If SystemOptions.ReturnSAlesByBarcode = False Then
        Retrive_Sales_invoice_data (val(Me.TxtInvID.text))
        End If
        
   FullGrid val(Me.TxtInvID.text)
    End If
With Me.FgC
.ColHidden(.ColIndex("IsRetCopon")) = False
End With
 'NewGrid.Calculate 1, , , True
 
         
             '       If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                            NewGrid.DtpBillDate_Change
                            NewGrid.Calculate 1, , , True
             '           End If
   
      
End Sub

Private Sub TxtInvSerial_Change()
If Me.TxtModFlg.text = "N" Then
RetriveReSalin
End If
End Sub

Private Sub TxtInvSerial_KeyPress(KeyAscii As Integer)
If SystemOptions.returnByBarCodeOnly = True Then
KeyAscii = 0
End If

End Sub

Private Sub TxtInvSerial_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Unload FrmBuySearch
           FrmBuySearch.DealingForm = GridTransType.InvoiceTransaction
     FrmBuySearch.Index = 131
     If SystemOptions.UserInterface = ArabicInterface Then
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ›« Ê—… „»Ì⁄«    "
         Else
                FrmBuySearch.Caption = "Search Sales Invoices "
         End If
            FrmBuySearch.show vbModal
            
End If
End Sub

Public Sub TxtItemCodeB_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub TxtItemCodeB_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch2
        FrmItemSearch2.RetrunType = 1
         FrmItemSearch2.show vbModal
    End If

End Sub

Private Sub TxtItemCodeB1_Change()
'If TxtShortName <> "" Then

' End If
 
End Sub

Private Sub TxtItemCodeB1_GotFocus()
TxtShortName.text = ""
If Trim(TxtShortName) <> "" Then
    SerchItems (TxtShortName)
End If
End Sub

Private Sub TxtItemCodeB1_KeyUp(KeyCode As Integer, Shift As Integer)
showComm
End Sub

Private Sub TxtItemCodeB1_LostFocus()
RelinVatGrid
End Sub

Private Sub txtItemCodeSearch2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
NewGrid.DcboItemCodeFromTextSerailCode
End If
End Sub

Private Sub txtItemCodeSearch2_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 7715
        FrmItemSearch.show vbModal
    End If

    If KeyCode = vbKeyF9 Then
                    
        FrmSearchSerial.XPTxtCode.text = txtItemCodeSearch2.text
        FrmSearchSerial.show
        FrmSearchSerial.Cmd_Click (0)
                    
    End If
End Sub

Private Sub TxtNetValue_Change()
    'If Me.TxtModFlg.text <> "E" Then
If SystemOptions.UserInterface = EnglishInterface Then
    FrmCustomerDisplay.LblInformation2.Caption = "  Total Amount   " & val(Me.TxtNetValue) & "" & "     SAR "
Else
FrmCustomerDisplay.LblInformation2.Caption = "  «·«Ã„«·Ì     " & val(Me.TxtNetValue) & "" & " —Ì‹«· ”⁄ÊœÌ  " 'vbNewLine

End If

    TxtRemainValue.text = val(Me.TxtPayedValue.text) - val(Me.TxtNetValue.text)
     TxtRemainValue.text = Round(TxtRemainValue.text, 2)
    'End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
       TxtPayedValue = val(Me.TxtNetValue.text)
    End If
    
   
If optsale(0).value = True Then
If inddx = 1 Then
 FgC.Clear flexClearScrollable, flexClearEverything
     FgC.rows = 1
CalculteValue
End If
End If

End Sub
Sub CalculteValue()

Dim i As Integer
  Dim Vlue As Double
  Dim TempVlue As Double
  Dim netvalue As Double
    Vlue = val(TxtNetValue.text)
    i = 1
     Do While Vlue > 0
    TempVlue = GetValue(Vlue, Me.XPDtbBill.value, netvalue)
    
    If TempVlue > 0 Then
    
    With FgC
   .rows = .rows + 1
    Vlue = Vlue - TempVlue
      .TextMatrix(i, .ColIndex("Serial")) = i
      .TextMatrix(i, .ColIndex("Vlue")) = i
       .TextMatrix(i, .ColIndex("Vlue")) = netvalue
       i = i + 1
    End With
    Else
    Vlue = TempVlue
   End If
    Loop
  
End Sub
Sub FullGrid(Optional TransID As Double)
Dim Rs4 As ADODB.Recordset
Dim sql As String
Dim i As Integer
FgC.Clear flexClearScrollable, flexClearEverything
      FgC.rows = 1
    
sql = "SELECT     dbo.TblTransCoupons.*"
sql = sql & " From dbo.TblTransCoupons"
If TransID = 0 Then
sql = sql & " Where (Transaction_ID = " & val(XPTxtBillID.text) & ")"
Else
sql = sql & " Where (Transaction_ID = " & TransID & ")"

End If
Set Rs4 = New ADODB.Recordset
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
With FgC
Rs4.MoveFirst
.rows = .rows + Rs4.RecordCount
For i = 1 To .rows - 1
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(Rs4("ID").value), 0, Rs4("ID").value)
.TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(Rs4("Transaction_ID").value), 0, Rs4("Transaction_ID").value)
.TextMatrix(i, .ColIndex("Vlue")) = IIf(IsNull(Rs4("Vlue").value), 0, Rs4("Vlue").value)
.TextMatrix(i, .ColIndex("Num")) = IIf(IsNull(Rs4("Num").value), "", Rs4("Num").value)
.TextMatrix(i, .ColIndex("Selcd")) = IIf(IsNull(Rs4("Selcd").value), 0, Rs4("Selcd").value)
.TextMatrix(i, .ColIndex("PerioDID")) = IIf(IsNull(Rs4("periodID").value), 0, Rs4("periodID").value)
.TextMatrix(i, .ColIndex("IsRetCopon")) = 0

Rs4.MoveNext
Next i
End With
End If
'Relim
End Sub
Function ChekGrid2(Optional Num As String) As Boolean
Dim i As Integer
With FgC
ChekGrid2 = False
For i = 1 To .rows - 1
If .cell(flexcpChecked, i, .ColIndex("Selcd2")) = flexChecked And Num = .TextMatrix(i, .ColIndex("Num")) Then
ChekGrid2 = True
Exit Function
End If
Next i
End With
End Function
Function ChekGrid(Optional Row As Long, Optional Num As String) As Boolean
Dim i As Integer
With FgC
ChekGrid = False
For i = 1 To .rows - 1
If i <> Row And .cell(flexcpChecked, i, .ColIndex("Selcd")) = flexChecked And Num = .TextMatrix(i, .ColIndex("Num")) Then
ChekGrid = True
Exit Function
End If
Next i
End With
End Function
Function GetValue(Optional Vlue As Double, Optional DateTrna As Date, Optional ByRef NetVlue As Double) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     dbo.TblCouponsDet.TypTrans, dbo.TblCouponsDet.Vlue, dbo.TblCouponsDet.FromVlue, dbo.TblCouponsDet.TOVlue, dbo.TblCoupons.FromDate2,"
sql = sql & "                      dbo.TblCoupons.ToDate2 "
sql = sql & " FROM         dbo.TblCouponsDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblCoupons ON dbo.TblCouponsDet.CoupID = dbo.TblCoupons.ID"
sql = sql & " WHERE     (dbo.TblCouponsDet.TypTrans = 0) AND (dbo.TblCouponsDet.FromVlue <= " & Vlue & ") AND (dbo.TblCouponsDet.TOVlue >= " & Vlue & ") AND"
sql = sql & "  (dbo.TblCoupons.FromDate <= " & SQLDate(DateTrna, True) & ") AND (dbo.TblCoupons.ToDate  >= " & SQLDate(DateTrna, True) & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetValue = IIf(IsNull(Rs3("FromVlue").value), 0, Rs3("FromVlue").value)
NetVlue = IIf(IsNull(Rs3("Vlue").value), 0, Rs3("Vlue").value)
Else
GetValue = 0
NetVlue = 0
End If
End Function
Function CheckPeriodCopon(Optional DateTrna As Date) As Boolean
Exit Function
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT    dbo.TblCouponsDet.ID, dbo.TblCouponsDet.TypTrans, dbo.TblCouponsDet.Transaction_ID, dbo.TblCouponsDet.Num, dbo.TblCoupons.FromDate2, dbo.TblCoupons.ToDate2 ,dbo.TblCouponsDet.Vlue"
sql = sql & " FROM         dbo.TblCouponsDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblCoupons ON dbo.TblCouponsDet.CoupID = dbo.TblCoupons.ID"
sql = sql & " WHERE     (dbo.TblCouponsDet.TypTrans = 1) AND (dbo.TblCouponsDet.NewTransaction_ID = 0 ) AND (dbo.TblCouponsDet.RetTransaction_ID = 0 )  "
sql = sql & "  and (dbo.TblCoupons.FromDate2 <= " & SQLDate(DateTrna, True) & ") AND (dbo.TblCoupons.ToDate2 >= " & SQLDate(DateTrna, True) & ")"
 Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckPeriodCopon = True
Else
CheckPeriodCopon = False
End If
End Function
Function CheckParcode4(Optional Num As String, Optional DateTrna As Date, Optional ByRef ID As Double, Optional ByRef Valu As Double) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT    dbo.TblCouponsDet.ID, dbo.TblCouponsDet.TypTrans, dbo.TblCouponsDet.Transaction_ID, dbo.TblCouponsDet.Num, dbo.TblCoupons.FromDate2, dbo.TblCoupons.ToDate2 ,dbo.TblCouponsDet.Vlue"
sql = sql & " FROM         dbo.TblCouponsDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblCoupons ON dbo.TblCouponsDet.CoupID = dbo.TblCoupons.ID"
sql = sql & " WHERE      (dbo.TblCouponsDet.TypTrans = 1) AND (dbo.TblCouponsDet.NewTransaction_ID = 0 ) AND (dbo.TblCouponsDet.RetTransaction_ID = 0 ) AND (dbo.TblCouponsDet.Transaction_ID <> 0 ) AND (dbo.TblCouponsDet.Num = N'" & Num & "')"
sql = sql & "  and (dbo.TblCoupons.FromDate2 <= " & SQLDate(DateTrna, True) & ") AND (dbo.TblCoupons.ToDate2 >= " & SQLDate(DateTrna, True) & ")"
 Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckParcode4 = True
ID = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
Valu = IIf(IsNull(Rs3("Vlue").value), 0, Rs3("Vlue").value)
Else
CheckParcode4 = False
ID = 0
Valu = 0
End If
End Function

Function CheckParcode3(Optional Num As String, Optional DateTrna As Date) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT    dbo.TblCouponsDet.ID, dbo.TblCouponsDet.TypTrans, dbo.TblCouponsDet.Transaction_ID, dbo.TblCouponsDet.Num, dbo.TblCoupons.FromDate2, dbo.TblCoupons.ToDate2 ,dbo.TblCouponsDet.Vlue"
sql = sql & " FROM         dbo.TblCouponsDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblCoupons ON dbo.TblCouponsDet.CoupID = dbo.TblCoupons.ID"
sql = sql & " WHERE      (dbo.TblCouponsDet.TypTrans = 1) AND (dbo.TblCouponsDet.NewTransaction_ID = 0 ) AND (dbo.TblCouponsDet.RetTransaction_ID = 0 ) AND (dbo.TblCouponsDet.Transaction_ID <> 0 ) AND (dbo.TblCouponsDet.Num = N'" & Num & "')"
sql = sql & "  and (dbo.TblCoupons.FromDate2 <= " & SQLDate(DateTrna, True) & ") AND (dbo.TblCoupons.ToDate2 >= " & SQLDate(DateTrna, True) & ")"
 Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckParcode3 = True
Else
CheckParcode3 = False
End If
End Function

Function CheckParcode2(Optional Num As String, Optional DateTrna As Date, Optional Vlue As Double, Optional ByRef ID As Double, Optional Transaction_ID As Double) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT    dbo.TblCouponsDet.ID, dbo.TblCouponsDet.TypTrans, dbo.TblCouponsDet.Transaction_ID, dbo.TblCouponsDet.Num, dbo.TblCoupons.FromDate2, dbo.TblCoupons.ToDate2 ,dbo.TblCouponsDet.Vlue"
sql = sql & " FROM         dbo.TblCouponsDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblCoupons ON dbo.TblCouponsDet.CoupID = dbo.TblCoupons.ID"
sql = sql & " WHERE   (dbo.TblCouponsDet.Vlue = " & Vlue & ") and   (dbo.TblCouponsDet.TypTrans = 1) AND (dbo.TblCouponsDet.NewTransaction_ID = 0 ) AND (dbo.TblCouponsDet.RetTransaction_ID = 0 ) AND (dbo.TblCouponsDet.Transaction_ID <> 0 ) AND (dbo.TblCouponsDet.Num = N'" & Num & "')"
sql = sql & "  and (dbo.TblCoupons.FromDate2 <= " & SQLDate(DateTrna, True) & ") AND (dbo.TblCoupons.ToDate2 >= " & SQLDate(DateTrna, True) & ")"
 Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
ID = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
CheckParcode2 = True
Else
CheckParcode2 = False
ID = 0
End If
End Function
Function CheckParcode(Optional Num As String, Optional DateTrna As Date, Optional Vlue As Double, Optional ByRef ID As Double, Optional Transaction_ID As Double) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT    dbo.TblCouponsDet.ID, dbo.TblCouponsDet.TypTrans, dbo.TblCouponsDet.Transaction_ID, dbo.TblCouponsDet.Num, dbo.TblCoupons.FromDate2, dbo.TblCoupons.ToDate2 ,dbo.TblCouponsDet.Vlue"
sql = sql & " FROM         dbo.TblCouponsDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblCoupons ON dbo.TblCouponsDet.CoupID = dbo.TblCoupons.ID"
sql = sql & " WHERE   (dbo.TblCouponsDet.Vlue = " & Vlue & ") and   (dbo.TblCouponsDet.TypTrans = 1) AND (dbo.TblCouponsDet.Transaction_ID = 0 or dbo.TblCouponsDet.Transaction_ID=" & Transaction_ID & ") AND (dbo.TblCouponsDet.Num = N'" & Num & "')"
sql = sql & "  and (dbo.TblCoupons.FromDate <= " & SQLDate(DateTrna, True) & ") AND (dbo.TblCoupons.ToDate >= " & SQLDate(DateTrna, True) & ")"
 Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
ID = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
CheckParcode = True
Else
CheckParcode = False
ID = 0
End If
End Function
Private Sub Fgc_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FgC
Select Case .ColKey(Col)
Case "Vlue"
Cancel = True
Case "Selcd"
If optsale(1).value = True Then
Cancel = False
Else
Cancel = True
End If
Case "Num"
If optsale(0).value = True Then
Cancel = False
.ComboList = ""
Else
Cancel = True
End If
Case "Num2"
.ComboList = ""
Case "IsRetCopon"
Cancel = True
End Select
End With
End Sub
Sub RelimSali()
 XPCboDiscountType.ListIndex = 0
XPTxtDiscountVal.text = 0

Dim TempNetValue As Double
If optsale(0).value = True Then
Dim i As Integer
Dim SmVal As Double
SmVal = 0
With FgC
For i = 1 To .rows - 1
If .cell(flexcpChecked, i, .ColIndex("Selcd2")) = flexChecked Then
SmVal = SmVal + val(.TextMatrix(i, .ColIndex("Vlue")))
End If
Next i
End With
If SmVal > 0 Then

XPCboDiscountType.ListIndex = 1
XPTxtDiscountVal.text = SmVal
Else
XPCboDiscountType.ListIndex = 0
XPTxtDiscountVal.text = 0

End If

 NewGrid.Calculate 1, , , True
 
LblDiscountsTotalView(1).Caption = Format(SmVal, "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

LblTotal.Caption = val(LblTotalAll.Caption) - val(LblDiscountsTotal.Caption) + IIf(SystemOptions.PriceWithVAT = True, val(TxtValueAdded.text) * -1, val(TxtValueAdded.text)) '- SmVal
 LblTotal.Caption = Round(LblTotal.Caption, 2)
LBLPayVal.Caption = val(TxtNetValue.text) + val(TxtValueAdded.text)
End If
End Sub

Sub Relim()
XPCboDiscountType.ListIndex = 0
XPTxtDiscountVal.text = 0

Dim TempNetValue As Double
If optsale(1).value = True Then
Dim i As Integer
Dim SmVal As Double
SmVal = 0
With FgC
For i = 1 To .rows - 1
If .cell(flexcpChecked, i, .ColIndex("IsRetCopon")) = flexUnchecked Then
SmVal = SmVal + val(.TextMatrix(i, .ColIndex("Vlue")))
End If
Next i
End With
If SmVal > 0 Then

XPCboDiscountType.ListIndex = 1
XPTxtDiscountVal.text = SmVal
Else
XPCboDiscountType.ListIndex = 0
XPTxtDiscountVal.text = 0

End If

 NewGrid.Calculate 1, , , True
 
LblDiscountsTotalView(1).Caption = Format(SmVal, "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

LblTotal.Caption = val(LblTotalAll.Caption) - val(LblDiscountsTotal.Caption) '- SmVal
 LblTotal.Caption = Round(LblTotal.Caption, 2)
LBLPayVal.Caption = TxtNetValue.text
End If
End Sub
Sub ChangCel(ByVal Row As Long, ByVal Col As Long)
'If Me.TxtModFlg.Text <> "R" Then
If optsale(0).value = True Then
Dim sql As String
Dim PerioDID As Double
With FgC
Select Case .ColKey(Col)
Case "Num"
If Row = 0 Then Exit Sub
If .TextMatrix(Row, .ColIndex("Num")) <> "" Then
If CheckParcode(.TextMatrix(Row, .ColIndex("Num")), Me.XPDtbBill.value, val(.TextMatrix(Row, .ColIndex("Vlue"))), PerioDID, val(XPTxtBillID.text)) = True Then
                If ChekGrid(Row, .TextMatrix(Row, .ColIndex("Num"))) = True Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "Â–« «·ﬂÊ»Ê‰ „” Œœ„ ·«Ì„ﬂ‰ «· ﬂ—«—"
                                Else
                                     MsgBox "This is coupon were used previously"
                                End If
                                .TextMatrix(Row, .ColIndex("Selcd")) = 0
                                .TextMatrix(Row, .ColIndex("Num")) = ""
                                .TextMatrix(Row, .ColIndex("Selcd2")) = 0
                                

              '  Exit Sub
                Else
                .TextMatrix(Row, .ColIndex("PerioDID")) = PerioDID
                .TextMatrix(Row, .ColIndex("Selcd")) = 1
                .TextMatrix(Row, .ColIndex("Selcd2")) = 0
                End If
ElseIf CheckParcode2(.TextMatrix(Row, .ColIndex("Num")), Me.XPDtbBill.value, val(.TextMatrix(Row, .ColIndex("Vlue"))), PerioDID, val(XPTxtBillID.text)) = True Then
 .TextMatrix(Row, .ColIndex("PerioDID")) = PerioDID
                .TextMatrix(Row, .ColIndex("Selcd")) = 1
                .TextMatrix(Row, .ColIndex("Selcd2")) = 1
               
                
Else
.TextMatrix(Row, .ColIndex("Selcd2")) = 0
                .TextMatrix(Row, .ColIndex("Selcd")) = 0
                .TextMatrix(Row, .ColIndex("Num")) = ""
                If SystemOptions.UserInterface = ArabicInterface Then
                                      MsgBox "Â–« «·ﬂÊ»Ê‰ „” Œœ„ „‰ ﬁ»· «Ê €Ì— „ÊÃÊœ Œ·«· Â–Â «·› —…"
                 Else
                                     MsgBox "This is coupon were used previously.or not in period"
                  End If
                                
              '  Exit Sub
                'msgbox"Wrong Bo"
End If
End If
End Select
End With
End If

RelimSali
End Sub
Private Sub Fgc_CellChanged(ByVal Row As Long, ByVal Col As Long)


'End If
End Sub
Sub SaveCopoun()
Dim i As Integer
Dim sql As String
  Dim RsDevsub As ADODB.Recordset
  Set RsDevsub = New ADODB.Recordset
  
If Me.TxtModFlg.text = "E" Then
Cn.Execute "Update  TblCouponsDet set Transaction_ID=0,BillNo=null where Transaction_ID=" & val(XPTxtBillID.text) & ""
End If
   sql = "SELECT  *  from TblTransCoupons Where (1 = -1)"
    RsDevsub.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
With FgC
For i = 1 To .rows - 1
If optsale(0).value = True Then
If .cell(flexcpChecked, i, .ColIndex("Selcd")) = flexChecked Then
       Dim StrRecID As String
       StrRecID = new_id("TblTransCoupons", "ID", "")
       RsDevsub.AddNew
       RsDevsub.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
       RsDevsub("Transaction_ID").value = val(XPTxtBillID.text)
       RsDevsub("Vlue").value = val(.TextMatrix(i, .ColIndex("Vlue")))
      If .cell(flexcpChecked, i, .ColIndex("Selcd2")) = flexChecked Then
       RsDevsub("Selcd2").value = 1
       Else
       RsDevsub("Selcd2").value = 1
       End If
       If .cell(flexcpChecked, i, .ColIndex("Selcd")) = flexChecked Then
       RsDevsub("Selcd").value = 1
       
       RsDevsub("Num").value = .TextMatrix(i, .ColIndex("Num"))
       Else
       RsDevsub("Num").value = Null
       RsDevsub("Selcd").value = 0
       End If
       
       RsDevsub("PerioDID").value = val(.TextMatrix(i, .ColIndex("PerioDID")))
       RsDevsub.update
    End If
End If
If .cell(flexcpChecked, i, .ColIndex("Selcd2")) = flexChecked Then
If optsale(0).value = True Then
sql = " update TblCouponsDet set NewBillNo='" & TxtNoteSerial1.text & "' "
sql = sql & " ,NewTransaction_ID=" & val(XPTxtBillID.text) & "  where id=" & val(.TextMatrix(i, .ColIndex("PerioDID"))) & ""
Cn.Execute sql
End If

ElseIf .cell(flexcpChecked, i, .ColIndex("Selcd")) = flexChecked Then
If optsale(0).value = True Then
sql = " update TblCouponsDet set BillNo='" & TxtNoteSerial1.text & "' "
sql = sql & " ,Transaction_ID=" & val(XPTxtBillID.text) & "  where id=" & val(.TextMatrix(i, .ColIndex("PerioDID"))) & ""
Cn.Execute sql
End If
End If

If optsale(1).value = True Then
sql = " update TblCouponsDet set  ReturnBillNo='" & TxtNoteSerial1.text & "' "
sql = sql & " ,RetTransaction_ID=" & val(XPTxtBillID.text) & "  where id=" & val(.TextMatrix(i, .ColIndex("PerioDID"))) & ""
Cn.Execute sql

End If
If .cell(flexcpChecked, i, .ColIndex("IsRetCopon")) = flexUnchecked Then
If optsale(1).value = True Then
sql = " update TblCouponsDet set IsRetCopon=1   where id=" & val(.TextMatrix(i, .ColIndex("PerioDID"))) & ""
Cn.Execute sql
End If
End If
Next i
End With
'FullGrid
End Sub
Sub DelCopoun()
Cn.Execute "Update  TblCouponsDet set Transaction_ID=0,BillNo=null where Transaction_ID=" & val(XPTxtBillID.text) & ""
Cn.Execute "Delete from TblTransCoupons where Transaction_ID=" & val(XPTxtBillID.text) & ""
End Sub

Private Sub TxtNetValue_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    TxtNetValue.ToolTipText = WriteNo(LblTotal.Caption, 0, True)
End Sub

Function Retrive_Sales_invoice_data(Transaction_ID As Long)
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Integer
    
    Dim rs As ADODB.Recordset
    
 '**************************************************************************
 
        StrSQL = "Select * from transactions where  Transaction_ID=" & Transaction_ID
  

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Function
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
  
'        Me.DCboStoreName.BoundText = IIf(IsNull(rs("storeid").value), "", rs("storeid").value)
'        Me.dcBranch.BoundText = IIf(IsNull(rs("Branchid").value), "", rs("Branchid").value)

XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), -1, val(rs("Trans_DiscountType").value))
   ' CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)

    XPTxtDiscountVal.text = IIf(IsNull(rs("Trans_Discount").value), "", (rs("Trans_Discount").value))
    
    
    TxtPhone(1).text = IIf(IsNull(rs("VATNO").value), "", (rs("VATNO").value))
    If Not (IsNull(rs("CashCustomerPhone").value)) Then
        Me.TxtPhone(0).text = rs("CashCustomerPhone").value
    Else
        Me.TxtPhone(0).text = ""
    End If


    If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.CashCustomerName.text = rs("CashCustomerName").value
    Else
        Me.CashCustomerName.text = ""
    End If
    
  End If
     '**************************************************************************
rs.Close
Set rs = Nothing
     
    
    StrSQL = "SELECT dbo.Transaction_Details.ItemSerial , TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & Transaction_ID

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        row_count = FG.rows
    
        If FG.TextMatrix(row_count - 1, FG.ColIndex("Code")) = "" Then
            row_count = row_count - 1
        End If
     
        FG.rows = RsDetails.RecordCount + row_count

        For Num = row_count To FG.rows - 1 'RsDetails.RecordCount
    
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no")), "", (RsDetails("order_no").value))
            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate")), "", (RsDetails("OrderArrivalDate").value))
            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
        FG.TextMatrix(Num, FG.ColIndex("Select")) = True

        FG.TextMatrix(Num, FG.ColIndex("ParrtNoCode")) = IIf(IsNull(RsDetails("ParrtNoCode")), "", (RsDetails("ParrtNoCode").value))
        FG.TextMatrix(Num, FG.ColIndex("ItemDetailedCode")) = IIf(IsNull(RsDetails("ItemDetailedCode")), "", (RsDetails("ItemDetailedCode").value))
        
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))
        FG.TextMatrix(Num, FG.ColIndex("PrintName")) = IIf(IsNull(RsDetails("ItemName")), "", (RsDetails("ItemName").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
            FG.TextMatrix(Num, FG.ColIndex("Valu")) = val(FG.TextMatrix(Num, FG.ColIndex("Price"))) * val(FG.TextMatrix(Num, FG.ColIndex("Count")))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
              
              FG.TextMatrix(Num, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", (RsDetails("ItemSerial").value))

            End If

            FG.TextMatrix(Num, FG.ColIndex("Vat")) = IIf(IsNull(RsDetails("Vat")), "", (RsDetails("Vat").value))
            FG.TextMatrix(Num, FG.ColIndex("Vatyo")) = IIf(IsNull(RsDetails("Vatyo")), "", (RsDetails("Vatyo").value))
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemCostPrice")) = IIf(IsNull(RsDetails("CostPrice")), "", (RsDetails("CostPrice").value))
            FG.TextMatrix(Num, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(Num, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(Num, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
            FG.TextMatrix(Num, FG.ColIndex("IsExpirDate")) = IIf(IsNull(RsDetails("IsExpirDate")), "", (RsDetails("IsExpirDate").value))
        
            RsDetails.MoveNext
        Next Num

    End If
    If NewGrid.Calculate(1, , False, True) = False Then
 
    End If
    RetriveValueAdded
End Function
 


Private Sub TXTOrDer_no_Change()

    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder Me.TXTOrDer_no
    End If

End Sub

Public Function NewBillFromOrder(orderNo As String)

    If Me.TxtModFlg = "R" Then
        Cmd_Click (0)
        Me.TXTOrDer_no.text = orderNo
        'txtorder_no_Change
        'RetriveOrder orderNo
    End If

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
     TxtRemainValue.text = Round(TxtRemainValue.text, 2)
     
    'End If

                                             
Dim MSTR As String
Dim MSTR1 As String
Dim MSTR2 As String


                                                If FrmCustomerDisplay.Visible = True And FramePay.Visible = True Then
                                               
                                                
                                               
                                                If SystemOptions.UserInterface = ArabicInterface Then
                                             
                                            MSTR1 = " «·„œ›Ê⁄  "
                                            
                                            MSTR2 = " «·„ »ﬁÌ  "
                                            Else
                                             
                                                MSTR1 = "Payed  "
                                                MSTR2 = "Remain  "
                                            End If
                                            
                                            
                                            
                                            FrmCustomerDisplay.LblInformation3.Caption = MSTR1 & TxtPayedValue & vbNewLine & MSTR2 & TxtRemainValue
                                             'FrmCustomerDisplay.LblInformation2.Caption = MSTR2 & TxtRemainValue
                                             

                                              End If

                                             
                                             
End Sub

Private Sub TxtPayedValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPayedValue.text, 0)
End Sub



Private Sub TxtPhone_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  CashCustomerName = GetCashCustomernamebyphone(TxtPhone(0).text)
    End If

    If KeyAscii = vbKeyReturn Then
                If Index = 0 Then
                XPCboDiscountType_Click
                End If
    End If
End Sub

Private Sub TxtPrice_KeyPress(KeyAscii As Integer)
Dim X As Double
End Sub

Private Sub TXTPrintInvoice_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF3 Then
Unload FrmBuySearch
           FrmBuySearch.DealingForm = GridTransType.InvoiceTransaction
     FrmBuySearch.Index = 1205
     If SystemOptions.UserInterface = ArabicInterface Then
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ›« Ê—… „»Ì⁄«    "
         Else
                FrmBuySearch.Caption = "Search Sales Invoices "
         End If
            FrmBuySearch.show vbModal
            
End If

If KeyCode = vbKeyReturn Then
printInvoice
   End If
End Sub
Public Function printInvoice()
Dim SaleReport As New ClsSaleReport
Dim getSallingID As Long
 

  SaleReport.ShowSallingData 0, 0, 0, val(Me.TxtPayedValue.text), val(Me.TxtRemainValue.text), 0, "", , "", , , , , TXTPrintInvoice, getSallingID
  TXTPrintInvoice.text = ""
End Function

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub

 

Function LoadSpecificItems()
Dim StrWhere As String
Dim StrSQL As String
    StrSQL = "SELECT ItemID,barCodeNO From TblItems  where 1=1  "
 StrWhere = ""
 
If TxtShortName.text <> "" Then
       StrSQL = StrSQL + " and ShortName like'%" & TxtShortName.text & "%'"
            StrSQL = StrSQL + " or  barCodeNO like'%" & TxtShortName.text & "%'"
            StrSQL = StrSQL + " or  ItemName like'%" & TxtShortName.text & "%'"
            StrSQL = StrSQL + " or  ItemNamee like'%" & TxtShortName.text & "%'"
            
            
End If

'If TxtShorCode.Text <> "" Then
'       StrSQL = StrSQL + " and barCodeNO like'%" & TxtShorCode.Text & "%'"
'End If

     If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = StrSQL + " Order BY barCodeNO "
    Else
        StrSQL = StrSQL + " Order BY barCodeNO "
    End If

fill_combo DCboItemsCode, StrSQL
    
    If SystemOptions.UserInterface = ArabicInterface Then

        StrSQL = "SELECT ItemID,ItemName From  TblItems  where 1=1"
    Else
        StrSQL = "SELECT ItemID,ItemNamee From  TblItems   where 1=1"
    End If
    
 If TxtShortName.text <> "" Then
'       StrSQL = StrSQL + " and ShortName like'%" & TxtShortName.Text & "%'"
            StrSQL = StrSQL + " or  barCodeNO like'%" & TxtShortName.text & "%'"
            StrSQL = StrSQL + " or  ItemName like'%" & TxtShortName.text & "%'"
            StrSQL = StrSQL + " or  ItemNamee like'%" & TxtShortName.text & "%'"

End If

'If TxtShorCode.Text <> "" Then
'       StrSQL = StrSQL + " and barCodeNO like'%" & TxtShorCode & "%'"
'End If

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = StrSQL + " Order BY ItemName  "
    Else
        StrSQL = StrSQL + " Order BY ItemName  "
    End If
        fill_combo DCboItemsName, StrSQL

        'GetComboData GetItemsNames, StrSQL
   
'      Dcombos.GetItemsCodes Me.DCboItemsCode, -1, -1, , val(DCboStoreName.BoundText), , TxtshortName.Text
'       Dcombos.GetItemsNames Me.DCboItemsName, -1, -1, , , val(DCboStoreName.BoundText), , TxtshortName.Text
        
        


End Function

Private Sub TxtShorCode_KeyDown(KeyCode As Integer, Shift As Integer)
'   LoadSpecificItems
'      If KeyCode = vbKeyReturn Then
'
'
'   DCboItemsName.SetFocus
'        SendKeys "{F4}"
'        End If
        
End Sub

 Sub SerchItems(Optional str As String)
 
Dim sql As String
Dim SQL1 As String
   
    SerchItemspUBLIC str, sql, SQL1
    fill_combo DCboItemsCode, sql
  fill_combo DCboItemsName, SQL1
        
         
End Sub
Sub SerchItemsxx(Optional str As String)
 
Set DCboItemsCode.RowSource = Nothing
Set DCboItemsName.RowSource = Nothing
If str <> "" Then
Dim sql As String
Dim SQL1 As String
 
Dim StrWhere As String
  Dim astrSplit2tems2() As String
  Dim j As Integer
  Dim nElements As Integer
  Dim SearchString As String
StrWhere = ""
SearchString = ""
sql = " select  ItemID,barCodeNO   from  dbo.TblItems where TblItems.IsArchive=0"
If SystemOptions.UserInterface = ArabicInterface Then
SQL1 = " select  ItemID,ItemName   from  dbo.TblItems where TblItems.IsArchive=0"
Else
SQL1 = " select  ItemID,ItemNamee   from  dbo.TblItems where TblItems.IsArchive=0"
End If
  If SystemOptions.ShowOnlyItemsOfSales = True Then
    SQL1 = SQL1 & "  and dbo.TblItems.GroupID in(   SELECT     GroupID"
    SQL1 = SQL1 & " From dbo.Groups  WHERE     (ISNULL(POSGroup, 0) = 1))"
  End If
  
  
  If SystemOptions.WorkWithLINKEDiActivity = True Then
   SQL1 = SQL1 & "  and dbo.TblItems.GroupID in(   SELECT     GroupID"
    SQL1 = SQL1 & " select GroupID from fullgroups ()  )"
 
End If


          astrSplit2tems2 = Split(str, " ")
          nElements = UBound(astrSplit2tems2) - LBound(astrSplit2tems2)
          If nElements = 0 Then
                             If SystemOptions.UserInterface = ArabicInterface Then
                            StrWhere = " and (ItemName Like N'%" & Trim(str) & "%' or barCodeNO Like N'%" & Trim(str) & "%' or shortName Like N'%" & Trim(str) & "%'  or fullcode Like N'%" & Trim(str) & "%') "
                    Else
                            StrWhere = " and (ItemNamee Like N'%" & Trim(str) & "%' or barCodeNO Like N'%" & Trim(str) & "%' or shortName Like N'%" & Trim(str) & "%' or fullcode Like N'%" & Trim(str) & "%' ) "
                    End If
                    
          End If
        If nElements > 0 Then
        
     '   StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(0)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(0)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(0)) & "%') "
        SearchString = ""
        For j = 0 To nElements
        
         SearchString = SearchString & "%" & Trim(astrSplit2tems2(j))
             '     SearchString = "%" & Trim(astrSplit2tems2(j)) & SearchString
                  
        '   StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(j)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(j)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(j)) & "%') "
        '   StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
         Next j
         SearchString = SearchString & "%"
                             If SystemOptions.UserInterface = ArabicInterface Then

             StrWhere = StrWhere + " and (ItemName Like '" & SearchString & "' or barCodeNO Like '" & SearchString & "' or shortName Like '" & SearchString & "') "
             Else
              StrWhere = StrWhere + " and (ItemNamee Like '" & SearchString & "' or barCodeNO Like '" & SearchString & "' or shortName Like '" & SearchString & "') "
             End If
        '-  StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
      
         End If
        
    sql = sql & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        sql = sql + " Order BY barCodeNO "
    Else
        sql = sql + " Order BY barCodeNO "
    End If


    SQL1 = SQL1 & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        SQL1 = SQL1 + " Order BY ItemName "
    Else
        SQL1 = SQL1 + " Order BY ItemNamee "
    End If
    
   End If
    fill_combo DCboItemsCode, sql
        fill_combo DCboItemsName, SQL1
                If str = "" Then
                                 sql = " select  ItemID,barCodeNO   from  dbo.TblItems where TblItems.IsArchive=0"
                                             If SystemOptions.ShowOnlyItemsOfSales = True Then
                                      sql = sql & "  and dbo.TblItems.GroupID in(   SELECT     GroupID"
                                       sql = sql & " From dbo.Groups  WHERE     (ISNULL(POSGroup, 0) = 1))"
                                    End If
                                    
                                    
                                      If SystemOptions.WorkWithLINKEDiActivity = True Then
   SQL1 = SQL1 & "  and dbo.TblItems.GroupID in(   SELECT     GroupID"
    SQL1 = SQL1 & " select GroupID from fullgroups ()  )"
 
End If


                                 If SystemOptions.UserInterface = ArabicInterface Then
                                 SQL1 = " select  ItemID,ItemName   from  dbo.TblItems where TblItems.IsArchive=0"
                                   If SystemOptions.ShowOnlyItemsOfSales = True Then
                                      SQL1 = SQL1 & "  and dbo.TblItems.GroupID in(   SELECT     GroupID"
                                      SQL1 = SQL1 & " From dbo.Groups  WHERE     (ISNULL(POSGroup, 0) = 1))"
                                   End If
                                   
                                   
  If SystemOptions.WorkWithLINKEDiActivity = True Then
    SQL1 = SQL1 & "  and dbo.TblItems.GroupID in(   SELECT     GroupID"
     SQL1 = SQL1 & " select GroupID from fullgroups ()  )"
 End If

                                     SQL1 = SQL1 + " Order BY ItemName "
                                 Else
                                 SQL1 = " select  ItemID,ItemNamee   from  dbo.TblItems where TblItems.IsArchive=0"
                                   If SystemOptions.ShowOnlyItemsOfSales = True Then
                                      SQL1 = SQL1 & "  and dbo.TblItems.GroupID in(   SELECT     GroupID"
                                       SQL1 = SQL1 & " From dbo.Groups  WHERE     (ISNULL(POSGroup, 0) = 1))"
                                    End If
                                    
                                    
                                      If SystemOptions.WorkWithLINKEDiActivity = True Then
   SQL1 = SQL1 & "  and dbo.TblItems.GroupID in(   SELECT     GroupID"
    SQL1 = SQL1 & " select GroupID from fullgroups ()  )"
 
End If

                                     SQL1 = SQL1 + " Order BY ItemNameE "
                                 End If
                                 
                                     fill_combo DCboItemsCode, sql
                                         fill_combo DCboItemsName, SQL1
                End If
   
       Exit Sub
       
If str <> "" Then
'Dim Sql As String
'Dim StrWhere As String
'  Dim astrSplit2tems2() As String
'  Dim j As Integer
'  Dim nElements As Integer
StrWhere = ""
If SystemOptions.UserInterface = ArabicInterface Then
sql = " select  ItemID,ItemName   from  dbo.TblItems where TblItems.IsArchive=0"
Else
sql = " select  ItemID,ItemNamee   from  dbo.TblItems where TblItems.IsArchive=0"
End If
          astrSplit2tems2 = Split(str, " ")
          nElements = UBound(astrSplit2tems2) - LBound(astrSplit2tems2)
        If nElements > 0 Then
        StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(0)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(0)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(0)) & "%') "
        For j = 1 To nElements - 1
        
           StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(j)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(j)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(j)) & "%') "
           StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
         Next j
         End If
    sql = sql & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        sql = sql + " Order BY ItemName "
    Else
        sql = sql + " Order BY ItemNamee "
    End If


   End If
   
        fill_combo DCboItemsName, sql
        
End Sub

Private Sub TxtShortName_KeyDown(KeyCode As Integer, Shift As Integer)
   'LoadSpecificItems
   SerchItems (TxtShortName)
        If KeyCode = vbKeyReturn Then
        
        
   DCboItemsName.SetFocus
        Sendkeys "{F4}"
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
                        rs.Find "Transaction_Serial='" & StrSearch & "'", , adSearchForward, adBookmarkFirst

                        If Not (rs.BOF Or rs.EOF) Then
                            Me.Retrive rs("Transaction_ID").value
                        Else
                            rs.Bookmark = VarBookMark
                            Msg = "Â–Â «·›« Ê—… €Ì— „ÊÃÊœ…...!!!"
                            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
'RelinVatGrid
'showComm
End Sub

Private Sub VatGrid_Click()
RelinVatGrid
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case Index

        Case 4

            If Not (rs.EOF Or rs.BOF) Then
                rs.MovePrevious

                If rs.BOF Then rs.MoveFirst
            End If

        Case 6

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
            End If

        Case 7
              Me.TxtModFlg.text = ""
        Dim StrSQL As String
     StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=21 "
     
     
       '     If SystemOptions.usertype <> UserAdminAll Then
       '         StrSQL = StrSQL & " AND   BranchId=" & Current_branch
       '     End If


     If SystemOptions.usertype <> UserAdminAll Then
 
       '   If SystemOptions.FixedCustomer = 1 Then
            StrSQL = StrSQL & " and  UserID = " & user_id
       '      End If
  
        Me.dcBranch.Enabled = True
      
      
    End If
    StrSQL = StrSQL & " and  (DATEDIFF(day, Transaction_Date, " & SQLDate(Date, True) & ") <= 61)"
            StrSQL = StrSQL & " Order by Transaction_ID"
                
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            
            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If
   Me.TxtModFlg.text = "R"

          '  If Not (rs.EOF Or rs.BOF) Then
          '      rs.MoveLast
          '  End If

        Case 5

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveNext

                If rs.EOF Then rs.MoveLast
            End If

    End Select
btnpay(0).Enabled = False
btnpay(1).Enabled = False
btnExit(2).Enabled = False
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
        '    If Cmd(3).Enabled = False Then Exit Sub
        '    Cmd_Click (3)
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
    
     If KeyCode = vbKeyF2 Then
        Call btnpay_Click(0)
        PayMode = True    ' ›⁄· Ê÷⁄ «·œ›⁄
        KeyCode = 0
    End If

    ' ·Ê ÷€ÿ Enter
    If KeyCode = vbKeyReturn Then
        If PayMode = True Then
            Call CMDPAy_Click(0)
            PayMode = False   ' —Ã¯⁄ «·Ê÷⁄ ·ÿ»Ì⁄ Â »⁄œ «· ‰›Ì–
            KeyCode = 0
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
    Dim colX As cColumn
    Dim itmX As cListItem
    Dim i    As Long
    Dim j    As Long

    If SystemOptions.CanPayWithoutPrint = True Then
        CMDPAy(1).Visible = True
    Else
        CMDPAy(1).Visible = False

    End If

    If SystemOptions.HideInfroCasher = True Then
        btnExit(9).Visible = False
        btnExit(8).Visible = False
    Else
        btnExit(9).Visible = True
        btnExit(8).Visible = True
    End If
    If SystemOptions.AllowItemsShortName = True Then
        lbl(90).Visible = True
        lbl(91).Visible = True
        lbl(92).Visible = True
        TxtShortName.Visible = True
        'TxtShorCode.Visible = True
        DCboItemsCode.Visible = True
        DCboItemsName.Visible = True
        TxtQuantity.Visible = True
        TxtPrice.Visible = True
        CmdAdd.Visible = True
    
    Else
        lbl(90).Visible = False
        lbl(91).Visible = False
        '    lbl(92).Visible = False
        TxtShortName.Visible = False
        'TxtShorCode.Visible = False
        DCboItemsCode.Visible = False
        DCboItemsName.Visible = False
        TxtQuantity.Visible = False
        TxtPrice.Visible = False
        CmdAdd.Visible = False
    
    End If
    
With Me.DefaultInvoicetype
            .Clear
            
             


            .AddItem " ›« Ê—… ÷—Ì»Ì…  "
            .ItemData(0) = 0
     
            .AddItem " ›« Ê—… ÷—Ì»Ì… „»”ÿ… "
            .ItemData(1) = 2
         
        End With
        
        
    If SystemOptions.WorkWithItemsDetails = True Then
    
        lbl(89).Visible = True
        TxtItemCodeB.Visible = True
        SearchCashCustomer(1).Visible = True

    Else
    
        lbl(90).Visible = True
        lbl(91).Visible = True
        lbl(92).Visible = True
        TxtShortName.Visible = True
        'TxtShorCode.Visible = True
        DCboItemsCode.Visible = True
        DCboItemsName.Visible = True
        TxtQuantity.Visible = True
        TxtPrice.Visible = True
        CmdAdd.Visible = True
        lbl(89).Visible = False
        TxtItemCodeB.Visible = False
        SearchCashCustomer(1).Visible = False
    End If
    
    '   lvwItems.BackgroundPicture = App.path & "\Garphics\wallpaper_Main11.jpg"
    'printtomanyprinter
    lvwMain.Visible = False
    Me.show 'Force to show window
    loadLogo
    Me.backcolor = RGB(220, 228, 243)
    optsale(0).backcolor = Me.backcolor
    optsale(1).backcolor = Me.backcolor
    Frame5.backcolor = Me.backcolor
    FramePay.backcolor = Me.backcolor
    FramePay.Visible = True

    'FillGridWithData
    FramePay.Visible = False

    TimeOut_InSec = 10
    Me.Refresh
    Dim cOptions As ClsCompanyInfo
    Set cOptions = New ClsCompanyInfo
    If SystemOptions.UserInterface = ArabicInterface Then
        lbl(84).Caption = cOptions.ArabCompanyName & CHR(13) & CurrentBranchName
    Else
        lbl(84).Caption = cOptions.EngCompanyName & CHR(13) & CurrentBranchNameE
    End If
 
    '    With lvwItems
    '        lvwItems.Listitems.Clear
    '        .Visible = False
    '        .CustomDraw = True
    '        .AutoArrange = True
    '        .ImageList(eLVLargeIcon) = GrouplImageList ' ilsIcons32
    '        .ImageList(eLVSmallIcon) = GrouplImageList ' ilsIcons16
    '        .ImageList(eLVTileImages) = GrouplImageList ' ilsIcons48
    '        .ImageList(eLVHeaderImages) = GrouplImageList ' ilsIcons16
    '
    '        ' Add column headers
    '        Set colX = .Columns.Add(, "NAME", "Name")
    '        colX.Tag = "Stores the name of the item"
    '        colX.IconIndex = 0
    '        Set colX = .Columns.Add(, "Code", "Code")
    '        colX.Tag = "Stores the date of the item"
    '        colX.IconIndex = 1
    '        Set colX = .Columns.Add(, "id", "id")
    '        colX.Tag = "Stores the size of the item"
    '        colX.Alignment = eLVColumnAlignRight
    '
    '        Set colX = .Columns.Add(, "ItemType", "ItemType")
    '       colX.Tag = "Stores the size of the item"
    '       colX.Alignment = eLVColumnAlignRight
      
    '    End With
 
    '    With lvwTables
    '        .Visible = False
    '        .CustomDraw = True
    '
    '        .AutoArrange = True
    '        .ImageList(eLVLargeIcon) = ilsIcons32
    '        .ImageList(eLVSmallIcon) = ilsIcons16
    '        .ImageList(eLVTileImages) = ilsIcons48
    '        .ImageList(eLVHeaderImages) = ilsIcons16
    '
    '        ' Set up image lists:
    '
    '        ' Add column headers
    '        Set colX = .Columns.Add(, "NAME", "Name")
    '        colX.Tag = "Stores the name of the item"
    '        colX.IconIndex = 0
    '        Set colX = .Columns.Add(, "DATE", "Date")
    '        colX.Tag = "Stores the date of the item"
    '        colX.IconIndex = 1
    '        Set colX = .Columns.Add(, "SIZE", "Size")
    '        colX.Tag = "Stores the size of the item"
    '        colX.Alignment = eLVColumnAlignRight
    '
    '        'For i = 1 To 3
    '        '    .Columns(i).ItemData = i * 100
    '        ' Next i
    '    End With
    '
   
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
   
    'FillGroups
    '   FillTables
    lbl(82).Caption = Date
    lbl(83).Caption = GetWeekdayName(DatePart("w", Date) + 1)

    lblLabel1(0).Width = Me.Width

    lblLabel1(0).AutoSize = True
    ' Load lblLabel1(1)
    ' lblLabel1(1).Visible = True
    '   Load lblLabel1(1).
    lblLabel1(1).Width = Me.Width
    lblLabel1(1).left = Me.Width

    'showmessage
    ' Me.left = (mdifrmmain.Width - Me.Width) / 2
    '    Me.top = (mdifrmmain.Height - Me.Height) / 2
    ScreenNameArabic = " ›« Ê—… «·„»Ì⁄«  "
    ScreenNameEnglish = " Sales Bill"
    ' RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
 
    first_run = True
    Dim StrSQL  As String
    Dim Num     As Integer
    Dim StrList As String
    Dim BGround As New ClsBackGroundPic
    Dim ShowTax As Boolean
    If SystemOptions.UserInterface = EnglishInterface Then

        With Me.CboRetrunType
            .Clear
            .AddItem "With bill "
            .AddItem "With out Bill"
        End With

    Else

        With Me.CboRetrunType
            .Clear
            .AddItem "»›« Ê—…"
            .AddItem "»œÊ‰ ›« Ê—…"
        End With

    End If
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
    Set NewGrid = New ClsGrid
    NewGrid.GridTrans = InvoiceTransaction
    Set NewGrid.Grid = FG
    
    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
    Ele(4).Visible = ShowTax
   
    '      Set NewGrid.TxtLotNo = Me.TxtLotNo

    Set NewGrid.TxtInvID = Me.XPTxtBillID
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.txtTotal = XPTxtSum
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    Set NewGrid.TxtValueAdded = TxtValueAdded
    Set NewGrid.TxtValueCash = XPTxtValue(0)
    Set NewGrid.TxtValueDelay = XPTxtValue(1)
    Set NewGrid.TxtValuechque = XPTxtValue(2)
    Set NewGrid.txt_Currency_rate = txt_Currency_rate
    Set NewGrid.txt_Currency_rate = txt_Currency_rate
    Set NewGrid.txtItemCodeSearch2 = txtItemCodeSearch2
    Set NewGrid.TxtItemCodeB = TxtItemCodeB
    Set NewGrid.TxtItemCodeB1 = TxtItemCodeB1
    Set NewGrid.Branch = Me.dcBranch
    Set NewGrid.LBLGross = LBLGross
    '--------------------------------------
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.TxtAddTax = Me.TxtTaxAddValue
    Set NewGrid.TxtStampTax = Me.TxtTaxStampValue
    Set NewGrid.TxtServiceTax = Me.TxtTaxServiceValue
    Set NewGrid.TxtShortName = Me.TxtShortName
    Set NewGrid.TxtNots = Me.Text1
    Set NewGrid.Customer = Me.DBCboClientName

    '------------------------------------------------
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.StoreName = Me.DCboStoreName
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.CmdAddSerialLIst = Me.CmdSearch
    Set NewGrid.VatGrid = VatGrid
    'Set NewGrid.CboDiscountType = CboDiscountType
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.TxtCashCustomerName = Me.CashCustomerName
    Set NewGrid.DcboEmp = Me.DcboEmp
    
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
    NewGrid.frmname = "frmsalebill2"
    NewGrid.FillGrid
    
    StrSQL = " select id,code from currency"
 
    fill_combo Me.DcCurrency, StrSQL

    FG.WallPaper = BGround.Picture
    '    AddTip
    XPTab301.CurrTab = 0
    XPDtbBill.value = Date

    If SystemOptions.UserInterface = ArabicInterface Then

        With XPCboDiscountType
            .Clear
            .AddItem "·«ÌÊÃœ Œ’„"
            .AddItem "Œ’„ »ﬁÌ„…"
            .AddItem "Œ’„ »‰”»…"
            .AddItem "Œ’„ ‰ﬁ«ÿ"
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
            .AddItem "«·ÿ«Ê·…"
            .AddItem "ÿ·»«  Œ«—ÃÌ…"
            .AddItem " Œœ„…  Ê’Ì· "
            .AddItem " Œœ„… ”Ì«—«  "
        End With
    
    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With XPCboDiscountType
            .Clear
            .AddItem "No Discount"
            .AddItem "Value Discount"
            .AddItem "Precetage Discount"
            .AddItem "Points"
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
        
        With CboPOSBillType
            .Clear
            .AddItem "table"
            .AddItem "out order"
            .AddItem " car "
            .AddItem "car2 "
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
        .rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.FgCheques
        .rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

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
    
    '----------------------------
    '    Dim rsOut As New ADODB.Recordset
    Dim Msg As String
    '    Set rsOut = New ADODB.Recordset
    '    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '
    '    If Not (rsOut.EOF Or rsOut.BOF) Then
    '
    '        If rsOut!checkout = True Then
    '            StrSQL = "SELECT * FROM Transactions WHERE 1=-1"
    '
    ''            If SystemOptions.usertype <> UserAdminAll Or val(Current_branch) <> 0 Then
    ''                StrSQL = StrSQL & " AND   BranchId=" & Current_branch
    ''            End If
    ''
    ''            StrSQL = StrSQL & " Order by Transaction_ID"
    '
    '        Else
    '
    '            StrSQL = "SELECT * FROM Transactions WHERE 1=-1"
    '
    ''            If SystemOptions.usertype <> UserAdminAll Or val(Current_branch) <> 0 Then
    ''                StrSQL = StrSQL & "  AND   BranchId=" & Current_branch
    ''            End If
    '
    ''            StrSQL = StrSQL & " Order by Transaction_ID"
    '
    '            Set rs = New ADODB.Recordset
    '            rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    '
    '            If Not (rs.BOF Or rs.EOF) Then
    '                rs.MoveLast
    '            End If
    '
    '            Retrive
    '            Me.TxtModFlg.Text = "R"
    '            InvType = 2
    '        End If
    '    End If
    '
    '
    '
    '
    StrSQL = "SELECT * FROM Transactions WHERE 1=-1"
    '

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast
    End If

    '   Retrive
    Me.TxtModFlg.text = "R"
    InvType = 21
 
    btnNew_Click 0
    If CheckPeriodCopon(Me.XPDtbBill.value) = True Then
        TXtCopon.Visible = True
        lbl(94).Visible = True
    Else
        TXtCopon.Visible = False
        lbl(94).Visible = False
    End If
    
    'CheckInputIdle 2
    Exit Sub
ErrTrap:
End Sub
Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i      As Integer
    Dim rs     As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "SELECT    dbo.TblPaymentType.IsDefault, dbo.TblPaymentType.PaymentID, dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId, dbo.TblPaymentType.Accountsus, "
    My_SQL = My_SQL & "  dbo.TblPaymentType.Accountcom, dbo.TblPaymentType.commision, dbo.TblPaymentType.PaymentNamee, dbo.BanksData.Account_Code AS bankAccount_Code"
    My_SQL = My_SQL & " FROM         dbo.TblPaymentType LEFT OUTER JOIN"
    My_SQL = My_SQL & " dbo.BanksData ON dbo.TblPaymentType.BankId = dbo.BanksData.BankID "
    My_SQL = My_SQL & " where (dbo.TblPaymentType.TypTran=2 or dbo.TblPaymentType.TypTran is null) "
    If SystemOptions.LinkUsersWithPayment = True Then
        My_SQL = My_SQL & " and dbo.TblPaymentType.PaymentID in (SELECT     PaynetID"
        My_SQL = My_SQL & " From dbo.TblPaymentUser"
        My_SQL = My_SQL & " Where (UserID = " & user_id & "))"
    End If
    My_SQL = My_SQL & " order by PaymentID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 2
            rs.MoveFirst
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(1, .ColIndex("PaymentName")) = " ‰ﬁœÌ"
            Else
                .TextMatrix(1, .ColIndex("PaymentName")) = " Cash"
            End If
               
            .TextMatrix(1, .ColIndex("PaymentID")) = 0
            Dim IsDefSet As Boolean
            For i = 2 To .rows - 1
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
                'IsDefSet = IIf(IsNull(rs!IsDefault & ""), False, IIf(rs!IsDefault, True, False))
                .TextMatrix(1, .ColIndex("IsDefault")) = 0
                If Not IsNull(rs("IsDefault").value) Then
                    If (rs!IsDefault) Then
                        .TextMatrix(1, .ColIndex("IsDefault")) = 1
                    
                    End If
                Else
                    .TextMatrix(1, .ColIndex("IsDefault")) = 0
                End If
                If .TextMatrix(1, .ColIndex("IsDefault")) = "1" Then
                    If Not IsDefSet Then
                        IsDefSet = True
                        .TextMatrix(i, .ColIndex("Value")) = TxtNetValue
                        Grid_AfterEdit i, .ColIndex("Value")
                    End If
                End If
              
                rs.MoveNext
            Next

            rs.Close
        End If

        '      .RowHeight(-1) = 300
    End With

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
            
            
              
            Me.btnNew(1).Enabled = False
            Me.btnNew(0).Enabled = True
            'Me.btnEdit.Enabled = True
             
            
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
            XPChkPayType(2).Enabled = False
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
                
                
          'Me.btnEdit.Enabled = False
              
             
             
            End If
        
            CboPayMentType.locked = True
            DtpDelayDate.Enabled = False

            If Not m_Menu1 Is Nothing Then
                m_Menu1.Enabled = False
            End If

            CmdINSTALLMENT.Enabled = False
            CmdCheque.Enabled = False

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
            Me.Cmd(7).Enabled = False
            Me.Cmd(9).Enabled = False
            Me.DcboEmp.Enabled = True

 
            Me.btnNew(1).Enabled = True
            Me.btnNew(0).Enabled = False
            'Me.btnEdit.Enabled = False
            
            
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
            FG.rows = FG.FixedRows
            FG.rows = 2
            XPCboDiscountType.locked = False
            Me.XPDtbBill.Enabled = True
            XPDtbBill.value = Date
            Me.DBCboClientName.locked = False
            CboPayMentType.locked = False
            Me.DCboStoreName.locked = False
            Me.XPTxtDiscountVal.locked = False
        
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
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
            m_Menu1.Enabled = True
            DtpDelayDate.value = Date
       
            CmdINSTALLMENT.Enabled = False
            CmdCheque.Enabled = False
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
            
             
            Me.btnNew(1).Enabled = True
            Me.btnNew(0).Enabled = False
            'Me.btnEdit.Enabled = False
            
            
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

            If Me.XPChkPayType(2).value = vbChecked Then
                CmdCheque.Enabled = True
            Else
                CmdCheque.Enabled = False
            End If

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
 
Private Sub CreateRecieveVoucher()
    ' On Error GoTo errortrap

On Error GoTo ErrTrap
IsVouc = False

    Dim UnitID As Long
    Dim i As Long

    If CboRetrunType.ListIndex = 1 Then

        With FG

            For i = 1 To FG.rows - 1

                If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
                    UnitID = IIf(FG.cell(flexcpData, i, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, i, FG.ColIndex("UnitID"))))

                    If val(FG.TextMatrix(i, FG.ColIndex("ItemCostPrice"))) = 0 Then
                                               
                        '       If ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod) = 0 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ  ﬂ·›Â «·»Ì⁄ ·Â Ê·„ Ì „  ÕœÌœ À„‰ «·‘—«¡ Ê·Ì” ·Â ﬁÌ„Â —’Ìœ «›  «ÕÌ… ·–·ﬂ ·« Ì„ﬂ‰ «‰‘«¡ ”‰œ «·«÷«›Â "
                        Else
                            MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                        End If

                        GoTo ErrTrap
                    End If
                End If

            Next i

        End With

    End If

    Dim groupAccount  As String

    If detect_inventory_work_type = 3 Then
   
        With FG

            For i = 1 To FG.rows - 1

                If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
                
                    ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                    groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                    If groupAccount = "Error" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                        Else
                            MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                        End If

                        On Error GoTo ErrTrap
                    End If
                End If

            Next i

        End With

    End If

    'CurrentVoucherNo = GetVoucherGLNO(Val(Text1.text))
    '  DeleteTransactiomsVoucher val(Text1.text)
    '   Dim RowNum As Integer

    '   For RowNum = 1 To Fg.Rows - 1
    '                    If Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) <> "" Then
    '
    '                     If CboRetrunType.ListIndex = 0 Then '„ﬁÌœ »›« Ê—…
    '
    ''
    '                 Else '€Ì— „ﬁÌœ »›« Ê—…
    '                     unitid = IIf(Fg.Cell(flexcpData, RowNum, Fg.ColIndex("UnitID")) = "", Null, (Fg.Cell(flexcpData, RowNum, Fg.ColIndex("UnitID"))))
    '                    Fg.TextMatrix(RowNum, Fg.ColIndex("ItemCostPrice")) = ModItemCostPrice.GetCostItemPrice(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , unitid)
    '
    '                 End If
    '
    '            End If
 
    'Next RowNum
 
    Dim MYWAER As String
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim MYinvnum As String
    Dim note_id As Long

    Dim RSTransDetails As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
 
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

  '  rs.Close
 
    Dim xyeas As Boolean
    xyeas = True

    If xyeas = True Then
 
        MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=20"))
        'mytext = TxtTransSerial.text

        '         rs!nots = mytext
        '         rs.update

        Dim Transaction_ID As Long
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
        Dim TxtNoteSerialV As String
            
        my_branch = Me.dcBranch.BoundText

        If TxtNoteSerialV = "" Then
            If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": GoTo ErrTrap
                
            Else
                       
                If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  "
                    GoTo ErrTrap
                    
                Else
                    TxtNoteSerialV = Notes_coding(val(my_branch), XPDtbBill.value)
                End If
            End If
        End If

        If TxtNoteSerial1V = "" Then
        TxtNoteSerial1V = Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20, , val(DCboStoreName.BoundText), , , , val(DCboUserName.BoundText))
        
            If TxtNoteSerial1V = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  «÷«›Â ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": GoTo ErrTrap
            Else
                       
                If TxtNoteSerial1V = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ «·«” ·«„ ÌœÊÌ« ﬂ„« Õœœ   ": GoTo ErrTrap
                Else
                 '   TxtNoteSerial1V = Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20)
                End If
            End If
        End If
                 
        If Trim(CurrentVoucherNo) <> "" And DateChanged <> True Then
            TxtNoteSerialV = CurrentVoucherNo '—ﬁ„ «·ﬁÌœ
            TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
        End If
           
        Dim sql As String
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        Text1.text = Transaction_ID

        sql = "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,nots2,NoteSerial,NoteSerial1,NoteId,BranchId,Closed)SELECT " & Transaction_ID & "," & MYTEXT & ",Transaction_Date,Transaction_Type = 20,CusID,StoreID,UserID,Emp_ID,nots=" & val(XPTxtBillID.text) & ",nots2='" & TxtNoteSerial1.text & "' ,NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId,1From Transactions Where Transaction_ID =" & val(XPTxtBillID.text) & " And Transaction_Type =9"
        Cn.Execute sql
        '
        Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,ClassId,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ProductionDate,ExpiryDate,LotNO)SELECT round (costprice,2),guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, price , ColorID,ItemSize,ClassId,UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ProductionDate,ExpiryDate,LotNO  From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text
        Text1.text = Transaction_ID
        'TxtIssueSerial.text = TxtNoteSerial1V
        'Create big notes
     
        Set RsNotesGeneral = New ADODB.Recordset
'        RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

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
        RsNotesGeneral("NoteType").value = 160
        RsNotesGeneral("Note_Value").value = Null
        RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
        'RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        RsNotesGeneral("Remark").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        
        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(9) '«–‰ wvt
        RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
        RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
        RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        RsNotesGeneral.update
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

        CREATE_VOUCHER_GEx Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, val(Me.dcBranch.BoundText)
        sql = " update Transactions Set NoteID = " & general_noteid & " where Transaction_ID = " & Transaction_ID
        
        Cn.Execute sql
    End If
 
    '
 
    IsVouc = True
    Exit Sub
ErrTrap:
IsVouc = False

End Sub





Function CREATE_VOUCHER_GEx(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchID As Integer)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim LngCurItemID As Double
    Dim LngUnitID As Long
    Dim UnitFactor As Double

    Dim TOTAL_COST As Variant

    With FG

        For i = 1 To FG.rows - 1
            LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
            LngUnitID = val(FG.cell(flexcpData, i, FG.ColIndex("UnitID")))
            
            GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
                TOTAL_COST = TOTAL_COST + (FG.TextMatrix(i, FG.ColIndex("Count")) * FG.TextMatrix(i, FG.ColIndex("ItemCostPrice")))
            End If

        Next i

    End With

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
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

            StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            StrTempDes = "”‰œ «÷«›Â  —ﬁ„ " & Me.TxtNoteSerial1.text & "»‰«¡ ⁄·Ï „—œÊœ«  „»Ì⁄« "
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")
            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "”‰œ    «÷«›Â —ﬁ„ " & TxtNoteSerial1V & "»‰«¡ ⁄·Ï „—œÊœ«  „»Ì⁄«  »—ﬁ„ " & Me.TxtNoteSerial1.text
            Else
                StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V & "»‰«¡ ⁄·Ï „—œÊœ«  „»Ì⁄«  »—ﬁ„ " & Me.TxtNoteSerial1.text
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With FG

                For i = 1 To FG.rows - 1

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

                        line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "”‰œ    «÷«›Â —ﬁ„ " & TxtNoteSerial1V & "»‰«¡ ⁄·Ï „—œÊœ«  „»Ì⁄«  »—ﬁ„ " & Me.TxtNoteSerial1.text
                        Else
                            StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V & "»‰«¡ ⁄·Ï „—œÊœ«  „»Ì⁄«  »—ﬁ„ " & Me.TxtNoteSerial1.text
                        End If
            
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·œ«∆‰
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

                StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄« 
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ    «÷«›Â —ﬁ„ " & TxtNoteSerial1V & "»‰«¡ ⁄·Ï „—œÊœ«  „»Ì⁄«  »—ﬁ„ " & Me.TxtNoteSerial1.text
                Else
                    StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
         
            ElseIf detect_inventory_work_type = 3 Then

                With FG

                    For i = 1 To FG.rows - 1

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

                            line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 1, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "”‰œ    «÷«›Â —ﬁ„ " & TxtNoteSerial1V & "»‰«¡ ⁄·Ï „—œÊœ«  „»Ì⁄«  »—ﬁ„ " & Me.TxtNoteSerial1.text
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
        End If
    End If

    Dim StrSQL  As String
    StrSQL = "UPDATE Transactions SET NOTS=" & val(Me.Text1.text) & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
    'sql = "update transactions set closed=1" & ",nots=" & Val(Me.XPTxtBillID.text) & ",nots2=" & Me.TxtNoteSerial1.text & " where  Transaction_ID= " & Val(Me.Text1.text)
    Cn.Execute StrSQL
    IsVouc = True
    Exit Function
ErrTrap:
IsVouc = False
End Function


Function SaveItemsData()
    Dim RsgGrantee    As New ADODB.Recordset
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    Dim AllDes As String
    Dim RowNum As Integer
    Dim StrSQL As String
    strFilterText = ","
    Set RsgGrantee = New ADODB.Recordset
    Cn.Execute "delete ItemsDetails   where Transaction_ID= " & val(Me.XPTxtBillID.text)
    
  '  RsgGrantee.Open "TBLRegularMaint", Cn, adOpenStatic, adLockOptimistic, adCmdTable

   StrSQL = "SELECT    * from  ItemsDetails Where (1 = -1)"
   RsgGrantee.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
 
    Dim strFilterText1 As String
      Dim UnitName As String
    Dim ttypename As String
     Dim typename As String
 
 
 
 
    Dim inty As Integer
    Dim intervalstr As String
Dim Name As String
Dim NameE As String
Dim Remarks As String
Dim NooFRows As Double
    
     Dim astrSplitItems1() As String
 
    strFilterText = "&&"
         strFilterText1 = "@@"
     
    For RowNum = 1 To FG.rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            If FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) <> "" Then
            
    RsgGrantee.AddNew
              RsgGrantee("ParrtNoCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode"))
            RsgGrantee("count").value = FG.TextMatrix(RowNum, FG.ColIndex("Count"))
                   RsgGrantee("unitid").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
          RsgGrantee("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            RsgGrantee("sizeid").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            RsgGrantee("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
  
          RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.text)
                    RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
  RsgGrantee("ItemDetailedCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode"))
        RsgGrantee("EffectN").value = -1
        If optsale(1).value = True Then   ' return sallimng
                                                     RsgGrantee("EffectN").value = 1
                                                     
                                                End If
                                                
               
                         
                         RsgGrantee.update
                  
                   
                    
            End If

        End If

    Next RowNum

End Function
Public Sub RetrivePending(Optional Lngid As Long = 0, _
                   Optional NoteSerial1 As String)
                  Cmd_Click (0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    Dim sql As String
    Dim Rs3 As ADODB.Recordset
    
    Dim i As Long

     On Error GoTo ErrTrap

       ' Gridgurantee.Clear flexClearScrollable, flexClearEverything
        
    With Me.FgInstallments
        .Clear flexClearScrollable, flexClearEverything
        .rows = .FixedRows
        LblPrecenType.Caption = ""
        LblPrecenValue.Caption = ""
        LblInstallTotal.Caption = ""
        LblInstallCount.Caption = ""
        LblFirstInstallDate.Caption = ""
        LblInstallmentType.Caption = ""
    End With
    
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
   sql = " SELECT    *"
  sql = sql & "  From dbo.Transactions"
  sql = sql & " Where (Transaction_Type = 70) And (Transaction_ID = " & Lngid & ")"
  sql = sql & "  ORDER BY Transaction_ID"
Set Rs3 = New ADODB.Recordset
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '---------------------------------------------
    If Rs3.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    ' »Ì«‰«  ÃœÌœ…
   ' txtAdvPay.Text = IIf(IsNull(Rs3("AdvPay").value), 0, Rs3("AdvPay").value)
    Me.DCPaymentNet.BoundText = IIf(IsNull(Rs3("PaymentNetid").value), "", Rs3("PaymentNetid").value)
    TxtNetValue.text = IIf(IsNull(Rs3("NetValue").value), "", (Rs3("NetValue").value))
    TxtPayedValue.text = IIf(IsNull(Rs3("PayedValue").value), "", (Rs3("PayedValue").value))
    TxtRemainValue.text = IIf(IsNull(Rs3("RemainValue").value), "", (Rs3("RemainValue").value))
   ' TxtLotNo.Text = IIf(IsNull(Rs3("lotNo").value), "", (Rs3("lotNo").value))
 
    TxtManualNo1.text = IIf(IsNull(Rs3("ManualNo1").value), "", (Rs3("ManualNo1").value))
    TxtManualNo2.text = IIf(IsNull(Rs3("ManualNo2").value), "", (Rs3("ManualNo2").value))
 
'   CBoBasedON.ListIndex = IIf(IsNull(Rs3("CBoBasedON").value), 0, (Rs3("CBoBasedON").value))
 '  Me.DCExtraAccount.BoundText = IIf(IsNull(Rs3("ExtraAccount").value), "", Rs3("ExtraAccount").value)

   ' If Me.DCExtraAccount.BoundText = "" Then
   '     TxtExtraValue.Text = 0
   ' Else
   '     TxtExtraValue.Text = IIf(IsNull(Rs3("ExtraValue").value), 0, (Rs3("ExtraValue").value))
   ' End If

 
 
 
    '‰ﬁ«ÿ «·»Ì⁄
    If Not IsNull(Rs3("POSBillType").value) Then
        CboPOSBillType.ListIndex = val(Rs3("POSBillType").value)
        LblStableID.Caption = IIf(IsNull(Rs3("STableID").value), -1, (Rs3("STableID").value))

    Else
        CboPOSBillType.ListIndex = -1
        LblStableID.Caption = -1

    End If
 CboPOSBillType.ListIndex = 1
   ' Me.DCCar.BoundText = IIf(IsNull(Rs3("CarId").value), "", Rs3("CarId").value)
  '  Me.DCDriver.BoundText = IIf(IsNull(Rs3("DriverId").value), "", Rs3("DriverId").value)

  '  lblSessionD.Caption = IIf(IsNull(Rs3("SessionD").value), -1, (Rs3("SessionD").value))

    If Not IsNull(Rs3("BillBasedOn").value) Then

        If Rs3("BillBasedOn").value = 0 Then
            BillBasedOn(0).value = True
            '   BillBasedOn_Click (0)
        ElseIf Rs3("BillBasedOn").value = 1 Then
            BillBasedOn(1).value = True
            '      BillBasedOn_Click (1)
        ElseIf Rs3("BillBasedOn").value = 2 Then
            BillBasedOn(2).value = True
            '      BillBasedOn_Click (2)
        End If
    
    Else

        BillBasedOn(0).value = True
        '  BillBasedOn_Click (0)
    End If
'DCCustomerLocation.BoundText = IIf(IsNull(Rs3("CustomerlocationID").value), "", Rs3("CustomerlocationID").value)

    dcBranch.BoundText = IIf(IsNull(Rs3("BranchId").value), "", Rs3("BranchId").value)
    DCDocTypes.BoundText = IIf(IsNull(Rs3("Doctype").value), "", Rs3("Doctype").value)
    Me.DcCurrency.BoundText = IIf(IsNull(Rs3("Currency_id").value), "", Rs3("Currency_id").value)
    txt_Currency_rate.text = IIf(IsNull(Rs3("Currency_rate").value), 1, (Rs3("Currency_rate").value))
 
    Me.TxtNoteSerial.text = IIf(IsNull(Rs3("NoteSerial").value), "", (Rs3("NoteSerial").value))

    Me.TxtNoteSerial1.text = IIf(IsNull(Rs3("NoteSerial1").value), "", (Rs3("NoteSerial1").value))
    
 ' If SystemOptions.UserInterface = ArabicInterface Then
 '             LblTitle.Caption = " ›« Ê—… —ﬁ„ : " & Me.TxtNoteSerial1.Text
 '   Else
 '            LblTitle.Caption = "Invoice NO : " & Me.TxtNoteSerial1.Text
 '   End If
 '
        
  '  DCPreFix.Text = IIf(IsNull(Rs3("Prefix").value), "", Rs3("Prefix").value)
    Me.oldtxtNoteSerial1.text = IIf(IsNull(Rs3("OldNoteSerial1").value), IIf(IsNull(Rs3("NoteSerial1").value), "", Rs3("NoteSerial1").value), Rs3("OldNoteSerial1").value)

    lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    Me.TXTNoteID.text = IIf(IsNull(Rs3("NoteID").value), "", (Rs3("NoteID").value))
    Text1.text = IIf(IsNull(Rs3("NotS").value), "", (Rs3("NotS").value))

    XPTxtBillID.text = IIf(IsNull(Rs3("Transaction_ID").value), "", val(Rs3("Transaction_ID").value))
    TxtPhone(3).text = IIf(IsNull(Rs3("Transaction_ID").value), "", val(Rs3("Transaction_ID").value))

    TxtTransSerial.text = IIf(IsNull(Rs3("Transaction_Serial").value), "", Rs3("Transaction_Serial").value)
    XPDtbBill.value = IIf(IsNull(Rs3("Transaction_Date").value), "", (Rs3("Transaction_Date").value))
  
   ' DpFrom.value = IIf(IsNull(Rs3("fromdate").value), XPDtbBill.value, (Rs3("fromdate").value))
   ' DpTo.value = IIf(IsNull(Rs3("todate").value), XPDtbBill.value, (Rs3("todate").value))
    
    
    XPCboDiscountType.ListIndex = IIf(IsNull(Rs3("Trans_DiscountType").value), -1, val(Rs3("Trans_DiscountType").value))
    CboPayMentType.ListIndex = IIf(IsNull(Rs3("PaymentType").value), 0, Rs3("PaymentType").value)
    XPTxtDiscountVal.text = IIf(IsNull(Rs3("Trans_Discount").value), "", (Rs3("Trans_Discount").value))
   ' '''/////////////02 03 2016
  '  LawFirmValue.Text = IIf(IsNull(Rs3("LawFirmValue").value), "", Rs3("LawFirmValue").value)
  '  Sandts.Text = IIf(IsNull(Rs3("Sandts").value), "", Rs3("Sandts").value)
  '  TotalQest.Text = IIf(IsNull(Rs3("TotalQest").value), 0, Rs3("TotalQest").value)
  '  QstValue.Text = IIf(IsNull(Rs3("QstValue").value), 0, Rs3("QstValue").value)
  '  QstNo.Text = IIf(IsNull(Rs3("QstNo").value), 0, Rs3("QstNo").value)
  '  QestStartDateH.value = IIf(IsNull(Rs3("QestStartDateH").value), ToHijriDate(Date), Rs3("QestStartDateH").value)
  '  QestEndtDateH.value = IIf(IsNull(Rs3("QestEndtDateH").value), ToHijriDate(Date), Rs3("QestEndtDateH").value)
  '  QestStartDate.value = IIf(IsNull(Rs3("QestStartDate").value), Date, Rs3("QestStartDate").value)
  '  QestEndtDate.value = IIf(IsNull(Rs3("QestEndtDate").value), Date, Rs3("QestEndtDate").value)
  '  ''///////////////
    Me.DBCboClientName.BoundText = IIf(IsNull(Rs3("CusID").value), "", Rs3("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(Rs3("UserID").value), "", Rs3("UserID").value)
    FG.Clear flexClearScrollable, flexClearEverything
    Me.DCboStoreName.BoundText = IIf(IsNull(Rs3("StoreID").value), "", Rs3("StoreID").value)
    Me.DcboEmp.BoundText = IIf(IsNull(Rs3("Emp_ID").value), "", Rs3("Emp_ID").value)
    XPTxtTaxValue.text = IIf(IsNull(Rs3("TaxValue").value), "", (Rs3("TaxValue").value))
    XPChkTAX.value = IIf(Rs3("TaxFound") = True, Checked, Unchecked)
    'Text1.text = IIf(IsNull(rs3("nots2").value), "", (rs3("nots2").value))
    Me.TXTOrDer_no.text = IIf(IsNull(Rs3("order_no").value), "", (Rs3("order_no").value))
  '  Me.DCGroupID.BoundText = IIf(IsNull(Rs3("LocationID").value), "", Rs3("LocationID").value)
   ' TxtPurchaseBill.Text = IIf(IsNull(Rs3("PurchaseBill").value), "", (Rs3("PurchaseBill").value))
 
    If IsNull(Rs3("BoxID").value) Then
        Me.DcboBox.BoundText = ""
    Else
        Me.DcboBox.BoundText = IIf(IsNull(Rs3("BoxID").value), "", Rs3("BoxID").value)
    End If

    If IsNull(Rs3("SaleType").value) Then
        Me.CboSaleType.ListIndex = 0
    Else
        Me.CboSaleType.ListIndex = IIf(Rs3("SaleType").value = 0, 0, 1)
    End If
    TxtPhone(1).text = IIf(IsNull(Rs3("VATNO").value), "", (Rs3("VATNO").value))
    If Not (IsNull(Rs3("CashCustomerPhone").value)) Then
        Me.TxtPhone(0).text = Rs3("CashCustomerPhone").value
    Else
        Me.TxtPhone(0).text = ""
    End If


    If Not (IsNull(Rs3("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.text = Rs3("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.text = ""
    End If


    'ChkInstall 11 10 2012
    If IsNull(Rs3("ChkInstall").value) Then
        Me.ChkInstall.value = vbUnchecked
    Else
        Me.ChkInstall.value = IIf(Rs3("ChkInstall").value = 0, vbUnchecked, vbChecked)
    End If

    '÷—»Ì… «·Œ’„ Ê«·≈÷«›…
    If Not IsNull(Rs3("TaxAddValue").value) Then
        If Rs3("TaxAddValue").value > 0 Then
            ChkTaxAdd.value = vbChecked
            Me.TxtTaxAddValue.text = Rs3("TaxAddValue").value
        End If
    End If

    '÷—»Ì… «·œ„€…
    If Not IsNull(Rs3("TaxStampValue").value) Then
        If Rs3("TaxStampValue").value > 0 Then
            ChkTaxStamp.value = vbChecked
            Me.TxtTaxStampValue.text = Rs3("TaxStampValue").value
        End If
    End If

    '÷—»Ì… «·Œœ„…
    If Not IsNull(Rs3("TaxServiceValue").value) Then
        If Rs3("TaxServiceValue").value > 0 Then
            ChkTaxSerivce.value = vbChecked
            Me.TxtTaxServiceValue.text = Rs3("TaxServiceValue").value
        End If
    End If

    TxtBillComment.text = IIf(IsNull(Rs3("TransactionComment").value), "", (Rs3("TransactionComment").value))
    ''//26 05 2015
   ' Me.txtManualNO.Text = IIf(IsNull(Rs3("ManualNO").value), "", (Rs3("ManualNO").value))
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = "SELECT  dbo.RetriveRecivedQty( Transaction_Details.Item_ID,'" & TxtNoteSerial1.text & "')  as resqty2, TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(Rs3("Transaction_ID").value)
    StrSQL = StrSQL + "order by id"

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For i = 1 To RsDetails.RecordCount
            FG.TextMatrix(i, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
            FG.TextMatrix(i, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
            FG.TextMatrix(i, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate").value), "", RsDetails("OrderArrivalDate").value)
            FG.cell(flexcpPicture, i, FG.ColIndex("Ser")) = ""
            FG.cell(flexcpData, i, FG.ColIndex("Ser")) = ""
            FG.TextMatrix(i, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(i, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim$(RsDetails("Item_ID").value))
            FG.TextMatrix(i, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))

            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(i, FG.ColIndex("HaveSerial")) = True

                '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«· «·Œ«’… »«·›« Ê—…
                If (RsDetails("Item_ID")) <> "" And RsDetails("ItemSerial") <> "" Then
                    StrSQL = "select * From ReplacedItems where ReturnID=" & XPTxtBillID.text
                    StrSQL = StrSQL + " and ItemID=" & RsDetails("Item_ID")
                    StrSQL = StrSQL + " and ItemSerial='" & RsDetails("ItemSerial") & "'"
                    Set RsReplace = New ADODB.Recordset
                    RsReplace.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (RsReplace.EOF Or RsReplace.BOF) Then
                        FG.cell(flexcpPicture, i, FG.ColIndex("Ser")) = mdifrmmain.ImgLstTree.ListImages("Request").Picture
                        FG.cell(flexcpData, i, FG.ColIndex("Ser")) = "X"
                    End If
                End If
            End If

            FG.TextMatrix(i, FG.ColIndex("ItemType")) = IIf(IsNull(RsDetails("ItemType").value), "", (RsDetails("ItemType").value))

            If RsDetails("ItemType").value = 1 Then
                FG.cell(flexcpPicture, i, FG.ColIndex("Ser")) = mdifrmmain.ImgLstTree.ListImages("Maintenance").Picture
            
            End If
            FG.TextMatrix(i, FG.ColIndex("ReminRequ")) = IIf(IsNull(RsDetails("ReminRequ")), "", (RsDetails("ReminRequ").value))
            FG.TextMatrix(i, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))
            FG.TextMatrix(i, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
        
       ' Fg.TextMatrix(i, Fg.ColIndex("DeliveredQty2")) = RetriveRecivedQty(Fg.TextMatrix(i, Fg.ColIndex("Code")), TxtNoteSerial1)
        
        
                  '  Fg.TextMatrix(i, Fg.ColIndex("DeliveredQty2")) = IIf(IsNull(RsDetails("resqty2")), 0, (RsDetails("resqty2").value))
        
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
            FG.TextMatrix(i, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(i, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(i, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
            FG.TextMatrix(i, FG.ColIndex("IsExpirDate")) = IIf(IsNull(RsDetails("IsExpirDate")), "", (RsDetails("IsExpirDate").value))
            'NEW DATA
            ' Fg.TextMatrix(i, Fg.ColIndex("LAXIS")) = IIf(IsNull(RsDetails("LAXIS")), "", (RsDetails("LAXIS").value))
           '  Fg.TextMatrix(i, Fg.ColIndex("LCYL")) = IIf(IsNull(RsDetails("LCYL")), "", (RsDetails("LCYL").value))
           '  Fg.TextMatrix(i, Fg.ColIndex("LSPH")) = IIf(IsNull(RsDetails("LSPH")), "", (RsDetails("LSPH").value))
           '  Fg.TextMatrix(i, Fg.ColIndex("RAXIS")) = IIf(IsNull(RsDetails("RAXIS")), "", (RsDetails("RAXIS").value))
           '  Fg.TextMatrix(i, Fg.ColIndex("RCYL")) = IIf(IsNull(RsDetails("RCYL")), "", (RsDetails("RCYL").value))
           '  Fg.TextMatrix(i, Fg.ColIndex("RSPH")) = IIf(IsNull(RsDetails("RSPH")), "", (RsDetails("RSPH").value))
             

            FG.TextMatrix(i, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
        
            FG.TextMatrix(i, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(i, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(i, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
                   If SystemOptions.showcostColorininvoice = True Then

            If val(FG.TextMatrix(i, FG.ColIndex("ItemProfit"))) = 0 Then
                Me.FG.cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = vbYellow
            ElseIf val(FG.TextMatrix(i, FG.ColIndex("ItemProfit"))) < 0 Then
                Me.FG.cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = vbRed
            Else
                Me.FG.cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = 0
            End If

            Else
              Me.FG.cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = 0
            End If


FG.cell(flexcpData, i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
        
            If SystemOptions.UserInterface = ArabicInterface Then
                FG.TextMatrix(i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            Else
                FG.TextMatrix(i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitNamee")), "", (RsDetails("UnitNamee").value))
            End If
   ' Fg.TextMatrix(i, Fg.ColIndex("ItemsDetailsNewidea")) = IIf(IsNull(RsDetails("ItemsDetailsNewidea")), "", (RsDetails("ItemsDetailsNewidea").value))
            FG.TextMatrix(i, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(i, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(i, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
       
            FG.TextMatrix(i, FG.ColIndex("GranteeType")) = IIf(IsNull(RsDetails("GranteeType")), "", (RsDetails("GranteeType").value))
            FG.TextMatrix(i, FG.ColIndex("GranteeStartDate")) = IIf(IsNull(RsDetails("GranteeStartDate")), "", (RsDetails("GranteeStartDate").value))
            FG.TextMatrix(i, FG.ColIndex("GranteeEndDate")) = IIf(IsNull(RsDetails("GranteeEndDate")), "", (RsDetails("GranteeEndDate").value))
            FG.TextMatrix(i, FG.ColIndex("RegularMaintenancedates")) = IIf(IsNull(RsDetails("RegularMaintenancedates")), "", (RsDetails("RegularMaintenancedates").value))
         '   Fg.TextMatrix(i, Fg.ColIndex("RegularMaintenanceIDS")) = IIf(IsNull(RsDetails("RegularMaintenanceIDS")), "", (RsDetails("RegularMaintenanceIDS").value))

            RsDetails.MoveNext
        
            If FG.rows > 10 Then
                If i = 8 Then FG.Refresh
            End If

        Next i

        '----------------------------
        Me.LblInvProfit.Caption = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("ItemProfit"), FG.rows - 1, FG.ColIndex("ItemProfit"))

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


          NewGrid.Calculate 1, , , True
          NewGrid.SentTypeVAT
   
   ' TxtFillData.Text = "F"
    '-----------------------------------------------------------------------------------------------
    Screen.MousePointer = vbDefault

 
    DoEvents
        
    Exit Sub

ErrTrap:
    Resume
    Screen.MousePointer = vbDefault
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

    On Error GoTo ErrTrap
    '---------------------------------------------
    'Here We Reset all Setting

    With Me.FgInstallments
        .Clear flexClearScrollable, flexClearEverything
        .rows = .FixedRows
        LblPrecenType.Caption = ""
        LblPrecenValue.Caption = ""
        LblInstallTotal.Caption = ""
        LblInstallCount.Caption = ""
        LblFirstInstallDate.Caption = ""
        LblInstallmentType.Caption = ""
    End With
    
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

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.Find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then

            With FG
                FG.rows = 1
   
            End With

            Exit Sub
        
        End If
    End If

    If NoteSerial1 <> "" Then

        rs.Find "noteserial1='" & NoteSerial1 & "'", , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.text = "T"
    Screen.MousePointer = vbArrowHourglass
    ' »Ì«‰«  ÃœÌœ…
    Me.DCPaymentNet.BoundText = IIf(IsNull(rs("PaymentNetid").value), "", rs("PaymentNetid").value)
    TxtNetValue.text = IIf(IsNull(rs("NetValue").value), "", (rs("NetValue").value))
    TxtPayedValue.text = IIf(IsNull(rs("PayedValue").value), "", (rs("PayedValue").value))
    TxtRemainValue.text = IIf(IsNull(rs("RemainValue").value), "", (rs("RemainValue").value))
 DefaultInvoicetype.ListIndex = IIf(IsNull(rs("Invoicetype").value), 0, rs("Invoicetype").value)
 Txtcard(0).text = IIf(IsNull(rs("CardId0").value), "", (rs("CardId0").value))
 Txtcard(1).text = IIf(IsNull(rs("CardId1").value), "", (rs("CardId1").value))
 zatcaStatus = IIf(IsNull(rs("zatcaStatus").value), 0, rs("zatcaStatus").value)
    TxtManualNo1.text = IIf(IsNull(rs("ManualNo1").value), "", (rs("ManualNo1").value))
    TxtManualNo2.text = IIf(IsNull(rs("ManualNo2").value), "", (rs("ManualNo2").value))
    'SessionD = IIf(IsNull(rs("SessionD").value), "", (rs("SessionD").value))
 
    '‰ﬁ«ÿ «·»Ì⁄
    If Not IsNull(rs("POSBillType").value) Then
        CboPOSBillType.ListIndex = val(rs("POSBillType").value)
        LblStableID.Caption = IIf(IsNull(rs("STableID").value), -1, (rs("STableID").value))

    Else
        CboPOSBillType.ListIndex = -1
        LblStableID.Caption = -1

    End If
 
    If CboPOSBillType.ListIndex = -1 Then
        LBLTable.Caption = ""
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

    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
    DCDocTypes.BoundText = IIf(IsNull(rs("Doctype").value), "", rs("Doctype").value)
    Me.DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
    txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    TxtPhone(1).text = IIf(IsNull(rs("VATNO").value), "", (rs("VATNO").value))
     If Not (IsNull(rs("CashCustomerPhone").value)) Then
        Me.TxtPhone(0).text = rs("CashCustomerPhone").value
    Else
        Me.TxtPhone(0).text = ""
    End If


    If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.CashCustomerName.text = rs("CashCustomerName").value
    Else
        Me.CashCustomerName.text = ""
    End If


    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)
    TxtPhone(2).text = Me.TxtNoteSerial1.text
    lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))
    Text1.text = IIf(IsNull(rs("NotS").value), "", (rs("NotS").value))

    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))

    TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), -1, val(rs("Trans_DiscountType").value))
    CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
    XPTxtDiscountVal.text = IIf(IsNull(rs("Trans_Discount").value), "", (rs("Trans_Discount").value))
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    FG.Clear flexClearScrollable, flexClearEverything
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    XPTxtTaxValue.text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)
    'Text1.text = IIf(IsNull(rs("nots2").value), "", (rs("nots2").value))
    Me.TXTOrDer_no.text = IIf(IsNull(rs("order_no").value), "", (rs("order_no").value))
    TxtValueAdded.text = IIf(IsNull(rs("VAT").value), 0, (rs("VAT").value))
 
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

    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
    StrSQL = StrSQL + "order by id"

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For i = 1 To RsDetails.RecordCount
            FG.TextMatrix(i, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
            FG.TextMatrix(i, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
            FG.TextMatrix(i, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate").value), "", RsDetails("OrderArrivalDate").value)
            FG.cell(flexcpPicture, i, FG.ColIndex("Ser")) = ""
            FG.cell(flexcpData, i, FG.ColIndex("Ser")) = ""
            FG.TextMatrix(i, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(i, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim$(RsDetails("Item_ID").value))
            FG.TextMatrix(i, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))

            FG.TextMatrix(i, FG.ColIndex("printed")) = IIf(IsNull(RsDetails("printed")), "", Trim(RsDetails("printed").value))
            FG.TextMatrix(i, FG.ColIndex("PrintName")) = IIf(IsNull(RsDetails("PrintName")), "", Trim(RsDetails("PrintName").value))
            
            
            FG.TextMatrix(i, FG.ColIndex("EmpID4")) = IIf(IsNull(RsDetails("EmpID4")), "", (RsDetails("EmpID4").value))
            
            FG.TextMatrix(i, FG.ColIndex("CusID2")) = IIf(IsNull(RsDetails("CusID2")), "", (RsDetails("CusID2").value))
            FG.TextMatrix(i, FG.ColIndex("SupplierID")) = IIf(IsNull(RsDetails("SupplierID")), "", (RsDetails("SupplierID").value))
            
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(i, FG.ColIndex("HaveSerial")) = True

                '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«· «·Œ«’… »«·›« Ê—…
                If (RsDetails("Item_ID")) <> "" And RsDetails("ItemSerial") <> "" Then
                    StrSQL = "select * From ReplacedItems where ReturnID=" & XPTxtBillID.text
                    StrSQL = StrSQL + " and ItemID=" & RsDetails("Item_ID")
                    StrSQL = StrSQL + " and ItemSerial='" & RsDetails("ItemSerial") & "'"
                    Set RsReplace = New ADODB.Recordset
                    RsReplace.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (RsReplace.EOF Or RsReplace.BOF) Then
                        FG.cell(flexcpPicture, i, FG.ColIndex("Ser")) = mdifrmmain.ImgLstTree.ListImages("Request").Picture
                        FG.cell(flexcpData, i, FG.ColIndex("Ser")) = "X"
                    End If
                End If
            End If

            FG.TextMatrix(i, FG.ColIndex("ItemType")) = IIf(IsNull(RsDetails("ItemType").value), "", (RsDetails("ItemType").value))

            If RsDetails("ItemType").value = 1 Then
                FG.cell(flexcpPicture, i, FG.ColIndex("Ser")) = mdifrmmain.ImgLstTree.ListImages("Maintenance").Picture
            
            End If

            FG.TextMatrix(i, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))
            FG.TextMatrix(i, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
        
            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                FG.TextMatrix(i, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Else
                FG.TextMatrix(i, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").value))
            End If
            FG.TextMatrix(i, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(i, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(i, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
            FG.TextMatrix(i, FG.ColIndex("IsExpirDate")) = IIf(IsNull(RsDetails("IsExpirDate")), "", (RsDetails("IsExpirDate").value))
            FG.TextMatrix(i, FG.ColIndex("Vat")) = IIf(IsNull(RsDetails("Vat")), "", (RsDetails("Vat").value))
            FG.TextMatrix(i, FG.ColIndex("Vatyo")) = IIf(IsNull(RsDetails("Vatyo")), "", (RsDetails("Vatyo").value))
            FG.TextMatrix(i, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(i, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(i, FG.ColIndex("guaranteeTime")) = IIf(IsNull(RsDetails("guaranteeTime")), "", (RsDetails("guaranteeTime").value))
            FG.TextMatrix(i, FG.ColIndex("ParrtNoCode")) = IIf(IsNull(RsDetails("ParrtNoCode")), "", (RsDetails("ParrtNoCode").value))
FG.TextMatrix(i, FG.ColIndex("ItemDetailedCode")) = IIf(IsNull(RsDetails("ItemDetailedCode")), "", (RsDetails("ItemDetailedCode").value))

            FG.TextMatrix(i, FG.ColIndex("ItemCostPrice")) = IIf(IsNull(RsDetails("CostPrice")), "", (RsDetails("CostPrice").value))
            FG.TextMatrix(i, FG.ColIndex("PofTransID")) = IIf(IsNull(RsDetails("CostTransID")), "", (RsDetails("CostTransID").value))
            FG.TextMatrix(i, FG.ColIndex("ItemProfit")) = IIf(IsNull(RsDetails("ItemProfit")), "", (RsDetails("ItemProfit").value))
            FG.TextMatrix(i, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
        
            FG.TextMatrix(i, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(i, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(i, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
            If val(FG.TextMatrix(i, FG.ColIndex("ItemProfit"))) = 0 Then
                Me.FG.cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = vbYellow
            ElseIf val(FG.TextMatrix(i, FG.ColIndex("ItemProfit"))) < 0 Then
                Me.FG.cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = vbRed
            Else
                Me.FG.cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = 0
            End If

            FG.cell(flexcpData, i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
        
            If SystemOptions.UserInterface = ArabicInterface Then
                FG.TextMatrix(i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            Else
                FG.TextMatrix(i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitNamee")), "", (RsDetails("UnitNamee").value))
            End If

            RsDetails.MoveNext
        
            If FG.rows > 10 Then
                If i = 8 Then FG.Refresh
            End If

        Next i

        '----------------------------
        Me.LblInvProfit.Caption = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("ItemProfit"), FG.rows - 1, FG.ColIndex("ItemProfit"))

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

    XPChkPayType(0).value = Unchecked
    XPChkPayType(1).value = Unchecked
    XPChkPayType(2).value = Unchecked
    XPTxtValue(0).text = ""
    XPTxtValue(1).text = ""
    XPTxtSerial(0).text = ""
    XPTxtSerial(1).text = ""
    XPTxtValue(1).Tag = ""
    DtpDelayDate.value = Date
    '----------------------------------------------------------------------------------------
    StrSQL = "Select * From Notes Where Transaction_ID=" & val(rs("Transaction_ID").value)
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsNotes.EOF Or RsNotes.BOF) Then

        For i = 1 To RsNotes.RecordCount

            If RsNotes("NoteType").value = 170 Then
                XPChkPayType(0).value = Checked
                XPChkPayType_Click (0)
                XPTxtValue(0).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtSerial(0).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim$(RsNotes("NoteSerial").value))
                Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").value), "", RsNotes("BoxID").value)
            End If

            If RsNotes("NoteType").value = 1 Then
                XPChkPayType(1).value = Checked
                XPChkPayType_Click (1)
                XPTxtValue(1).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").value), "", (RsNotes("NoteID").value))
                XPTxtSerial(1).text = IIf(IsNull(RsNotes("NoteSerial").value), "", (RsNotes("NoteSerial").value))
                DtpDelayDate.value = IIf(IsNull(RsNotes("DueDate").value), "", (RsNotes("DueDate").value))
            End If

            If RsNotes("NoteType").value = 2 Then
                XPChkPayType(2).value = Checked
                XPChkPayType_Click (2)
            End If

            RsNotes.MoveNext
        Next i

    End If

    Set RsNotes = New ADODB.Recordset
    StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.NoteSerial," & "Notes.Note_Value, Notes.BankID,BanksData.BankName , Notes.ChqueNum, Notes.DueDate "
    StrSQL = StrSQL + " FROM Notes INNER JOIN BanksData ON Notes.BankID = BanksData.BankID "
    StrSQL = StrSQL + " Where NoteType=2 AND NOTES.Transaction_ID=" & val(rs("Transaction_ID").value)
    StrSQL = StrSQL + " Order BY Notes.NoteID"
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.FgCheques
        .rows = .FixedRows

        If Not (RsNotes.BOF Or RsNotes.EOF) Then
            .rows = .FixedRows + RsNotes.RecordCount

            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("CheckValue")) = IIf(IsNull(RsNotes("Note_Value").value), "", RsNotes("Note_Value").value)
                .TextMatrix(i, .ColIndex("CheckNumber")) = IIf(IsNull(RsNotes("ChqueNum").value), "", RsNotes("ChqueNum").value)
                .TextMatrix(i, .ColIndex("BankID")) = IIf(IsNull(RsNotes("BankID").value), "", RsNotes("BankID").value)
                .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(RsNotes("BankName").value), "", RsNotes("BankName").value)

                If Not IsNull(RsNotes("DueDate").value) Then
                    .TextMatrix(i, .ColIndex("DueDate")) = DisplayDate(RsNotes("DueDate").value)
                Else
                    .TextMatrix(i, .ColIndex("DueDate")) = ""
                End If

                RsNotes.MoveNext
            Next i

        End If

        .AutoSize 0, .Cols - 1, False
        SumChecks
    End With
   
    TxtFillData.text = "F"
    '-----------------------------------------------------------------------------------------------
    Dim SngRelatedNotesValues As Single
    Me.CmdNotes.Visible = ShowRelatedNotes(val(Me.XPTxtBillID.text), 0, SngRelatedNotesValues)
    Me.CmdNotes.Tag = SngRelatedNotesValues

    SngRelatedNotesValues = 0
    Me.CmdRetruns.Visible = ShowRelatedTransactions(val(Me.XPTxtBillID.text), 0, SngRelatedNotesValues)
    Me.CmdRetruns.Tag = SngRelatedNotesValues

    '-----------------------------------------------------------------------------------------------
    
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    showComm
    FillVoucherGrid
    FillOrderGrid

    '    Else
    '        CmdINSTALLMENT.Enabled = False
    '        CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
    
    '  End If
    'Else
    'FgInstallments.Clear

    '⁄—÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
    If XPTxtValue(1).Tag <> "" Then
        StrSQL = "Select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
        Set RsTest = New ADODB.Recordset
        RsTest.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTest.EOF Or RsTest.BOF) Then
            CmdINSTALLMENT.Enabled = True
            CmdINSTALLMENT.Caption = "⁄—÷ «·√ﬁ”«ÿ «·„”Ã·…"
            LngPartID = RsTest("PartID").value
            Me.LblPrecenType.Tag = RsTest("InterestType").value

            If RsTest("InterestType").value = 0 Then
                LblPrecenType.Caption = "‰”»… „∆ÊÌ…"
            ElseIf RsTest("InterestType").value = 1 Then
                LblPrecenType.Caption = "ﬁÌ„… À«» …"
            ElseIf RsTest("InterestType").value = 2 Then
                LblPrecenType.Caption = "·«ÌÊÃœ"
            End If

            Me.LblPrecenValue.Caption = RsTest("InterestVal").value
            'LblDiscount.Caption = IIf(IsNull(RsTest("Discount").value), "", RsTest("Discount").value)
            'Me.LblAdvPayment.Caption = IIf(IsNull(RsTest("AdvPayment").value), "", RsTest("AdvPayment").value)
        
            Me.LblInstallTotal.Caption = RsTest("Total").value
            Me.LblInstallCount.Caption = RsTest("InstallCount").value
            Me.LblFirstInstallDate.Caption = DisplayDate(RsTest("FirstInstallDate").value)
            Me.LblInstallmentType.Tag = RsTest("InstallmentType").value

            If RsTest("InstallmentType").value = 0 Then
                LblInstallmentType.Caption = "ÌÊ„"
            ElseIf RsTest("InstallmentType").value = 1 Then
                LblInstallmentType.Caption = "‘Â—"
            ElseIf RsTest("InstallmentType").value = 2 Then
                LblInstallmentType.Caption = "”‰…"
            End If

            Me.LblInstallSeprator.Caption = RsTest("InstallSeprator").value
            Me.LblStartValue.Caption = IIf(IsNull(RsTest("StartValue").value), "", RsTest("StartValue").value)
            LblDiscount.Caption = IIf(IsNull(RsTest("Discount").value), "", RsTest("Discount").value)
            Me.LblAdvPayment.Caption = IIf(IsNull(RsTest("AdvPayment").value), "", RsTest("AdvPayment").value)
        
            Set RsPartDetails = New ADODB.Recordset
            StrSQL = "Select * From InstallMentDetails Where PartID=" & LngPartID
            RsPartDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            'fill Installments Grid
            If Not (RsPartDetails.BOF Or RsPartDetails.EOF) Then
                RsPartDetails.MoveFirst

                With Me.FgInstallments
                    .rows = .FixedRows + RsPartDetails.RecordCount

                    For i = .FixedRows To .rows - 1
                        .TextMatrix(i, .ColIndex("QestID")) = IIf(IsNull(RsPartDetails("QestID").value), "", RsPartDetails("QestID").value)
                        .TextMatrix(i, .ColIndex("Serial")) = IIf(IsNull(RsPartDetails("QeqtNum").value), "", RsPartDetails("QeqtNum").value)
                        .TextMatrix(i, .ColIndex("QeqtNum")) = IIf(IsNull(RsPartDetails("QeqtNum").value), "", RsPartDetails("QeqtNum").value)
                    
                        .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(RsPartDetails("Value").value), "", RsPartDetails("Value").value)

                        If Not IsNull(RsPartDetails("DueDate").value) Then
                            .TextMatrix(i, .ColIndex("Due_Date")) = DisplayDate(RsPartDetails("DueDate").value)
                        Else
                            .TextMatrix(i, .ColIndex("Due_Date")) = ""
                        End If

                        RsPartDetails.MoveNext
                    Next i

                End With

            End If

            showComm
        Else
            CmdINSTALLMENT.Enabled = False
            CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
    
        End If

'
         'NewGrid.Calculate 1, , , True
         ' NewGrid.SentTypeVAT
          
    End If
  '  RetriveValueAdded
         NewGrid.Calculate 1, , , True
       NewGrid.SentTypeVAT
          
    '›« Ê—… «·Œœ„« 
    If CheckBillType = 0 Then
        Command2.backcolor = &HC0C0C0
        Command2.Enabled = False

        If SystemOptions.UserInterface = ArabicInterface Then
            Command2.Caption = "  ›« Ê—… Œœ„«  Ê·Ì” ·Â« ”‰œ ’—› "
        Else
            Command2.Caption = " Services Invoices"
        
        End If

        Exit Sub

    End If

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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
Else

   Msg = "Undo Invoice Register"
            Msg = Msg & CHR(13) & "are you sure? ..!!"
            
End If
            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
Else
   Msg = "Undo Invoice Register"
            Msg = Msg & CHR(13) & "are you sure? ..!!"
End If

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                rs.Find "Transaction_ID='" & val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst

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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
            MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If Me.CboPayMentType.ListIndex = 0 Then

        '›« Ê—… ‰ﬁœÌ…
        If CheckBoxAccount(val(Me.DcboBox.BoundText), val(Me.XPTxtValue(0).text), XPDtbBill.value, False) = False Then
            Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
            Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï Õ”«»«  «·Œ“‰…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
 
    IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

    If IntRes = vbYes Then
        If Not rs.RecordCount < 1 Then
            Cn.BeginTrans
            BegainTrans = True
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & rs("Transaction_ID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
             StrSQL = "Delete From TblSalesPoints Where TransID=" & val(Me.XPTxtBillID.text)
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
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            rs.MoveFirst

            If rs.RecordCount < 1 Then
                clear_all Me
                TxtModFlg_Change
                XPTxtCurrent.Caption = 0
                XPTxtCount.Caption = 0
                VatGrid.Clear flexClearScrollable, flexClearEverything
                VatGrid.rows = 1
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
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title

    If BegainTrans = True Then
        rs.CancelUpdate
        Cn.RollbackTrans
        BegainTrans = False
    End If

End Sub
Sub RetriveValueAdded()
Dim sql As String
Dim i As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
    VatGrid.Clear flexClearScrollable, flexClearEverything
    VatGrid.rows = 1
sql = " SELECT     dbo.TransactionValueAdded.Transaction_Type, dbo.TransactionValueAdded.Transaction_ID, dbo.TransactionValueAdded.Vat, dbo.TransactionValueAdded.Vatyo,"
sql = sql & " dbo.TransactionValueAdded.ItemID , dbo.TblItems.itemname, dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee ,dbo.TransactionValueAdded.selectd ,dbo.TransactionValueAdded.Valu "
sql = sql & " FROM         dbo.TransactionValueAdded LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems ON dbo.TransactionValueAdded.ItemID = dbo.TblItems.ItemID"
sql = sql & " Where (dbo.TransactionValueAdded.Transaction_Type = 21) And (dbo.TransactionValueAdded.Transaction_ID = " & val(TxtInvID.text) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With Me.VatGrid
rs2.MoveFirst
.rows = .rows + rs2.RecordCount
For i = 1 To .rows - 1
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

Function Getmin(discValue As Double)
Dim RowNum As Double
Dim MinId As Double
 Dim MinValue As Double
 Dim LineID As Double
 MinId = 1
 RowNum = 1
 LineID = 1
MinValue = FG.TextMatrix(1, FG.ColIndex("Valu"))
If lbl(57).Visible = True Then
             For RowNum = 1 To FG.rows - 1
            
                        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(RowNum, FG.ColIndex("SpecialOffer"))) = 1 And FG.TextMatrix(RowNum, FG.ColIndex("MinID")) <> "1" Then
                             
                             If FG.TextMatrix(RowNum, FG.ColIndex("Valu")) <= MinValue Then
                                 MinId = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                                 MinValue = FG.TextMatrix(RowNum, FG.ColIndex("Valu"))
                               LineID = RowNum
                            End If
                        
                        End If
            Next RowNum

End If
   FG.TextMatrix(LineID, FG.ColIndex("MinID")) = 1
'Getmin = MinID

    With FG
          If discValue = 100 Then '„Ã«‰Ì
              .cell(flexcpData, LineID, .ColIndex("DiscountType")) = 4
                                          .TextMatrix(LineID, .ColIndex("DiscountType")) = 4
                                          .TextMatrix(LineID, .ColIndex("DiscountVal")) = 0
          Else
              .cell(flexcpData, LineID, .ColIndex("DiscountType")) = 2
                                          .TextMatrix(LineID, .ColIndex("DiscountType")) = 2
                                          .TextMatrix(LineID, .ColIndex("DiscountVal")) = discValue
          
         End If
     End With
End Function
Private Sub SavePoints(Optional ItemID As Double, Optional Price As Double)
Dim sql As String
Dim GroupID As Double
Dim NoPoint As Double
Dim Equation1 As Double
Dim Equation2 As Double
Dim Balance As Double
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select * from TblSalesPoints where 1=-1"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
rs2.AddNew
rs2("TransID").value = val(Me.XPTxtBillID.text)
rs2("RecordDate").value = XPDtbBill.value
GetGroupInfor ItemID, GroupID, Equation1, Equation2
rs2("GroupID").value = GroupID
rs2("ItemID").value = ItemID
NoPoint = Price * Equation1
Balance = NoPoint * Equation2
Balance = Round(Balance, 2)
rs2("Price").value = Price
rs2("NoPoint").value = NoPoint
If optsale(0).value = True Then
rs2("DebtValue").value = 0
rs2("CreditValue").value = Balance
Else
rs2("DebtValue").value = Balance
rs2("CreditValue").value = 0
End If
rs2("Balance").value = Balance
rs2("Mobile").value = TxtPhone(0).text
rs2.update
End Sub
Private Sub SavePointsPayd()
Dim sql As String
Dim GroupID As Double
Dim NoPoint As Double
Dim Equation1 As Double
Dim Equation2 As Double
Dim Balance As Double
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select * from TblSalesPoints where 1=-1"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
rs2.AddNew
rs2("TransID").value = val(Me.XPTxtBillID.text)
rs2("RecordDate").value = XPDtbBill.value
If val(XPCboDiscountType.ListIndex) = 3 Then
If optsale(0).value = True Then
rs2("DebtValue").value = val(XPTxtDiscountVal.text)
rs2("CreditValue").value = 0
Else
rs2("DebtValue").value = 0
rs2("CreditValue").value = val(XPTxtDiscountVal.text)
End If
End If
rs2("Balance").value = val(XPTxtDiscountVal.text)
rs2("Mobile").value = TxtPhone(0).text
rs2.update
End Sub
Sub GetGroupInfor(Optional ItemID As Double, Optional ByRef GroupID As Double, Optional ByRef Equation1 As Double, Optional ByRef Equation2 As Double)
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     dbo.Groups.GroupID,  dbo.Groups.Equation1,"
sql = sql & "                      dbo.Groups.Equation2"
sql = sql & " FROM         dbo.Groups RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID"
sql = sql & " Where (dbo.TblItems.ItemID = " & ItemID & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GroupID = IIf(IsNull(Rs3("GroupID").value), 0, Rs3("GroupID").value)
Equation1 = IIf(IsNull(Rs3("Equation1").value), 0, Rs3("Equation1").value)
Equation2 = IIf(IsNull(Rs3("Equation2").value), 0, Rs3("Equation2").value)
Else
GroupID = 0
Equation1 = 0
Equation2 = 0
End If
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
    On Error GoTo ErrTrap

    Me.FG.FinishEditing True

    DoEvents

 
    Screen.MousePointer = vbArrowHourglass

    If Trim(Me.TxtTransSerial.text) = "" Then
        Msg = "ÌÃ» ≈œŒ«· —ﬁ„ «·›« Ê—…...!!"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
      '  TxtTransSerial.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    Else

        If Me.TxtModFlg.text = "N" Then
    
            BolTemp = UniqueTransSerial(Trim(Me.TxtTransSerial.text), 2)
        ElseIf Me.TxtModFlg.text = "E" Then
            BolTemp = UniqueTransSerial(Trim(Me.TxtTransSerial.text), 2, val(Me.XPTxtBillID.text))
        End If

        BolTemp = True

        If BolTemp = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "—ﬁ„ «·›« Ê—… „”Ã· „”»ﬁ« ›Ï «·»—‰«„Ã.." & CHR(13)
                Msg = Msg & "Ê·«Ì„ﬂ‰  ﬂ—«— —ﬁ„ «·›« Ê—…"
            Else
                Msg = "This Bill No Already Exist" & CHR(13)
        
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         '   TxtTransSerial.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    If DcCurrency.BoundText = "" Then
    
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "«Œ — «·⁄„·… «Ê·« "
        Else
            Msg = "Select Currency First"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
     '   Dccurrency.SetFocus
        Sendkeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    
    End If
   
    If val(DBCboClientName.BoundText) = 0 Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·"
    '    Else
    '        Msg = "Select Customer First"
    '    End If
'
     '   MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
' DBCboClientName.SetFocus
'        SendKeys "{F4}"
'        Screen.MousePointer = vbDefault
'        Exit Sub
    End If

     If DCboStoreName.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
        Else
            Msg = "Select Inventory First"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
       ' DCboStoreName.SetFocus
       Sendkeys "{F4}"
       
        Screen.MousePointer = vbDefault
         Exit Sub
    End If

     If Trim(DcboEmp.BoundText) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·»«∆⁄/«·„‰œÊ»..!!!"
        Else
            Msg = "Must Specify SalesPerson/Saller..!!!"
        End If

        'MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title'
       
    messageResult = MsgBoxPause(hWnd, Msg, App.Title, vbExclamation, MsssageSeconde)
        
      
      
    '    DcboEmp.SetFocus
            Sendkeys "{F4}"
        Screen.MousePointer = vbDefault
             Exit Sub
     End If

    If XPDtbBill.value = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ  «—ÌŒ «·»Ì⁄"
        Else
            Msg = "Specify Date"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 '       XPDtbBill.SetFocus
        Sendkeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If CboPayMentType.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
        Else
            Msg = "Specify Payment Method"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
      '  CboPayMentType.SetFocus
        Sendkeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If XPChkPayType(0).value = vbChecked Then
            
            
            If Me.DcboBox.BoundText = "" Then
      Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
     
          '  If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
                DcboBox.Enabled = False
                DCboStoreName.Enabled = True
                DcboEmp.Enabled = False
          
                Me.dcBranch.BoundText = userbranchid
                Me.DCboStoreName.BoundText = dstore
                Me.DcboBox.BoundText = dBox
                Me.DcboEmp.BoundText = EmpID
 

          '  End If
        End If
        If Me.DcboBox.BoundText = "" Then
                     
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·Œ“‰…...!!!"" "
                'MsgBox "ÌÃ»  ÕœÌœ «”„ «·Œ“‰…...!!!", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                'MsgBox "Must Specify Box...!!!", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
              Msg = "Must Specify Box...!!!"
            End If
messageResult = MsgBoxPause(hWnd, Msg, App.Title, vbExclamation, MsssageSeconde)

     '       DcboBox.SetFocus
             Sendkeys "{F4}"
                   
             Screen.MousePointer = vbDefault
             Exit Sub
        End If
    
    End If

    '----------------------------------------------
    If val(Me.XPTxtValue(1).text) > 0 Then
        If ChkInstall.value = vbChecked Then
            If val(Me.LblInstallTotal.Caption) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» Õ”«» «·√ﬁ”«ÿ ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
                Else
            
                    Msg = "Must Calculate Installment Before Save..!!!"
                End If

                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.XPTab301.CurrTab = 1
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If val(Me.LblInstallTotal.Caption) <> val(Me.XPTxtValue(1).text) Then
                Me.XPTxtValue(1).text = val(Me.LblInstallTotal.Caption)
            End If
            
        End If
    End If

    '-----------------------------------------
    If XPChkPayType(2).value = vbChecked Then
        If val(Me.lbl(18).Caption) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈œŒ«· «·‘Ìﬂ«  ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
            Else
                Msg = "Must Enter Cheque Before Save..!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Me.XPTab301.CurrTab = 1
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If Dcbanks.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = Msg + "ÌÃ»  ÕœÌœ «”„ «·»‰ﬂ Â–« «·Œ’„ " & CHR(13)
            Else
                Msg = Msg + "Specify Bank Name " & CHR(13)
            End If
        
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            'Dcbanks.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
    
            Dim rsbank As New ADODB.Recordset
            Set rsbank = New ADODB.Recordset
            rsbank.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       
            If Not (rsbank.EOF Or rsbank.BOF) Then
                If rsbank!banks_Accounts = True Then
                    bank_account = get_bank_Account(val(Me.Dcbanks.BoundText), "Account_Code1")
                Else
                    bank_account = get_bank_Account(val(Me.Dcbanks.BoundText), "Account_Code")
                End If
            End If
        
        End If
    End If

    If XPChkTAX.value = Checked Then
        If XPTxtTaxValue.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» «œŒ«· ﬁÌ„… ÷—Ì»… «·„»Ì⁄« "
            Else
                Msg = "Enter Sales Tax"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
          '  XPTxtTaxValue.SetFocus
            'Fg.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
        If XPTxtDiscountVal.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "≈–« ﬂ«‰ Â‰«ﬂ Œ’„ ⁄·Ï «·›« Ê—… " & CHR(13)
                Msg = Msg + "ÌÃ»  ÕœÌœ ﬁÌ„… Â–« «·Œ’„ " & CHR(13)
                Msg = Msg + "√Ê √Œ Ì«— ·« ÌÊÃœ Œ’„ "
            Else
                Msg = Msg + " Must Enter Discount Value " & CHR(13)
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPCboDiscountType.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

 
 

    
    
    '«·ﬂ‘› ⁄·Ï ÊÃÊœ ﬂ· Õ”«»«  «·›« Ê—…
    If CheckAccounts = False Then
        Exit Sub
    End If
    
    Me.XPTab301.CurrTab = 0

    If CboPayMentType.ListIndex = 1 And Me.XPChkPayType(0).value = Unchecked And Me.XPChkPayType(2).value = Unchecked Then
 
    End If
 
    If NewGrid.CheckDataEntered = False Then
        Exit Sub
    End If

'          Exit Sub
    ' „—«Ã⁄Â «”⁄«— «· ﬂ·›…
        
        If optsale(1).value = True Then
        'DeleteTransactiomsVoucher val(Text1.text)
        Dim UnitID As Long
        Dim MsgBoxResult As Integer
        Dim DblItemCostPrice  As Double
      
       
        
        If SystemOptions.AllowReturnWithoutCost = True Then
                       For RowNum = 1 To FG.rows - 1
                       If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                    UnitID = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", 0, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                    FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = val(ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Text1.text), UnitID, val(Me.DCboStoreName.BoundText)))
                    End If
                    
                               If FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = 0 Then
                                FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = val(FG.TextMatrix(RowNum, FG.ColIndex("Price")))
                        End If
                        
                    Next RowNum
     End If
If SystemOptions.AllowReturnWithoutCost = False Then

        For RowNum = 1 To FG.rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                
                If CboRetrunType.ListIndex = 0 Then '„ﬁÌœ »›« Ê—…
                    If val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))) = 0 Then
                        MsgBox "«·’‰›   " & FG.TextMatrix(RowNum, FG.ColIndex("Name")) & " €Ì— „Õœœ ”⁄—  ﬂ·› Â Ê·–·ﬂ ·« Ì„ﬂ‰ « „«„ ⁄„·ÌÂ «·«—Ã«⁄ "
                                              
                        Exit Sub
                    End If
                                 
                Else '€Ì— „ﬁÌœ »›« Ê—…
                    UnitID = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))

                    If val(ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Text1.text), UnitID)) = 0 Then
                        'If Val(ModItemCostPrice.GetCostItemPrice(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod)) = 0 Then
                        MsgBoxResult = MsgBox("«·’‰›   " & FG.TextMatrix(RowNum, FG.ColIndex("Name")) & " €Ì— „Õœœ ”⁄—  ﬂ·› Â —»„« ·⁄œ„ ÊÃÊœ ﬂ„Ì… Ê·–·ﬂ ·« Ì„ﬂ‰ « „«„ ⁄„·ÌÂ «·«—Ã«⁄ " & CHR(13) & "Â·  —Ìœ Õ”«»  ﬂ·› … ⁄·Ï «”«” «Œ— ”‰œ ’—› «‰ ÊÃœ ‰⁄„ «Ê ·« ", vbYesNo)

                        If MsgBoxResult = vbYes Then
                            FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = getLastCostPriceForItems(FG.TextMatrix(RowNum, FG.ColIndex("Code")), UnitID)
                        Else
                            MsgBoxResult = MsgBox("«·’‰›   " & FG.TextMatrix(RowNum, FG.ColIndex("Name")) & " €Ì— „Õœœ ”⁄—  ﬂ·› Â —»„« ·⁄œ„ ÊÃÊœ ﬂ„Ì… Ê·–·ﬂ ·« Ì„ﬂ‰ « „«„ ⁄„·ÌÂ «·«—Ã«⁄ " & CHR(13) & "Â·  —Ìœ Õ”«»  ﬂ·› … ⁄·Ï «”«” «‰ ÌﬂÊ‰ ‰›” ”⁄— «·„—œÊœ«  ‰⁄„ / ·« «ﬁÊ„ »√œŒ«· ”⁄— ÌœÊÌ ", vbYesNo)

                            If MsgBoxResult = vbYes Then
                                FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = FG.TextMatrix(RowNum, FG.ColIndex("Price"))
                            Else
                                DblItemCostPrice = InputBox("«œŒ· «·”⁄— ··’‰›" & FG.TextMatrix(RowNum, FG.ColIndex("Name")))
                                FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = val(DblItemCostPrice)
                            End If
                                                                             
                        End If
                                                    
                    Else
                        FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = val(ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Text1.text), UnitID))
                        '   Exit Sub
                    End If
                                            
                End If
                      
            End If

        Next RowNum
        End If
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

    If Me.XPChkPayType(2).value = vbChecked Then
        DblNotesTotal = DblNotesTotal + val(Me.lbl(18).Caption)
    End If

    If CboPayMentType.ListIndex = 1 And Me.XPChkPayType(2).value = Unchecked Then
        XPChkPayType(1).value = 1
        '  XPTxtValue(1).text = Val(LblTotalAll.Caption)
        XPTxtValue(1).text = val(LblTotal.Caption)

    Else

        If CboPayMentType.ListIndex = 1 And Me.XPChkPayType(2).value = vbChecked Then
            XPChkPayType(1).value = 0

        Else
            XPChkPayType(0).value = 1
            '  XPTxtValue(0).text = Val(LblTotalAll.Caption)
            XPTxtValue(0).text = val(LblTotal.Caption)

        End If
    End If

    If Due_Date > DtpDelayDate.value Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "ÌÃ» «‰ ÌﬂÊ‰  «—ÌŒ «·«” Õﬁ«ﬁ «ﬂ»— „‰ «Ê Ì”«ÊÏ  «—ÌŒ «Œ— ﬁ”ÿ"
        Else
            MsgBox "Installment Date must be >= today date"
        End If

        Exit Sub
    End If

    CurrentVoucherNo = ""
    CurrentVoucherSerialNo = ""

    'Create big notes
    my_branch = val(Me.dcBranch.BoundText)

    If TxtNoteSerial.text = "" Then
             '       If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
                  '      MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                 '   Else
                                   '
                 '                   If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                 '                       MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                 '                   Else
                 '                       TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
                 '                   End If
                 '   End If
    End If

    my_branch = val(Me.dcBranch.BoundText)
'
'If optsale(0).value = True Then
'    If TxtNoteSerial1.Text = "" Then
'        If Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21, , , , , , val(DCboUserName.BoundText)) = "error" Then
'            MsgBox " ·« Ì„ﬂ‰ «÷«›…   ›« Ê—… „»Ì⁄«  ÃœÌœ… ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
'        Else
'
'            If Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21, , , , , , val(DCboUserName.BoundText)) = "" Then
'                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
'            Else
'
'            End If
'        End If
'    End If
'
'Else
'
'    If TxtNoteSerial1.Text = "" Then
'        If Voucher_coding(val(my_branch), XPDtbBill.value, 14, 220, , 9, , , , , , val(DCboUserName.BoundText)) = "error" Then
'            MsgBox " ·« Ì„ﬂ‰ «÷«›…   ›« Ê—… „»Ì⁄«  ÃœÌœ… ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
'        Else
'
'            If Voucher_coding(val(my_branch), XPDtbBill.value, 14, 220, 9, , , , , , , val(DCboUserName.BoundText)) = "" Then
'                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
'            Else
'
'            End If
'        End If
'    End If
'
'End If

   ' Set RsNotesGeneral = New ADODB.Recordset
   ' RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
   '  StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   'RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

On Error GoTo ErrTrap

    Screen.MousePointer = vbArrowHourglass
    Cn.BeginTrans
    TransBegine = True
    
    If Me.TxtModFlg.text = "N" Then
        Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
    Else
    
    
        'StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & Val(rs("Transaction_ID").value)
        'Cn.Execute StrSqlDel, , adExecuteNoRecords
        '        MsgBox Val(rs("Transaction_ID").value)
        StrSqlDel = "delete From Notes where Transaction_ID=" & val(Me.XPTxtBillID.text) ' Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
           StrSQL = "Delete From  dbo.ItemsDetails  Where Transaction_ID=" & val(Me.XPTxtBillID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "delete From Notes where noteid=" & val(TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords

        CurrentVoucherNo = GetVoucherGLNO(val(Text1.text), CurrentVoucherSerialNo)
        DeleteTransactiomsVoucher val(Text1.text)
        
        general_noteid = val(TXTNoteID.text)
    End If
    Cn.Execute " delete   From dbo.Transactions  Where (Transaction_Type = 70) And (Transaction_ID = " & val(TxtPhone(3).text) & ")"
   ' RsNotesGeneral.AddNew
   ' RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
   ' general_noteid = RsNotesGeneral("NoteID").value
   ' TXTNoteID.text = general_noteid
    ' RsNotesGeneral("Transaction_ID").value = Val(XPTxtBillID.text)
   ' RsNotesGeneral("NoteDate").value = XPDtbBill.value
   ' RsNotesGeneral("NoteType").value = 170
   ' RsNotesGeneral("Note_Value").value = val(LblTotal.Caption)
   ' RsNotesGeneral("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
   ' RsNotesGeneral("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
   ' RsNotesGeneral("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
'
'    RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
'    RsNotesGeneral("numbering_type1").value = sand_numbering_type(7) '  ›« Ê—… »Ì⁄
'    RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
'    RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
'    RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
    'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
'    RsNotesGeneral.update

    '---------------------------------
    Set RSTransDetails = New ADODB.Recordset
  '  RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
    
    Set RsNotes = New ADODB.Recordset
  '  RsNotes.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


    If SystemOptions.SysRegisterState <> Registered And SystemOptions.SysRegisterState <> DevelopVersion Then
        If rs.RecordCount > 50 Then
            'Exit Sub
        End If
    End If

    

    If Me.TxtModFlg.text = "N" Then
        
     
If optsale(0).value = True Then
     TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21, , , , , , val(DCboUserName.BoundText))
 Else
 TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 14, 220, , 9, , , , , , val(DCboUserName.BoundText))
 End If
 
        rs.AddNew
XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))

    ElseIf Me.TxtModFlg.text = "E" Then
        StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(Me.XPTxtBillID.text) 'Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        StrSqlDel = "delete From Notes where Transaction_ID=" & val(Me.XPTxtBillID.text)  'Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
         StrSQL = "Delete From TblTransactionPayments Where Transaction_ID=" & val(Me.XPTxtBillID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblSalesPoints Where TransID=" & val(Me.XPTxtBillID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
    End If
   rs("Transaction_ID").value = val(XPTxtBillID.text)
    rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
   
   rs.update
   rs.Resync adAffectCurrent
    If lbl(57).Visible = True Then
   rs("SpecialOffer").value = 1
   Else
   rs("SpecialOffer").value = 0
   End If
   'CopyNO
   rs("VAT").value = val(TxtValueAdded.text)
   rs("CopyNO").value = 1
   rs("CardId0").value = Trim(Txtcard(0).text)
   rs("CardId1").value = Trim(Txtcard(1).text)
    
    
    rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
    rs("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
        If Trim$(Me.CashCustomerName.text) <> "" Then
        rs("CashCustomerName").value = Trim$(Me.CashCustomerName.text)
    Else
        rs("CashCustomerName").value = Null
    End If
     rs("VATNO").value = IIf(Trim(Me.TxtPhone(1).text) = "", Null, Trim(Me.TxtPhone(1).text))
    If Trim$(Me.TxtPhone(0).text) <> "" Then
        rs("CashCustomerPhone").value = Trim$(Me.TxtPhone(0).text)
    Else
        rs("CashCustomerPhone").value = Null
    End If
    
    rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '

    If CboPayMentType.ListIndex = 0 Then
        rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
    Else
        rs("BoxID").value = Null
      
    End If
      
    rs("NoteId").value = val(TXTNoteID.text)
    rs("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.text) = "", "", Trim(Me.TxtTransSerial.text))
    rs("Transaction_Date").value = XPDtbBill.value
    '********************
  If optsale(0).value = True Then
    rs("Transaction_Type").value = 21
   Else
     rs("Transaction_Type").value = 9
   End If
   
    rs("UserID").value = user_id
    rs("nots").value = ""

    rs("Currency_id").value = IIf(DcCurrency.BoundText = "", Null, val(DcCurrency.BoundText))
    rs("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.text), 1, txt_Currency_rate.text)

    If XPCboDiscountType.ListIndex = -1 Then
        rs("Trans_DiscountType").value = 0
    Else
        rs("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If

    rs("Trans_Discount").value = IIf(XPTxtDiscountVal.text = "", Null, val(XPTxtDiscountVal.text))
    rs("CusID").value = 2 '  IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
    rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
    rs("order_no") = IIf(TXTOrDer_no.text = "", Null, val(TXTOrDer_no.text))

    If CboPayMentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else
        rs("PaymentType").value = val(CboPayMentType.ListIndex)
    End If

    rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
    rs("TaxValue").value = IIf(XPTxtTaxValue.text = "", Null, val(XPTxtTaxValue.text))
    rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)
    rs("ReturnSerial").value = IIf(TxtInvSerial.text = "", Null, TxtInvSerial.text)

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

  '  If Trim$(Me.TxtCashCustomerName.text) <> "" Then
  '      rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.text)
  '  Else
  '      rs("CashCustomerName").value = Null
  '  End If

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
  
  rs("PPointID").value = PPointID
  
    If BillBasedOn(0).value = True Then
        rs("BillBasedOn").value = 0
    ElseIf BillBasedOn(1).value = True Then
        rs("BillBasedOn").value = 1
    ElseIf BillBasedOn(2).value = True Then
        rs("BillBasedOn").value = 2
    End If
    
    '‰ﬁ«ÿ «·»Ì⁄
    If CboPOSBillType.ListIndex = 0 Then
        rs("POSBillType").value = 0
        rs("STableID").value = val(LblStableID.Caption)
    Else
        rs("POSBillType").value = val(CboPOSBillType.ListIndex)
        rs("STableID").value = Null
    End If

    rs("SessionD").value = SessionD
     rs("CoponValue").value = val(LblDiscountsTotalView(1).Caption)
        rs("Transaction_NetValue").value = val(lblInstComm.Caption) + val(LblTotal.Caption) '+ val(Me.TxtValueAdded.Text)

    rs.update

SaveCopoun
If val(XPCboDiscountType.ListIndex) = 3 And val(XPTxtDiscountVal.text) Then
SavePointsPayd
End If
SaveValueAdded
Dim mTaxExemptTotal As Double
    For RowNum = 1 To FG.rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then

            'Check Repeat Serial
            If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                StrSQL = "select * From Transaction_Details where ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                StrSQL = StrSQL + " and Transaction_ID =" & XPTxtBillID.text
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsTemp.EOF Or RsTemp.BOF) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & CHR(13)
                        Msg = Msg + FG.cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & CHR(13)
                        Msg = Msg + " „ √œŒ«·Â ·ﬁÿ⁄… √Œ—Ï ›Ì Â–Â «·›« Ê—…"
                    Else
                        Msg = "Item Serial " & CHR(13)
                        Msg = Msg + FG.cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & CHR(13)
                        Msg = Msg + "Duplicated in this Bill"
                    End If

                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    RsTemp.Close
                    XPTab301.CurrTab = 0
                    FG.Row = RowNum
                    FG.Col = FG.ColIndex("name")
                    FG.ShowCell RowNum, FG.ColIndex("name")
                  '  Fg.SetFocus
                
                    TransBegine = False
                    Cn.RollbackTrans
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If

                RsTemp.Close
            End If

            If IsEmpty(Me.FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))) Then
                If val(Me.FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))) = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = " ÌÃ»  ÕœÌœ ÊÕœ… «·ﬂ„Ì… «·Œ«’… »«·’‰›" & CHR(13)
                    Else
                        Msg = " Must Select Item Unit For Item :" & CHR(13)
                    End If

                    Msg = Msg + FG.cell(flexcpTextDisplay, RowNum, FG.ColIndex("Name")) & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTab301.CurrTab = 0
                    FG.Row = RowNum
                    FG.Col = FG.ColIndex("UnitID")
                    FG.ShowCell RowNum, FG.ColIndex("UnitID")
                    FG.SetFocus
                    Screen.MousePointer = vbDefault
                    GoTo ErrTrap
                End If
            End If

            RSTransDetails.AddNew
            RSTransDetails("ProductionDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
            RSTransDetails("ExpiryDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
            RSTransDetails("LotNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("LotNO")))
            RSTransDetails("IsExpirDate").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("IsExpirDate")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("IsExpirDate"))))
            RSTransDetails("printed").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("printed")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("printed")))
            RSTransDetails("PrintName").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("PrintName")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("PrintName")))
            
    
            RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))

            RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
            RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
            RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))

            RSTransDetails("Transaction_ID").value = val(XPTxtBillID.text)
            RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
      If optsale(1).value = True Then
       RSTransDetails("FLgReturn").value = -1
      End If
            'RSTransDetails("Quantity").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
            '            RSTransDetails("ItemName").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Name")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Name"))))
'            If Not FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "" Then
'                StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Name"))
'                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'                If Not (RsTemp.EOF Or RsTemp.BOF) Then
'                    If RsTemp("HaveSerial").value = True Then
'                        RSTransDetails("ItemSerial").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("Serial")))
'                    End If
'                End If
'
'                RsTemp.Close
'            End If
            '''
            RSTransDetails("Vat").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vat")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Vat"))))
            RSTransDetails("Vatyo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vatyo")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Vatyo"))))
            
            RSTransDetails("ShowPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
            RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
            
            RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
            
            RSTransDetails("guaranteeTime").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
          RSTransDetails("ParrtNoCode").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode"))))
  RSTransDetails("ItemDetailedCode").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode"))))
  
                        
            RSTransDetails("EmpID4").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("EmpID4")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("EmpID4"))))
              RSTransDetails("CusID2").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("CusID2")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("CusID2"))))
                RSTransDetails("SupplierID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("SupplierID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("SupplierID"))))
                        

            
            RSTransDetails("CostTransID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("PofTransID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("PofTransID"))))
            RSTransDetails("ItemProfit").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit"))))
        
            RSTransDetails("UnitID").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
          
            If SystemOptions.TypicalProduction = False Then
RSTransDetails("CostPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))))
                'RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , RSTransDetails("UnitID").value)

                If RSTransDetails("CostPrice").value = 0 Then
                  '  RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , LastPurPriceType, , , XPDtbBill.value, , RSTransDetails("UnitID").value)
                    
                End If
                  
            Else
                RSTransDetails("CostPrice").value = 0
            
            End If
            
               If optsale(1).value = True Then   ' return sallimng
                    RSTransDetails("CostPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))))
           
             
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
            LngUnitID = val(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
            DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
           '     RSTransDetails("OpeningSalesQty").value = RSTransDetails("Quantity").value
           '     RSTransDetails("OpeningSalesValue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Valu")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Valu"))))
                RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
            
            End If

            SngTemp = SngTemp + (val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))) * RSTransDetails("quantity").value)
         
            If Me.XPCboDiscountType.ListIndex = 1 Then
                TotalBillDiscount = IIf(XPTxtDiscountVal.text = "", Null, (XPTxtDiscountVal.text))
                     'XPTxtDiscountVal
                     
            ElseIf XPCboDiscountType.ListIndex = 2 Then

                If XPTxtDiscountVal.text <> "" Then
                '    TotalBillDiscount = IIf(XPTxtDiscountVal.Text = "", Null, (XPTxtDiscountVal.Text)) * val(LblTotalAll.Caption) / 100
                   TotalBillDiscount = IIf(XPTxtDiscountVal.text = "", Null, (XPTxtDiscountVal.text)) * val(LBLGross.Caption) / 100
                                                         
                Else
                    TotalBillDiscount = 0
                End If
            End If

            'TotalDiscountPerLine = ((RSTransDetails("SHOWprice") * RSTransDetails("SHOWQTY")) / LblTotalAll.Caption) * (TotalBillDiscount)
           
           
      '
      '     TotalDiscountPerLine = Fg.TextMatrix(RowNum, Fg.ColIndex("Valu")) / LblTotalAll.Caption * (TotalBillDiscount)
           
       If LblTotalAll.Caption > 0 Then
           If val(FG.TextMatrix(RowNum, FG.ColIndex("Valu"))) > 0 Then
           
          TotalDiscountPerLine = FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / (LBLGross) * TotalBillDiscount
           
         TotalDiscountPerLine = Round(TotalDiscountPerLine, 2)
           Else
           TotalDiscountPerLine = 0
           End If
         If val(FG.TextMatrix(RowNum, FG.ColIndex("itemtype"))) = 1 Then
                                                                                
         'ItemsServiceTotalsnew = ItemsServiceTotalsnew + TotalDiscountPerLine + val(Fg.TextMatrix(RowNum, Fg.ColIndex("discountvalue")))
         Else
         'ItemsGoodsTotalsnew = ItemsGoodsTotalsnew + TotalDiscountPerLine + val(Fg.TextMatrix(RowNum, Fg.ColIndex("discountvalue")))
         End If
 Else
 TotalDiscountPerLine = 0
 End If
     
     
            RSTransDetails("TotalDiscountPerLine") = Round(TotalDiscountPerLine, 20)
       '                       Dim OldQty As Double
       '      Dim OldCost As Double
       '       Dim NewQty As Double
       '        Dim NewCost As Double
               
'getItemCostData XPDtbBill.value, RSTransDetails("Item_ID").value, val(DCboStoreName.BoundText), val(Me.XPTxtBillID.Text), OldQty, OldCost, NewQty, NewCost
'       RSTransDetails("OldQty").value = NewQty
'       RSTransDetails("OldCost").value = NewCost
'
'      RSTransDetails("NewQty").value = RSTransDetails("OldQty").value - RSTransDetails("Quantity").value
'       RSTransDetails("NewCost").value = RSTransDetails("OldCost").value ' ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
'
       
            RSTransDetails.update
            '-------------
        If TxtPhone(0).text <> "" And SystemOptions.AllowWorkCustomerPoints = True Then
        SavePoints val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))), val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) * val(FG.TextMatrix(RowNum, FG.ColIndex("Price")))
        End If
        End If

    Next RowNum

    
'************************************************************************************
   Set RSTransDetails1 = New ADODB.Recordset
   StrSQL = "SELECT   * from dbo.TblTransactionPayments Where (1 = -1)"
   RSTransDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
      
        PayDes = ""
    For RowNum = 1 To Grid.rows - 1
            
                       If Grid.TextMatrix(RowNum, Grid.ColIndex("Value")) <> "" Then
                        
                                    'Check Repeat Serial
                                     
If PayDes <> "" Then
          PayDes = PayDes & CHR(13) & Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName")) & ":" & Grid.TextMatrix(RowNum, Grid.ColIndex("value"))
          If Grid.TextMatrix(RowNum, Grid.ColIndex("CardNo")) <> "" Then
          PayDes = PayDes & CHR(13) & "  —ﬁ„ «· ›ÊÌ÷:  " & Grid.TextMatrix(RowNum, Grid.ColIndex("CardNo"))
          End If
          
 Else
           PayDes = Grid.TextMatrix(RowNum, Grid.ColIndex("PaymentName")) & ":" & Grid.TextMatrix(RowNum, Grid.ColIndex("value"))
        If Grid.TextMatrix(RowNum, Grid.ColIndex("CardNo")) <> "" Then
               PayDes = PayDes & CHR(13) & "  —ﬁ„ «· ›ÊÌ÷:  " & Grid.TextMatrix(RowNum, Grid.ColIndex("CardNo"))
          End If
          
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
                                           
                                                If optsale(1).value = True Then   ' return sallimng
                                                    RSTransDetails1("Effect").value = -1
                                                  Else
                                                     RSTransDetails1("Effect").value = 1
                                                End If
                                                
                                           RSTransDetails1.update
                                  
                             
                    End If
                    If RowNum < FG.rows - 1 Then
                    If FG.ValueMatrix(RowNum, FG.ColIndex("chkTaxExempt")) = True Then
                    mTaxExemptTotal = mTaxExemptTotal + ((val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) * val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) - val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountValue")))
                
                    End If
                    End If
   Next RowNum
'***************************************************************************************
  'wael
            
   rs!TotalTaxExempt = mTaxExemptTotal
   rs.update
   savenewelectroncic
    'wael
    'LblValueAdded.Tag = mTaxExemptTotal
'createVoucher
  TransBegine = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
 If optsale(0).value = True Then
    If SystemOptions.autoIssueVoucher = True Then
        CreateIssueVoucher
    End If
    
  Else 'return sales
  
   
    CreateRecieveVoucher
    If Not IsVouc Then MsgBox "ÕœÀ Œÿ√ «À‰«¡ «‰‘«¡ «–‰ «·«” ·«„ ": GoTo ErrTrap
  End If
     
 'SaveItemsData
 SavecustomerData TxtPhone(0).text, Me.CashCustomerName
    
    Cn.CommitTrans

    TransBegine = False
    '----------------------------------------------------------------
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    
    Savetemp
    '·√‰‰« ﬁ„‰« »≈÷«›… Õ—ﬂ… „‰ ‰Ê⁄ „Œ ·›…
    StrSQL = "SELECT * FROM Transactions WHERE  1=-1" ' & InvType
         
    'If SystemOptions.usertype <> UserAdminAll Then
'    StrSQL = StrSQL & "  AND   BranchId=" & Current_branch
'
    Set rs = New ADODB.Recordset
   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    Me.Retrive val(Me.XPTxtBillID.text)
    '----------------------------------------------------------------

    CuurentLogdata
    If isFromExcel Then Exit Sub

    Select Case Me.TxtModFlg.text
    
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Data Was Saved do you want Another Entry" & CHR(13)
                
            End If
            
           ' XPBtnMove_Click (2)
  CboPOSBillType.ListIndex = 1
  If Pay_Print <> 1 Then
            If SystemOptions.Save_options = 1 Or SystemOptions.Save_options = 2 Then
                PrintReport

                DoEvents
                DoEvents
                DoEvents
        
            ElseIf CboPOSBillType.ListIndex > 0 Then
            If optsale(0).value = True Then
                PrintReport , 1, ""
             Else
                 PrintReport , 1, "›« Ê—… „— Ã⁄"
             End If
 End If
                '----------------------------------------------------------------
                '·√‰‰« ﬁ„‰« »≈÷«›… Õ—ﬂ… „‰ ‰Ê⁄ „Œ ·›…
         '       StrSQL = "SELECT * FROM Transactions WHERE  1=-1" ' & InvType
         
                'If SystemOptions.usertype <> UserAdminAll Then
         '       StrSQL = StrSQL & "  AND   BranchId=" & Current_branch

         '       Set rs = New ADODB.Recordset
         '       rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
         '       Me.Retrive val(Me.XPTxtBillID.text)
  
                DoEvents
                DoEvents
                DoEvents
        
               ' Cmd_Click (0)
                 btnNew_Click (0)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
'            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton1, App.title) = vbYes Then
            If 1 = 1 Then
            
                Cmd_Click (0)
                Screen.MousePointer = vbDefault
                
            Else
                TxtModFlg.text = "R"
            End If

 
 
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
           
        rs.Resync adAffectCurrent
        RsNotes.Resync adAffectCurrent
        RSTransDetails.Resync adAffectCurrent
    End If

    'Resume
'    If rs.EditMode <> adEditNone Then
'        rs.CancelUpdate
'    End If
'
'    If Not RsNotes Is Nothing Then
'        If RsNotes.EditMode <> adEditNone Then
'            RsNotes.CancelUpdate
'        End If
'    End If
'
'    If Not RSTransDetails Is Nothing Then
'        If RSTransDetails.EditMode <> adEditNone Then
'            RSTransDetails.CancelUpdate
'        End If
'    End If

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
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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

Function GetBalance() As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT        SUM(CreditValue - DebtValue) AS vlaue"
sql = sql & " From dbo.TblSalesPoints"
sql = sql & " WHERE        (Mobile = N'" & TxtPhone(0).text & "') "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetBalance = IIf(IsNull(rs2("vlaue").value), 0, rs2("vlaue").value)
Else
GetBalance = 0
End If
End Function
Private Sub XPCboDiscountType_Change()
    XPCboDiscountType_Click
End Sub

Private Sub XPCboDiscountType_Click()
    On Error GoTo ErrTrap

    If XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = -1 Then
    
        XPTxtDiscountVal.Enabled = False
        XPTxtDiscountVal.text = ""
    Else
    
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.text = ""
    End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            If Me.TxtModFlg.text = "N" Then
                If XPCboDiscountType.ListIndex = 3 Then
                    lbl(8).Visible = True
                    XPTxtDiscountVal.Visible = True
                    lbl(8).Visible = True
        
                If TxtPhone(0).text = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ «·ÃÊ«·"
                    Else
                        MsgBox "Please Enter No.Mobile"
                    End If
                    TxtPhone(0).SetFocus
                    Exit Sub
                End If
        End If
       'Wael Discount
       ' XPTxtDiscountVal.Text = GetBalance
    End If
    End If
    
        If FG.TextMatrix(1, FG.ColIndex("Code")) <> "" Then
            NewGrid.Calculate 1, , , True
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

            If XPChkPayType(2).value = Checked And Me.TxtModFlg.text <> "R" Then
                Me.CmdCheque.Enabled = True
            Else
                Me.CmdCheque.Enabled = False
                Me.lbl(18).Caption = 0
                Me.lbl(19).Caption = 0
                Me.FgCheques.rows = Me.FgCheques.FixedRows
            End If

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

Private Sub PrintReport(Optional PrinterTarget As Boolean = False, _
                        Optional pos As Integer = 0, _
                        Optional sTitle As String, Optional View As Integer = 0, Optional printername As String = "")

 Dim CopOInfro As String
 Dim i As Integer
 Dim CountCop As Integer
 Dim SmCop As Double

If optsale(0).value = True Then
CopOInfro = ""
CountCop = 0
SmCop = 0
        With FgC
             For i = 1 To .rows - 1
             If .cell(flexcpChecked, i, .ColIndex("Selcd")) = flexChecked Then
             CountCop = CountCop + 1
             SmCop = SmCop + val(.TextMatrix(i, .ColIndex("Vlue")))
             End If
             Next i
             End With
           If SmCop <> 0 And CountCop <> 0 Then
           If SystemOptions.UserInterface = ArabicInterface Then
             CopOInfro = CopOInfro & "·œÌﬂ ⁄œœ "
             CopOInfro = CopOInfro & " " & CountCop
             CopOInfro = CopOInfro & " " & "ﬁ”Ì„…"
             CopOInfro = CopOInfro & " " & " »≈Ã„«·Ì"
             CopOInfro = CopOInfro & " " & SmCop
             Else
             CopOInfro = CopOInfro & "You Have "
             CopOInfro = CopOInfro & " " & CountCop
             CopOInfro = CopOInfro & " " & "Coupon"
             CopOInfro = CopOInfro & " " & " Total"
             CopOInfro = CopOInfro & " " & SmCop
             End If
           End If
      End If
'‰ﬁ«ÿ «·»Ì⁄
View = 0
    If View = 0 Then
    'Cmd_Click (1)
   ' Cmd_Click (2)

    DoEvents
    DoEvents
    DoEvents


                 
                
                
         
    'TxtPayedValue = val(Me.LBLPayVal)
 TxtRemainValue.text = val(Me.TxtPayedValue) - val(Me.TxtNetValue.text)
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

    If SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3 Then
        P_Target = PrinterTarget
    Else
        P_Target = WindowTarget
    End If

    ShowType = GetSetting(StrAppRegPath, "View_Type", "SallReportType", 1)
ShowType = 2
    If ShowType = 1 Then
        If XPTxtBillID.text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingDataDetailed XPTxtBillID.text, 4, , , LblTotal, TxtSearchCode.text, TxtBillComment.text, val(lblInstComm.Caption)
            '    SaleReport.ShowSallingData XPTxtBillID.text, 4, , val(Me.TxtPayedValue.text), val(Me.TxtRemainValue.text), pos, sTitle

            '  If MDIFrmMain.MnuInvPrintReceipt.Checked = True Then
            '      SaleReport.PrintInvoiceReceipt Val(XPTxtBillID.text), P_Target
            '  End If
        End If
 Cn.Execute "update Transactions set Printed =1   where Transaction_ID=" & val(XPTxtBillID.text)
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
                  If optsale(0).value = True Then
              sTitle = ""
             Else
                 sTitle = "›« Ê—… „— Ã⁄"
             End If
                     Cn.Execute "update Transactions set PayDes ='" & PayDes & "'   where Transaction_ID=" & val(XPTxtBillID.text)
Dim noofCopies1 As Integer
Dim xi As Integer

       noofCopies1 = SystemOptions.NOOFPRINTCOPIESSALES
         If noofCopies1 = 0 Then noofCopies1 = 1
          For xi = 1 To noofCopies1
              DoEvents
    DoEvents
    
            SaleReport.ShowSallingData XPTxtBillID.text, 0, , val(Me.TxtPayedValue.text), val(Me.TxtRemainValue.text), pos, sTitle, , CopOInfro, , , , , , , vbYes
            DoEvents
        Next xi
            '      P_Target = PrinterTarget
        
            'ÿ»«⁄… ≈Ì’«· ≈” ·«„ «·‰ﬁœÌ…
   
        End If

    ElseIf ShowType = 3 Then

        If XPTxtBillID.text <> "" Then
            StrPath = GetSetting(StrAppRegPath, "PrintReport", "ReportPath", App.path & "\Bill_Template\SaleMain.drp")

            If StrPath = "" Then
                Msg = "⁄›Ê« : Â‰«ﬂ Œÿ√„« ›Ì „”«— «· ﬁ—Ì— "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
'If View = 0 Then
    clear_all Me
' End If
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

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

        Select Case IntResult

            Case vbYes
                Cancel = True
                SaveData
              '  Unload customer_screen

            Case vbCancel
                Cancel = True
             '   Unload customer_screen
        End Select

      '  Unload customer_screen
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
    Dim fullcode As String
 
    GetCustomersDetail val(DBCboClientName.BoundText), , fullcode, 1
    TxtSearchCode.text = fullcode

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
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
            Msg = "·ﬁœ  „ «” »œ«· «·ﬁÿ⁄… : " & FG.cell(flexcpTextDisplay, FG.Row, FG.ColIndex("Name")) & CHR(13)
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

    For RowNum = 1 To FG.rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            StrSQL = "select * From QryDelPurchase where Transaction_Date >=" & SQLDate(XPDtbBill.value, True) & ""
            StrSQL = StrSQL + " and Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))
            StrSQL = StrSQL + " and Transaction_Type=9"

            If FG.cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                End If
            End If

            Set RsSalle = New ADODB.Recordset
            RsSalle.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsSalle.EOF Or RsSalle.BOF) Then
                If FG.cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then

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
isFound = False
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

    Dcombos.GetDocTypebyid Me.DCDocTypes, 21, val(Me.dcBranch.BoundText)

    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
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

    StrSQL = "  select  BankID,BankName  from BanksData   "
    fill_combo Dcbanks, StrSQL
      
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(94).Caption = "Ex.Coupons"
    btnExit(6).Caption = "Items Request"
    btnExit(7).Caption = "Items Expiry Date"
    btnExit(8).Caption = "Network Reports"
    btnExit(9).Caption = " General Reports"
'    lbl(57).Caption = "Cash.visa"
btnExit(5).Caption = "Delete"
 lbl(59).Caption = " Payed"
    lbl(60).Caption = "Changed"
    LblDiscountsTotalView(3).Caption = "VAT"
    LblDiscountsTotalView(5).Caption = "VAT"
    LblDiscountsTotalView(4).Caption = "Data VAT"
    ChecVAT.RightToLeft = False
    ChecVAT.Caption = "Select"
 '''//////////////
     lbl(95).Caption = "Barcode"
    lbl(58).Caption = "Total"
    lbl(59).Caption = "Paid"
    lbl(60).Caption = "Remaining"
    FramePay.Caption = "Payments Data"
    btnExit(3).Caption = "Pending"
    btnExit(2).Caption = "Call Up"
    
    CMDPAy(0).Caption = "Pay+Print"
    CMDPAy(1).Caption = "Pay"
    With Grid
    .TextMatrix(0, .ColIndex("PaymentName")) = "Payments"
    .TextMatrix(0, .ColIndex("Value")) = "Value"
    .TextMatrix(0, .ColIndex("CardNo")) = "Card No."
    End With
    With VatGrid
.TextMatrix(0, .ColIndex("index")) = "Serial"
.TextMatrix(0, .ColIndex("select")) = "Select"
.TextMatrix(0, .ColIndex("Code")) = "Item Code"
.TextMatrix(0, .ColIndex("Name")) = "Item Name"
.TextMatrix(0, .ColIndex("Vatyo")) = "Percentage"
.TextMatrix(0, .ColIndex("Vat")) = "Value"
.TextMatrix(0, .ColIndex("Valu")) = "Item Value"
End With

btnNew(3).Caption = "Show Customer Points"
 ''////////
lbl(57).Caption = "Special Offer"
lbl(93).Caption = "User"
lbl(91).Caption = "Quick Srch."
'lbl(92).Caption = "Quick Srch."
lbl(90).Caption = "Code"

    Label3.Caption = "Branch"
    Frame1.Caption = "Info"
    lbl(56).Caption = "Order No."
    lbl(58).Caption = " Total"
    lbl(59).Caption = " Payed"
    lbl(60).Caption = " Changed"
    lbl(63).Caption = " Total Qty"
    Frame2.Caption = "Color Map"
    lbl(68).Caption = " Profit"
optsale(0).Caption = "Sales"
optsale(1).Caption = "Refund"
btnExit(0).Caption = "Exit"
btnExit(1).Caption = "Admin Login"

    Label1.Caption = "Doc Type"
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
    lbl(10).Caption = "Dis.Type"
    lbl(8).Caption = "Value"
    lbl(22).Caption = "Profit Value"
    lbl(23).Caption = "Profit Perce"

    lbl(3).Caption = "Total:"
    lbl(49).Caption = "Net "
    lbl(50).Caption = "Disc"
    lbl(1).Caption = "Cashier"
    lbl(2).Caption = "Rec. Count:"
    lbl(92).Caption = "Coupons"
'*****************************************************
lbl(85).Caption = "Invoice"
lbl(86).Caption = "Invoice No"
ALLButton9.Caption = "POS Close"
'ALLButton1.Caption = "Refund"
lbl(87).Caption = "Mob."
lbl(88).Caption = "Cust. Name"
lbl(89).Caption = "Item Code."
btnNew(0).Caption = "New"
'btnEdit.Caption = "Modify"
btnNew(1).Caption = "Undo"
'btnPending.Caption = "Pending"
btnpay(0).Caption = "Pay"
btnpay(1).Caption = "Pay"
CMDPAy(0).Caption = "Pay+Print"
CMDPAy(1).Caption = "Pay"
btnExit(4).Caption = "Search"
FramePay.Caption = "Pay Data"
    With FgC
        .TextMatrix(0, .ColIndex("Vlue")) = "Value"
        .TextMatrix(0, .ColIndex("Num")) = "Coupon No."
        .TextMatrix(0, .ColIndex("Num2")) = "Coupon No."
        .TextMatrix(0, .ColIndex("IsRetCopon")) = "Delivery"
       .TextMatrix(0, .ColIndex("Serial")) = "Serial"
    
    End With
    
lbl(69).Caption = "Totals."
 lbl(50).Caption = "Disc."
lbl(71).Caption = "Net."

'*****************************************************

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
    Me.Cmd(7).Caption = "Print"
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
    
    lbl(18).Caption = "Check NO."
    lbl(17).Caption = "Bank Name"
    lbl(19).Caption = "Due Date"
    CmdINSTALLMENT.Caption = "INSTALLMENT"
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
    CmdCheque.Caption = "Register"

    With FgInstallments
        .TextMatrix(0, .ColIndex("QestID")) = "ID"
        .TextMatrix(0, .ColIndex("value")) = "value"
        .TextMatrix(0, .ColIndex("Due_Date")) = "Due_Date"
 
    End With

    With FG
        .TextMatrix(0, .ColIndex("order_no")) = "ORD/INV NO."
        .TextMatrix(0, .ColIndex("Select")) = "Select"
    End With

    With FgCheques
 
        .TextMatrix(0, .ColIndex("CheckValue")) = "Value"
        .TextMatrix(0, .ColIndex("CheckNumber")) = "Cheque Number"
        .TextMatrix(0, .ColIndex("BankName")) = "Bank Name"
        .TextMatrix(0, .ColIndex("DueDate")) = "Due Date"
        .TextMatrix(0, .ColIndex("ReleaseDate")) = "Release Date"
 
    End With

    XPChkPayType(2).Caption = "Cheques"
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

End Sub
Sub SaveValueAdded()
Dim i As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

sql = "Select * from  TransactionValueAdded where 1=-1"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With Me.VatGrid
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
rs2.AddNew
rs2("Transaction_ID").value = val(Me.XPTxtBillID.text)
rs2("Transaction_Type").value = 21
rs2("ItemID").value = val(.TextMatrix(i, .ColIndex("ItemID")))
rs2("Vatyo").value = val(.TextMatrix(i, .ColIndex("Vatyo")))
rs2("Vat").value = val(.TextMatrix(i, .ColIndex("Vat")))
rs2("Valu").value = val(.TextMatrix(i, .ColIndex("Valu")))
If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
rs2("selectd").value = 1
Else
rs2("selectd").value = 1
End If

rs2.update
End If
Next i
End With
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

    With Me.FgCheques

        If .rows > 1 Then
            Me.lbl(19).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("CheckNumber"), .rows - 1, .ColIndex("CheckNumber"))
            Me.lbl(18).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CheckValue"), .rows - 1, .ColIndex("CheckValue"))
        Else
            Me.lbl(19).Caption = 0
            Me.lbl(18).Caption = 0
        End If

    End With

End Sub

Private Sub ClearNotes()

    LblPrecenType.Caption = 0
    LblPrecenValue.Caption = 0
    LblInstallTotal.Caption = 0
    LblInstallCount.Caption = 0
    LblFirstInstallDate.Caption = ""
    LblInstallSeprator.Caption = ""
    LblInstallmentType.Caption = ""
    LblStartValue.Caption = ""
    Me.LblDiscount.Caption = 0
    Me.LblAdvPayment.Caption = 0
    lbl(19).Caption = ""
    lbl(18).Caption = ""
End Sub









Sub Savetemp()
    
    
    
   ' XPTxtCurrent.Caption = rs.AbsolutePosition
   ' XPTxtCount.Caption = rs.RecordCount
 
 
            
SaveQRCode "transactions", "Transaction_ID", val(XPTxtBillID), TxtNoteSerial1.text, (XPDtbBill.value), _
        (LblTotal.Caption), Picture1, 0, (TxtValueAdded.text), (LblTotal.Caption)



       
End Sub

Function createVoucher()
    Dim bankDes As String
    Dim AccountCode As String
    Dim AccountCode1 As String
 Dim StrSQL As String
    Dim NoteID As Long
    Dim sql As String
 
Dim ReAccount_Code_dynamic  As String
 ReAccount_Code_dynamic = get_account_code_branch(3, my_branch)
bankDes = "”‰œ ﬁ»÷ ⁄„Ê„Ì"
    '//////////////////////////////////////Notes////////////////////////////////////
    Dim line_no As Integer
    Dim RsNotes As New ADODB.Recordset
  '  RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (1 = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
   
   
    
        
        Dim NoteDate As Date
        Dim NoteSerial As String
        Dim Notevalue As Double

    If Me.TxtModFlg.text = "E" Then
                  
        sql = "Delete notes where NoteID=" & val(Me.TXTNoteID.text)
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(TXTNoteID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords

    Else
        TXTNoteID.text = CStr(new_id("Notes", "NoteID", "", True))

   
   

            CreateNotes NoteID, (XPDtbBill.value), val(dcBranch.BoundText), 170, 0, NoteSerial, TxtNoteSerial1, "Transactions", "Transaction_ID", val(XPTxtBillID.text), TxtNoteSerial1.text, ToHijriDate(XPDtbBill.value), TxtManualNo1.text
            general_noteid = NoteID
    End If

TXTNoteID = NoteID
TxtNoteSerial = NoteSerial
    
                
   
                
    line_no = 0
Dim i As Integer
  
'*********************************ﬁÌœ «À»«  «·„»Ì⁄«  **********************************************
        
  


Dim LngDevID  As Long
Dim debitorcredit As Integer
Dim Tvalue As Double
Dim mTaxTobacco As Double
        With Me.Grid
 
            For i = 1 To .rows - 1

                If .TextMatrix(i, .ColIndex("PaymentID")) <> "" And val(.TextMatrix(i, .ColIndex("Value"))) <> 0 Then
              If val(.TextMatrix(i, .ColIndex("PaymentID"))) = 0 Then
                   AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcboBox.BoundText))
               Else
                    AccountCode = (.TextMatrix(i, .ColIndex("Accountsus")))
               End If
               
                    bankDes = "«Ìœ«⁄   „‰ ...   " & (.TextMatrix(i, .ColIndex("PaymentName")))
                    line_no = line_no + 1
  
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
Tvalue = val(.TextMatrix(i, .ColIndex("Value")))
        If Tvalue > 0 Then
        debitorcredit = 0
        Else
        debitorcredit = 1
        Tvalue = Abs(Tvalue)
        End If
'
'                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, Tvalue, debitorcredit, .TextMatrix(i, .ColIndex("PaymentName")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbBill.value, user_id, , , , Tvalue, , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                        GoTo ErrTrap
'
'                    End If
         
                End If

            Next i


Dim salesAccount As String
Dim ReturnsalesAccount As String

Dim VatsalesAccount As String
Dim VatReturnsalesAccount As String
Dim vaTAccount As String
Dim VATValue As Double
Dim X As Boolean
 salesAccount = get_account_code_branch(2, val(Me.dcBranch.BoundText))
 ReturnsalesAccount = get_account_code_branch(3, val(Me.dcBranch.BoundText))
 X = GetValueAddedAccount(XPDtbBill.value, , VatsalesAccount, 1, 21)
 X = GetValueAddedAccount(XPDtbBill.value, VatReturnsalesAccount, , 1, 9)
Dim AccountVATCreitRe As String
 GetValueAddedAccount XPDtbBill.value, AccountVATCreitRe, , 1, 9
 PercentgValueAdded XPDtbBill.value, , , 9
 
   
Dim Percetage As Double
PercentgValueAddedAccount_Transec XPDtbBill.value, 21, 0, , Percetage

               'beforeVat = val(.TextMatrix(i, .ColIndex("value"))) / (1 + Percetage / 100) '1.05
               'Vat = beforeVat * Percetage / 100 ' 0.05
               
 
'
'            For i = 1 To .Rows - 1
'
'                If .TextMatrix(i, .ColIndex("PaymentID")) <> "" And val(.TextMatrix(i, .ColIndex("Value"))) <> 0 Then
'               Tvalue = val(.TextMatrix(i, .ColIndex("beforeVat")))
'               VATValue = val(.TextMatrix(i, .ColIndex("Vat")))
'               If Tvalue > 0 Then ' „»Ì⁄« 
'                 AccountCode = salesAccount
'                vaTAccount = VatsalesAccount
'                 debitorcredit = 1
'
'               Else
'                AccountCode = ReturnsalesAccount
'                vaTAccount = VatReturnsalesAccount
'               Tvalue = Abs(Tvalue)
'               VATValue = Abs(VATValue)
'               debitorcredit = 0
'               End If
'            '        AccountCode = (.TextMatrix(i, .ColIndex("Accountcom")))
'                    bankDes = "«Ìœ«⁄   „‰    " & (.TextMatrix(i, .ColIndex("PaymentName")))
'                    line_no = line_no + 1
'
'                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'
'                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, Tvalue, debitorcredit, .TextMatrix(i, .ColIndex("PaymentName")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbBill.value, user_id, , , , Tvalue, , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                        GoTo ErrTrap
'
'                    End If
'
'                 line_no = line_no + 1
'
'                   If ModAccounts.AddNewDev(LngDevID, line_no, vaTAccount, VATValue, debitorcredit, .TextMatrix(i, .ColIndex("PaymentName")) & bankDes & " ﬁ „÷«›…", val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbBill.value, user_id, , , , Tvalue, , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                        GoTo ErrTrap
'
'                    End If
'                End If
'
'            Next i
'
'
        End With
'
'
'
'*******************************************
'
'  With VSFlexGrid1
'
'    If VSFlexGrid1.rows > 1 Then
'            If val(.TextMatrix(1, .ColIndex("CollectedValue"))) > 0 And (.TextMatrix(1, .ColIndex("PaymentID"))) = 0 Then
'            Dim RsDev  As ADODB.Recordset
'            Set RsDev = New ADODB.Recordset
'            '   RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'            StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
'            RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'
'            '«·ÿ—› «·œ«∆‰      «·’‰œÊﬁ   «·€—⁄Ì «·ﬂ«‘Ì—
'            ' AccountCode = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.Dcbank.BoundText))
'            AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcboBox.BoundText))
'
'            RsDev.AddNew
'            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
'            RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
'            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'            RsDev("Account_Code").value = AccountCode
'            RsDev("Value").value = val(.TextMatrix(1, .ColIndex("salesValue")))
'            RsDev("Credit_Or_Debit").value = 0
'
'            RsDev("RecordDate").value = Me.XPDtbBill.value
'            RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
'            RsDev("Double_Entry_Vouchers_Description").value = "«À»«  «·„»Ì⁄«  «·‰›œÌ…"   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
'            RsDev("Double_Entry_Vouchers_Descriptione").value = "«À»«  «·„»Ì⁄«  «·‰›œÌ…"
'
'            RsDev("UserID").value = user_id
'            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'
'            RsDev.update
'
'
'
'            line_no = line_no + 1
'
'            AccountCode = salesAccount
'            vaTAccount = VatsalesAccount
'            debitorcredit = 1
'            RsDev.AddNew
'            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
'            RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
'            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'            RsDev("Account_Code").value = AccountCode
'            RsDev("Value").value = (val(.TextMatrix(1, .ColIndex("salesValue")) - val(.TextMatrix(1, .ColIndex("chkTaxExempt")))) / (1 + Percetage / 100)) 'AAA
'            'beforeVat = val(.TextMatrix(i, .ColIndex("value"))) / (1 + Percetage / 100) '1.05
'               'Vat = beforeVat * Percetage / 100 ' 0.05
'            RsDev("Credit_Or_Debit").value = 1
'
'            RsDev("RecordDate").value = Me.XPDtbBill.value
'            RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
'            RsDev("Double_Entry_Vouchers_Description").value = "«À»«  «·„»Ì⁄«  «·‰›œÌ…"   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
'            RsDev("Double_Entry_Vouchers_Descriptione").value = "«À»«  «·„»Ì⁄«  «·‰›œÌ…"
'
'            RsDev("UserID").value = user_id
'            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'
'            RsDev.update
'
'
'
'            line_no = line_no + 1
'            If Percetage <> 0 Then
'                AccountCode = salesAccount
'                vaTAccount = VatsalesAccount
'                debitorcredit = 1
'                RsDev.AddNew
'                RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
'                RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
'                RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'                RsDev("Account_Code").value = vaTAccount
'                RsDev("Value").value = val(.TextMatrix(1, .ColIndex("salesValue")) - val(.TextMatrix(1, .ColIndex("chkTaxExempt")))) - ((val(.TextMatrix(1, .ColIndex("salesValue")) - val(.TextMatrix(1, .ColIndex("chkTaxExempt")))) / (1 + Percetage / 100)))
'                RsDev("Credit_Or_Debit").value = 1
'                '- val(.TextMatrix(1, .ColIndex("chkTaxExempt")))
'                RsDev("RecordDate").value = Me.XPDtbBill.value
'                RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
'                RsDev("Double_Entry_Vouchers_Description").value = "«À»«  «·„»Ì⁄«  «·‰›œÌ…"   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
'                RsDev("Double_Entry_Vouchers_Descriptione").value = "«À»«  «·„»Ì⁄«  «·‰›œÌ…"
'
'                RsDev("UserID").value = user_id
'                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'
'                RsDev.update
'
'            End If
'            bankDes = "«Ìœ«⁄ ‰ﬁœÌ „‰   «·’‰œÊﬁ: " & DcGeneralBox.text
'            '«·ÿ—› «·„œÌ‰      «·’‰œÊﬁ «·⁄„Ê„Ì ﬂ«‘
'            ' AccountCode = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.Dcbank.BoundText))
'            AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcGeneralBox.BoundText))
'
'            RsDev.AddNew
'            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
'            RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
'            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'            RsDev("Account_Code").value = AccountCode
'            RsDev("Value").value = val(.TextMatrix(1, .ColIndex("NetValue")))
'            RsDev("Credit_Or_Debit").value = 0
'
'            RsDev("RecordDate").value = Me.XPDtbBill.value
'            RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
'            RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
'            RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
'
'            RsDev("UserID").value = user_id
'            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'
'            RsDev.update
'
'
'            line_no = line_no + 1
'              bankDes = "«Ìœ«⁄ ‰ﬁœÌ «·Ï  «·’‰œÊﬁ:  " & DcboBox.text
'            '«·ÿ—› «·œ«∆‰      «·’‰œÊﬁ   «·€—⁄Ì «·ﬂ«‘Ì—
'            ' AccountCode = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.Dcbank.BoundText))
'            AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcboBox.BoundText))
'
'            RsDev.AddNew
'            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
'            RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
'            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'            RsDev("Account_Code").value = AccountCode
'            RsDev("Value").value = val(.TextMatrix(1, .ColIndex("NetValue")))
'            RsDev("Credit_Or_Debit").value = 1
'
'            RsDev("RecordDate").value = Me.XPDtbBill.value
'            RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
'            RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
'            RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
'
'            RsDev("UserID").value = user_id
'            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'
'            RsDev.update
'
'
'    End If
'End If
'End With
'
'
'
'
''*******************************
'With VSFlexGrid1
'If val(.TextMatrix(1, .ColIndex("ReturnValue"))) > 0 And (.TextMatrix(1, .ColIndex("PaymentID"))) = 0 Then
'
'
'                  line_no = line_no + 1
'            bankDes = "«À»«  «·„—œÊœ«  «·‰ﬁœÌ…   «·’‰œÊﬁ: " & DcboBox.text
'            '«·ÿ—› «·œ«∆‰      «·’‰œÊﬁ   «·€—⁄Ì «·ﬂ«‘Ì—
'            ' AccountCode = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.Dcbank.BoundText))
'            AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcboBox.BoundText))
'
'            RsDev.AddNew
'            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
'            RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
'            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'            RsDev("Account_Code").value = AccountCode
'            RsDev("Value").value = val(.TextMatrix(1, .ColIndex("ReturnValue")))
'            RsDev("Credit_Or_Debit").value = 0
'
'            RsDev("RecordDate").value = Me.XPDtbBill.value
'            RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
'            RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
'            RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
'
'            RsDev("UserID").value = user_id
'            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'
'            RsDev.update
'
'
' line_no = line_no + 1
' bankDes = "«À»«  «·„—œÊœ«  «·‰ﬁœÌ…   «·’‰œÊﬁ: " & DcGeneralBox.text
' ' bankDes = "«Ìœ«⁄ ‰ﬁœÌ «·Ï    " & DcGeneralBox.Text   '«·ÿ—› «·„œÌ‰      «·’‰œÊﬁ «·⁄„Ê„Ì ﬂ«‘
'            ' AccountCode = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.Dcbank.BoundText))
'            AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcGeneralBox.BoundText))
'
'            RsDev.AddNew
'            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
'            RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
'            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'            RsDev("Account_Code").value = AccountCode
'            RsDev("Value").value = val(.TextMatrix(1, .ColIndex("ReturnValue")))
'            RsDev("Credit_Or_Debit").value = 1
'
'            RsDev("RecordDate").value = Me.XPDtbBill.value
'            RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
'            RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
'            RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
'
'            RsDev("UserID").value = user_id
'            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'
'            RsDev.update
'
'
'            line_no = line_no + 1
'
'
'  AccountCode = ReturnsalesAccount
'                vaTAccount = VatsalesAccount
'                 debitorcredit = 0
'                 bankDes = " «À»«  «·„—œÊœ«  «·‰ﬁœÌ…"
'        RsDev.AddNew
'        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'        RsDev("branch_id").value = val(Me.dcBranch.BoundText)
'        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
'        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'        RsDev("Account_Code").value = AccountCode
'        RsDev("Value").value = (val(.TextMatrix(1, .ColIndex("ReturnValue"))) / (1 + Percetage / 100))
'        RsDev("Credit_Or_Debit").value = 0
'
'        RsDev("RecordDate").value = Me.XPDtbBill.value
'        RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
'        RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
'        RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
'
'        RsDev("UserID").value = user_id
'        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'
'        RsDev.update
'
'
'
' line_no = line_no + 1
'
'  AccountCode = salesAccount
'                vaTAccount = AccountVATCreitRe
'                 debitorcredit = 0
'                 bankDes = "«À»«  «·„—œÊœ«  «·‰ﬁœÌ… Õ”«» «·ﬁÌ„… «·„÷«›…"
'        RsDev.AddNew
'        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'        RsDev("branch_id").value = val(Me.dcBranch.BoundText)
'        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
'        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'        RsDev("Account_Code").value = AccountVATCreitRe
'        RsDev("Value").value = val(.TextMatrix(1, .ColIndex("ReturnValue"))) - (val(.TextMatrix(1, .ColIndex("ReturnValue"))) / (1 + Percetage / 100))
'        RsDev("Credit_Or_Debit").value = 0
'
'        RsDev("RecordDate").value = Me.XPDtbBill.value
'        RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
'        RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
'        RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
'
'        RsDev("UserID").value = user_id
'        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'
'        RsDev.update
'                                                  line_no = line_no + 1
'   AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcboBox.BoundText))
'       bankDes = "««À»«  «·„—œÊœ«  «·‰ﬁœÌ…Ìœ«⁄ ‰ﬁœÌ «·Ï    " & DcboBox.text
'        RsDev.AddNew
'        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'        RsDev("branch_id").value = val(Me.dcBranch.BoundText)
'        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
'        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'        RsDev("Account_Code").value = AccountCode
'        RsDev("Value").value = val(.TextMatrix(1, .ColIndex("ReturnValue")))
'        RsDev("Credit_Or_Debit").value = 1
'
'        RsDev("RecordDate").value = Me.XPDtbBill.value
'        RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
'        RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
'        RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
'
'        RsDev("UserID").value = user_id
'        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'
'        RsDev.update
'
'
'
'     End If
'     End With
''     line_no = line_no + 1
''bankDes = "«Ìœ«⁄ ‰ﬁœÌ „‰    " & DcboBox.Text
''        '«·ÿ—› «·„œÌ‰      «·’‰œÊﬁ «·⁄„Ê„Ì ﬂ«‘
''       ' AccountCode = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.Dcbank.BoundText))
''       AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcGeneralBox.BoundText))
''
''        RsDev.AddNew
''        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
''        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
''        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
''        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
''        RsDev("Account_Code").value = AccountCode
''        RsDev("Value").value = val(.TextMatrix(1, .ColIndex("NetValue")))
''        RsDev("Credit_Or_Debit").value = 0
''
''        RsDev("RecordDate").value = Me.XPDtbBill.value
''        RsDev("Notes_ID").value = val(Me.TxtNoteID.Text)   '(XPTxtID.text)
''        RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
''        RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
''
''        RsDev("UserID").value = user_id
''        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
''
''        RsDev.update
''
''
'' line_no = line_no + 1
''
''        '«·ÿ—› «·œ«∆‰      «·’‰œÊﬁ   «·€—⁄Ì «·ﬂ«‘Ì—
''       ' AccountCode = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.Dcbank.BoundText))
''       AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcboBox.BoundText))
''
''        RsDev.AddNew
''        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
''        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
''        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
''        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
''        RsDev("Account_Code").value = AccountCode
''        RsDev("Value").value = val(.TextMatrix(1, .ColIndex("NetValue")))
''        RsDev("Credit_Or_Debit").value = 1
''
''        RsDev("RecordDate").value = Me.XPDtbBill.value
''        RsDev("Notes_ID").value = val(Me.TxtNoteID.Text)   '(XPTxtID.text)
''        RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
''        RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
''
''        RsDev("UserID").value = user_id
''        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
''
''        RsDev.update
''
''
''    End If
''End If
'
'
'Dim mValue As Double
'Dim mValue2 As Double
' Dim mValTaxEx As Double
'  Dim mValTaxExValue As Double
' Dim mValue3 As Double
'    'ﬁÌœ «‰Ê«⁄ «·œ›⁄ «·«Œ—Ì
'    If VSFlexGrid1.rows > 2 Then
'
'
'
'
'        With VSFlexGrid1
'
'            For i = 2 To .rows - 1
'
'                If .TextMatrix(i, .ColIndex("PaymentID")) <> "" And val(.TextMatrix(i, .ColIndex("NetValue"))) > 0 Then
'               mTaxTobacco = val(.TextMatrix(i, .ColIndex("TaxTobacco")))
'                    AccountCode = (.TextMatrix(i, .ColIndex("Account_Code")))
'                    bankDes = "«À»«  «·„»Ì⁄«  ÿ—ﬁ œ›⁄ :    " & (.TextMatrix(i, .ColIndex("PaymentName")))
'                    line_no = line_no + 1
'
'                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'
'
'                    If SystemOptions.IsBlue Then
'                     '   mValue = val(.TextMatrix(i, .ColIndex("NetValue"))) - val(.TextMatrix(i, .ColIndex("CommissionValue")))
'                        mValue = val(.TextMatrix(i, .ColIndex("salesvalue")))
'                        mValue3 = val(.TextMatrix(i, .ColIndex("salesvalue")))
'                    Else
'                        mValue = val(.TextMatrix(i, .ColIndex("NetValue")))
'                        mValue3 = val(.TextMatrix(i, .ColIndex("NetValue")))
'                    End If
'                    Dim mVaaa As Double
'
'                    If mTaxTobacco <> 0 Then
'                        mValTaxEx = val(.TextMatrix(i, .ColIndex("chkTaxExempt")))
'                        mValue2 = ((mValue - val(mValTaxEx)) / (1 + (Percetage / 100))) + val(mValTaxEx)
'
'                        mValTaxExValue = mValue2 * mTaxTobacco / 100
'
'                        mVaaa = mValue + mValTaxExValue + (mValue - ((mValue - val(mValTaxEx)) / (1 + (Percetage / 100))) + val(mValTaxEx))
'                    End If
'                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, mValue3, 0, .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbBill.value, user_id, , , , .TextMatrix(i, .ColIndex("NetValue")), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                        GoTo ErrTrap
'
'                    End If
'                    mValue = mValue3
'
'         AccountCode = salesAccount
'                vaTAccount = VatsalesAccount
'
'                    line_no = line_no + 1
'
'
'                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'                    If mTaxTobacco <> 0 Then
'                        mValue2 = mValue / 2 / 1.075
'                    Else
'
'                        mValTaxEx = val(.TextMatrix(i, .ColIndex("chkTaxExempt")))
'                        mValue2 = ((mValue - val(mValTaxEx)) / (1 + (Percetage / 100))) + val(mValTaxEx)
'                    End If
'
'                  '  val (mValue - val(.TextMatrix(i, .ColIndex("chkTaxExempt"))))
'                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(mValue2), 1, .TextMatrix(i, .ColIndex("Remarks")) & "ﬁÌ„… „÷«›…", val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbBill.value, user_id, , , , .TextMatrix(i, .ColIndex("NetValue")), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                        GoTo ErrTrap
'
'                    End If
'
'        vaTAccount = VatsalesAccount
'
'                    line_no = line_no + 1
'
'                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'
'                    If mTaxTobacco <> 0 Then
'                        mValue2 = mValue - (mValue2 * 2)
'                        If ModAccounts.AddNewDev(LngDevID, line_no, vaTAccount, mValue2, 1, .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbBill.value, user_id, , , , .TextMatrix(i, .ColIndex("NetValue")), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                            GoTo ErrTrap
'
'                        End If
'                    Else
'                        mValue2 = val(mValue - val(.TextMatrix(i, .ColIndex("chkTaxExempt"))))
'
'                        If ModAccounts.AddNewDev(LngDevID, line_no, vaTAccount, mValue2 - (val(mValue2) / (1 + Percetage / 100)), 1, .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbBill.value, user_id, , , , .TextMatrix(i, .ColIndex("NetValue")), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                            GoTo ErrTrap
'
'                        End If
'
'
'                    End If
'
'                    If mTaxTobacco <> 0 Then
'
'
'                       vaTAccount = Trim((.TextMatrix(i, .ColIndex("AccTaxTobacco"))))
'
'                       line_no = line_no + 1
'
'                       LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'                        mValTaxExValue = mValue / 2 / 1.075  'mValue2 * mTaxTobacco / 100
'                       If ModAccounts.AddNewDev(LngDevID, line_no, vaTAccount, mValTaxExValue, 1, .TextMatrix(i, .ColIndex("Remarks")) & "÷—Ì»… «· »€", val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbBill.value, user_id, , , , .TextMatrix(i, .ColIndex("NetValue")), , , , "÷—Ì»… «· »€", , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                           GoTo ErrTrap
'
'                       End If
'
'
'                    End If
'
'
'
'                End If
'                If .TextMatrix(i, .ColIndex("PaymentID")) <> "" And val(.TextMatrix(i, .ColIndex("CommissionValue"))) > 0 Then
'                    AccountCode = (.TextMatrix(i, .ColIndex("Accountcom")))
'                    bankDes = "««À»«  ⁄„Ê·Â «·»‰ﬂ :  " & (.TextMatrix(i, .ColIndex("PaymentName")))
'                    line_no = line_no + 1
'
'                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'
'                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(.TextMatrix(i, .ColIndex("CommissionValue"))), 0, .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbBill.value, user_id, , , , .TextMatrix(i, .ColIndex("CommissionValue")), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                    GoTo ErrTrap
'
'                    End If
'
'
'                    AccountCode = (.TextMatrix(i, .ColIndex("Account_Code")))
'                    bankDes = "««À»«  ⁄„Ê·Â «·»‰ﬂ :  " & (.TextMatrix(i, .ColIndex("PaymentName")))
'                    line_no = line_no + 1
'
'                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'
'                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(.TextMatrix(i, .ColIndex("CommissionValue"))), 1, .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbBill.value, user_id, , , , .TextMatrix(i, .ColIndex("CommissionValue")), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                    GoTo ErrTrap
'
'                    End If
'
'
'
'                    AccountCode = vaTAccount
'                    bankDes = "«·ﬁÌ„… «·„÷«›… ··⁄„Ê·… " & (.TextMatrix(i, .ColIndex("PaymentName")))
'                    line_no = line_no + 1
'
'                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'
'                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(.TextMatrix(i, .ColIndex("CommissionValue"))) - (val(.TextMatrix(i, .ColIndex("CommissionValue"))) / (1 + Percetage / 100)), 0, .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbBill.value, user_id, , , , .TextMatrix(i, .ColIndex("CollectedValue")), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                    GoTo ErrTrap
'
'                    End If
'
'
'
'                    AccountCode = (.TextMatrix(i, .ColIndex("Accountcom")))
'                    bankDes = "«·ﬁÌ„… «·„÷«›… ··⁄„Ê·…" & (.TextMatrix(i, .ColIndex("PaymentName")))
'                    line_no = line_no + 1
'
'                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'
'                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(.TextMatrix(i, .ColIndex("CommissionValue"))) - (val(.TextMatrix(i, .ColIndex("CommissionValue"))) / (1 + Percetage / 100)), 1, .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbBill.value, user_id, , , , .TextMatrix(i, .ColIndex("CollectedValue")), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                    GoTo ErrTrap
'
'                    End If
'
'
'                End If
'            Next i
'
'
'
'                     For i = 2 To .rows - 1
'
'                If .TextMatrix(i, .ColIndex("PaymentID")) <> "" And val(.TextMatrix(i, .ColIndex("CollectedValue"))) > 0 Then
'
'                    AccountCode = (.TextMatrix(i, .ColIndex("Accountsus")))
'                    bankDes = "«Ìœ«⁄   „‰    " & (.TextMatrix(i, .ColIndex("PaymentName")))
'                    line_no = line_no + 1
'
'                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
''
''                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(.TextMatrix(i, .ColIndex("CollectedValue"))), 1, .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbBill.value, user_id, , , , .TextMatrix(i, .ColIndex("CollectedValue")), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
''                        GoTo ErrTrap
''
''                    End If
'
'            AccountCode = (.TextMatrix(i, .ColIndex("Account_Code")))
'                    bankDes = "«Ìœ«⁄   „‰    " & (.TextMatrix(i, .ColIndex("PaymentName")))
'                    line_no = line_no + 1
'
'                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'
''                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(.TextMatrix(i, .ColIndex("CollectedValue"))), 1, .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbBill.value, user_id, , , , .TextMatrix(i, .ColIndex("CollectedValue")), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
''                        GoTo ErrTrap
''
''                    End If
'
'                        End If
'
'            Next i
'
'        End With
'
'    End If
  
  
    updateNotesValueAndNobytext (val(Me.TXTNoteID.text))

ErrTrap:
End Function






Function savenewelectroncic()
   'vat data
    Dim InvoiceTypeCodeID As Integer
    rs("CIBAN").value = "" 'TXTIban.text
    'vat data
      rs("RecTime").value = Time
            

  If val(DCDocTypes.BoundText) <> 0 Then
  'wAEL
    getDocAccounts val(DCDocTypes.BoundText), , , , , , , , , , , , InvoiceTypeCodeID
  Else
 InvoiceTypeCodeID = 388
  End If
  rs("InvoiceTypeCodeID").value = InvoiceTypeCodeID
 
 
 
 If val(Me.DefaultInvoicetype.ListIndex) = 0 Then
   
   
    If Export = 1 Then
    rs("InvoiceTypeCodename").value = "0100100"
    Else
      rs("InvoiceTypeCodename").value = "0100000"
   End If
   
   
   
   
   Else
    rs("InvoiceTypeCodename").value = "0200000"
   End If

   rs("DocumentCurrencyCode").value = DcCurrency.text
   rs("TaxCurrencyCode").value = DcCurrency.text
  rs("ActualDeliveryDate").value = Date
 rs("LatestDeliveryDate").value = Date
Dim PaymentMeansCode As String
         
            '10 In cash
            '30 Credit
            '42 Payment to bank account
            '48 Bank card
            '1 Instrument not defined(Free text)
            Dim paymentnote
        If CboPayMentType.ListIndex = 0 Then ' ‰ﬁœ«
                  PaymentMeansCode = "10"
                      paymentnote = "Payment by Cash"
        ElseIf CboPayMentType.ListIndex = 1 Then ' ¬Ã·
                 PaymentMeansCode = "30"
                 paymentnote = "Payment by Credit"
         ElseIf CboPayMentType.ListIndex = 2 Or CboPayMentType.ListIndex = 3 Then  '  ÕÊÌ· »‰ﬂÌ
                    If SystemOptions.AllowSalesMultyPayed = True Then
                     PaymentMeansCode = "48" 'ﬂ«— 
                      paymentnote = "Payment by Bank Card"
                    Else
                    PaymentMeansCode = "42" '»‰ﬂ Õ”«»
                    paymentnote = "Payment by Bank Account"
                    End If
         
         End If
         
         rs("PaymentMeansCode").value = PaymentMeansCode
      
rs("paymentnote").value = paymentnote
rs.update
End Function

