VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form System_manger333 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬂÊÌœ «·„” ‰œ« "
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16170
   Icon            =   "System_manger333.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   16170
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "System_manger333.frx":6852
      Left            =   17760
      List            =   "System_manger333.frx":685C
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   840
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   18720
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   1560
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   19080
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   1440
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "System_manger333.frx":6871
      Left            =   18720
      List            =   "System_manger333.frx":6881
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   48
      Top             =   3360
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   20520
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Text            =   "modflag"
      Top             =   4440
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2655
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   840
      Width           =   16035
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   2040
         Width           =   7695
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   3960
            TabIndex        =   9
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox Check2 
            Caption         =   "6"
            Height          =   195
            Left            =   1080
            TabIndex        =   10
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " ﬂÊÌœ ÿ»ﬁ« ··„Œ“‰"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4200
            TabIndex        =   29
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ì„·∆ «’›«—"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1440
            TabIndex        =   28
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   2055
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   0
         Width           =   7695
         Begin VB.CommandButton Command1 
            Caption         =   " ÿ»Ìﬁ «· ﬂÊÌœ ⁄·Ì ﬂ· «·”‰œ« "
            Height          =   495
            Left            =   360
            TabIndex        =   64
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3720
            TabIndex        =   5
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3720
            TabIndex        =   6
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3720
            TabIndex        =   7
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            ItemData        =   "System_manger333.frx":689A
            Left            =   3720
            List            =   "System_manger333.frx":68AA
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox TxtPrefix 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   480
            TabIndex        =   8
            Top             =   1680
            Width           =   4815
         End
         Begin VB.ComboBox CBOYearDigit 
            Height          =   315
            ItemData        =   "System_manger333.frx":68CF
            Left            =   480
            List            =   "System_manger333.frx":68D9
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Ì»œ√ „‰"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5760
            TabIndex        =   61
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Ì‰ ÂÌ ›Ì"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5760
            TabIndex        =   60
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·Œ«‰« "
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   6000
            TabIndex        =   46
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «· —ﬁÌ„"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6120
            TabIndex        =   45
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ã“¡ «·À«» "
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   5880
            TabIndex        =   44
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " ‰”Ìﬁ «·”‰…"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2400
            TabIndex        =   43
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Height          =   2655
         Left            =   7920
         TabIndex        =   25
         Top             =   0
         Width           =   7815
         Begin VB.CheckBox ChkIsSerialByUser 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”—Ì«· ÿ»ﬁ« ··„” Œœ„ ›Ï «·Õ—ﬂ« "
            Height          =   285
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   2310
            Width           =   3285
         End
         Begin VB.CheckBox chkIsCodeByBranch 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· ﬂÊÌœ ÿ»ﬁ« ··›—⁄"
            Height          =   285
            Left            =   6060
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox txtBreaks 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1080
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Text            =   "/"
            Top             =   1980
            Width           =   1365
         End
         Begin VB.CheckBox chkIsBreaks 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· ﬂÊÌœ »›Ê«’·"
            Height          =   285
            Left            =   6060
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   1980
            Width           =   1575
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            Height          =   915
            Left            =   360
            TabIndex        =   57
            Top             =   600
            Width           =   7095
            Begin VB.OptionButton Option2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›—⁄ „Õœœ"
               Height          =   315
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   480
               Width           =   1575
            End
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂ· «·›—Ê⁄ "
               Height          =   315
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   150
               Width           =   1455
            End
            Begin MSDataListLib.DataCombo dcBranch 
               Height          =   315
               Left            =   120
               TabIndex        =   1
               Top             =   480
               Width           =   4875
               _ExtentX        =   8599
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.TextBox TxtSerial 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   4560
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   240
            Width           =   1665
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            ItemData        =   "System_manger333.frx":68EE
            Left            =   480
            List            =   "System_manger333.frx":69EB
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   1620
            Width           =   4815
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‘ﬂ· «·›«’·"
            Height          =   315
            Index           =   35
            Left            =   2610
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   2010
            Width           =   1035
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ  "
            Height          =   315
            Index           =   3
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·”‰œ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5760
            TabIndex        =   42
            Top             =   1620
            Width           =   1695
         End
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   16305
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "System_manger333.frx":7001
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   675
         TabIndex        =   20
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "System_manger333.frx":739B
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1275
         TabIndex        =   21
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "System_manger333.frx":7735
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1800
         TabIndex        =   22
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "System_manger333.frx":7ACF
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ﬂÊÌœ «·„” ‰œ« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   12120
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   2520
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   15360
         Picture         =   "System_manger333.frx":7E69
         Stretch         =   -1  'True
         Top             =   120
         Width           =   615
      End
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1545
      Left            =   120
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6720
      Width           =   16035
      _cx             =   28284
      _cy             =   2725
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   36
         Top             =   600
         Width           =   15855
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   14400
            TabIndex        =   13
            Top             =   240
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
            ButtonStyle     =   1
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
            ButtonImage     =   "System_manger333.frx":8C17
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   9720
            TabIndex        =   15
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            ButtonStyle     =   1
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
            ButtonImage     =   "System_manger333.frx":F479
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   12240
            TabIndex        =   14
            Top             =   240
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
            ButtonStyle     =   1
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
            ButtonImage     =   "System_manger333.frx":F813
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7200
            TabIndex        =   16
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
            ButtonImage     =   "System_manger333.frx":16075
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   4920
            TabIndex        =   17
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·„Õœœ"
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
            ButtonImage     =   "System_manger333.frx":1640F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
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
            ButtonImage     =   "System_manger333.frx":169A9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   330
            Left            =   2640
            TabIndex        =   63
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·ﬂ·"
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
            ButtonImage     =   "System_manger333.frx":16D43
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   3855
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   1
            Left            =   690
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   975
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   255
            Width           =   675
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   540
         End
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   15600
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   -600
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "»ÕÀ"
         BackColor       =   14737632
         FontSize        =   9.75
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "System_manger333.frx":1D5A5
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   16080
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   0
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ÕœÌÀ"
         BackColor       =   14871017
         FontSize        =   9.75
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "System_manger333.frx":1D93F
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   11520
         TabIndex        =   39
         Top             =   120
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton ISButton3 
         Height          =   330
         Left            =   9360
         TabIndex        =   12
         ToolTipText     =   " ÕœÌœ «·ﬂ·"
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ÕœÌœ «·ﬂ·"
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
         ButtonImage     =   "System_manger333.frx":1DCD9
         ButtonImageDisabled=   "System_manger333.frx":2453B
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton ISButton1 
         Height          =   330
         Left            =   7200
         TabIndex        =   62
         ToolTipText     =   " ÕœÌœ «·ﬂ·"
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "«·€«¡ «· ÕœÌœ"
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
         ButtonImage     =   "System_manger333.frx":43725
         ButtonImageDisabled=   "System_manger333.frx":49F87
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   14760
         TabIndex        =   40
         Top             =   120
         Width           =   900
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   3075
      Left            =   90
      TabIndex        =   11
      Top             =   3600
      Width           =   15975
      _cx             =   28178
      _cy             =   5424
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
      BackColorAlternate=   16777088
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
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"System_manger333.frx":69171
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
      ExplorerBar     =   1
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   615
         Left            =   2760
         TabIndex        =   41
         Top             =   1680
         Visible         =   0   'False
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   19080
      TabIndex        =   51
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   1200
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
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   18480
      TabIndex        =   52
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   20400
      Top             =   3960
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
            Picture         =   "System_manger333.frx":6943D
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "System_manger333.frx":697D7
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "System_manger333.frx":69B71
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "System_manger333.frx":69F0B
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "System_manger333.frx":6A2A5
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "System_manger333.frx":6A63F
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "System_manger333.frx":6A9D9
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "System_manger333.frx":6AF73
            Key             =   "BuyValue"
         EndProperty
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
      Left            =   19800
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "System_manger333"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim II As Long

Private Sub Command1_Click()
 Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
    If Option2.value = True Then
                If dcBranch.Text = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                        MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «”„ «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
                                        dcBranch.SetFocus
                                         
                                Else
                                        MsgBox "Write Branch Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                      
                                      dcBranch.SetFocus
                            End If
                              Exit Sub
                 End If
         End If
     '''''''''''''''''''''''''''''''''
 
    '+++++++++++++++++++++++++++++++++++++++++++++++
       If Combo2.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· ‰Ê⁄ «· —ﬁÌ„", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Combo2.SetFocus
            Exit Sub
            Else
            MsgBox "Write Numbering type ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            Combo2.SetFocus
         End If
     End If
    ''''''''''''''''''''''''''''''''''''''''''''''
      If Me.Combo2.ListIndex > 0 Then
        If Text3.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· »œ«Ì… «· ﬂÊÌœ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Text3.SetFocus
             Exit Sub
             Else
            MsgBox "Write Start Codeing ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Text3.SetFocus
            Exit Sub
           End If
            End If
     End If
     ''''''''''''''''''''''''''''''''''''''''''
     If Text2.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· ‰Â«Ì… «· ﬂÊÌœ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Text2.SetFocus
             Exit Sub
             Else
            MsgBox "Write End Codeing ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Text2.SetFocus
            Exit Sub
            End If
     End If
    '+++++++++++++++++++++++++++++++++++++++++++++++
       If Text4.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· ⁄œœ «·Œ«‰«  ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Text4.SetFocus
             Exit Sub
      Else
            MsgBox "Write Number of Digits", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Text4.SetFocus
            Exit Sub
            End If
     End If
        
   
       If Option1.value = True Then 'ALL BRANCH
FiLLRecWithAllVoucher
Else
FiLLRecWithAllVoucher val(dcBranch.BoundText)
End If
End Sub

   Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    
  
    conection = "select * from sanad_numbering order by  id "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
   'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.dcBranch
    'Wael
    If SystemOptions.IsByNewCoding Then
        chkIsBreaks.Enabled = True
        chkIsCodeByBranch.Enabled = True
        ChkIsSerialByUser.Enabled = True
        txtBreaks.Enabled = True
    Else
        chkIsBreaks.Enabled = False
        chkIsCodeByBranch.Enabled = False
        ChkIsSerialByUser.Enabled = False
        txtBreaks.Enabled = False
        chkIsBreaks.value = vbUnchecked
        chkIsCodeByBranch.value = vbUnchecked
        ChkIsSerialByUser.value = vbUnchecked
        txtBreaks = ""
    End If
    'Wael
    
     AdditemTocCmp
     BtnLast_Click
     ShowTip
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
      If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If
    Option1.value = True
    Me.Refresh
ErrTrap:
End Sub
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
   On Error GoTo ErrTrap
    RsSavRec.Fields("branch_no").value = val(Me.dcBranch.BoundText)
    RsSavRec.Fields("BranchName").value = Me.dcBranch.Text
    RsSavRec.Fields("sanad_no").value = IIf(val(Combo1.ListIndex) <> -1, val(Combo1.ListIndex), Null)
    RsSavRec.Fields("sanad_type").value = IIf(Combo1.Text <> "", Trim(Combo1.Text), Null)
    RsSavRec.Fields("start_at").value = IIf(Text3.Text <> "", Trim(Text3.Text), Null)
    RsSavRec.Fields("end_at").value = IIf(Text2.Text <> "", Trim(Text2.Text), Null)
    '''''''''''''''''''''''''
    RsSavRec.Fields("numbering_id").value = IIf(val(Combo2.ListIndex) <> -1, val(Combo2.ListIndex), Null)
    RsSavRec.Fields("numbering_type").value = IIf(Combo2.Text <> "", Trim(Combo2.Text), Null)
    '' '''''''''''''''''''''
    If CBOYearDigit.ListIndex = 0 Then
    RsSavRec.Fields("YearDigit").value = 2
    End If
    If CBOYearDigit.ListIndex = 1 Then
    RsSavRec.Fields("YearDigit").value = 4
    End If
    RsSavRec.Fields("no_of_digit").value = IIf(Text4.Text <> "", Trim(Text4.Text), Null)
    RsSavRec.Fields("Prefix").value = IIf(TxtPrefix.Text <> "", Trim(TxtPrefix.Text), Null)
          
    If Check1.value = vbChecked Then
    RsSavRec.Fields("StoreCoding").value = 1
    Else
    RsSavRec.Fields("StoreCoding").value = 0
    End If
               
    If Check2.value = vbChecked Then
    RsSavRec.Fields("zeros").value = 1
    Else
    RsSavRec.Fields("zeros").value = 0
    End If
    
    
    
        'Wael
        RsSavRec.Fields("Breaks").value = IIf(txtBreaks.Text <> "", Trim(txtBreaks.Text), Null)
        
        If chkIsBreaks.value = vbChecked Then
            RsSavRec.Fields("IsBreaks").value = 1
        Else
            RsSavRec.Fields("IsBreaks").value = 0
        End If
               
        
        If chkIsCodeByBranch.value = vbChecked Then
            RsSavRec.Fields("IsCodeByBranch").value = 1
        Else
            RsSavRec.Fields("IsCodeByBranch").value = 0
        End If
               
        If ChkIsSerialByUser.value = vbChecked Then
            RsSavRec.Fields("IsSerialByUser").value = 1
        Else
            RsSavRec.Fields("IsSerialByUser").value = 0
        End If
               
            
               
                       
               
                  

                             
                                 'Wael
    RsSavRec.update
    '''''''''''''''''''''''''''''''''''
    Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                GetData
                TxtModFlg = "R"
                If SystemOptions.UserInterface = ArabicInterface Then
             Else
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                GetData
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
                Me.Refresh
                FiLLTXT
                GetData
                End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                GetData
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                GetData
                TxtModFlg = "R"
            End If
       End Select
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub
   Public Sub FiLLRecWithAll()
    On Error GoTo ErrTrap
    Dim Rs7 As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    Dim StrRecID As String
    Set Rs7 = New ADODB.Recordset
    sql = "select* from TblBranchesData "
    Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs7.RecordCount <= 0 Then Exit Sub
    Rs7.MoveFirst
    For i = 1 To Rs7.RecordCount
    If chek(Rs7.Fields("branch_id").value, val(Combo1.ListIndex), Me.TxtPrefix.Text) = False Then
    StrRecID = new_id("sanad_numbering", "id", "")
    RsSavRec.AddNew
    RsSavRec.Fields("branch_no").value = Rs7.Fields("branch_id").value
    If SystemOptions.UserInterface = ArabicInterface Then
    RsSavRec.Fields("BranchName").value = Rs7.Fields("branch_name").value
    Else
    RsSavRec.Fields("BranchName").value = Rs7.Fields("branch_namee").value
    End If
    ''''''''''''''''''''''''''''''''''''''''
    RsSavRec.Fields("sanad_no").value = IIf(val(Combo1.ListIndex) <> -1, val(Combo1.ListIndex), Null)
    RsSavRec.Fields("sanad_type").value = IIf(Combo1.Text <> "", Trim(Combo1.Text), Null)
    RsSavRec.Fields("start_at").value = IIf(Text3.Text <> "", Trim(Text3.Text), Null)
    RsSavRec.Fields("end_at").value = IIf(Text2.Text <> "", Trim(Text2.Text), Null)
    '''''''''''''''''''''''''
    RsSavRec.Fields("numbering_id").value = IIf(val(Combo2.ListIndex) <> -1, val(Combo2.ListIndex), Null)
    RsSavRec.Fields("numbering_type").value = IIf(Combo2.Text <> "", Trim(Combo2.Text), Null)
    '' '''''''''''''''''''''
    
        'Wael
        RsSavRec.Fields("Breaks").value = IIf(txtBreaks.Text <> "", Trim(txtBreaks.Text), Null)
        
        If chkIsBreaks.value = vbChecked Then
            RsSavRec.Fields("IsBreaks").value = 1
        Else
            RsSavRec.Fields("IsBreaks").value = 0
        End If
               
        
        If chkIsCodeByBranch.value = vbChecked Then
            RsSavRec.Fields("IsCodeByBranch").value = 1
        Else
            RsSavRec.Fields("IsCodeByBranch").value = 0
        End If
               
        If ChkIsSerialByUser.value = vbChecked Then
            RsSavRec.Fields("IsSerialByUser").value = 1
        Else
            RsSavRec.Fields("IsSerialByUser").value = 0
        End If
               
            
               
                       
               
                  

                             
                                 'Wael
    
    
    If CBOYearDigit.ListIndex = 0 Then
    RsSavRec.Fields("YearDigit").value = 2
    End If
    If CBOYearDigit.ListIndex = 1 Then
    RsSavRec.Fields("YearDigit").value = 4
    End If
    RsSavRec.Fields("no_of_digit").value = IIf(Text4.Text <> "", Trim(Text4.Text), Null)
    RsSavRec.Fields("Prefix").value = IIf(TxtPrefix.Text <> "", Trim(TxtPrefix.Text), Null)
          
    If Check1.value = vbChecked Then
    RsSavRec.Fields("StoreCoding").value = 1
    Else
    RsSavRec.Fields("StoreCoding").value = 0
    End If
               
    If Check2.value = vbChecked Then
    RsSavRec.Fields("zeros").value = 1
    Else
    RsSavRec.Fields("zeros").value = 0
    End If
    RsSavRec.update
    Rs7.MoveNext
    Else
    Rs7.MoveNext
    End If
    Next i
    '''''''''''''''''''''''''''''''''''
    Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                GetData
                TxtModFlg = "R"
                If SystemOptions.UserInterface = ArabicInterface Then
             Else
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                GetData
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
                Me.Refresh
                FiLLTXT
                GetData
                End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                GetData
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                GetData
                TxtModFlg = "R"
            End If
       End Select
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub
   
   Public Sub FiLLRecWithAllVoucher(Optional branch_id As Integer = 0)
    On Error GoTo ErrTrap
    Dim Rs7 As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    Dim StrRecID As String
    Set Rs7 = New ADODB.Recordset
    sql = "select* from TblBranchesData "
    If branch_id <> 0 Then
    sql = sql & " WHERE branch_id=" & branch_id
    End If
    Dim j As Integer
    Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs7.RecordCount <= 0 Then Exit Sub
    Rs7.MoveFirst
    For i = 1 To Rs7.RecordCount
             For j = 0 To Combo1.ListCount - 1
                        If chek(Rs7.Fields("branch_id").value, j, Me.TxtPrefix.Text) = False And (j <> 11 And j <> 44 And j <> 47) Then
                        
                        
                        StrRecID = new_id("sanad_numbering", "id", "")
                        RsSavRec.AddNew
                        RsSavRec.Fields("branch_no").value = Rs7.Fields("branch_id").value
                                If SystemOptions.UserInterface = ArabicInterface Then
                                RsSavRec.Fields("BranchName").value = Rs7.Fields("branch_name").value
                                Else
                                RsSavRec.Fields("BranchName").value = Rs7.Fields("branch_namee").value
                                End If
                        ''''''''''''''''''''''''''''''''''''''''
                        RsSavRec.Fields("sanad_no").value = j ' IIf(val(Combo1.ListIndex) <> -1, val(Combo1.ListIndex), Null)
                        RsSavRec.Fields("sanad_type").value = Combo1.List(j) ' IIf(Combo1.text <> "", Trim(Combo1.text), Null)
                        RsSavRec.Fields("start_at").value = IIf(Text3.Text <> "", Trim(Text3.Text), Null)
                        RsSavRec.Fields("end_at").value = IIf(Text2.Text <> "", Trim(Text2.Text), Null)
                        
                       
                        
                        '''''''''''''''''''''''''
                        RsSavRec.Fields("numbering_id").value = IIf(val(Combo2.ListIndex) <> -1, val(Combo2.ListIndex), Null)
                        RsSavRec.Fields("numbering_type").value = IIf(Combo2.Text <> "", Trim(Combo2.Text), Null)
                        '' '''''''''''''''''''''
                                If CBOYearDigit.ListIndex = 0 Then
                                RsSavRec.Fields("YearDigit").value = 2
                                End If
                                If CBOYearDigit.ListIndex = 1 Then
                                RsSavRec.Fields("YearDigit").value = 4
                                End If
                        RsSavRec.Fields("no_of_digit").value = IIf(Text4.Text <> "", Trim(Text4.Text), Null)
                        RsSavRec.Fields("Prefix").value = IIf(TxtPrefix.Text <> "", Trim(TxtPrefix.Text), Null)
                              
                                If Check1.value = vbChecked Then
                                RsSavRec.Fields("StoreCoding").value = 1
                                Else
                                RsSavRec.Fields("StoreCoding").value = 0
                                End If
                               'Wael
                                If chkIsBreaks.value = vbChecked Then
                                    RsSavRec.Fields("IsBreaks").value = 1
                                Else
                                    RsSavRec.Fields("IsBreaks").value = 0
                                End If
                                   
                                If chkIsCodeByBranch.value = vbChecked Then
                                    RsSavRec.Fields("IsCodeByBranch").value = 1
                                Else
                                    RsSavRec.Fields("IsCodeByBranch").value = 0
                                End If
                                  
                                If ChkIsSerialByUser.value = vbChecked Then
                                    RsSavRec.Fields("IsSerialByUser").value = 1
                                Else
                                    RsSavRec.Fields("IsSerialByUser").value = 0
                                End If
                                     
                                   
                                 RsSavRec.Fields("Breaks").value = IIf(txtBreaks.Text <> "", Trim(txtBreaks.Text), Null)
                                 'Wael
                                   
                                   
                                If Check2.value = vbChecked Then
                                RsSavRec.Fields("zeros").value = 1
                                Else
                                RsSavRec.Fields("zeros").value = 0
                                End If
                                RsSavRec.update
                                 
                        Else
                     
                        
                        End If
                    Next j
        Rs7.MoveNext
    Next i
    '''''''''''''''''''''''''''''''''''
    Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                GetData
                TxtModFlg = "R"
                If SystemOptions.UserInterface = ArabicInterface Then
             Else
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                GetData
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
                Me.Refresh
                FiLLTXT
                GetData
                End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                GetData
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                GetData
                TxtModFlg = "R"
            End If
       End Select
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub
' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    ProgressBar1.Visible = True
    TxtSerial.Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    dcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("branch_no").value), "", RsSavRec.Fields("branch_no").value)
   ' Dcbranch.text = IIf(IsNull(RsSavRec.Fields("BranchName").value), "", RsSavRec.Fields("BranchName").value)
    Combo1.ListIndex = IIf(IsNull(RsSavRec.Fields("sanad_no").value), "", RsSavRec.Fields("sanad_no").value)
    Text3.Text = IIf(IsNull(RsSavRec.Fields("start_at").value), "", RsSavRec.Fields("start_at").value)
    Text2.Text = IIf(IsNull(RsSavRec.Fields("end_at").value), "", RsSavRec.Fields("end_at").value)
    Combo2.Text = IIf(IsNull(RsSavRec.Fields("numbering_type").value), "", RsSavRec.Fields("numbering_type").value)
    Combo2.ListIndex = IIf(IsNull(RsSavRec.Fields("numbering_id").value), -1, RsSavRec.Fields("numbering_id").value)
    
    
      'Wael
    txtBreaks.Text = IIf(IsNull(RsSavRec.Fields("Breaks").value), "", RsSavRec.Fields("Breaks").value)
    If RsSavRec.Fields("IsBreaks").value = True Then
        chkIsBreaks.value = vbChecked
    Else
        chkIsBreaks.value = vbUnchecked
     End If
    If RsSavRec.Fields("IsCodeByBranch").value = True Then
        chkIsCodeByBranch.value = vbChecked
    Else
        chkIsCodeByBranch.value = vbUnchecked
     End If
                             
    If RsSavRec.Fields("IsSerialByUser").value = True Then
        ChkIsSerialByUser.value = vbChecked
    Else
        ChkIsSerialByUser.value = vbUnchecked
     End If
                             
                                 'Wael
    
    If RsSavRec.Fields("YearDigit").value = 2 Then
    CBOYearDigit.ListIndex = 0
    ElseIf RsSavRec.Fields("YearDigit").value = 4 Then
    CBOYearDigit.ListIndex = 1
    Else
    CBOYearDigit.ListIndex = -1
    End If
      
    Text4.Text = IIf(IsNull(RsSavRec.Fields("no_of_digit").value), "", RsSavRec.Fields("no_of_digit").value)
    TxtPrefix.Text = IIf(IsNull(RsSavRec.Fields("Prefix").value), "", RsSavRec.Fields("Prefix").value)
     If RsSavRec.Fields("StoreCoding").value = True Then
     Check1.value = vbChecked
     Else
     Check1.value = vbUnchecked
     End If
     
     If RsSavRec.Fields("zeros").value = True Then
     Check2.value = vbChecked
     Else
     Check2.value = vbUnchecked
     End If
     LabCurrRec.Caption = RsSavRec.AbsolutePosition
     LabCountRec.Caption = RsSavRec.RecordCount
     ProgressBar1.Visible = False
 ProgressBar1.value = 0
ErrTrap:
  ProgressBar1.Visible = False
 ProgressBar1.value = 0
End Sub
    Public Sub GetData()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim Rs1 As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    'Wael
    sql = "SELECT     id, sanad_no, numbering_id, sanad_type, numbering_type,Breaks,IsNull(IsCodeByBranch,0) as IsCodeByBranch,IsNull(IsSerialByUser,0) IsSerialByUser,IsNull(IsBreaks,0) IsBreaks, branch_no, no_of_digit, start_at, zeros, departement, end_at, BranchName, Prefix, StoreCoding, YearDigit"
    'Wael
    sql = sql & "  From dbo.sanad_numbering"
        BolBegine = False
       StrWhere = ""
    ''''' COMBOW BOX SEARCH
        If (Me.Option2.value = True) Then
        If Me.dcBranch.Text <> "" And (val(dcBranch.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  branch_no =" & Me.dcBranch.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where branch_no =" & Me.dcBranch.BoundText & ""
       End If
     End If
     End If
        '-----------------------------------
    sql = sql & StrWhere
    sql = sql & " Order By id"
    Set Rs1 = New ADODB.Recordset
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Rs1.BOF Or Rs1.EOF Then
           Exit Sub
   Else
        With Me.Grid
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = Rs1.RecordCount + .FixedRows
             Rs1.MoveFirst
             For i = .FixedRows To .Rows - 1
                 .TextMatrix(i, .ColIndex("Ser")) = i
                 
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs1("id").value), "", Rs1("id").value): ProgressBar1.value = 20
                 
                 .TextMatrix(i, .ColIndex("BrinchName")) = IIf(IsNull(Rs1("BranchName").value), "", Rs1("BranchName").value): ProgressBar1.value = 30
                 .TextMatrix(i, .ColIndex("Type")) = IIf(IsNull(Rs1("sanad_type").value), "", Rs1("sanad_type").value): ProgressBar1.value = 50
                 .TextMatrix(i, .ColIndex("start")) = IIf(IsNull(Rs1("start_at").value), "", Rs1("start_at").value): ProgressBar1.value = 20
                 .TextMatrix(i, .ColIndex("endwith")) = IIf(IsNull(Rs1("end_at").value), "", Rs1("end_at").value): ProgressBar1.value = 40
                 .TextMatrix(i, .ColIndex("no_of_digit")) = IIf(IsNull(Rs1("no_of_digit").value), "", Rs1("no_of_digit").value): ProgressBar1.value = 30
                 .TextMatrix(i, .ColIndex("Year")) = IIf(IsNull(Rs1("YearDigit").value), "", Rs1("YearDigit").value): ProgressBar1.value = 60
                 .TextMatrix(i, .ColIndex("numberr")) = IIf(IsNull(Rs1("no_of_digit").value), "", Rs1("no_of_digit").value): ProgressBar1.value = 80
                 .TextMatrix(i, .ColIndex("payment")) = IIf(IsNull(Rs1("Prefix").value), "", Rs1("Prefix").value): ProgressBar1.value = 90
                 .TextMatrix(i, .ColIndex("CodingStore")) = IIf(IsNull(Rs1("StoreCoding").value), "", Rs1("StoreCoding").value)
                 .TextMatrix(i, .ColIndex("zeros")) = IIf(IsNull(Rs1("zeros").value), "", Rs1("zeros").value)
                  If (Not IsNull(Rs1("YearDigit").value)) Then
                  If SystemOptions.UserInterface = ArabicInterface Then
                   If Rs1("YearDigit").value = 2 Then
                 .TextMatrix(i, .ColIndex("Year")) = "2 Œ«‰…"
                  ElseIf Rs1("YearDigit").value = 4 Then
                 .TextMatrix(i, .ColIndex("Year")) = "4 Œ«‰« "
                  End If
                  Else
                  If Rs1("YearDigit").value = 2 Then
                 .TextMatrix(i, .ColIndex("Year")) = "2 digit"
                  ElseIf Rs1("YearDigit").value = 4 Then
                 .TextMatrix(i, .ColIndex("Year")) = "4 digit"
                  End If
                  End If
                  End If
                 .TextMatrix(i, .ColIndex("no_of_digit")) = IIf(IsNull(Rs1("numbering_type").value), "", Rs1("numbering_type").value)
               .TextMatrix(i, .ColIndex("Sanad_No")) = IIf(IsNull(Rs1("Sanad_No").value), "", Rs1("Sanad_No").value)
               
               
               'Wael
               .TextMatrix(i, .ColIndex("Breaks")) = IIf(IsNull(Rs1("Breaks").value), "", Rs1("Breaks").value)
                If CBool(Rs1!IsBreaks & "") Then
                    .TextMatrix(i, .ColIndex("IsBreaks")) = 1
                Else
                     .TextMatrix(i, .ColIndex("IsBreaks")) = 0
                End If
                
                If CBool(Rs1!IsCodeByBranch & "") Then
                    .TextMatrix(i, .ColIndex("IsCodeByBranch")) = 1
                Else
                    .TextMatrix(i, .ColIndex("IsCodeByBranch")) = 0
                End If
               
                If CBool(Rs1!IsSerialByUser & "") Then
                    .TextMatrix(i, .ColIndex("IsSerialByUser")) = 1
                Else
                    .TextMatrix(i, .ColIndex("IsSerialByUser")) = 0
                End If
               
                  

                             
                                 'Wael
               
                  Rs1.MoveNext
                 LabCurrRec.Caption = Rs1.AbsolutePosition
                 LabCountRec.Caption = Rs1.RecordCount
                 ProgressBar1.Visible = False
                 ProgressBar1.value = 0
               Next i
          End With
    End If
End Sub

Private Sub Grid_Click()
   On Error GoTo ErrTrap
     FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id")))
     Combo2_Click
ErrTrap:
End Sub
Private Sub Combo2_Change()
Combo2_Click
End Sub
Private Sub ISButton1_Click()
UNSelectedAllRows
End Sub
 Private Sub Option1_Click()
  On Error GoTo ErrTrap
   dcBranch.Enabled = False
   Select Case Me.TxtModFlg.Text
        Case "N"
             Exit Sub
        Case "E"
            Exit Sub
        Case "R"
        GetData
       End Select
    Exit Sub
ErrTrap:
End Sub
Private Sub Option2_Click()
  On Error GoTo ErrTrap
  dcBranch.Enabled = True
   Select Case Me.TxtModFlg.Text
        Case "N"
             Exit Sub
        Case "E"
            Exit Sub
        Case "R"
        GetData
       End Select
    Exit Sub
ErrTrap:
End Sub
Private Sub Dcbranch_Click(Area As Integer)
  On Error GoTo ErrTrap
   Select Case Me.TxtModFlg.Text
        Case "N"
             Exit Sub
        Case "E"
            Exit Sub
        Case "R"
        GetData
       End Select
  Exit Sub
ErrTrap:
End Sub
Function chek(Optional brnch As Integer = 0, Optional Sand As Integer = 0, Optional Prefix As String = "") As Boolean
       Dim bol As Boolean
       bol = False
       Dim rs As ADODB.Recordset
       Dim i As Integer
       Set rs = New ADODB.Recordset
       Dim My_SQL As String
       
       
    If Prefix = "" Then
    My_SQL = "select * from  sanad_numbering where branch_no=" & val(brnch) & "   and  sanad_no= " & Sand
    Else
    My_SQL = "select * from  sanad_numbering where branch_no=" & val(brnch) & "   and  sanad_no= " & Sand & "   and  Prefix='" & Prefix & "'"
    End If
       rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly
       If rs.RecordCount > 0 Then
       bol = True
       Else
       bol = False
       End If
       chek = bol
End Function

' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
    If Option2.value = True Then
        If dcBranch.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «”„ «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            dcBranch.SetFocus
            Exit Sub
            Else
            MsgBox "Write Branch Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            dcBranch.SetFocus
         End If
         End If
         End If
     '''''''''''''''''''''''''''''''''
      If Combo1.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· ‰Ê⁄ «·”‰œ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Combo1.SetFocus
            Exit Sub
            Else
            MsgBox "Write RECEIPT Type ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            Combo1.SetFocus
         End If
     End If
    '+++++++++++++++++++++++++++++++++++++++++++++++
       If Combo2.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· ‰Ê⁄ «· —ﬁÌ„", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Combo2.SetFocus
            Exit Sub
            Else
            MsgBox "Write Numbering type ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            Combo2.SetFocus
         End If
     End If
    ''''''''''''''''''''''''''''''''''''''''''''''
      If Me.Combo2.ListIndex > 0 Then
        If Text3.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· »œ«Ì… «· ﬂÊÌœ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Text3.SetFocus
             Exit Sub
             Else
            MsgBox "Write Start Codeing ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Text3.SetFocus
            Exit Sub
           End If
            End If
     End If
     ''''''''''''''''''''''''''''''''''''''''''
     If Text2.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· ‰Â«Ì… «· ﬂÊÌœ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Text2.SetFocus
             Exit Sub
             Else
            MsgBox "Write End Codeing ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Text2.SetFocus
            Exit Sub
            End If
     End If
    '+++++++++++++++++++++++++++++++++++++++++++++++
       If Text4.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· ⁄œœ «·Œ«‰«  ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Text4.SetFocus
             Exit Sub
      Else
            MsgBox "Write Number of Digits", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Text4.SetFocus
            Exit Sub
            End If
     End If
       '+++++++++++++++++++++++++++++++++++++++++++++++
 '      If TxtPrefix.text = "" Then
 '       If SystemOptions.UserInterface = ArabicInterface Then
 '           MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·Ã“¡ «·À«» ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
 '           TxtPrefix.SetFocus
 '           Exit Sub
 '           Else
 '           MsgBox "Write Fixed Section ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 '           Exit Sub
 '           TxtPrefix.SetFocus
 '        End If
 ''    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++
    If Option2.value = True Then
       Dim rs As ADODB.Recordset
      Dim i As Integer
      Set rs = New ADODB.Recordset
      Dim My_SQL As String
      My_SQL = "select * from  sanad_numbering where branch_no=" & val(dcBranch.BoundText) & "   and  sanad_no= " & Combo1.ListIndex
      
      
      
          If TxtPrefix.Text = "" Then
        My_SQL = "select * from  sanad_numbering where branch_no=" & val(dcBranch.BoundText) & "   and  sanad_no= " & Combo1.ListIndex
    Else
        My_SQL = "select * from  sanad_numbering where branch_no=" & val(dcBranch.BoundText) & "   and  sanad_no= " & Combo1.ListIndex & " and  Prefix='" & TxtPrefix & "'"

    End If
    
    
      rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly
      For i = 1 To rs.RecordCount
      If rs.RecordCount > 0 Then
      If SystemOptions.UserInterface = ArabicInterface Then
      MsgBox "⁄›Ê« ...Â–« «·‰Ê⁄ „‰ «·”‰œ«  „Õœœ  —ﬁÌ„Â „‰ ﬁ»· ·«⁄«œ… «· —ﬁÌ„ ﬁ„ »Õ–›Â À„ ﬁ„ »«÷«› Â „⁄ ‰Ê⁄ «· —ﬁÌ„ «·ÃœÌœ „—… «Œ—Ï", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
    '   MsgBox "⁄›Ê« ...Â–« «·‰Ê⁄ „‰ «·”‰œ«  „Õœœ  —ﬁÌ„Â „‰ ﬁ»· ·«⁄«œ… «· —ﬁÌ„ ﬁ„ »Õ–›Â À„ ﬁ„ »«÷«› Â „⁄ ‰Ê⁄ «· —ﬁÌ„ «·ÃœÌœ „—… «Œ—Ï ›Ì «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
      Exit Sub
      dcBranch.SetFocus
      Else
      MsgBox "this voucher type alraedy defined with numbering method to change delete it and then try to add it again", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
      Exit Sub
      dcBranch.SetFocus
      End If
      End If
      rs.MoveNext
      Next i
      End If
    ''''''''''''''''''''''''''''''''''''''''''''
 
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text
            '------------------------------ new record ----------------------------
        Case "N"
                  '------------------------- save record -----------------------------
        If Me.Option2.value = True Then
'          AddNewRecored
          AddNewRec
          Else
          FiLLRecWithAll
         ' AddNewRecored
        'AddNewRecAll
         End If
        '  BtnLast_Click
        Case "E"
            '----------------------------- save edit -------------------------------
       With Me.Grid
        If .Row <= 0 Then
          MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·”ÿ— «·„ÿ·Ê»  ⁄œÌ·Â", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
           Exit Sub
           Else
            FiLLRec
           End If
       End With
    End Select
    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("sanad_numbering", "id", "")
    RsSavRec.AddNew
    FiLLRec
ErrTrap:
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim RsSavRec As ADODB.Recordset
   On Error GoTo ErrTrap
    Set RsSavRec = New ADODB.Recordset
   My_SQL = "sanad_numbering"
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If RsSavRec.RecordCount > 0 Then
        TxtSerial.Text = RsSavRec.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If
   RsSavRec.Close
ErrTrap:
End Sub
Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
     FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id")))
ErrTrap:
End Sub
 Private Sub ISButton2_Click()
 DeletAllRrc
 End Sub
 Private Sub ISButton3_Click()
 SelectedAllRows
 End Sub
' change id search
Private Sub TxtSerial_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
    FiLLTXT
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "id=" & RecId, , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
   End If
    Exit Function
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If
  End Function
  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click()
   On Error GoTo ErrTrap
    Unload Me
ErrTrap:
End Sub
' undo sub
 Private Sub BtnUndo_Click()
 On Error GoTo ErrTrap
    FindRec val(TxtSerial.Text)
    Me.TxtModFlg.Text = "R"
    FiLLTXT
    GetData
    BtnLast_Click
ErrTrap:
End Sub
' delet sub
Private Sub btnDelete_Click()
    On Error GoTo ErrTrap
     Dim bol As Boolean
     bol = False
     Dim StrMSG As String
     Dim ID As Double
     Dim k As Integer
     Dim i As Integer
     With Me.Grid
       If .Rows < 2 Then Exit Sub
            Dim X As Integer
             If SystemOptions.UserInterface = EnglishInterface Then
            X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
            Else
             X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
            End If
            If X = vbNo Then Exit Sub
             If TxtSerial.Text = "" Then
             If SystemOptions.UserInterface = EnglishInterface Then
              X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
              Else
              X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
              End If
              Else
              k = Grid.Rows - 1
                For i = 1 To Grid.Rows - 1
                If .Cell(flexcpChecked, k, .ColIndex("checkid")) = flexChecked Then
                ID = val(Grid.TextMatrix(k, Grid.ColIndex("id")))
                StrSQL = "Delete From sanad_numbering Where id=" & ID & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                Grid.RemoveItem k
                bol = True
                End If
                k = k - 1
                Next i
                cleargriid
                FiLLTXT
                GetData
                If bol = True Then
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
                Else
                X = MsgBox("⁄›Ê« .... ·„  ﬁ„ » ÕœÌœ «·”Ã· √Ê «·”Ã·«  «·„—«œ Õ–›Â«", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
                End If
                End If
                TxtModFlg = "R"
     End With
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
           Cn.Errors.Clear
    End Select
End Sub
Private Sub DeletAllRrc()
On Error GoTo ErrTrap
   Dim StrMSG As String
     Dim ID As Double
     Dim i As Integer
     With Me.Grid
        If .Rows < 2 Then Exit Sub
             Dim X As Integer
             If SystemOptions.UserInterface = EnglishInterface Then
             X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
              Else
             X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
             End If
             If X = vbNo Then Exit Sub
             If TxtSerial.Text = "" Then
             If SystemOptions.UserInterface = EnglishInterface Then
              X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
              Else
              X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
              End If
              Else
              
              For i = 1 To Grid.Rows - 1
              If val(.TextMatrix(i, .ColIndex("checkid"))) = True Then
              ID = val(Grid.TextMatrix(i, Grid.ColIndex("id")))
              StrSQL = "Delete From sanad_numbering Where id=" & ID & ""
              Cn.Execute StrSQL, , adExecuteNoRecords
              Else
              X = MsgBox("⁄›Ê« .... ·„  ﬁ„ » ÕœÌœ «·ﬂ·", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
              Exit Sub
              End If
              Next i
              Me.Grid.Clear flexClearScrollable, flexClearEverything
              Me.Grid.Rows = 2
              X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
              End If
       End With
ErrTrap:
  Select Case Err.Number
        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
           Cn.Errors.Clear
    End Select
End Sub
Sub ChekSelectedRows()
    On Error GoTo ErrTrap
    Dim Selrow As Integer
    Dim IdRow As Integer
    Dim DelRow As Integer
    Dim RowSelected As Integer
    Dim Rs1 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset
    Dim sql As String
    sql = "SELECT     id, sanad_no, numbering_id, sanad_type, numbering_type, branch_no, no_of_digit, start_at, zeros, departement, end_at, BranchName, Prefix, StoreCoding, YearDigit"
    sql = sql & "  From dbo.sanad_numbering"
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
    Rs1.MoveFirst
    End If
    With Me.Grid
                  Selrow = True
                  For DelRow = .FixedRows To .Rows - 1
                  If .Cell(flexcpChecked, DelRow, .ColIndex("checkid")) = flexChecked Then
                ' .TextMatrix(DelRow, .ColIndex("checkid")) = Selrow
                  Rs1.Find "id=" & val(.TextMatrix(DelRow, .ColIndex("id")))
                  Rs1.delete
                  Else
                  Rs1.MoveNext
                  End If
                  Next DelRow
    End With
         Exit Sub
ErrTrap:
  End Sub
  Sub SelectedAllRows()
  On Error GoTo ErrTrap
    Dim Selrow As Integer
    Dim IdRow As Integer
    Dim DelRow As Integer
    Dim RowSelected As Integer
    Dim Rs1 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset
    Dim sql As String
    sql = "SELECT     id, sanad_no, numbering_id, sanad_type, numbering_type, branch_no, no_of_digit, start_at, zeros, departement, end_at, BranchName, Prefix, StoreCoding, YearDigit"
    sql = sql & "  From dbo.sanad_numbering"
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
    Rs1.MoveFirst
    End If
    With Me.Grid
                  Selrow = True
                  For DelRow = .FixedRows To .Rows - 1
                  If val(.TextMatrix(DelRow, .ColIndex("checkid"))) = vbUnchecked Then
                 .TextMatrix(DelRow, .ColIndex("checkid")) = Selrow
                 .RowSel = DelRow
                  Else
                  Rs1.MoveNext
                  End If
                  Next DelRow
    End With
         Exit Sub
ErrTrap:
  End Sub
 Sub UNSelectedAllRows()
    On Error GoTo ErrTrap
    Dim Selrow As Integer
    Dim IdRow As Integer
    Dim DelRow As Integer
    Dim Rs1 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset
    Dim sql As String
    sql = "SELECT     id, sanad_no, numbering_id, sanad_type, numbering_type, branch_no, no_of_digit, start_at, zeros, departement, end_at, BranchName, Prefix, StoreCoding, YearDigit"
    sql = sql & "  From dbo.sanad_numbering"
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
    Rs1.MoveFirst
    End If
    With Me.Grid
                  Selrow = False
                  For DelRow = .FixedRows To .Rows - 1
                  If val(.TextMatrix(DelRow, .ColIndex("checkid"))) = True Then
                  .TextMatrix(DelRow, .ColIndex("checkid")) = Selrow
                 .RowSel = False
                  Else
                  Rs1.MoveNext
                 End If
               Next DelRow
    End With
         Exit Sub
ErrTrap:
  End Sub
' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
               btnSave_Click
        Case vbCancel
              Cancel = True
        End Select
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Terminate()
     ' Set FrmVacancy = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If
        RsSavRec.Close
        Set RsSavRec = Nothing
    End If
ErrTrap:
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub EditRec(StrTable As String, _
                   RecId As String)
     FiLLRec
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        ISButton2.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
       ' ISButton1.Enabled = False
    ElseIf TxtModFlg.Text = "R" Then
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        ISButton2.Enabled = False
        If TxtSerial.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
            ISButton2.Enabled = True
    End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.Text = "E" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        ISButton2.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If
End Sub
' move btowen recored
Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial.Text)
        Me.TxtModFlg.Text = "R"
        End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
     End If
BegnieWork:
    RsSavRec.MoveLast
     FiLLTXT
     Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial.Text <> "" Then
        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
      '  Me.Dcbranch.BoundText = branch_id
        Frm2.Enabled = True
        Me.Combo1.SetFocus
        End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
                    If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
                 End If
    End Select
End Sub
Private Sub btnNew_Click()
 On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
 '   clear_all Me
    cleargriid
    ClearTXTandCO
    TxtModFlg.Text = "N"
    Me.DCboUserName.BoundText = user_id
'    CmbType.ListIndex = CmbType.ListIndex + 1
    Text3.Text = 1
Text2.Text = 0
Text4.Text = 3

CBOYearDigit.ListIndex = 0
    Option1.value = True
  '  Combo1.ListIndex = Combo1.ListIndex + 1
    
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
         clear_all Me
        Exit Sub
    End If
BegnieWork:
     If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext
        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If
      FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
     FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
'Information for camand
'++++++++++++++++++++++++++++++++++++++
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With
ErrTrap:
End Sub
' short cut for keys
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrTrap
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            btnNew_Click
        Else
            SendKeys "{TAB}"
        End If
    End If
    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If
    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If
    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If
    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If
    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If
    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If
    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
        BtnFirst_Click
    End If
    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If
    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If
    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
    End If
    'End If
    Exit Sub
ErrTrap:
End Sub
' print Events
'++++++++++++++++++++++++++++++++++++++++++
Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
    Me.Caption = "Bonds Coding"
    ' labell name
    '''''''''''''' next
    Me.Label1(2).Caption = "Bonds Coding"
    Me.Label1(3).Caption = "Code"
    Me.Option1.Caption = "All"
    Me.Option2.Caption = "Selected Branch"
    Me.Label4.Caption = "Receipt Type"
    Me.Label7.Caption = "Started from"
    Me.Label9.Caption = "End in"
    Me.Label10.Caption = "Numbering type"
    Me.Label33.Caption = "Year Code"
    Me.Label1(0).Caption = "Number Of Digits"
    Me.Label8.Caption = "Fixed Section"
    Me.Label5.Caption = "Coding According the Store"
    Me.Label6.Caption = "Full Zero"
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    ISButton3.Caption = "Select All"
    ISButton1.Caption = "Un Select All"
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete Selected"
    ISButton2.Caption = "Delete All"
    btnCancel.Caption = "Exit"
                  
    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
      .TextMatrix(0, .ColIndex("Sanad_No")) = "Sanad_No"
        
        .TextMatrix(0, .ColIndex("checkid")) = "Selected"
        .TextMatrix(0, .ColIndex("id")) = "Code"
        .TextMatrix(0, .ColIndex("BrinchName")) = "Branch"
        .TextMatrix(0, .ColIndex("Type")) = "Receipt Type"
        .TextMatrix(0, .ColIndex("start")) = "Started from"
        .TextMatrix(0, .ColIndex("endwith")) = "End in"
        .TextMatrix(0, .ColIndex("no_of_digit")) = "Numbering type"
        .TextMatrix(0, .ColIndex("Year")) = "Year Code"
        .TextMatrix(0, .ColIndex("numberr")) = "Number Of Digits"
        .TextMatrix(0, .ColIndex("payment")) = "Fixed Section"
        .TextMatrix(0, .ColIndex("CodingStore")) = "Coding According the Store"
        .TextMatrix(0, .ColIndex("zeros")) = "Full Zero"
        End With
ErrTrap:
End Sub
Private Sub cleargriid()
On Error GoTo ErrTrap
Me.Grid.Rows = 1
ErrTrap:
End Sub
 Private Sub ClearTXTandCO()
 
On Error GoTo ErrTrap
TxtSerial.Text = ""
  Exit Sub
dcBranch.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text3.Text = ""
Text2.Text = ""
Text4.Text = ""
TxtPrefix.Text = ""
CBOYearDigit.Text = ""
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end
' key press
Private Sub Dcbranch_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  Combo1.SetFocus
  End If
ErrTrap:
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  Combo2.SetFocus
  End If
ErrTrap:
End Sub
Private Sub CBOYearDigit_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  Text3.SetFocus
  End If
ErrTrap:
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
   If KeyAscii = 13 Then
    Text2.SetFocus
    Else
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    KeyAscii = 0
    End If
    End If
ErrTrap:
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  Text4.SetFocus
 Else
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    KeyAscii = 0
    End If
    End If
ErrTrap:
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  TxtPrefix.SetFocus
  Else
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    KeyAscii = 0
    End If
    End If
ErrTrap:
End Sub
Private Sub Combo2_Click()
On Error GoTo ErrTrap
     If val(Combo2.ListIndex) = 0 Or val(Combo2.ListIndex) = -1 Then
        Text3.Enabled = False
        Text3.Text = 1
    Else
        Text3.Enabled = True
    End If
    If Combo2.ListIndex = 2 Or Combo2.ListIndex = 3 Then
       CBOYearDigit.Visible = True
       Label33.Visible = True
    Else
        CBOYearDigit.Visible = False
        Label33.Visible = False
        CBOYearDigit.ListIndex = -1
    End If
ErrTrap:
    End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  If CBOYearDigit.Visible = False Then
  Text4.SetFocus
  Else
  CBOYearDigit.SetFocus
  End If
  End If
ErrTrap:
End Sub
Private Sub TxtPrefix_KeyPress(KeyAscii As Integer)
On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  Call btnSave_Click
  End If
ErrTrap:
End Sub
Private Sub AdditemTocCmp()
 On Error GoTo ErrTrap
   If SystemOptions.UserInterface = EnglishInterface Then
    With Me.Combo2
        .Clear
        .AddItem "Manual"
        .AddItem "Automatic"
        .AddItem "Monthly"
        .AddItem "Yearly"
      End With
        With Me.CBOYearDigit
        .Clear
        .AddItem "2 Digit"
        .AddItem "4 Digit"
        End With
    Else
    With Me.Combo2
        .Clear
        .AddItem "ÌœÊÌ"
        .AddItem "¬·Ì"
        .AddItem "„ ’· ‘Â—Ì"
        .AddItem "„ ’· ”‰ÊÌ"
      End With
        With Me.CBOYearDigit
        .Clear
        .AddItem "2 Œ«‰…"
        .AddItem "4 Œ«‰« "
      End With
    End If
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end








