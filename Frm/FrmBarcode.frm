VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{A3550A07-56EC-11D3-8DC5-00409503C9B8}#1.0#0"; "axbarcode.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{85FD608E-54A8-11D4-8ED4-00E07D815373}#1.0#0"; "MBClrPkr.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmBarcode 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ’„Ì„ «·»«—ﬂÊœ"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10410
   HelpContextID   =   30
   Icon            =   "FrmBarcode.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   10410
   Begin C1SizerLibCtl.C1Elastic ELeMain 
      Height          =   7470
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10410
      _cx             =   18362
      _cy             =   13176
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
      Align           =   5
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
      Begin VB.CheckBox chkIsSomeItemWeight 
         Alignment       =   1  'Right Justify
         Caption         =   "ÌÊÃœ «’‰«›  ⁄„· »«·Ê“‰"
         Height          =   405
         Left            =   8490
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   1110
         Width           =   1725
      End
      Begin VB.Frame Frame1 
         Height          =   1485
         Left            =   3180
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1440
         Width           =   7035
         Begin VB.TextBox txtWeightTo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   630
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   990
            Width           =   885
         End
         Begin VB.TextBox txtWeightFrom 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4140
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   990
            Width           =   885
         End
         Begin VB.TextBox txtCodeTo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   630
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   600
            Width           =   885
         End
         Begin VB.TextBox txtCodeFrom 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4140
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   600
            Width           =   885
         End
         Begin VB.TextBox txtOrNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   630
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   240
            Width           =   885
         End
         Begin VB.TextBox txtFromNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4140
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   240
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "«·Ï"
            Height          =   225
            Index           =   22
            Left            =   2010
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "Ê“‰ «·’‰› Ì»œ√ „‰ "
            Height          =   225
            Index           =   21
            Left            =   5340
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "«·Ï"
            Height          =   225
            Index           =   20
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   570
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "ﬂÊœ «·’‰› Ì»œ√ „‰ "
            Height          =   225
            Index           =   19
            Left            =   5370
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   570
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "√Ê"
            Height          =   225
            Index           =   18
            Left            =   2070
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   210
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   " »œ√ »"
            Height          =   225
            Index           =   17
            Left            =   5700
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   240
            Width           =   1005
         End
      End
      Begin AXBARCODELib.Axbarcode Axbarcode1 
         Height          =   585
         Left            =   3300
         TabIndex        =   54
         Top             =   2940
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   1032
         _StockProps     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   3225
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   4200
         Width           =   6975
         Begin VB.CheckBox margins 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√ŸÂ— ⁄·«„«  «·„Õ«–«…"
            Height          =   285
            Left            =   4770
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   2520
            Width           =   1785
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "··Ì„Ì‰"
            Height          =   255
            Index           =   2
            Left            =   5580
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   2190
            Width           =   765
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‰ ’›"
            Height          =   255
            Index           =   1
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   2190
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "··Ì”«—"
            Height          =   255
            Index           =   0
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   2190
            Width           =   765
         End
         Begin VB.CheckBox nominal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÕÃ„ ÿ»Ì⁄Ï"
            Height          =   195
            Left            =   3030
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   1560
            Width           =   1365
         End
         Begin VB.CheckBox showtext 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√ŸÂ— «·ﬂÊœ"
            Height          =   225
            Left            =   4740
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   1545
            Value           =   1  'Checked
            Width           =   1665
         End
         Begin VB.CheckBox autocheck 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂ‘›  ·ﬁ«∆Ì ··—ﬁ„"
            Height          =   255
            Left            =   1170
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   1560
            Width           =   1665
         End
         Begin VB.TextBox bearthick 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   690
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   510
            Width           =   585
         End
         Begin VB.TextBox margin 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   660
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   1170
            Width           =   585
         End
         Begin VB.CheckBox Exbearers 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÕœÊœ „„ œ…"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   690
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   840
            Width           =   1815
         End
         Begin VB.CheckBox BothBearers 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœ √⁄·Ï Ê√”›·"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   690
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox CboScaleType 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   3900
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1080
            Width           =   2445
         End
         Begin VB.TextBox Heightt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   5310
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   480
            Width           =   555
         End
         Begin VB.TextBox Widtht 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   3630
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   510
            Width           =   555
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   345
            Index           =   9
            Left            =   4590
            TabIndex        =   46
            Top             =   2790
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "16"
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBarcode.frx":038A
            ColorButton     =   12632256
            DrawFocusRectangle=   0   'False
         End
         Begin MBColorPicker.ColorPicker CPic 
            Height          =   345
            Index           =   0
            Left            =   270
            TabIndex        =   47
            ToolTipText     =   "·Ê‰ Œÿ «·»«—ﬂÊœ"
            Top             =   1920
            Width           =   915
            _ExtentX        =   1773
            _ExtentY        =   556
            CustomButtonText=   " Œ’Ì’"
            Color           =   17
            Style           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumColors       =   64
            Color1          =   0
            Color2          =   128
            Color3          =   32768
            Color4          =   32896
            Color5          =   8388608
            Color6          =   8388736
            Color7          =   8421376
            Color8          =   12632256
            Color9          =   8421504
            Color10         =   255
            Color11         =   65280
            Color12         =   65535
            Color13         =   16711680
            Color14         =   16711935
            Color15         =   16776960
            Color18         =   12632319
            Color19         =   12640511
            Color20         =   12648447
            Color21         =   12648384
            Color22         =   16777152
            Color23         =   16761024
            Color24         =   16761087
            Color25         =   14737632
            Color26         =   8421631
            Color27         =   8438015
            Color28         =   8454143
            Color29         =   8454016
            Color30         =   16777088
            Color31         =   16744576
            Color32         =   16744703
            Color33         =   12632256
            Color34         =   255
            Color35         =   33023
            Color36         =   65535
            Color37         =   65280
            Color38         =   16776960
            Color39         =   16711680
            Color40         =   16711935
            Color41         =   8421504
            Color42         =   192
            Color43         =   16576
            Color44         =   49344
            Color45         =   49152
            Color46         =   12632064
            Color47         =   12582912
            Color48         =   12583104
            Color49         =   4210752
            Color50         =   128
            Color51         =   16512
            Color52         =   32896
            Color53         =   32768
            Color54         =   8421376
            Color55         =   8388608
            Color56         =   8388736
            Color57         =   0
            Color58         =   64
            Color59         =   4210816
            Color60         =   16448
            Color61         =   16384
            Color62         =   4210688
            Color63         =   4194304
            Color64         =   4194368
         End
         Begin MBColorPicker.ColorPicker CPic 
            Height          =   345
            Index           =   1
            Left            =   270
            TabIndex        =   48
            ToolTipText     =   "·Ê‰ Œ·›Ì… «·»«—ﬂÊœ"
            Top             =   2310
            Width           =   915
            _ExtentX        =   1773
            _ExtentY        =   556
            CustomButtonText=   " Œ’Ì’"
            Style           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumColors       =   64
            Color1          =   0
            Color2          =   128
            Color3          =   32768
            Color4          =   32896
            Color5          =   8388608
            Color6          =   8388736
            Color7          =   8421376
            Color8          =   12632256
            Color9          =   8421504
            Color10         =   255
            Color11         =   65280
            Color12         =   65535
            Color13         =   16711680
            Color14         =   16711935
            Color15         =   16776960
            Color18         =   12632319
            Color19         =   12640511
            Color20         =   12648447
            Color21         =   12648384
            Color22         =   16777152
            Color23         =   16761024
            Color24         =   16761087
            Color25         =   14737632
            Color26         =   8421631
            Color27         =   8438015
            Color28         =   8454143
            Color29         =   8454016
            Color30         =   16777088
            Color31         =   16744576
            Color32         =   16744703
            Color33         =   12632256
            Color34         =   255
            Color35         =   33023
            Color36         =   65535
            Color37         =   65280
            Color38         =   16776960
            Color39         =   16711680
            Color40         =   16711935
            Color41         =   8421504
            Color42         =   192
            Color43         =   16576
            Color44         =   49344
            Color45         =   49152
            Color46         =   12632064
            Color47         =   12582912
            Color48         =   12583104
            Color49         =   4210752
            Color50         =   128
            Color51         =   16512
            Color52         =   32896
            Color53         =   32768
            Color54         =   8421376
            Color55         =   8388608
            Color56         =   8388736
            Color57         =   0
            Color58         =   64
            Color59         =   4210816
            Color60         =   16448
            Color61         =   16384
            Color62         =   4210688
            Color63         =   4194304
            Color64         =   4194368
         End
         Begin ImpulseButton.ISButton CmdDef 
            Height          =   435
            Left            =   270
            TabIndex        =   49
            Top             =   2700
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " Õ„Ì· «·√› —«÷Ì« "
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
            ButtonImage     =   "FrmBarcode.frx":0724
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "·Ê‰ «·Œ·›Ì…"
            Height          =   225
            Index           =   16
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   2340
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "·Ê‰ «·ŒÿÊÿ"
            Height          =   225
            Index           =   15
            Left            =   1470
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   1950
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "‰Ê⁄ «·Œÿ"
            Height          =   225
            Index           =   14
            Left            =   5580
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   2850
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "„Õ«–«… ‰’ «·ﬂÊœ"
            Height          =   225
            Index           =   13
            Left            =   4890
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   1890
            Width           =   1785
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "’›— =·«ÌÊÃœ Õœ √œ‰Ï"
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   12
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   360
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   2
            Left            =   4080
            Picture         =   "FrmBarcode.frx":0ABE
            Top             =   2190
            Width           =   240
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   3
            Left            =   5310
            Picture         =   "FrmBarcode.frx":0E48
            Top             =   2190
            Width           =   240
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   4
            Left            =   6390
            Picture         =   "FrmBarcode.frx":11D2
            Top             =   2190
            Width           =   240
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Â«„‘"
            Height          =   255
            Index           =   11
            Left            =   1410
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”„ﬂ «·Õœ"
            Height          =   255
            Index           =   10
            Left            =   1410
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "‰Ê⁄ «·„ﬁÌ«”"
            Height          =   225
            Index           =   9
            Left            =   5340
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   810
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "⁄—÷"
            Height          =   225
            Index           =   8
            Left            =   4170
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   540
            Width           =   495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "„·„"
            Height          =   225
            Index           =   7
            Left            =   3090
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   510
            Width           =   495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "„·„"
            Height          =   225
            Index           =   6
            Left            =   4770
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   510
            Width           =   495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "≈— ›«⁄"
            Height          =   225
            Index           =   5
            Left            =   5910
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   510
            Width           =   495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "ÕÃ„ «·«” Ìﬂ—"
            Height          =   225
            Index           =   3
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   150
            Width           =   1215
         End
      End
      Begin VB.PictureBox PicCopy 
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         ScaleHeight     =   285
         ScaleWidth      =   525
         TabIndex        =   22
         Top             =   3420
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.PictureBox Pic 
         AutoRedraw      =   -1  'True
         Height          =   345
         Left            =   720
         RightToLeft     =   -1  'True
         ScaleHeight     =   285
         ScaleWidth      =   525
         TabIndex        =   21
         Top             =   3420
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   60
         Top             =   1260
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox error1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   3300
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   3810
         Width           =   3765
      End
      Begin VB.TextBox TxtMsg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2700
         TabIndex        =   11
         Top             =   420
         Width           =   2775
      End
      Begin VB.ComboBox CboBarcodes 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2700
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   60
         Width           =   2775
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   7395
         Left            =   7080
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   3225
         _cx             =   5689
         _cy             =   13044
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   2
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ›Ÿ ’Ê—…..."
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
            ButtonImage     =   "FrmBarcode.frx":155C
            ColorButton     =   14871017
            ColorHoverText  =   16711680
            DrawFocusRectangle=   0   'False
            RightToLeft     =   -1  'True
            ColorToggledHoverText=   16711680
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   2
            Left            =   60
            TabIndex        =   3
            Top             =   2970
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "≈⁄œ«œ ÿ«»⁄… »«—ﬂÊœ"
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
            ButtonImage     =   "FrmBarcode.frx":18F6
            ColorButton     =   14871017
            ColorHoverText  =   16711680
            Alignment       =   1
            DrawFocusRectangle=   0   'False
            RightToLeft     =   -1  'True
            ColorToggledHoverText=   16711680
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   4
            Top             =   1050
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   873
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "‰”Œ ﬂ‹ wmf"
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
            ButtonImage     =   "FrmBarcode.frx":1C90
            ColorButton     =   14871017
            ColorHoverText  =   16711680
            DrawFocusRectangle=   0   'False
            RightToLeft     =   -1  'True
            ColorToggledHoverText=   16711680
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   5
            Top             =   750
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "‰”Œ ﬂ‹ bmp"
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
            ButtonImage     =   "FrmBarcode.frx":202A
            ColorButton     =   14871017
            ColorHoverText  =   16711680
            DrawFocusRectangle=   0   'False
            RightToLeft     =   -1  'True
            ColorToggledHoverText=   16711680
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   6
            Top             =   6060
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
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
            ButtonImage     =   "FrmBarcode.frx":23C4
            ColorButton     =   14871017
            ColorHoverText  =   16711680
            DrawFocusRectangle=   0   'False
            RightToLeft     =   -1  'True
            ColorToggledHoverText=   16711680
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   7
            Left            =   60
            TabIndex        =   7
            Top             =   3390
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… »«—ﬂÊœ"
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
            ButtonImage     =   "FrmBarcode.frx":275E
            ColorButton     =   14871017
            ColorHoverText  =   16711680
            Alignment       =   1
            DrawFocusRectangle=   0   'False
            RightToLeft     =   -1  'True
            ColorToggledHoverText=   16711680
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   8
            Left            =   60
            TabIndex        =   8
            Top             =   4170
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "„⁄«‰Ì… ÿ«»⁄… ⁄«œÌ…"
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
            ButtonImage     =   "FrmBarcode.frx":2AF8
            ColorButton     =   14871017
            ColorHoverText  =   16711680
            Alignment       =   1
            DrawFocusRectangle=   0   'False
            RightToLeft     =   -1  'True
            ColorToggledHoverText=   16711680
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   20
            Left            =   60
            TabIndex        =   9
            Top             =   5310
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   661
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
            ButtonImage     =   "FrmBarcode.frx":2E92
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   285
         Index           =   10
         Left            =   2370
         TabIndex        =   12
         Top             =   30
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         ButtonStyle     =   1
         ButtonPositionImage=   2
         Caption         =   ""
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBarcode.frx":322C
         ColorButton     =   12632256
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   0
         Left            =   6810
         Picture         =   "FrmBarcode.frx":35C6
         Top             =   3510
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   2055
         Left            =   630
         Stretch         =   -1  'True
         Top             =   1290
         Width           =   2505
      End
      Begin VB.Image Img 
         Height          =   1215
         Index           =   5
         Left            =   0
         Picture         =   "FrmBarcode.frx":3950
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Õ«·…"
         Height          =   225
         Index           =   4
         Left            =   5970
         TabIndex        =   19
         Top             =   3570
         Width           =   735
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   1
         Left            =   6810
         Picture         =   "FrmBarcode.frx":42DE
         Top             =   3510
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   7
         Left            =   2400
         Picture         =   "FrmBarcode.frx":4668
         Top             =   390
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   6
         Left            =   2370
         Picture         =   "FrmBarcode.frx":4BF2
         Top             =   390
         Width           =   240
      End
      Begin VB.Label lblMax 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   1770
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   780
         Width           =   3705
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·»«—ﬂÊœ «·„” Œœ„"
         Height          =   315
         Index           =   0
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   60
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·’‰› «·„—«œ"
         Height          =   315
         Index           =   1
         Left            =   5460
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   420
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·»Ì«‰«  «· Ï  ﬁ»· "
         Height          =   315
         Index           =   2
         Left            =   5490
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   810
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog Cdg 
      Left            =   2010
      Top             =   3330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const twipFactor = 1440

Private Const WM_PAINT = &HF

Private Const WM_PRINT = &H317

Private Const PRF_CLIENT = &H4&    ' Draw the window's client area.

Private Const PRF_CHILDREN = &H10& ' Draw all visible child windows.

Private Const PRF_OWNED = &H20&    ' Draw all owned windows.

Private m_ReportsNumber As Integer

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

Private Sub autocheck_Click()

    DoBarcode ' update barcode image
End Sub

Private Sub bearthick_Change()

    DoBarcode ' update barcode image
End Sub

Private Sub BothBearers_Click()

    DoBarcode ' update barcode image
End Sub

Private Sub CboBarcodes_Click()

    DoBarcode ' update barcode image
    GetType
End Sub

Private Sub CboScaleType_Click()

    DoBarcode

    With CboScaleType

        If .ItemData(.ListIndex) = 6 Then
            lbl(6).Caption = "„·„"
            lbl(7).Caption = "„·„"
        ElseIf .ItemData(.ListIndex) = 7 Then
            lbl(6).Caption = "”„"
            lbl(7).Caption = "”„"
        ElseIf .ItemData(.ListIndex) = 5 Then
            lbl(6).Caption = "»Ê’…"
            lbl(7).Caption = "»Ê’…"
        ElseIf .ItemData(.ListIndex) = 2 Then
            lbl(6).Caption = "‰ﬁÿ…"
            lbl(7).Caption = "‰ﬁÿ…"
        ElseIf .ItemData(.ListIndex) = 1 Then
            lbl(6).Caption = " ÊÌ»"
            lbl(7).Caption = " ÊÌ»"
        End If

    End With

End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim i As Integer
    Dim Msg As String
    Dim StrSavePath As String
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If Axbarcode1.Orientation = 0 Then
                Axbarcode1.Orientation = 1
            ElseIf Axbarcode1.Orientation = 1 Then
                Axbarcode1.Orientation = 3
            ElseIf Axbarcode1.Orientation = 3 Then
                Axbarcode1.Orientation = 2
            ElseIf Axbarcode1.Orientation = 2 Then
                Axbarcode1.Orientation = 0
            End If

            DoBarcode

        Case 1

            If Axbarcode1.Picture = 0 Then
                GetMsgs 158, vbExclamation
                Exit Sub
            End If

            With cdg
                .CancelError = False
                .filter = "Metafile (*.wmf)|*.wmf|Bitmap (*.bmp)|*.bmp|Paintbrush (*.pcx)|*.pcx|Encapsulated PostScript (*.eps)|*.eps|Portable Network Graphic (*.png)|*.png| GIF (*.gif)|*.gif"  ' choose formats to include
                'Specify default filter
                .Flags = cdlOFNExtensionDifferent + cdlOFNLongNames + cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNHideReadOnly
                'CommonDialog1.FilterIndex = 1
                .ShowSave
                StrSavePath = .filename

                If (Len(StrSavePath) > 1) Then i = Axbarcode1.saveimage(StrSavePath)
            End With

        Case 2
            cdg.CancelError = False
            cdg.ShowPrinter

        Case 3

            '‰”Œ ’Ê—… „‰ «·»«—ﬂÊœ ≈·Ï Õ«›Ÿ… «·ÊÌ‰œÊ“
            'MeteFile
            If Axbarcode1.Picture = 0 Then
                GetMsgs 158, vbExclamation
                Exit Sub
            End If

            i = Axbarcode1.CopyImage()

            If i = 0 Then
                Msg = "⁄›Ê«"
                Msg = Msg & CHR(13) & "›‘·  ‰”Œ ’Ê—… «·»«—ﬂÊ ≈·Ï"
                Msg = Msg & CHR(13) & "Õ«›Ÿ… «·ÊÌ‰œÊ“"
                Msg = Msg & CHR(13) & "»—Ã«¡ „—«Ã⁄… «·ﬂÊœ "
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                Msg = " ‰ÃÕ  ⁄„·Ì… ‰”Œ ’Ê—… «·»«—ﬂÊ ≈·Ï Õ«›Ÿ… «·ÊÌ‰œÊ“"
                Msg = Msg & CHR(13) & "(MeteFile)⁄·Ï «·‰”ﬁ "
                Msg = Msg & CHR(13) & "Ì„ﬂ‰ﬂ «·√‰ › Õ «Ï »—‰«„Ã „À· »—‰«„Ã «·—”«„  "
                Msg = Msg & CHR(13) & "√Ê »—‰«„Ã „Ìﬂ—Ê”Ê›  Ê——œ "
                Msg = Msg & CHR(13) & "Ê⁄„· ·’ﬁ ··’Ê—… „‰ Õ«›Ÿ… «·ÊÌ‰œÊ“ ›Ï «Ï „” ‰œ"
                MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If

        Case 5

            '‰”Œ ’Ê—… „‰ «·»«—ﬂÊœ ≈·Ï Õ«›Ÿ… «·ÊÌ‰œÊ“
            'Bitmap
            If Axbarcode1.Picture = 0 Then
                Msg = "⁄›Ê«"
                Msg = Msg & CHR(13) & "·«Ì„ﬂ‰ ‰”Œ ’Ê—… «·»«—ﬂÊœ ÊÂÏ ›«—€…..!"
                Msg = Msg & CHR(13) & "»—Ã«¡ ﬂ «»… ﬂÊœ «Ê „—«Ã⁄… «·ﬂÊœ «·„œŒ·. "
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            i = Axbarcode1.CopyBitmap()

            If i = 0 Then
                Msg = "⁄›Ê«"
                Msg = Msg & CHR(13) & "›‘·  ‰”Œ ’Ê—… «·»«—ﬂÊ ≈·Ï"
                Msg = Msg & CHR(13) & "Õ«›Ÿ… «·ÊÌ‰œÊ“"
                Msg = Msg & CHR(13) & "»—Ã«¡ „—«Ã⁄… «·ﬂÊœ "
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                Msg = " ‰ÃÕ  ⁄„·Ì… ‰”Œ ’Ê—… «·»«—ﬂÊ ≈·Ï Õ«›Ÿ… «·ÊÌ‰œÊ“"
                Msg = Msg & CHR(13) & "(Bitmap)⁄·Ï «·‰”ﬁ "
                Msg = Msg & CHR(13) & "Ì„ﬂ‰ﬂ «·√‰ › Õ «Ï »—‰«„Ã „À· »—‰«„Ã «·—”«„  "
                Msg = Msg & CHR(13) & "√Ê »—‰«„Ã „Ìﬂ—Ê”Ê›  Ê——œ "
                Msg = Msg & CHR(13) & "Ê⁄„· ·’ﬁ ··’Ê—… „‰ Õ«›Ÿ… «·ÊÌ‰œÊ“ ›Ï «Ï „” ‰œ"
                MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If

        Case 4

            If Axbarcode1.Orientation = 0 Then
                Axbarcode1.Orientation = 2
            ElseIf Axbarcode1.Orientation = 1 Then
                Axbarcode1.Orientation = 0
            ElseIf Axbarcode1.Orientation = 3 Then
                Axbarcode1.Orientation = 1
            ElseIf Axbarcode1.Orientation = 2 Then
                Axbarcode1.Orientation = 3
            End If

            DoBarcode

        Case 6
            Unload Me

        Case 7
            PrintBarcode1

        Case 8
            PrintBarcode2

        Case 9

            With cdg
                .CancelError = False
                .Flags = cdlCFBoth  ' choose fonts to include
                .FontName = Axbarcode1.FontName
                .fontsize = Axbarcode1.fontsize
                .FontBold = Axbarcode1.FontBold
                .FontItalic = Axbarcode1.FontItalic
                .ShowFont
                Axbarcode1.FontName = .FontName
                Axbarcode1.fontsize = .fontsize
                Axbarcode1.FontBold = .FontBold
                Axbarcode1.FontItalic = .FontItalic
            End With

            DoBarcode ' update barcode image

        Case 10
            ShowInfo

        Case 20
            SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
            SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Cmd_MouseEnter(Index As Integer)
    Cmd(Index).backcolor = &HC0FFFF
End Sub

Private Sub Cmd_MouseLeave(Index As Integer)
    Cmd(Index).backcolor = &HE2E9E9
End Sub

Private Sub CmdDef_Click()
    Dim Msg As String
    Dim IntRes As Integer
    On Error GoTo ErrTrap

    Msg = "”Ê› Ì „  Õ„Ì· «·√⁄œ«œ«  «·√› —«÷Ì…" & CHR(13)
    Msg = Msg + "Â· «‰  „ «ﬂœ „‰ «·√” „—«— .øø"
    IntRes = MsgBox(Msg, vbOKCancel + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

    If IntRes = vbOK Then
        Heightt.Text = "20"
        Widtht.Text = "40"
        
        TxtFromNo = 20
        txtOrNo = ""
        txtCodeFrom = "3"
        txtCodeTo = "7"
        txtWeightFrom = "8"
        txtWeightTo = "12"
        
        showtext.value = vbChecked
        Opt(1).value = True
        margins.value = vbChecked
        CPic(0).Color = vbBlack
        margin.Text = "0"
        autocheck.value = vbUnchecked
        Exbearers.value = vbUnchecked
        BothBearers.value = vbUnchecked
        bearthick.Text = "0"
        Axbarcode1.ForeColor = CPic(0).Color
        CPic(1).Color = vbWhite
        Axbarcode1.backcolor = CPic(1).Color
        Axbarcode1.fontsize = 10
        Axbarcode1.Font = "Arial"
        Axbarcode1.Font.Bold = False
        Axbarcode1.Font.Italic = False
        CboScaleType.ListIndex = 0

        DoBarcode
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Command1_Click()

    DoBarcode
End Sub

Private Sub CPic_Change(Index As Integer, _
                        ByVal NewColor As stdole.OLE_COLOR)

    Select Case Index

        Case 0
            Axbarcode1.ForeColor = CPic(0).Color

        Case 1
            Axbarcode1.backcolor = CPic(1).Color
    End Select

End Sub

Private Sub Exbearers_Click()

    DoBarcode ' update barcode image
End Sub

Private Sub Form_Activate()

    DoBarcode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (6)
        End If
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass
    'If SystemOptions.UserInterface = EnglishInterface Then
    '    SetInterface Me
    '    'Axbarcode1.left = Me.ScaleWidth - (Axbarcode1.Width + Axbarcode1.left)
    'End If
    margin.Text = "0"
    bearthick.Text = "0"
    Heightt.Text = "20.0"
    Widtht.Text = "40.0"
    
    TxtFromNo = 20
    txtOrNo = ""
    txtCodeFrom = "3"
    txtCodeTo = "7"
    txtWeightFrom = "8"
    txtWeightTo = "12"

    
    SetBarCode
    Me.CPic(0).Color = Axbarcode1.ForeColor
    Me.CPic(1).Color = Axbarcode1.backcolor

    With Me.CboScaleType
        .Clear
        .AddItem "„·„Ì —", 0
        .ItemData(0) = 6
        .AddItem "”‰ „Ì —", 1
        .ItemData(1) = 7
        .AddItem "»Ê’…", 2
        .ItemData(2) = 5
        .AddItem "‰ﬁÿ…", 3
        .ItemData(3) = 2
        .AddItem " ÌÊÌ»", 4
        .ItemData(4) = 1
        '    .AddItem "Windows HIMETRIC (0.01 mm)", 5
        '    .ItemData(5) = 0
        '    .AddItem "Windows TEXT", 6
        '    .ItemData(6) = 3
        '    .AddItem "Windows HIENGLISH (0.001 inches)", 7
        '    .ItemData(7) = 4
        .ListIndex = 0
    End With

    BtnsStatus False
    AddTip
    Resize_Form Me
    BarcodeSetting 2

    DoBarcode
    Screen.MousePointer = vbDefault
    TxtMsg_Change
    Exit Sub
ErrTrap:
End Sub

Private Sub DoBarcode()
    Dim i As Single, j As Single
    On Error GoTo ErrTrap

    With CboBarcodes

        If .ListIndex = -1 Then
            Axbarcode1.CodeType = .ListIndex
        Else
            Axbarcode1.CodeType = .ItemData(CboBarcodes.ListIndex)
        End If

    End With

    i = val(Heightt.Text)
    'If (I < 1) Then I = 20
    Axbarcode1.ImageHeight = i
    j = val(Widtht.Text)
    Axbarcode1.ImageWidth = j

    'If (Option1.Value = True) Then
    '    If (j < 1) Then j = 30
    'j = Val(Widtht.Text)
    'Axbarcode1.ImageWidth = j
    '    Axbarcode1.Xunit = 0
    'End If
    'If (Option2.Value = True) Then
    'j = Val(Widtht.Text)
    'If (j < 10) Then j = 10
    '    Axbarcode1.Xunit = j
    'End If
    With CboScaleType

        If .ListIndex = -1 Then
            Axbarcode1.ScaleMode = 6
        Else
            Axbarcode1.ScaleMode = .ItemData(.ListIndex)
        End If

    End With

    If (nominal.value > 0) Then Axbarcode1.NominalSize = 100 Else Axbarcode1.NominalSize = 0
    If autocheck.value > 0 Then
        Axbarcode1.AutoParity = True
        Axbarcode1.ShowCheckDigit = True
    Else
        Axbarcode1.AutoParity = False
        Axbarcode1.ShowCheckDigit = False
    End If

    If Opt(0).value = True Then
        Axbarcode1.JustifyText = 1
    ElseIf Opt(1).value = True Then
        Axbarcode1.JustifyText = 0
    ElseIf Opt(2).value = True Then
        Axbarcode1.JustifyText = 2
    End If

    'Axbarcode1.CodeType = CboBarcodes.ListIndex
    Axbarcode1.showtext = showtext.value
    Axbarcode1.ShowLightMargins = margins.value
    Axbarcode1.ShowBearerBars = BothBearers.value
    Axbarcode1.ExtendBearers = Exbearers.value
    Axbarcode1.BearerBarThickness = val(bearthick.Text)
    Axbarcode1.MarginSize = val(margin.Text)
    Axbarcode1.Caption = TxtFromNo & txtMsg.Text

    DoEvents
    lblMax.Caption = Axbarcode1.Nrequired

    error1.Text = GetBarCodeErr(Axbarcode1.errorCode)

    If Axbarcode1.errorCode = 0 Then
        Img(0).Visible = True
        Img(1).Visible = False
        Img(7).Visible = True
        Img(6).Visible = False
        Timer1.Enabled = False
    Else
        Img(0).Visible = False
        Img(1).Visible = True
        Img(7).Visible = False
        Img(6).Visible = True
        Timer1.Enabled = True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub SetBarCode()
    On Error GoTo ErrTrap

    Dim X As String
    Dim j  As Integer
    Dim IntLoop As Integer
    X = "            "
    CboBarcodes.Clear

    For j = 0 To 100
        X = Axbarcode1.GetTypeNameB(j)
    
        If (Len(X) > 0) Then
            CboBarcodes.AddItem (X)

            For IntLoop = 0 To CboBarcodes.ListCount - 1

                If CboBarcodes.List(IntLoop) = X Then
                    CboBarcodes.ItemData(IntLoop) = j
                    Exit For
                End If

            Next IntLoop

        Else
            Exit For
        End If

    Next j

    'CboBarcodes.ListIndex = 8
    'justification = 0
    'margin.Text = "0"
    'bearthick.Text = "0"
    'Heightt.Text = "20.0"
    'Widtht.Text = "40.0"
    'DoBarcode ' update barcode image
    Exit Sub
ErrTrap:
End Sub

Private Function GetBarCodeErr(IntErrCode As Integer) As String
    On Error GoTo ErrTrap

    Select Case IntErrCode

        Case 0
            GetBarCodeErr = "·«ÌÊÃœ √Œÿ«¡" 'no Error

        Case 1
            GetBarCodeErr = "Œÿ√ ›Ï ÿÊ· «·ﬂÊœ" 'Wrong code length

        Case 2
            GetBarCodeErr = "Â–« «·ﬂÊœ €Ì— „⁄—›" 'Unrecognised code type

        Case 3
            GetBarCodeErr = "Œÿ√ ›Ï ÿÊ· «·ﬂÊœ" ' Wrong add-on code length

        Case 4
            GetBarCodeErr = "Õ—Ê› €Ì— „ﬁ»Ê·… ›Ï «·ﬂÊœ" 'Illegal character in code

        Case 5
            GetBarCodeErr = "Œÿ√ ›Ï «·ﬂÊœ «·„÷„‰" ' Error in embedded code

        Case 6
            GetBarCodeErr = "«·Œÿ «·‰« Ã ⁄—÷Â «ﬁ· „‰ ÊÕœ… Ê«Õœ…" 'Generated line width less than 1 unit

        Case 7
            GetBarCodeErr = "‰Ê⁄ «·›Ê‰  €Ì— „ﬁ»Ê·" 'Invalid text font

        Case 8
            GetBarCodeErr = "Invalid device context" '8        Invalid device context

        Case 9
            GetBarCodeErr = "Œÿ√ ›Ï «·‰’ «·„⁄—Ê÷" 'Invalid Caption property

        Case 10
            GetBarCodeErr = "Œÿ√ ›Ï Õ›Ÿ «·„·›" 'Error writing disk file
    End Select

    If IntErrCode = 0 Then
        BtnsStatus True
    Else
        BtnsStatus False
    End If

    Exit Function
ErrTrap:
End Function

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    If FrmBarcode.ReportsNumber > 0 Then
        'MsgBox "Close Reports"
        Unload DataRptBarcode
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    BarcodeSetting 1
End Sub

Private Sub Heightt_Change()

    DoBarcode ' update barcode image
End Sub

Private Sub margin_Change()

    DoBarcode ' update barcode image
End Sub

Private Sub margins_Click()

    DoBarcode ' update barcode image
End Sub

Private Sub nominal_Click()

    DoBarcode ' update barcode image
End Sub

Private Sub Opt_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
            Opt(0).value = True
            Opt(1).value = False
            Opt(2).value = False

        Case 1
            Opt(1).value = True
            Opt(0).value = False
            Opt(2).value = False

        Case 2
            Opt(2).value = True
            Opt(1).value = False
            Opt(0).value = False
    End Select

    DoBarcode
    Exit Sub
ErrTrap:
End Sub

Private Sub showtext_Click()
    On Error GoTo ErrTrap

    'Me.Ele(2).Enabled = CBool(showtext.Value)
    Me.Opt(0).Enabled = CBool(showtext.value)
    Me.Opt(1).Enabled = CBool(showtext.value)
    Me.Opt(2).Enabled = CBool(showtext.value)
    Me.margins.Enabled = CBool(showtext.value)
    'lbl(9).Enabled = CBool(showtext.Value)
    Cmd(9).Enabled = CBool(showtext.value)
    Me.Img(2).Enabled = CBool(showtext.value)
    Me.Img(3).Enabled = CBool(showtext.value)
    Me.Img(4).Enabled = CBool(showtext.value)

    DoBarcode ' update barcode image
    Exit Sub
ErrTrap:
End Sub

Private Sub Timer1_Timer()
    Img(7).Visible = False
    Img(6).Visible = Not Img(6).Visible
End Sub

Private Sub TxtModFlg_Change()

    Select Case Me.TxtModFlg.Text

        Case "N"
    
        Case "E"

        Case "R"
    
    End Select

End Sub

Private Sub TxtMsg_Change()

    If Trim(txtMsg.Text) = "" Then
        BtnsStatus False
    Else
        BtnsStatus True
    End If

    DoBarcode ' update barcode image
End Sub

Private Sub TxtMsg_KeyPress(KeyAscii As Integer)

    If KeyAscii >= Asc(("«")) And KeyAscii <= Asc(("Ï")) Then
        KeyAscii = 0
    Else
        KeyAscii = KeyAscii
    End If

End Sub

Private Sub Widtht_Change()

    DoBarcode ' update barcode image
End Sub

Private Sub AddTip()
    On Error GoTo ErrTrap

    Dim Msg As String
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim i As Integer
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, " œÊÌ— «·√” ﬂÌ—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 15000
        .DelayTime = 300
        Msg = "œÊ—«‰ «·√” ﬂÌ—  90 œ—Ã… ›Ï " & Wrap & "≈ Ã«Â ⁄ﬁ«—» «·”«⁄…" & Wrap & "„·ÕÊŸ…:- Ì „ ÿ»⁄ «·√” ﬂÌ—" & Wrap & "»«·Õ«·… «·⁄«œÌ… "
        .AddControl Cmd(0), Msg, True
    End With

    With TTP
        .Create Me.hwnd, " œÊÌ— «·√” ﬂÌ—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 15000
        .DelayTime = 300
        Msg = "œÊ—«‰ «·√” ﬂÌ—  90 œ—Ã… ›Ï " & Wrap & " ⁄ﬂ” ≈ Ã«Â ⁄ﬁ«—» «·”«⁄…" & Wrap & "„·ÕÊŸ…:- Ì „ ÿ»⁄ «·√” ﬂÌ—" & Wrap & "»«·Õ«·… «·⁄«œÌ…"
        .AddControl Cmd(4), Msg, True
    End With

    With TTP
        .Create Me.hwnd, "Õ›Ÿ ’Ê—… „‰ «·»«—ﬂÊœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 15000
        .DelayTime = 300
        Msg = "≈÷€ÿ Â‰« Õ Ï ÌŸÂ— ·ﬂ „—»⁄ ÕÊ«—  " & Wrap & "·Õ›Ÿ ’Ê—… „‰  ’„Ì„ «·»«—ﬂÊœ ﬂ„·›" & Wrap & "⁄·Ï ÃÂ«“ﬂ- ÊÌ„ﬂ‰ﬂ Õ›Ÿ Â–Â «·’Ê—…" & Wrap & "Metafile (*.wmf)" & Wrap & "Bitmap (*.bmp)" & Wrap & "Paintbrush (*.pcx)" & Wrap & "Encapsulated PostScript (*.eps)" & Wrap & "Portable Network Graphic (*.png)" & Wrap & "GIF (*.gif)"
        .AddControl Cmd(1), Msg, True
    End With

    'MetaFile‰”Œ ’Ê—… „‰ «·»«—ﬂÊœ ≈·Ï Õ«›Ÿ… «·ÊÌ‰œÊ“  ⁄·Ï «·‰”ﬁ
    With TTP
        .Create Me.hwnd, "‰”Œ ’Ê—… ≈·Ï «·Õ«›Ÿ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 15000
        .DelayTime = 300
        Msg = "Â·  —Ìœ ‰ﬁ· ’Ê—… „‰ «·»«—ﬂÊœ ≈·Ï √Ï " & Wrap & "»—‰«„Ã...øø" & Wrap & "≈÷€ÿ Â‰« Õ Ï Ì „ ‰”Œ ’Ê—… „‰ «·»«—ﬂÊœ" & Wrap & "≈·Ï Õ«›Ÿ… «·ÊÌ‰œÊ“ ⁄·Ï «·‰”ﬁ'MetaFile" & Wrap & "ÊÌ„ﬂ‰ﬂ »⁄œ –·ﬂ › Õ «Ï »—‰«„Ã „‰ »—«„Ã" & Wrap & "«·ÊÌ‰œÊ“ „À· »—‰«„Ã „Ìﬂ—Ê”Ê›  Ê——œ" & Wrap & "√Ê »—‰«„Ã «·—”«„ -À„ ⁄„· ·’ﬁ ......." & Wrap & "”Ê›  Ãœ «‰ ’Ê—… «·»«—ﬂÊœ ﬁœ Ê÷⁄  „‰ " & Wrap & "Õ«›Ÿ… «·ÊÌ‰œÊ“ ≈·Ï Â–« «·»—‰«„Ã." & Wrap & "" & Wrap & "„·ÕÊŸ…:- Â–« «·‰Ê⁄ „‰ «·’Ê— Ì„ﬂ‰  €Ì— " & Wrap & "ÕÃ„Â »”ÂÊ·… œ«Œ· «·»—‰«„Ã «·„—«œ ‰ﬁ· " & Wrap & "«·’Ê—… ·Â."
        .AddControl Cmd(3), Msg, True
    End With

    With TTP
        .Create Me.hwnd, "‰”Œ ’Ê—… ≈·Ï «·Õ«›Ÿ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 15000
        .DelayTime = 300
        Msg = "Â·  —Ìœ ‰ﬁ· ’Ê—… „‰ «·»«—ﬂÊœ ≈·Ï √Ï " & Wrap & "»—‰«„Ã...øø" & Wrap & "≈÷€ÿ Â‰« Õ Ï Ì „ ‰”Œ ’Ê—… „‰ «·»«—ﬂÊœ" & Wrap & "≈·Ï Õ«›Ÿ… «·ÊÌ‰œÊ“ ⁄·Ï «·‰”ﬁ'Bitmap" & Wrap & "ÊÌ„ﬂ‰ﬂ »⁄œ –·ﬂ › Õ «Ï »—‰«„Ã „‰ »—«„Ã" & Wrap & "«·ÊÌ‰œÊ“ „À· »—‰«„Ã „Ìﬂ—Ê”Ê›  Ê——œ" & Wrap & "√Ê »—‰«„Ã «·—”«„ -À„ ⁄„· ·’ﬁ ......." & Wrap & "”Ê›  Ãœ «‰ ’Ê—… «·»«—ﬂÊœ ﬁœ Ê÷⁄  „‰ " & Wrap & "Õ«›Ÿ… «·ÊÌ‰œÊ“ ≈·Ï Â–« «·»—‰«„Ã."
        .AddControl Cmd(5), Msg, True
    End With

    With TTP
        .Create Me.hwnd, "Œ—ÊÃ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ „‰ ‘«‘…  ’„Ì„ «·»«—ﬂÊœ"
        .AddControl Cmd(6), Msg, True
    End With

    With TTP
        .Create Me.hwnd, " ≈⁄œ«œ ÿ«»⁄… «·»«—ﬂÊœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 15000
        .DelayTime = 300
        Msg = "·Ê «‰ ·œÌﬂ ÿ«»⁄… „Œ’’… ··»«—ﬂÊœ" & Wrap & "›≈÷€ÿ Â‰« Õ Ï ÌŸÂ— ·ﬂ „—»⁄ ÕÊ«—" & Wrap & "≈⁄œ«œ «·ÿ«»⁄«  Õ Ï  ﬁÊ„ »≈Œ Ì«—" & Wrap & "Â–Â «·ÿ«»⁄… Ê ÕœÌœ Œ’«∆Â«" & Wrap & "„·ÕÊŸ…:- ÌÃ» «‰  ﬂÊ‰ Â–Â «·ÿ«»⁄…" & Wrap & "ÂÏ «·ÿ«»⁄… «·√› —«÷Ì…-›Ï Õ«·… ÊÃÊœ" & Wrap & "√ﬂÀ— „‰ ÿ«»⁄… ·œÌﬂ-ÊÌ„ﬂ‰ﬂ ⁄„· –·ﬂ" & Wrap & "»«·÷€ÿ „— Ì‰ ⁄·Ï Â–Â «·ÿ«»⁄…"
        .AddControl Cmd(2), Msg, True
    End With

    With TTP
        .Create Me.hwnd, "«·ÿ»«⁄… »≈” Œœ«„ ÿ«»⁄… «·»«—ﬂÊœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 15000
        .DelayTime = 300
        Msg = "≈÷€ÿ Â‰« Õ Ï   „ ⁄„·Ì… «·ÿ»«⁄… »Ê«”ÿ…" & Wrap & "ÿ«»⁄… »«—ﬂÊœ „Œ’’…" & Wrap & "„·ÕÊŸ…:- ÌÃ» «‰  ﬂÊ‰ Â–Â «·ÿ«»⁄…" & Wrap & "ÂÏ «·ÿ«»⁄… «·√› —«÷Ì…-›Ï Õ«·… ÊÃÊœ" & Wrap & "√ﬂÀ— „‰ ÿ«»⁄… ·œÌﬂ-ÊÌ„ﬂ‰ﬂ ⁄„· –·ﬂ" & Wrap & "»«·÷€ÿ „— Ì‰ ⁄·Ï Â–Â «·ÿ«»⁄…"
        .AddControl Cmd(7), Msg, True
    End With

    With TTP
        .Create Me.hwnd, "«·ÿ»«⁄… »≈” Œœ«„ ÿ«»⁄… ⁄«œÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 15000
        .DelayTime = 300
        Msg = "≈–« ﬂ«‰  ·œÌﬂ ÿ«»⁄… ⁄«œÌ… Ê —Ìœ " & Wrap & "≈” Œœ«„Â« ›Ï ⁄„·Ì… ÿ»«⁄… «·»«—ﬂÊœ" & Wrap & "!!!...." & Wrap & "›≈÷€ÿ Â‰« Õ Ï Ì „ ŸÂÊ— ‘«‘… ÿ»«⁄…" & Wrap & "«·»«—ﬂÊœ «· Ï   ÌÕ ·ﬂ «Œ Ì«— " & Wrap & "«·ÿ«»⁄… ÊÕÃ„ «·Ê—ﬁ Ê ’„Ì„ «·’›Õ…" & Wrap & "Ê«·ÂÊ«„‘ Ê«·›Ê«’· »Ì‰ «·√” ﬂÌ—« ...."
        .AddControl Cmd(8), Msg, True
    End With

    With TTP
        .Create Me.hwnd, "‰Ê⁄ «·»«—ﬂÊœ «·„” Œœ„", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        i = CboBarcodes.ListCount
        Msg = "«Œ — ‰Ê⁄ «·»«—ﬂÊœ «·–Ï  —Ìœ ≈” Œœ«„Â ›Ï  ’„Ì„ «·√” ﬂÌ—." & Wrap & ""
        Msg = Msg + "·œÌﬂ ⁄œœ " & i & " »«—ﬂÊœ „Œ ·› ."
        .AddControl CboBarcodes, Msg, True
    End With

    With TTP
        .Create Me.hwnd, "ﬂÊœ «·’‰› «·„—«œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Â‰« Ì„ﬂ‰ﬂ ﬂ «»… ﬂÊœ «·’‰› " & Wrap & " «·–Ï  —Ìœ  ’„Ì„ «·√” ﬂÌ— ·Â."
        .AddControl txtMsg, Msg, True
    End With

    With TTP
        .Create Me.hwnd, "Õ«·… (—”«·…) «·»«—ﬂÊœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Â‰«  ⁄—÷ «·—”«∆· «· Ï  »Ì‰ Õ«·… «·ﬂÊœ «·„œŒ·..." & Wrap & "Â· ÂÊ ’ÕÌÕ Ê„ﬁ»Ê· „⁄ «·»«—ﬂÊœ «·„” Œœ„ Ê«·„Õœœ." & Wrap & "«„ »Â √Œÿ«¡ Ê›Ï Õ«·… ÊÃÊœ √Œÿ«¡.. ⁄—÷ ·ﬂ „«ÂÊ" & Wrap & "Â–« «·Œÿ√."
        .AddControl error1, Msg, True
    End With

    With TTP
        .Create Me.hwnd, "„ﬁœ«— ≈— ›«⁄ «·√” ﬂÌ—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Â‰« Ì„ﬂ‰ﬂ  ÕœÌœ ≈— ›«⁄ „⁄Ì‰ ··√” ﬂÌ— ÊÂ–«" & Wrap & " «·√— ›«⁄ Ì„ﬂ‰ «‰ Ìﬁ«” »«·„·„Ì — «Ê »«·”‰ „Ì—." & Wrap & "«·ﬁÌ„… «·√› —«÷Ì… ÂÏ 20 „·„ =2 ”„ ."
        .AddControl Heightt, Msg, True
    End With

    With TTP
        .Create Me.hwnd, "„ﬁœ«— ⁄—÷ «·√” ﬂÌ—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Â‰« Ì„ﬂ‰ﬂ  ÕœÌœ ⁄—÷ „⁄Ì‰ ··√” ﬂÌ— ÊÂ–«" & Wrap & " «·⁄—÷ Ì„ﬂ‰ «‰ Ìﬁ«” »«·„·„Ì — «Ê »«·”‰ „Ì—." & Wrap & "«·ﬁÌ„… «·√› —«÷Ì… ÂÏ 40 „·„ =4 ”„ ."
        .AddControl Widtht, Msg, True
    End With

    With TTP
        .Create Me.hwnd, "‰Ê⁄ «·„ﬁÌ«” «·„” Œœ„(„·„Ì —)", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Â–« «·ŒÌ«— ÌÃ⁄· „ﬁœ«— ⁄—÷ «·√” ﬂÌ— Ìﬁ«”" & Wrap & "»«·„·„Ì —."
        '.AddControl Option1, Msg, True
    End With

    With TTP
        .Create Me.hwnd, "‰Ê⁄ «·„ﬁÌ«” «·„” Œœ„(”‰ „Ì —)", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Â–« «·ŒÌ«— ÌÃ⁄· „ﬁœ«— ⁄—÷ «·√” ﬂÌ— Ìﬁ«”" & Wrap & "»«·”‰ „Ì —."
        '.AddControl Option2, Msg, True
    End With

    With TTP
        .Create Me.hwnd, "√ŸÂ— ⁄·«„«  «·„Õ«–«…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "» ›⁄Ì· Â–« «·ŒÌ«— ÌﬁÊ„ «·»—‰«„Ã »ÿ»«⁄… ⁄·«„«  «·„Õ«–«…" & Wrap & "⁄·Ï «·√” ﬂÌ— · »Ì‰ Â· «·‰’ «·„ÿ»Ê⁄ „Õ«–«Ï(Ì„Ì‰-Ì”«—-Ê”ÿ)."
        .AddControl margins, Msg, True
    End With

    With TTP
        .Create Me.hwnd, "ÕœÊœ „„ œ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "» ›⁄Ì· Â–« «·ŒÌ«— ÌﬁÊ„ «·»—‰«„Ã »„œ ÕœÊœ" & Wrap & "≈·Ï ‰Â«Ì… ÕÃ„ «·√” ﬂÌ—(Ê·Ì” ⁄·Ï „ﬁœ«— «·ŒÿÊÿ)"
        .AddControl Exbearers, Msg, True
    End With

    With TTP
        .Create Me.hwnd, "Õœ √⁄·Ï Ê√”›·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "» ›⁄Ì· Â–« «·ŒÌ«— ÌﬁÊ„ «·»—‰«„Ã »⁄„·" & Wrap & "«·Õœ √⁄·Ï «·√” ﬂÌ— Ê√”›·Â"
        .AddControl BothBearers, Msg, True
    End With

    With TTP
        .Create Me.hwnd, "„ﬁœ«— «·Â«„‘", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«ﬂ » Â‰« ﬁÌ„… «·Â«„‘ ··√” ﬂÌ— «·„ÿ»Ê⁄" & Wrap & "«·ﬁÌ„… «·√› —«’Ì… ’›— ( ·«ÌÊÃœ Â«„‘)"
        .AddControl margin, Msg, True
    End With

    With TTP
        .Create Me.hwnd, "„ﬁœ«— ”„ﬂ «·Õœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«ﬂ » Â‰« ﬁÌ„… Õœ «·√” ﬂÌ— «·„ÿ»Ê⁄" & Wrap & "«·ﬁÌ„… «·√› —«’Ì… ’›— ( ·«ÌÊÃœ Õœ)"
        .AddControl bearthick, Msg, True
    End With

    '
    With TTP
        .Create Me.hwnd, "√ŸÂ— «·ﬂÊœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "» ›⁄Ì· Â–« «·ŒÌ«— ÌﬁÊ„ «·»—‰«„Ã »ÿ»«⁄…" & Wrap & "ﬂÊœ «·’‰›(√”›· ŒÿÊÿ «·»«—ﬂÊœ)."
        .AddControl showtext, Msg, True
    End With

    With TTP
        .Create Me.hwnd, "ÕÃ„ ÿ»Ì⁄Ï", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "» ›⁄Ì· Â–« «·ŒÌ«— ÌﬁÊ„ «·»—‰«„Ã " & Wrap & "»Ê÷⁄ «·ﬁÌ„ «·√› —«÷Ì… · ’„Ì„ «·√” ﬂÌ—"
        .AddControl nominal, Msg, True
    End With

    '
    With TTP
        .Create Me.hwnd, "„Õ«–«… ··Ì”«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Â–« «·ŒÌ«— ÌÃ⁄· ﬂÊœ «·’‰› ›Ï Õ«·…" & Wrap & " ŸÂÊ—Â „Õ«–«Ï ‰«ÕÌ… «·Ì”«—"
        .AddControl Opt(0), Msg, True
    End With

    With TTP
        .Create Me.hwnd, "„Õ«–«… ›Ï «·„‰ ’›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Â–« «·ŒÌ«— ÌÃ⁄· ﬂÊœ «·’‰› ›Ï Õ«·…" & Wrap & " ŸÂÊ—Â „Õ«–«Ï ›Ï „‰ ’› «·√” ﬂÌ—."
        .AddControl Opt(1), Msg, True
    End With

    With TTP
        .Create Me.hwnd, "„Õ«–«… ··Ì„Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Â–« «·ŒÌ«— ÌÃ⁄· ﬂÊœ «·’‰› ›Ï Õ«·…" & Wrap & " ŸÂÊ—Â „Õ«–«Ï ‰«ÕÌ… «·Ì„Ì‰"
        .AddControl Opt(2), Msg, True
    End With

    '
    '    With TTP
    '        .Create Me.hwnd, "·Ê‰ «·»«—ﬂÊœ «·„ÿ»Ê⁄", 1, 15204351, -2147483630, True
    '        .MaxWidth = 4000
    '        .VisibleTime = 10000
    '        .DelayTime = 300
    '        Msg = "≈÷€ÿ Â‰« · Œ «— ·Ê‰ „⁄Ì‰ " & Wrap & _
    '                "··»«—ﬂÊœ «·„ÿ»Ê⁄"
    '        .AddControl CPic(0), Msg, True
    '    End With
    CPic(0).ToolTipText = "≈÷€ÿ Â‰« · Œ «— ·Ê‰ „⁄Ì‰ ··»«—ﬂÊœ «·„ÿ»Ê⁄"
    '    With TTP
    '        .Create Me.hwnd, "·Ê‰ Œ·›Ì… «·√” ﬂÌ—", 1, 15204351, -2147483630, True
    '        .MaxWidth = 4000
    '        .VisibleTime = 10000
    '        .DelayTime = 300
    '        Msg = "≈÷€ÿ Â‰« · Œ «— ·Ê‰ „⁄Ì‰ " & Wrap & _
    '                "·Œ·›Ì… «·√” ﬂÌ—"
    '        .AddControl CPic(1), Msg, True
    '    End With
    CPic(1).ToolTipText = "≈÷€ÿ Â‰« · Œ «— ·Ê‰ „⁄Ì‰ Œ·›Ì… «·√” ﬂÌ—"

    With TTP
        .Create Me.hwnd, " Œ’Ì’ «·Œÿ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "≈÷€ÿ Â‰« · Œ «— ‰Ê⁄ «·Œÿ ÊÕÃ„Â " & Wrap & "«·–Ï Ì” Œœ„ ›Ï ÿ»«⁄… ﬂÊœ «·’‰›"
        .AddControl Cmd(9), Msg, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub PrintBarcode1()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Printers.count = 0 Then
        Msg = "·« ÊÃœ ÿ«»⁄«  „⁄—›… ›Ï «·ÃÂ«“"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    DoBarcode

    If Me.Axbarcode1.Picture = 0 Then
        GetMsgs 159, vbExclamation
        Exit Sub
    End If

    Printer.ScaleMode = 6  ' sets printer scale to mm
    Printer.PaintPicture Axbarcode1.Picture, 20, 20, Axbarcode1.PictureWidth, Axbarcode1.PictureHeight
    Printer.NewPage
    Printer.EndDoc
    Exit Sub
ErrTrap:
End Sub

Private Sub PrintBarcode2()
    Dim Msg As String

    If Printers.count = 0 Then
        Msg = "·« ÊÃœ ÿ«»⁄«  „⁄—›… ›Ï «·ÃÂ«“"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    DoBarcode

    If Me.Axbarcode1.Picture = 0 Then
        GetMsgs 159, vbExclamation
        Exit Sub
    End If

 '   FrmDesOptions.show
 '   FrmDesOptions.ZOrder 0
End Sub

Public Property Get ReportsNumber() As Integer
    ReportsNumber = m_ReportsNumber
End Property

Public Property Let ReportsNumber(ByVal vNewValue As Integer)
    m_ReportsNumber = vNewValue
End Property

Private Sub GetType()
    On Error GoTo ErrTrap

    Dim Msg As String

    With Axbarcode1

        If .CodeType = 0 Then
            Msg = "13 —ﬁ„ ›ﬁÿ"
        ElseIf .CodeType = 1 Then
            Msg = "8 —ﬁ„ ›ﬁÿ"
        ElseIf .CodeType = 2 Then
            Msg = "15 —ﬁ„ ›ﬁÿ"
        ElseIf .CodeType = 3 Then
            Msg = "18 —ﬁ„ ›ﬁÿ"
        ElseIf .CodeType = 4 Then
            Msg = "12 —ﬁ„ ›ﬁÿ"
        ElseIf .CodeType = 5 Then
            Msg = "7 —ﬁ„ ›ﬁÿ"
        ElseIf .CodeType = 6 Then
            Msg = "14 —ﬁ„ ›ﬁÿ"
        ElseIf .CodeType = 7 Then
            Msg = "6 —ﬁ„ ›ﬁÿ"
        ElseIf .CodeType = 8 Then
            'any
            Msg = "Õ—Ê› √‰Ã·Ì“Ì…( ﬂ»Ì—…) Ê√—ﬁ«„" & CHR(13)
            Msg = Msg & " * - + / $"
        ElseIf .CodeType = 9 Then
            'Code 128  any*
            Msg = "«Ï „‰ «·√—ﬁ«„ «Ê «·Õ—Ê› -≈‰Ã·Ì“Ì…(ﬂ»Ì—… √Ê ’€Ì—…)" & CHR(13)
            Msg = Msg & " Ìﬁ»· «·Õ—Ê› «·⁄—»Ì…(·ﬂ‰ ·«  ŸÂ— ⁄·Ï «·„·’ﬁ)"
        ElseIf .CodeType = 10 Then
            'EAN/UCC-128  any*
            Msg = "«Ï „‰ «·√—ﬁ«„ «Ê «·Õ—Ê› -≈‰Ã·Ì“Ì…(ﬂ»Ì—… √Ê ’€Ì—…)" & CHR(13)
            Msg = Msg & ""
        ElseIf .CodeType = 11 Then
            '2 of 5  any numbers
            Msg = "√—ﬁ«„ ›ﬁÿ(»Õœ √ﬁ’Ï 37 —ﬁ„)" & CHR(13)
            Msg = Msg & ""
        ElseIf .CodeType = 13 Then
            '3 of 9  any
            Msg = "√—ﬁ«„ ÊÕ—Ê› (»Õœ √ﬁ’Ï 32 —ﬁ„)" & CHR(13)
            Msg = Msg & "Õ—Ê› ≈‰Ã·Ì“Ì… ( ﬂ»Ì—… ›ﬁÿ)"
        ElseIf .CodeType = 14 Then
            'Code B
            Msg = "√—ﬁ«„ ›ﬁÿ (»Õœ √ﬁ’Ï 80 —ﬁ„)"
        ElseIf .CodeType = 15 Then
            'Code 11
            Msg = "√—ﬁ«„ ›ﬁÿ (»Õœ √ﬁ’Ï 80 —ﬁ„)"
        ElseIf .CodeType = 16 Then
            'Codabar
            Msg = "√—ﬁ«„ Ê«·Õ—Ê› A B C D E N T * " & CHR(13)
            Msg = Msg & "(»Õœ √ﬁ’Ï 43 Œ«‰…)"
        ElseIf .CodeType = 17 Then
            'MSI
            Msg = "√—ﬁ«„ ÊÕ—Ê› (»Õœ √ﬁ’Ï 80 —ﬁ„)" & CHR(13)
            Msg = Msg & "Õ—Ê› ≈‰Ã·Ì“Ì… ( ﬂ»Ì—… ›ﬁÿ)"
        ElseIf .CodeType = 18 Then
            'Ext. Code 39
            Msg = "(√—ﬁ«„ ÊÕ—Ê› (»Õœ √ﬁ’Ï 32 —ﬁ„)" & CHR(13)
            Msg = Msg & "(Õ—Ê› √‰Ã·Ì“Ì…(ﬂ»Ì—… √Ê ’€Ì—…"
        ElseIf .CodeType = 19 Then
            'UPCA+2
            Msg = "14 —ﬁ„"
        ElseIf .CodeType = 20 Then
            'UPCA+5
            Msg = "17 —ﬁ„"
        ElseIf .CodeType = 21 Then
            'EAN8+2
            Msg = "10 —ﬁ„"
        ElseIf .CodeType = 22 Then
            'EAN8 5
            Msg = "13 —ﬁ„"
        ElseIf .CodeType = 23 Then
            'UPCE 2
            Msg = "9 —ﬁ„"
        ElseIf .CodeType = 24 Then
            'UPCE+5
            Msg = "12 —ﬁ„"
        ElseIf .CodeType = 25 Then
            'Telepen standard
            Msg = "√—ﬁ«„ ÊÕ—Ê› √‰Ã·Ì“Ì…(ﬂ»Ì—… √Ê ’€Ì—…)" & CHR(13)
            Msg = "»Õœ √ﬁ’Ï 32 Œ«‰…"
        ElseIf .CodeType = 28 Then
            'PostNet type A
            Msg = "5 √—ﬁ«„"
        ElseIf .CodeType = 29 Then
            'PostNet type C
            Msg = "9 √—ﬁ«„"
        ElseIf .CodeType = 30 Then
            'PostNet type C
            Msg = "11 √—ﬁ«„"
        ElseIf .CodeType = 36 Then
            'Code 93
            Msg = "√—ﬁ«„ ÊÕ—Ê› √‰Ã·Ì“Ì…(ﬂ»Ì—…)" & CHR(13)
            Msg = Msg & "»Õœ √ﬁ’Ï 80 Œ«‰…"
        ElseIf .CodeType = 58 Then
            'Japan Post
            Msg = "√—ﬁ«„ ÊÕ—Ê› √‰Ã·Ì“Ì…(ﬂ»Ì—…)" & CHR(13)
            Msg = Msg & "»Õœ √ﬁ’Ï 32 Œ«‰…"
        End If

        lblMax.Caption = Msg
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub ShowInfo()
    Dim StrDate As String
    Dim StrInfo As String
    On Error GoTo ErrTrap

    If Axbarcode1.CodeType = 0 Then
    ElseIf Axbarcode1.CodeType = 1 Then
    Else
    End If

    'FrmInfo.lbl(1).Caption = Axbarcode1.GetTypeNameB(Axbarcode1.CodeType)
    'FrmInfo.Show vbModal
    Exit Sub
ErrTrap:
End Sub

Private Sub BtnsStatus(BolStatus As Boolean)
    On Error GoTo ErrTrap
    Cmd(1).Enabled = BolStatus
    Cmd(5).Enabled = BolStatus
    Cmd(3).Enabled = BolStatus
    Cmd(8).Enabled = BolStatus
    Cmd(7).Enabled = BolStatus
    Exit Sub
ErrTrap:
End Sub

Private Sub BarcodeSetting(IntMode As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer

    If IntMode = 1 Then ' Save
        SaveData
        
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "CodeType", Me.CboBarcodes.ListIndex
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "CodeMsg", Trim(Me.txtMsg.Text)
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "BarcodeHeight", val(Me.Heightt.Text)
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "BarcodeWidth", val(Me.Widtht.Text)
'
'        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "FromNo", txtFromNo
'        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "OrNo", txtOrNo
'        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "CodeFrom", txtCodeFrom
'        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "CodeTo", txtCodeTo
'        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "WeightFrom", txtWeightFrom
'        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "WeightTo", txtWeightTo
 
        
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "ScaleMode", val(Me.CboScaleType.ListIndex)
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "ShowText", showtext.value
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "margins", val(Me.margins.value)

        For i = 0 To Opt.count - 1

            If Opt(i).value = True Then
                SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "Opt", i
                Exit For
            End If

        Next i

        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "FontSize", Me.Axbarcode1.fontsize
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "FontBold", Me.Axbarcode1.FontBold
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "FontItalic", Me.Axbarcode1.FontItalic
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "FontName", Me.Axbarcode1.FontName
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "ForeColor", Me.CPic(0).Color
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "BackColor", Me.CPic(1).Color
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "BothBearers", Me.BothBearers.value
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "bearthick", bearthick.Text
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "Exbearers", Me.Exbearers.value
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "margin", Me.margin.Text
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "autocheck", autocheck.value
        

        'SaveSetting SystemOptions.SysRegsAppPath, "BarcodeDesgin", "CodeFrom", txtCodeFrom
                        

    ElseIf IntMode = 2 Then
        Retrive
        
        Me.CboBarcodes.ListIndex = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "CodeType", Me.CboBarcodes.ListIndex)
        txtMsg.Text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "CodeMsg", "")
'               txtFromNo.Text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "FromNo", 20)
'       txtOrNo.Text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "OrNo", "0")
'        txtCodeFrom.Text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "CodeFrom", 3)
'        txtCodeTo.Text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "CodeTo", 7)
'
'        txtWeightFrom.Text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "WeightFrom", 8)
'        txtWeightTo.Text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "WeightTo", 12)

        Heightt.Text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "BarcodeHeight", 20)
        Widtht.Text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "BarcodeWidth", 40)
        CboScaleType.ListIndex = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "ScaleMode", val(Me.CboScaleType.ListIndex))
        showtext.value = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "ShowText", showtext.value)
        margins.value = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "margins", Me.margins.value)
        





        i = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "Opt", i)
        Opt(i).value = True
        Me.Axbarcode1.fontsize = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "FontSize", Me.Axbarcode1.fontsize)
        Me.Axbarcode1.FontBold = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "FontBold", Me.Axbarcode1.FontBold)
        Me.Axbarcode1.FontItalic = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "FontItalic", Me.Axbarcode1.FontItalic)
        Me.Axbarcode1.FontName = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "FontName", Me.Axbarcode1.FontName)
        Me.CPic(0).Color = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "ForeColor", Me.CPic(0).Color)
        Me.CPic(1).Color = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "BackColor", Me.CPic(1).Color)
        Axbarcode1.ForeColor = CPic(0).Color
        Axbarcode1.backcolor = CPic(1).Color
        Me.BothBearers.value = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "BothBearers", Me.BothBearers.value)
        bearthick.Text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "bearthick", bearthick.Text)
        Me.Exbearers.value = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "Exbearers", Me.Exbearers.value)
        Me.margin.Text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "margin", Me.margin.Text)
        autocheck.value = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeDesgin", "autocheck", autocheck.value)
    
    End If

    Exit Sub
ErrTrap:
End Sub
Private Sub SaveData()
Dim s As String
Dim rsDummy As New ADODB.Recordset

s = "Select * from  TblOptions "
rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
If Not rsDummy.EOF Then
    rsDummy!FromNo = val(TxtFromNo)
   
    If Me.chkIsSomeItemWeight.value = vbChecked Then
        rsDummy("IsSomeItemWeight").value = 1
    ElseIf Me.chkIsSomeItemWeight.value = vbUnchecked Then
        rsDummy("IsSomeItemWeight").value = 0
    End If
    
    
    rsDummy!OrNo = val(txtOrNo)
    rsDummy!CodeFrom = val(txtCodeFrom)
    rsDummy!CodeTo = val(txtCodeTo)
    rsDummy!WeightFrom = val(txtWeightFrom)
    rsDummy!WeightTo = val(txtWeightTo)
    rsDummy.update
End If


End Sub
Private Sub Retrive()
Dim s As String
Dim rsDummy As New ADODB.Recordset

s = "Select * from  TblOptions "


rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
If Not rsDummy.EOF Then
    TxtFromNo = rsDummy!FromNo & ""
    txtOrNo = rsDummy!OrNo & ""
    txtCodeFrom = rsDummy!CodeFrom & ""
    txtCodeTo = rsDummy!CodeTo & ""
    txtWeightFrom = rsDummy!WeightFrom & ""
    txtWeightTo = rsDummy!WeightTo & ""
    If rsDummy("IsSomeItemWeight").value = vbTrue Then
        Me.chkIsSomeItemWeight.value = vbChecked
    Else
        Me.chkIsSomeItemWeight.value = vbUnchecked
    End If
        
End If
   
End Sub
Private Sub XPPanel306_GotFocus()

End Sub
