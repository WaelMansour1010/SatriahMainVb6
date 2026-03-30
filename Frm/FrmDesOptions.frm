VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{85FD608E-54A8-11D4-8ED4-00E07D815373}#1.0#0"; "MBClrPkr.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDesOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ»«⁄… «·»«—ﬂÊœ"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   Icon            =   "FrmDesOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   5190
   Begin ImpulseButton.ISButton Cmdyes 
      Default         =   -1  'True
      Height          =   405
      Left            =   900
      TabIndex        =   1
      Top             =   5340
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   714
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„Ê«›ﬁ"
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
      ButtonImage     =   "FrmDesOptions.frx":038A
      ColorButton     =   14871017
      ColorHighlight  =   4194304
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton CmdNo 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   60
      TabIndex        =   2
      Top             =   5340
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   714
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈·€«¡"
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
      ButtonImage     =   "FrmDesOptions.frx":0724
      ColorButton     =   14871017
      ColorHighlight  =   4194304
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   4245
      Left            =   0
      TabIndex        =   3
      Top             =   930
      Width           =   5145
      _cx             =   9075
      _cy             =   7488
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
      BackColor       =   -2147483633
      ForeColor       =   0
      FrontTabColor   =   -2147483633
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   16711680
      Caption         =   "≈⁄œ«œ «·’›Õ…|»Ì«‰«  √Œ—Ï"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   1
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   2
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Picture(0)      =   "FrmDesOptions.frx":0ABE
      Picture(1)      =   "FrmDesOptions.frx":0E58
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3780
         Index           =   1
         Left            =   5790
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   45
         Width           =   5055
         _cx             =   8916
         _cy             =   6668
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
         Begin VB.CheckBox ChkShow 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "≈ŸÂ«— «· ⁄·Ìﬁ «·”›·Ï"
            Height          =   315
            Index           =   1
            Left            =   2970
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   2010
            Width           =   2025
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E0E0E0&
            Height          =   1515
            Index           =   3
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   2190
            Width           =   4995
            Begin VB.TextBox TxtHeight 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   1
               Left            =   3210
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Text            =   "10"
               Top             =   1065
               Width           =   555
            End
            Begin VB.TextBox TxtComment 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   345
               Index           =   1
               Left            =   2070
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   180
               Width           =   2835
            End
            Begin MBColorPicker.ColorPicker CPic 
               Height          =   345
               Index           =   2
               Left            =   30
               TabIndex        =   21
               ToolTipText     =   "·Ê‰ Œÿ «·»«—ﬂÊœ"
               Top             =   765
               Width           =   735
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
            Begin MBColorPicker.ColorPicker CPic 
               Height          =   345
               Index           =   3
               Left            =   30
               TabIndex        =   22
               ToolTipText     =   "·Ê‰ Œÿ «·»«—ﬂÊœ"
               Top             =   1110
               Width           =   735
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
            Begin MSComctlLib.Toolbar TBr 
               Height          =   330
               Index           =   1
               Left            =   2700
               TabIndex        =   23
               Top             =   570
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               ImageList       =   "ImgListToolBar"
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   7
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Bold"
                     ImageKey        =   "Bold"
                     Style           =   1
                  EndProperty
                  BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Italic"
                     ImageKey        =   "Italic"
                     Style           =   1
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Under"
                     ImageKey        =   "Under"
                     Style           =   1
                  EndProperty
                  BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Style           =   3
                  EndProperty
                  BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Left"
                     ImageKey        =   "Left"
                     Style           =   2
                  EndProperty
                  BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Center"
                     ImageKey        =   "Center"
                     Style           =   2
                     Value           =   1
                  EndProperty
                  BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Right"
                     ImageKey        =   "Right"
                     Style           =   2
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   345
               Index           =   2
               Left            =   90
               TabIndex        =   24
               Top             =   390
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "..."
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
               ButtonImage     =   "FrmDesOptions.frx":11F2
               ColorButton     =   12632256
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "„·„"
               Height          =   255
               Index           =   21
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   1050
               Width           =   315
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "≈— ›«⁄ «· ⁄·Ìﬁ"
               Height          =   225
               Index           =   20
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   1080
               Width           =   1065
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "·Ê‰ «·‰’"
               Height          =   285
               Index           =   22
               Left            =   840
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   1140
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "·Ê‰ «·Œ·›Ì…"
               Height          =   285
               Index           =   23
               Left            =   840
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   840
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "‰Ê⁄ «·Œÿ"
               Height          =   285
               Index           =   24
               Left            =   870
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   480
               Width           =   855
            End
         End
         Begin VB.CheckBox ChkShow 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "≈ŸÂ«— «· ⁄·Ìﬁ «·⁄·ÊÏ"
            Height          =   375
            Index           =   0
            Left            =   2970
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   90
            Width           =   2025
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Height          =   1695
            Index           =   2
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   300
            Width           =   4965
            Begin VB.TextBox TxtHeight 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Index           =   0
               Left            =   3180
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Text            =   "10"
               Top             =   1140
               Width           =   525
            End
            Begin VB.TextBox TxtComment 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   345
               Index           =   0
               Left            =   1890
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   180
               Width           =   3015
            End
            Begin MBColorPicker.ColorPicker CPic 
               Height          =   345
               Index           =   0
               Left            =   30
               TabIndex        =   8
               ToolTipText     =   "·Ê‰ Œÿ «·»«—ﬂÊœ"
               Top             =   885
               Width           =   735
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
            Begin MBColorPicker.ColorPicker CPic 
               Height          =   345
               Index           =   1
               Left            =   30
               TabIndex        =   9
               ToolTipText     =   "·Ê‰ Œÿ «·»«—ﬂÊœ"
               Top             =   1260
               Width           =   735
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   345
               Index           =   1
               Left            =   90
               TabIndex        =   10
               Top             =   510
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "..."
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
               ButtonImage     =   "FrmDesOptions.frx":158C
               ColorButton     =   12632256
               DrawFocusRectangle=   0   'False
            End
            Begin MSComctlLib.Toolbar TBr 
               Height          =   330
               Index           =   0
               Left            =   2670
               TabIndex        =   11
               Top             =   600
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               ImageList       =   "ImgListToolBar"
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   7
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Bold"
                     ImageKey        =   "Bold"
                     Style           =   1
                  EndProperty
                  BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Italic"
                     ImageKey        =   "Italic"
                     Style           =   1
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Under"
                     ImageKey        =   "Under"
                     Style           =   1
                  EndProperty
                  BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Style           =   3
                  EndProperty
                  BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Left"
                     ImageKey        =   "Left"
                     Style           =   2
                  EndProperty
                  BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Center"
                     ImageKey        =   "Center"
                     Style           =   2
                     Value           =   1
                  EndProperty
                  BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Right"
                     ImageKey        =   "Right"
                     Style           =   2
                  EndProperty
               EndProperty
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "„·„"
               Height          =   255
               Index           =   17
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   1140
               Width           =   315
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "≈— ›«⁄ «· ⁄·Ìﬁ"
               Height          =   225
               Index           =   16
               Left            =   3750
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   1170
               Width           =   1065
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "·Ê‰ «·‰’"
               Height          =   285
               Index           =   18
               Left            =   840
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   1290
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "·Ê‰ «·Œ·›Ì…"
               Height          =   285
               Index           =   19
               Left            =   840
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   990
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "‰Ê⁄ «·Œÿ"
               Height          =   285
               Index           =   15
               Left            =   870
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   630
               Width           =   855
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3780
         Index           =   0
         Left            =   45
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   45
         Width           =   5055
         _cx             =   8916
         _cy             =   6668
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
         BackColor       =   14737632
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
         Begin VB.Frame Fra 
            BackColor       =   &H00E0E0E0&
            Caption         =   "ÂÊ«„‘"
            Height          =   1515
            Index           =   1
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   2220
            Width           =   3765
            Begin VB.TextBox TxtMargnRight 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   1950
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   180
               Width           =   675
            End
            Begin VB.TextBox TxtMargnLeft 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   1950
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   510
               Width           =   675
            End
            Begin VB.TextBox TxtMargnTop 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   180
               Width           =   615
            End
            Begin VB.TextBox TxtMargnBottom 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   510
               Width           =   615
            End
            Begin ImpulseButton.ISButton CmdDef 
               Height          =   315
               Left            =   1950
               TabIndex        =   48
               Top             =   840
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   " Õ„Ì· «·√› —«÷Ì« "
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmDesOptions.frx":1926
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Â«„‘ √Ì”—"
               Height          =   255
               Index           =   11
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   540
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Â«„‘ √Ì„‰"
               Height          =   255
               Index           =   10
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   210
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Â«„‘ ⁄·ÊÌ"
               Height          =   255
               Index           =   12
               Left            =   780
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   210
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Â«„‘ ”›·Ì"
               Height          =   255
               Index           =   13
               Left            =   780
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   540
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "„·ÕÊŸ…:- ﬁÌ„… «·Â«„‘ «·„œŒ·…  ﬂ » »«·‹ ”‰ „Ì — "
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
               Index           =   14
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   1200
               Width           =   3465
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E0E0E0&
            Caption         =   "›Ê«’·"
            Height          =   825
            Index           =   0
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   1410
            Width           =   3795
            Begin VB.TextBox TxtVerticalBand 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   750
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   150
               Width           =   645
            End
            Begin VB.TextBox TxtHorizontalBand 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   750
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   450
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "„·„"
               Height          =   285
               Index           =   8
               Left            =   450
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   450
               Width           =   195
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "„·„"
               Height          =   285
               Index           =   9
               Left            =   450
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   150
               Width           =   195
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "ﬁÌ„… «·›«’· «·—√”Ì"
               Height          =   255
               Index           =   6
               Left            =   1410
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   150
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "ﬁÌ„… «·›«’· «·√›ﬁÌ"
               Height          =   285
               Index           =   7
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   450
               Width           =   1425
            End
            Begin VB.Image Image1 
               Height          =   240
               Index           =   5
               Left            =   2940
               Picture         =   "FrmDesOptions.frx":1CC0
               Top             =   465
               Width           =   240
            End
            Begin VB.Image Image1 
               Height          =   240
               Index           =   6
               Left            =   2940
               Picture         =   "FrmDesOptions.frx":204A
               Top             =   165
               Width           =   240
            End
         End
         Begin VB.CheckBox ChkDefault 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "«·ÿ«»⁄… «·√› —«÷Ì… ··ÊÌ‰œÊ“"
            Height          =   285
            Left            =   60
            MaskColor       =   &H00C0E0FF&
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   390
            Width           =   2115
         End
         Begin VB.ComboBox CboPaperSize 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   1110
            TabIndex        =   34
            Top             =   750
            Width           =   2595
         End
         Begin VB.TextBox TxtNO 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   345
            Left            =   2700
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   1080
            Width           =   990
         End
         Begin VB.ComboBox CboPrinters 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   1440
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   30
            Width           =   2205
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   315
            Index           =   0
            Left            =   960
            TabIndex        =   54
            Top             =   30
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   2
            Caption         =   ""
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmDesOptions.frx":23D4
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   3
            Left            =   2190
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   390
            Width           =   1425
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "⁄œœ «·„·’ﬁ« "
            Height          =   285
            Index           =   5
            Left            =   3750
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   1140
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "„ﬁ«” «·Ê—ﬁ"
            Height          =   225
            Index           =   4
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   780
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "«·„‘€·"
            Height          =   345
            Index           =   2
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   390
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "«·ÿ«»⁄…"
            Height          =   195
            Index           =   1
            Left            =   3750
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   60
            Width           =   795
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   4680
            Picture         =   "FrmDesOptions.frx":276E
            Top             =   30
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   4710
            Picture         =   "FrmDesOptions.frx":2AF8
            Top             =   810
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   2
            Left            =   4710
            Picture         =   "FrmDesOptions.frx":2E82
            Top             =   1140
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   4
            Left            =   4560
            Picture         =   "FrmDesOptions.frx":320C
            Top             =   1470
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   3
            Left            =   4680
            Picture         =   "FrmDesOptions.frx":3596
            Top             =   2190
            Width           =   240
         End
      End
   End
   Begin ImpulseAniLabel.ISAniLabel LblAddNewPrinter 
      Height          =   285
      Left            =   30
      TabIndex        =   59
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      BackColor       =   14737632
      Alignment       =   2
      Caption         =   "≈÷«›… ÿ«»⁄… ÃœÌœ…"
      ImageCount      =   0
   End
   Begin MSComctlLib.ImageList ImgListToolBar 
      Left            =   30
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDesOptions.frx":3920
            Key             =   "Under"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDesOptions.frx":3CBA
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDesOptions.frx":4054
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDesOptions.frx":43EE
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDesOptions.frx":4788
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDesOptions.frx":4B22
            Key             =   "Center"
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   5160
      Y1              =   5250
      Y2              =   5250
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "·« ÊÃœ ÿ«»⁄«  „⁄—›… ›Ï «·ÃÂ«“"
      Height          =   285
      Index           =   0
      Left            =   2790
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   570
      Width           =   2325
   End
   Begin VB.Label XPLblHeader 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÿ»«⁄… «·»«—ﬂÊœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   525
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "FrmDesOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function OpenPrinter _
                Lib "winspool.drv" _
                Alias "OpenPrinterA" (ByVal pPrinterName As String, _
                                      phPrinter As Long, _
                                      pDefault As Any) As Long

Private Declare Function ClosePrinter _
                Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Private Declare Function PrinterProperties _
                Lib "winspool.drv" (ByVal hWnd As Long, _
                                    ByVal hPrinter As Long) As Long

Private Declare Function DeviceCapabilities _
                Lib "winspool.drv" _
                Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, _
                                             ByVal lpPort As String, _
                                             ByVal iIndex As Long, _
                                             lpOutput As Any, _
                                             lpDevMode As Any) As Long

Private Const DC_PAPERNAMES = 16

Private Const DC_PAPERS = 2

Private Const DC_PAPERSIZE = 3

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Type PaperDim
    PaperWidth As Long
    PaperHeight As Long
End Type

Private Sub CboPaperSize_Click()

    With CboPaperSize

        If .ListIndex <> -1 Then
            'WriteCaps .itemdata(.ListIndex)
            WriteCaps .ListIndex + 1
        End If

    End With

End Sub

Private Sub CboPrinters_Click()
    GetPrinterPaperSizes (CboPrinters.text)
End Sub

Private Sub ChkShow_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
            TxtComment(0).Enabled = CBool(ChkShow(Index).value)
            TBr(0).Enabled = CBool(ChkShow(Index).value)
            TxtHeight(0).Enabled = CBool(ChkShow(Index).value)
            lbl(17).Enabled = CBool(ChkShow(Index).value)
            lbl(18).Enabled = CBool(ChkShow(Index).value)
            lbl(19).Enabled = CBool(ChkShow(Index).value)
            lbl(23).Enabled = CBool(ChkShow(Index).value)
            lbl(16).Enabled = CBool(ChkShow(Index).value)
            CPic(0).Enabled = CBool(ChkShow(Index).value)
            CPic(1).Enabled = CBool(ChkShow(Index).value)
            Cmd(1).Enabled = CBool(ChkShow(Index).value)

        Case 1
            TxtComment(1).Enabled = CBool(ChkShow(Index).value)
            TBr(1).Enabled = CBool(ChkShow(Index).value)
            TxtHeight(1).Enabled = CBool(ChkShow(Index).value)
            lbl(20).Enabled = CBool(ChkShow(Index).value)
            lbl(21).Enabled = CBool(ChkShow(Index).value)
            lbl(22).Enabled = CBool(ChkShow(Index).value)
            lbl(24).Enabled = CBool(ChkShow(Index).value)
            lbl(25).Enabled = CBool(ChkShow(Index).value)
            CPic(2).Enabled = CBool(ChkShow(Index).value)
            CPic(3).Enabled = CBool(ChkShow(Index).value)
            Cmd(2).Enabled = CBool(ChkShow(Index).value)
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim i As Integer
    Dim ObjPrinter As Object
    Dim hPrinter As Long
    Dim OldChar As Integer
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If Me.CboPrinters.ListIndex > -1 Then
                Set ObjPrinter = GetPrinter(Me.CboPrinters.text)

                If Not ObjPrinter Is Nothing Then
                    Screen.MousePointer = vbArrowHourglass
                    OpenPrinter ObjPrinter.DeviceName, hPrinter, ByVal 0&
                    PrinterProperties Me.hWnd, hPrinter
                    ClosePrinter hPrinter
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If

        Case 1
            FrmBarcode.Cdg.CancelError = False
            FrmBarcode.Cdg.Flags = cdlCFBoth + cdlCFLimitSize + cdlCFEffects + cdlCFANSIOnly
            FrmBarcode.Cdg.FontName = TxtComment(0).Font.name
            FrmBarcode.Cdg.FontBold = TxtComment(0).Font.Bold
            FrmBarcode.Cdg.FontItalic = TxtComment(0).Font.Italic
            FrmBarcode.Cdg.FontUnderline = TxtComment(0).Font.Underline
            FrmBarcode.Cdg.FontSize = TxtComment(0).Font.Size
            OldChar = TxtComment(0).Font.Charset
            FrmBarcode.Cdg.color = CPic(1).color
            FrmBarcode.Cdg.Min = 8
            FrmBarcode.Cdg.Max = 14
            FrmBarcode.Cdg.ShowFont
            TxtComment(0).Font.name = FrmBarcode.Cdg.FontName
            TxtComment(0).Font.Bold = FrmBarcode.Cdg.FontBold
            TxtComment(0).Font.Italic = FrmBarcode.Cdg.FontItalic
            TxtComment(0).Font.Underline = FrmBarcode.Cdg.FontUnderline
            TxtComment(0).Font.Size = FrmBarcode.Cdg.FontSize
            TxtComment(0).Font.Charset = OldChar
            TBr(0).Buttons("Bold").value = IIf(TxtComment(0).Font.Bold = True, tbrPressed, tbrUnpressed)
            TBr(0).Buttons("Under").value = IIf(TxtComment(0).Font.Underline = True, tbrPressed, tbrUnpressed)
            TBr(0).Buttons("Italic").value = IIf(TxtComment(0).Font.Italic = True, tbrPressed, tbrUnpressed)
            CPic(1).color = FrmBarcode.Cdg.color

        Case 2
            FrmBarcode.Cdg.CancelError = False
            FrmBarcode.Cdg.Flags = cdlCFBoth + cdlCFLimitSize + cdlCFEffects + cdlCFANSIOnly
            FrmBarcode.Cdg.FontName = TxtComment(1).Font.name
            FrmBarcode.Cdg.FontBold = TxtComment(1).Font.Bold
            FrmBarcode.Cdg.FontItalic = TxtComment(1).Font.Italic
            FrmBarcode.Cdg.FontUnderline = TxtComment(1).Font.Underline
            FrmBarcode.Cdg.FontSize = TxtComment(1).Font.Size
            OldChar = TxtComment(1).Font.Charset
            FrmBarcode.Cdg.color = CPic(3).color
            FrmBarcode.Cdg.Min = 8
            FrmBarcode.Cdg.Max = 14
            FrmBarcode.Cdg.ShowFont
            TxtComment(1).Font.name = FrmBarcode.Cdg.FontName
            TxtComment(1).Font.Bold = FrmBarcode.Cdg.FontBold
            TxtComment(1).Font.Italic = FrmBarcode.Cdg.FontItalic
            TxtComment(1).Font.Underline = FrmBarcode.Cdg.FontUnderline
            TxtComment(1).Font.Size = FrmBarcode.Cdg.FontSize
            TxtComment(1).Font.Charset = OldChar
            TBr(1).Buttons("Bold").value = IIf(TxtComment(1).Font.Bold = True, tbrPressed, tbrUnpressed)
            TBr(1).Buttons("Under").value = IIf(TxtComment(1).Font.Underline = True, tbrPressed, tbrUnpressed)
            TBr(1).Buttons("Italic").value = IIf(TxtComment(1).Font.Italic = True, tbrPressed, tbrUnpressed)
            CPic(3).color = FrmBarcode.Cdg.color
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdDef_Click()
    DefaultMargins
End Sub

Private Sub CmdNo_Click()
    Unload Me
End Sub

Private Sub Cmdyes_Click()
    On Error GoTo ErrTrap
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim IntNoRows As Integer
    Dim IntNoCols As Integer
    Dim LngNoLabels As Long
    Dim SngFactorMM_Twip As Single
    Dim SngImgWidth As Single   'By Twip
    Dim SngImgHeight As Single  'By Twip
    Dim SngImgLeft As Single
    Dim SngTopLblHeight As Single
    Dim SngBotLblHeight As Single

    Dim SngPageDim As PaperDim
    Dim SngPageWidth As Single
    Dim SngVerSpacing As Single ' «·„”«›… «·—√”Ì… »Ì‰ «·√” ﬂÌ—« 
    Dim sngHorSpacing As Single ' «·„”«›… «·√›ﬁÌ… »Ì‰ «·√” ﬂÌ—« 
    Dim IntCounter As Integer
    'Dim Rpt As DataRptBarcode

    'Convert From mm to Twip
    'Why 576 ......>
    'Why 10  ......>
    'A twip is a unit of length equal to 1/20 of a printer's point,
    'and a printer's point is 1/72 of an inch. There are approximately 1440
    'twips to a logical inch or 567 twips to a logical centimeter
    '(the length of a screen item measuring one inch or one centimeter
    'when printed).
    If CboPrinters.ListIndex = -1 Then

    End If

    If Me.CboPaperSize.ListIndex = -1 Then
        GetMsgs 157, vbExclamation

        If CboPaperSize.Enabled = True Then
            CboPaperSize.SetFocus
        End If

        Exit Sub
    End If

    LngNoLabels = val(TXTNO.text) 'Number Of lables

    If LngNoLabels = 0 Then
        GetMsgs 160, vbExclamation
        TXTNO.SetFocus
        Exit Sub
    End If

    '--------Initialize variables
    SngFactorMM_Twip = 567 / 10

    SngVerSpacing = val(Me.TxtVerticalBand.text) * SngFactorMM_Twip
    sngHorSpacing = val(Me.TxtHorizontalBand.text) * SngFactorMM_Twip
    FrmBarcode.Axbarcode1.Refresh
    SngImgHeight = FrmBarcode.Axbarcode1.ImageHeight * SngFactorMM_Twip
    SngImgWidth = FrmBarcode.Axbarcode1.ImageWidth * SngFactorMM_Twip

    SngTopLblHeight = val(Me.TxtHeight(0).text) * SngFactorMM_Twip
    SngBotLblHeight = val(Me.TxtHeight(1).text) * SngFactorMM_Twip

    'SngImgHeight = FrmBarcode.Axbarcode1.PictureHeight * SngFactorMM_Twip
    'SngImgWidth = FrmBarcode.Axbarcode1.PictureWidth * SngFactorMM_Twip
    '----------------------------
    'Set Rpt = New DataRptBarcode
    rs.Fields.Append "LblNumber", adInteger
    rs.Open

    With DataRptBarcode
        'Set The Main Report Properties
        .Sections("Section4").Visible = False
        .Sections("Section3").Visible = False
        .Sections("Section2").Visible = False
        .Sections("Section5").Visible = False
        .Sections("Section4").Height = 0
        .Sections("Section3").Height = 0
        .Sections("Section2").Height = 0
        .Sections("Section5").Height = 0
        'Convert all From Cm to Twip
        .LeftMargin = val(Me.TxtMargnLeft.text) * 567
        .RightMargin = val(Me.TxtMargnRight.text) * 567
        .TopMargin = val(Me.TxtMargnTop.text) * 567
        .BottomMargin = val(Me.TxtMargnBottom.text) * 567

        With CboPaperSize
            'Get the Seelcted page width by mm
            SngPageDim = GetPaperSizeDim(.ListIndex + 1, Me.CboPrinters.text, .ItemData(.ListIndex))
        End With

        'Convert this width into twip
        SngPageWidth = SngPageDim.PaperWidth * SngFactorMM_Twip
        .ReportWidth = ((SngPageWidth)) - (.LeftMargin + .RightMargin + 1)

        With .Sections("Section1")
            '«·√— ›«⁄ Ì”«ÊÏ ≈— ›«⁄ «·’Ê—… + «·„”«›… «·—√”Ì… «·›«’·…
            IntCounter = 0
            SngImgLeft = 0
            IntNoCols = 0

            If ChkShow(0).value = vbChecked Then

                Do
                    .Controls("lbl" & IntCounter).top = 0
                    .Controls("lbl" & IntCounter).left = SngImgLeft
                    .Controls("lbl" & IntCounter).Width = SngImgWidth
                    .Controls("lbl" & IntCounter).Height = SngTopLblHeight
                    .Controls("lbl" & IntCounter).Caption = TxtComment(0).text
                    Set .Controls("lbl" & IntCounter).Font = TxtComment(0).Font
                    .Controls("lbl" & IntCounter).Alignment = TxtComment(0).Alignment
                    .Controls("lbl" & IntCounter).BackColor = Me.CPic(0).color
                    .Controls("lbl" & IntCounter).ForeColor = Me.CPic(1).color
                    .Controls("lbl" & IntCounter).Visible = True
                    SngImgLeft = .Controls("lbl" & IntCounter).left + SngImgWidth + sngHorSpacing
                    IntCounter = IntCounter + 1
                Loop While (SngImgLeft + SngImgWidth) < DataRptBarcode.ReportWidth

                For i = IntCounter To 14
                    .Controls("lbl" & i).left = 0
                    .Controls("lbl" & i).top = 0
                    .Controls("lbl" & IntCounter).Font.Bold = False
                    .Controls("lbl" & IntCounter).Font.Italic = False
                    .Controls("lbl" & IntCounter).Font.Underline = False
                    .Controls("lbl" & i).Width = SngImgWidth
                    .Controls("lbl" & i).Height = 0
                    .Controls("lbl" & i).Visible = False
                Next i

            Else
                SngTopLblHeight = 0

                For i = 0 To 14
                    .Controls("lbl" & i).top = 0
                    .Controls("lbl" & i).left = 0
                    .Controls("lbl" & i).Font.Bold = False
                    .Controls("lbl" & i).Font.Italic = False
                    .Controls("lbl" & i).Font.Underline = False
                    .Controls("lbl" & i).Width = 0
                    .Controls("lbl" & i).Height = 0
                    .Controls("lbl" & i).Visible = False
                Next i

            End If

            '------------------------------------
            IntCounter = 0
            SngImgLeft = 0
            IntNoCols = 0

            Do
                .Controls("Img" & IntCounter).left = SngImgLeft
                .Controls("Img" & IntCounter).top = SngTopLblHeight
                .Controls("Img" & IntCounter).Width = SngImgWidth
                .Controls("Img" & IntCounter).Height = SngImgHeight
                .Controls("Img" & IntCounter).Visible = True
                .Controls("Img" & IntCounter).SizeMode = rptSizeZoom
                Set .Controls("Img" & IntCounter).Picture = Nothing
                Set .Controls("Img" & IntCounter).Picture = FrmBarcode.Axbarcode1.Picture
                SngImgLeft = .Controls("Img" & IntCounter).left + SngImgWidth + sngHorSpacing
                IntCounter = IntCounter + 1
                IntNoCols = IntNoCols + 1
            Loop While (SngImgLeft + SngImgWidth) < DataRptBarcode.ReportWidth

            If SystemOptions.SysVersion = DemoVersion Then
                .Controls("LblDemo").Visible = True
                .Controls("LblDemo").left = .Controls("Img" & 0).left
                .Controls("LblDemo").top = .Controls("Img" & 0).top
                .Controls("LblDemo").Width = SngImgLeft
                .Controls("LblDemo").Height = .Controls("Img" & 0).Height
                '.Controls("LblDemo").Font.Size = 36
            Else
                .Controls("LblDemo").Visible = False
                .Controls("LblDemo").left = .Controls("Img" & 0).left
                .Controls("LblDemo").top = .Controls("Img" & 0).top
                .Controls("LblDemo").Height = .Controls("Img" & 0).Height
                '.Controls("LblDemo").Font.Size = 8
            End If

            'Hide the other images
            For i = IntCounter To 14
                .Controls("Img" & i).left = .Controls("Img" & 0).left
                .Controls("Img" & i).top = .Controls("Img" & 0).top
                .Controls("Img" & i).Width = .Controls("Img" & 0).Width
                .Controls("Img" & i).Height = .Controls("Img" & 0).Height
                .Controls("Img" & i).Visible = False
            Next i

            '------------------------------------
            IntCounter = 0
            SngImgLeft = 0

            If ChkShow(1).value = vbChecked Then

                Do
                    .Controls("lblB" & IntCounter).top = SngTopLblHeight + SngImgHeight
                    .Controls("lblB" & IntCounter).left = SngImgLeft
                    .Controls("lblB" & IntCounter).Width = SngImgWidth
                    .Controls("lblB" & IntCounter).Height = SngBotLblHeight
                    .Controls("lblB" & IntCounter).Caption = TxtComment(1).text
                    Set .Controls("lblB" & IntCounter).Font = TxtComment(1).Font
                    .Controls("lblB" & IntCounter).Alignment = TxtComment(1).Alignment
                    .Controls("lblB" & IntCounter).BackColor = Me.CPic(2).color
                    .Controls("lblB" & IntCounter).ForeColor = Me.CPic(3).color
                    .Controls("lblB" & IntCounter).Visible = True
                    SngImgLeft = .Controls("lblB" & IntCounter).left + SngImgWidth + sngHorSpacing
                    IntCounter = IntCounter + 1
                Loop While (SngImgLeft + SngImgWidth) < DataRptBarcode.ReportWidth

                For i = IntCounter To 14
                    .Controls("lblB" & i).left = 0
                    .Controls("lblB" & i).top = 0
                    .Controls("lbl" & i).Font.Bold = False
                    .Controls("lbl" & i).Font.Italic = False
                    .Controls("lbl" & i).Font.Underline = False
                    .Controls("lblB" & i).Width = SngImgWidth
                    .Controls("lblB" & i).Height = SngImgHeight
                    .Controls("lblB" & i).Visible = False
                Next i

            Else
                SngBotLblHeight = 0

                For i = 0 To 14
                    .Controls("lblB" & i).top = 0
                    .Controls("lblB" & i).left = 0
                    .Controls("lblB" & i).Font.Bold = False
                    .Controls("lblB" & i).Font.Italic = False
                    .Controls("lblB" & i).Font.Underline = False
                    .Controls("lblB" & i).Width = 0
                    .Controls("lblB" & i).Height = 0
                    .Controls("lblB" & i).Visible = False
                Next i

            End If

            '------------------------------------
            '.Height = SngTopLblHeight + SngBotLblHeight + SngImgHeight + SngVerSpacing
            .Height = .Controls("lblB" & 0).Height + .Controls("lbl" & 0).Height + SngBotLblHeight + SngImgHeight + SngVerSpacing
            IntNoRows = (LngNoLabels \ IntNoCols)

            If LngNoLabels Mod IntNoCols > 0 Then
                IntNoRows = IntNoRows + 1
            End If

            For i = 1 To IntNoRows
                rs.AddNew "LblNumber", i
                rs.update
            Next

        End With

        Set .DataSource = rs
        .WindowState = vbMaximized
        .Show vbModeless
        .ZOrder 0
        '    Rs.Close
        '    Set Rs = Nothing
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        TabMain.SetFocus

        If KeyCode = vbKeyTab Then
            If TabMain.CurrTab = 0 Then
                TabMain.CurrTab = 1

                If ChkShow(0).Enabled = True Then
                    ChkShow(0).SetFocus
                End If

            Else
                TabMain.CurrTab = 0
                CboPrinters.SetFocus
            End If
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            CmdNo_Click
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    'vbPRPSA4
    Dim i As Integer
    Dim IntDefIndex As Integer
    On Error GoTo ErrTrap
    CenterForm Me

    FormPostion Me, GetPostion
    AddTip
    Me.TabMain.CurrTab = 0

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Me.Icon = FrmBarcode.Icon
    CboPrinters.Clear
    CboPaperSize.Clear

    If Printers.count > 0 Then
        lbl(3).Visible = False

        For i = 0 To Printers.count - 1
            CboPrinters.AddItem Printers(i).DeviceName

            If Printer.DeviceName = Printers(i).DeviceName Then
                IntDefIndex = i
                lbl(3).Caption = Printers(i).DriverName
            End If

        Next i

        CboPrinters.ListIndex = IntDefIndex
    Else
        lbl(0).Visible = True
        DisableAll
    End If

    'AddPaperSize
    DefaultMargins
    BarcodeSetting 2

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.Caption = "Barcode Print"
    Me.TabMain.TabCaption(0) = "Page Setup"
    lbl(3).Caption = "No Printers Installed in System"
    lbl(2).Caption = "Paper Size"
    lbl(1).Caption = "Lables Number"
    lbl(0).Caption = "Printer"
    lbl(12).Caption = "Vertical Spacing"
    lbl(13).Caption = "Horizontal Spacing"
    lbl(14).Caption = "mm."
    lbl(15).Caption = "mm."
    lbl(23).Caption = "mm."
    lbl(24).Caption = "mm."
    lbl(5).Caption = "RightMargin"
    lbl(6).Caption = "LeftMargin"
    lbl(7).Caption = "TopMargin"
    lbl(8).Caption = "BottomMargin"
    CmdDef.Caption = "Load Default"
    lbl(9).Caption = "Note:- Margins are in Cm."
    lbl(11).Caption = "Driver Name : "
    ChkDefault.Caption = "Set as Default Printer"
    Cmdyes.Caption = "&OK"
    CmdNo.Caption = "&Cancel"
    LblAddNewPrinter.Caption = "Add New Printer..."
    Me.TabMain.TabCaption(1) = "Other Data"
    ChkShow(0).Caption = "Show Top Comment"
    lbl(17).Caption = "Comment Height"
    lbl(18).Caption = "Background Color"
    lbl(19).Caption = "Forecolor Color"
    ChkShow(1).Caption = "Show Bottom Comment"
    lbl(20).Caption = "Comment Height"
    lbl(22).Caption = "Background Color"
    lbl(21).Caption = "Forecolor Color"

    For i = 0 To CPic.UBound
        CPic(i).CustomButtonText = "Custom..."
    Next i

    Exit Sub
ErrTrap:
End Sub

Private Sub AddPaperSize()
    On Error GoTo ErrTrap

    With Me.CboPaperSize
        .AddItem "10x14", 0
        .ItemData(0) = vbPRPS10x14
    
        .AddItem "11x17", 1
        .ItemData(1) = vbPRPS11x17
    
        .AddItem "A3"
        .ItemData(2) = vbPRPSA3
    
        .AddItem "A4"
        .ItemData(3) = vbPRPSA4
    
        .AddItem "A4Small"
        .ItemData(4) = vbPRPSA4Small
    
        .AddItem "A5"
        .ItemData(5) = vbPRPSA5
        .AddItem "B4"
        .ItemData(6) = vbPRPSB4
    
        .AddItem "B5"
        .ItemData(7) = vbPRPSB5
    
        .AddItem "CSheet"
        .ItemData(8) = vbPRPSCSheet
    
        .AddItem "DSheet"
        .ItemData(9) = vbPRPSDSheet
    
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub WriteCaps(Lngid As Long)
    On Error GoTo ErrTrap
    Dim StrCaps As String

    Dim StrWidth As String
    Dim StrHeight As String
    Dim StrUnitIn As String
    Dim StrUnitCm As String
    Dim StrScale As String
    Dim Dims As PaperDim

    Dim Wrap As String
    Wrap = Chr(13) + Chr(10)

    If SystemOptions.UserInterface = EnglishInterface Then
        StrWidth = "Width: "
        StrHeight = "Height: "
        StrUnitIn = "Inch "
        StrUnitCm = "mm"
        StrScale = "Scale: "
    Else
        StrWidth = "⁄—÷: "
        StrHeight = "≈— ›«⁄: "
        StrUnitIn = "»Ê’…"
        StrUnitCm = "„·„"
        StrScale = "«·„ﬁ«” »‹ :"
    End If

    Dims = GetPaperSizeDim(Lngid, Me.CboPrinters.text, CboPaperSize.ItemData(CboPaperSize.ListIndex))
    StrCaps = StrWidth & Dims.PaperWidth & Wrap & StrHeight & Dims.PaperHeight & Wrap & StrScale & StrUnitCm
    Me.lbl(4).Caption = StrCaps
    Exit Sub
ErrTrap:
End Sub

Private Sub AddTip()

    Dim Msg As String
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim i As Integer
    Dim BolRtl As Boolean
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True

        With TTP
            .Create Me.hWnd, "√Œ — ÿ«»⁄…", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "√Œ — «·ÿ«»⁄… «· Ï  —Ìœ ≈” Œœ«„Â« ›Ï" & Wrap & "⁄„·Ì… «·ÿ»«⁄…." & Wrap & "" & Wrap & "„·ÕÊŸ…:-»⁄œ ≈Œ Ì«—ﬂ ··ÿ«»⁄… ”Ê› " & Wrap & "ÌﬁÊ„ «·»—‰«„Ã »«·√” ⁄·«„ ⁄‰ ÕÃ„ «·Ê—ﬁ" & Wrap & "«·–Ï  œ⁄„Â Â–Â «·ÿ«»⁄… Ê⁄—÷Â ﬁÏ   " & Wrap & "(ﬁ«∆„… ÕÃ„ «·Ê—ﬁ)."
            .AddControl CboPrinters, Msg, True
        End With

        With TTP
            .Create Me.hWnd, "√Œ — ÕÃ„ «·Ê—ﬁ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "Â‰«  ⁄—÷ Ã„Ì⁄ √ÕÃ«„ «·Ê—ﬁ «·–Ï" & Wrap & " œ⁄„Â «·ÿ«»⁄… «·„Œ «—…. "
            .AddControl CboPaperSize, Msg, True
        End With

        With TTP
            .Create Me.hWnd, "ÿ«»⁄… ≈› —«÷Ì…", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "» ›⁄Ì· Â–« «·ŒÌ«— ”Ê› ÌﬁÊ„  " & Wrap & "«·»—‰«„Ã »Ã⁄· «·ÿ«»⁄… «·„Œ «—…" & Wrap & "ÂÏ «·ÿ«»⁄… «·√› —«÷Ì… ›Ï ‰Ÿ«„ " & Wrap & "«·ÊÌ‰œÊ“"
            .AddControl ChkDefault, Msg, True
        End With

        With TTP
            .Create Me.hWnd, "≈÷«›… ÿ«»⁄… ÃœÌœ…", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "≈÷€ÿ Â‰« Õ Ï ÌﬁÊ„ «·»—‰«„Ã   " & Wrap & "» ‘€Ì· „⁄«·Ã ≈÷«›… «·ÿ«»⁄« " & Wrap & "«·ÃœÌœ… «·„ÊÃÊ ›Ï «·ÊÌ‰œÊ“ "
            .AddControl LblAddNewPrinter, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "⁄œœ «·„·’ﬁ« ( «·√” ﬂÌ—)", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "√œŒ· Â‰« ⁄œœ «·„·’ﬁ« («·√” ﬂÌ—« ) «· Ï" & Wrap & " —Ìœ ÿ»«⁄ Â«."
            .AddControl TXTNO, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "ﬁÌ„… «·›«’· «·—√”Ï", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "„ﬁœ«— «·„”«›… «·›«’·… »Ì‰ ﬂ· „·’ﬁ" & Wrap & " Ê«·„·’ﬁ «·„Ã«Ê— ·Â ›Ï ‰›” «·⁄„Êœ." & Wrap & "" & Wrap & "»„⁄‰Ï „ﬁœ«— «·„”«›… «·›«’·… »Ì‰ ﬂ·" & Wrap & "„·’ﬁ Ê«·„·’ﬁ «·„ÊÃÊœ ›ÊﬁÂ √Ê  Õ Â"
            .AddControl TxtVerticalBand, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "ﬁÌ„… «·›«’· «·√›ﬁÏ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "„ﬁœ«— «·„”«›… «·›«’·… »Ì‰ ﬂ· „·’ﬁ " & Wrap & " Ê«·„·’ﬁ «·„Ã«Ê— ·Â ›Ï ‰›” «·’‰›." & Wrap & "" & Wrap & "»„⁄‰Ï «·„”«›… »Ì‰ ﬂ· „·’ﬁ Ê«·„·’ﬁ " & Wrap & " ⁄·Ï Ì„Ì‰Â √Ê Ì”«—Â ›Ï ‰›” «·’‰›"
            .AddControl TxtHorizontalBand, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, " Õ„Ì· ≈› —«÷Ì«  «·ÂÊ«„‘", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "»«·÷€ÿ Â‰« ”Ê› ÌﬁÊ„ «·»—‰«„Ã " & Wrap & "»Ê÷⁄ «·ﬁÌ„ «·√› —«÷Ì… ··ÂÊ«„‘ " & Wrap & "ﬂ„« Ì„ﬂ‰ﬂ «· ⁄œÌ· ›Ï Â–Â «·ﬁÌ„" & Wrap & "" & Wrap & "„·ÕÊŸ…:-" & Wrap & "Â–Â «·ﬁÌ„  ﬁ«” »«·”‰ „Ì—" & Wrap & "" & Wrap & "«·ﬁÌ„ «·√› —«÷Ì… ÂÏ :-" & Wrap & "«·Â«„‘ «·√Ì„‰: " & Round(1440 / 567, 2) & " ”„." & Wrap & "«·Â«„‘ «·√Ì”—: " & Round(1440 / 567, 2) & " ”„." & Wrap & "«·Â«„‘ «·⁄·ÊÏ: " & Round(1440 / 567, 2) & " ”„." & Wrap & "«·Â«„‘ «·”›·Ï: " & Round(1440 / 567, 2) & " ”„."
            .AddControl CmdDef, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "ﬁÌ„… «·Â«„‘ «·√Ì„‰", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "ﬁÌ„… «·Â«„‘ «·√Ì„‰ ··Ê—ﬁ…" & Wrap & "«·ﬁÌ„… «·√› —«÷Ì…: " & Round(1440 / 567, 2) & " ”„."
            .AddControl TxtMargnRight, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "ﬁÌ„… «·Â«„‘ «·√Ì”—", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "ﬁÌ„… «·Â«„‘ «·√Ì”— ··Ê—ﬁ…" & Wrap & "«·ﬁÌ„… «·√› —«÷Ì…: " & Round(1440 / 567, 2) & " ”„."
            .AddControl TxtMargnLeft, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "ﬁÌ„… «·Â«„‘ «·⁄·ÊÏ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "ﬁÌ„… «·Â«„‘ «·⁄·ÊÏ ··Ê—ﬁ…" & Wrap & "«·ﬁÌ„… «·√› —«÷Ì…: " & Round(1440 / 567, 2) & " ”„."
            .AddControl TxtMargnTop, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "ﬁÌ„… «·Â«„‘ «·”›·Ï", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "ﬁÌ„… «·Â«„‘ «·”›·Ï ··Ê—ﬁ…" & Wrap & "«·ﬁÌ„… «·√› —«÷Ì…: " & Round(1440 / 567, 2) & " ”„."
            .AddControl TxtMargnBottom, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "≈ŸÂ«— «· ⁄·Ìﬁ «·⁄·ÊÏ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "» ›⁄Ì· Â–« «·ŒÌ«— Ì„ﬂ‰ﬂ ≈ŸÂ«— Œ«‰…" & Wrap & " ﬂ· „·’ﬁ ·ﬂ «»… ›ÌÂ« «Ì… »Ì«‰«  " & Wrap & "„À· «”„ «·’‰› «Ê «”„ «·‘—ﬂ… «Ê " & Wrap & " «·„‰‘√…."
            .AddControl ChkShow(0), Msg, True
        End With

        With TTP
            .Create Me.hWnd, "‰Ê⁄ «·Œÿ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "≈÷€ÿ Â‰« ·ÌŸÂ— „—»⁄ √Œ Ì«— ‰Ê⁄ " & Wrap & "«·Œÿ ·· ⁄·Ìﬁ «·⁄·ÊÏ Êﬂ·–·ﬂ ÕÃ„Â."
            .AddControl Cmd(1), Msg, True
        End With

        With TTP
            .Create Me.hWnd, "√— ›«⁄ Œ«‰… «· ⁄·Ìﬁ «·⁄·ÊÏ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "«ﬂ » ﬁÌ„… ≈— ›«⁄ «· ⁄·Ìﬁ «·⁄·ÊÏ " & Wrap & "√⁄·Ï «·„·’ﬁ «·ﬁÌ„… «·√› —«÷Ì…=2 „·„"
            .AddControl TxtHeight(0), Msg, True
        End With
    
        With TTP
            .Create Me.hWnd, " Œ’«∆’ «·ÿ«»⁄…", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "⁄—÷ Œ’«∆’ «·ÿ«»⁄… «·„Õœœ…"
            .AddControl Cmd(0), Msg, True
        End With
    
        With TTP
            .Create Me.hWnd, " ÕœÌœ Œ’«∆’ «·‰’", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = " ÕœÌœ Œ’«∆’ «·‰’ „À· «·„Œ«–«… ··Ì„Ì‰" & Wrap & "√Ê «·Ì”«— Ê«Ì÷« Œ’«∆’ «·Œÿ"
            .AddControl Me.TBr(0), Msg, True
        End With

    Else
        BolRtl = False

        With TTP
            .Create Me.hWnd, "Select Printer", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "Select the printer which you want to" & Wrap & " use in Print." & Wrap & "" & Wrap & "Note:-When you choose a printer  " & Wrap & "the Programe will show all supported" & Wrap & "paper size in the paper size Combo List."
            .AddControl CboPrinters, Msg, False
        End With

        With TTP
            .Create Me.hWnd, "Select Paper Size", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "All the supported paper size" & Wrap & "-with your printer-are listed" & Wrap & "here choose the paper size."
            .AddControl CboPaperSize, Msg, False
        End With

        With TTP
            .Create Me.hWnd, "Add New Printer", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "Click here to lunch the" & Wrap & "Add New Printer wizard."
            .AddControl LblAddNewPrinter, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Default Printer", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "When enable this check the programe" & Wrap & "will set the selected printer as " & Wrap & "the default windows printer."
            .AddControl ChkDefault, Msg, False
        End With

        With TTP
            .Create Me.hWnd, "Lables Number", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "Enter number of the lables" & Wrap & "which you want to print."
            .AddControl TXTNO, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Vertical Spacing", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "The distance betweens labels" & Wrap & "it the same column."
            .AddControl TxtVerticalBand, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Horizonral Spacing", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "The distance betweens labels" & Wrap & "it the same row."
            .AddControl TxtHorizontalBand, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Load Default Margins", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "Click here to load the Default Margins" & Wrap & "and you can edit or change this Default" & Wrap & "Values." & Wrap & "" & Wrap & "Note:-" & Wrap & "This values are in centimeter." & Wrap & "" & Wrap & "Default Values :-" & Wrap & "Right Margin: " & Round(1440 / 567, 2) & " Cm." & Wrap & "Left Margin: " & Round(1440 / 567, 2) & " Cm." & Wrap & "Top Margin: " & Round(1440 / 567, 2) & " Cm." & Wrap & "Bottom Margin: " & Round(1440 / 567, 2) & " Cm."
                
            .AddControl CmdDef, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Right Margin", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "Right Margin" & Wrap & "Defalut Value: " & Round(1440 / 567, 2) & " Cm."
            .AddControl TxtMargnRight, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Left Margin", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "Left Margin" & Wrap & "Defalut Value: " & Round(1440 / 567, 2) & " Cm."
            .AddControl TxtMargnLeft, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Top Margin", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "Top Margin" & Wrap & "Defalut Value: " & Round(1440 / 567, 2) & " Cm."
            .AddControl TxtMargnTop, Msg, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bottom Margin", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "Bottom Margin" & Wrap & "Defalut Value: " & Round(1440 / 567, 2) & " Cm."
            .AddControl TxtMargnBottom, Msg, BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DefaultMargins()
    Me.TxtMargnBottom.text = Round(1440 / 567, 2)
    Me.TxtMargnLeft.text = Round(1440 / 567, 2)
    Me.TxtMargnRight.text = Round(1440 / 567, 2)
    Me.TxtMargnTop.text = Round(1440 / 567, 2)
End Sub

Private Function GetPrinter(Optional StrPrintrName As String = "") As Object
    Dim i As Integer
    On Error GoTo ErrTrap

    If StrPrintrName = "" Then
        Set GetPrinter = Printer
    Else

        For i = 0 To Printers.count - 1

            If Printers(i).DeviceName = Trim(StrPrintrName) Then
                Set GetPrinter = Printers(i)
                Exit Function
            End If

        Next i

    End If

    Exit Function
ErrTrap:
    Set GetPrinter = Nothing
End Function

Private Sub GetPrinterPaperSizes(Optional StrPrinterName As String = "")
    Dim ObjPrinter As Object
    Dim ret As Long, PaperSizes() As Byte, PaperSizesNum() As Integer, Cnt As Long, AllNames As String
    Dim lStart As Long, lEnd As Long
    Dim PapersDim() As POINTAPI
    On Error GoTo ErrTrap
    Set ObjPrinter = GetPrinter(StrPrinterName)

    If Not ObjPrinter Is Nothing Then
        ret = DeviceCapabilities(ObjPrinter.DeviceName, ObjPrinter.Port, DC_PAPERNAMES, ByVal 0&, ByVal 0&)
        'resize the array
        ReDim PaperSizes(1 To ret * 64) As Byte
        'retrieve all the available paper names
        Call DeviceCapabilities(ObjPrinter.DeviceName, ObjPrinter.Port, DC_PAPERNAMES, PaperSizes(1), ByVal 0&)
        Me.CboPaperSize.Clear
        'convert the retrieved byte array to a string
        AllNames = StrConv(PaperSizes, vbUnicode)

        'loop through the string and search for the names of the papers
        Do
            lEnd = InStr(lStart + 1, AllNames, Chr$(0), vbBinaryCompare)

            If (lEnd > 0) And (lEnd - lStart - 1 > 0) Then
                Me.CboPaperSize.AddItem Mid$(AllNames, lStart + 1, lEnd - lStart - 1)
            End If

            lStart = lEnd
        Loop Until lEnd = 0

        '==========================
        ret = DeviceCapabilities(ObjPrinter.DeviceName, ObjPrinter.Port, DC_PAPERS, ByVal 0&, ByVal 0&)
        ReDim PaperSizesNum(1 To ret) As Integer
        Call DeviceCapabilities(ObjPrinter.DeviceName, ObjPrinter.Port, DC_PAPERS, PaperSizesNum(1), ByVal 0&)

        For Cnt = 1 To ret
            'Put in the item data property
            'the paper size value
            Me.CboPaperSize.ItemData(Cnt - 1) = (PaperSizesNum(Cnt))

            If ObjPrinter.PaperSize = PaperSizesNum(Cnt) Then
                Me.CboPaperSize.ListIndex = Cnt - 1
                Exit For
            End If

        Next

        '==========================
    End If

    Exit Sub
ErrTrap:
End Sub

Private Function GetPaperSizeDim(LngPaperIndex As Long, _
                                 Optional StrPrinterName As String = "", _
                                 Optional LngPaperSize As Long = 0) As PaperDim
    
    Dim ObjPrinter As Object
    Dim ret As Long, Cnt As Long
    Dim PapersDim() As POINTAPI
    Set ObjPrinter = GetPrinter(StrPrinterName)
    On Error GoTo ErrTrap

    If Not ObjPrinter Is Nothing Then
        If LngPaperSize = 0 Then
            LngPaperSize = ObjPrinter.PaperSize
        End If

        ret = DeviceCapabilities(ObjPrinter.DeviceName, ObjPrinter.Port, DC_PAPERSIZE, ByVal 0&, ByVal 0&)
        ReDim PapersDim(1 To ret) As POINTAPI
        Call DeviceCapabilities(ObjPrinter.DeviceName, ObjPrinter.Port, DC_PAPERSIZE, PapersDim(1), ByVal 0&)
        GetPaperSizeDim.PaperWidth = (PapersDim(LngPaperIndex).x)
        'to retrun by mm
        GetPaperSizeDim.PaperWidth = GetPaperSizeDim.PaperWidth / 10
        GetPaperSizeDim.PaperHeight = (PapersDim(LngPaperIndex).Y)
        'to retrun by mm
        GetPaperSizeDim.PaperHeight = GetPaperSizeDim.PaperHeight / 10
    End If

    Exit Function
ErrTrap:
    GetPaperSizeDim.PaperWidth = -1
    GetPaperSizeDim.PaperHeight = -1
End Function

Private Sub DisableAll()
    Dim i As Integer
    On Error GoTo ErrTrap
    CboPrinters.Enabled = False
    CboPaperSize.Enabled = False
    ChkDefault.Enabled = False
    TXTNO.Enabled = False
    TxtMargnBottom.Enabled = False
    TxtMargnLeft.Enabled = False
    TxtMargnRight.Enabled = False
    TxtMargnTop.Enabled = False

    CmdDef.Enabled = False
    Cmdyes.Enabled = False

    For i = lbl.LBound To lbl.UBound
        lbl(i).Enabled = False
    Next i

    For i = Image1.LBound To Image1.UBound
        Image1(i).Enabled = False
    Next i

    lbl(3).Enabled = True
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    BarcodeSetting 1
End Sub

Private Sub Label1_Click()

End Sub

Private Sub LblAddNewPrinter_Click()
    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass
    Shell "rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL AddPrinter", vbNormalFocus
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
End Sub

Private Sub BarcodeSetting(IntMode As Integer)
    'IntMode=1 -----> Save
    'IntMOde=2 -----> Load
    On Error GoTo ErrTrap

    If IntMode = 1 Then
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TxtNO", val(Me.TXTNO.text)
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TxtVerticalBand", val(Me.TxtVerticalBand.text)
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TxtHorizontalBand", val(Me.TxtHorizontalBand.text)
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TxtMargnRight", val(Me.TxtMargnRight.text)
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TxtMargnLeft", val(Me.TxtMargnLeft.text)
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TxtMargnTop", val(Me.TxtMargnTop.text)
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TxtMargnBottom", val(Me.TxtMargnBottom.text)
    
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "ChkShow0", ChkShow(0).value
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "ChkShow1", ChkShow(1).value
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopComment", TxtComment(0).text
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomComment", TxtComment(1).text
     
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopHeight", TxtHeight(0).text
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomHeight", TxtHeight(1).text
    
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "Cpic0", CPic(0).color
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "Cpic1", CPic(1).color
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "Cpic2", CPic(2).color
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "Cpic3", CPic(3).color
    
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopCommentAlign", TxtComment(0).Alignment
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomCommentAlign", TxtComment(1).Alignment
    
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopCommentBold", TxtComment(0).Font.Bold
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomCommentBold", TxtComment(1).Font.Bold
    
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopCommentItalic", TxtComment(0).Font.Italic
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomCommentItalic", TxtComment(1).Font.Italic
    
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopCommentUnder", TxtComment(0).Font.Underline
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomCommentUnder", TxtComment(1).Font.Underline
       
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopCommentFontName", TxtComment(0).Font.name
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopCommentFontSize", TxtComment(0).Font.Size
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomCommentFontName", TxtComment(1).Font.name
        SaveSetting SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomCommentFontSize", TxtComment(1).Font.Size
    ElseIf IntMode = 2 Then
        Me.TXTNO.text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TxtNO", 100)
        Me.TxtVerticalBand.text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TxtVerticalBand", 2)
        Me.TxtHorizontalBand.text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TxtHorizontalBand", 2)
        Me.TxtMargnRight.text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TxtMargnRight", TxtMargnRight.text)
        Me.TxtMargnLeft.text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TxtMargnLeft", TxtMargnLeft.text)
        Me.TxtMargnTop.text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TxtMargnTop", TxtMargnTop.text)
        Me.TxtMargnBottom.text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TxtMargnBottom", TxtMargnBottom.text)
    
        ChkShow(0).value = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "ChkShow0", vbUnchecked)
        ChkShow_Click 0
        ChkShow(1).value = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "ChkShow1", vbUnchecked)
        ChkShow_Click 1
    
        TxtComment(0).text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopComment", "")
        TxtComment(1).text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomComment", "")
    
        TxtHeight(0).text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopHeight", "2")
        TxtHeight(1).text = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomHeight", "2")
    
        CPic(0).color = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "Cpic0", vbWhite)
        CPic(1).color = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "Cpic1", vbBlack)
    
        CPic(2).color = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "Cpic2", vbWhite)
        CPic(3).color = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "Cpic3", vbBlack)
    
        TxtComment(0).Alignment = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopCommentAlign", vbCenter)

        If TxtComment(0).Alignment = vbRightJustify Then
            TBr(0).Buttons("Right").value = tbrPressed
        ElseIf TxtComment(0).Alignment = vbLeftJustify Then
            TBr(0).Buttons("Left").value = tbrPressed
        Else
            TBr(0).Buttons("Center").value = tbrPressed
        End If

        TxtComment(1).Alignment = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomCommentAlign", vbCenter)

        If TxtComment(1).Alignment = vbRightJustify Then
            TBr(1).Buttons("Right").value = tbrPressed
        ElseIf TxtComment(1).Alignment = vbLeftJustify Then
            TBr(1).Buttons("Left").value = tbrPressed
        Else
            TBr(1).Buttons("Center").value = tbrPressed
        End If
    
        TxtComment(0).Font.Bold = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopCommentBold", False)
        TBr(0).Buttons("Bold").value = IIf(TxtComment(0).Font.Bold = True, tbrPressed, tbrUnpressed)
        TxtComment(1).Font.Bold = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomCommentBold", False)
        TBr(1).Buttons("Bold").value = IIf(TxtComment(1).Font.Bold = True, tbrPressed, tbrUnpressed)
    
        TxtComment(0).Font.Italic = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopCommentItalic", False)
        TBr(0).Buttons("Italic").value = IIf(TxtComment(0).Font.Italic = True, tbrPressed, tbrUnpressed)
        TxtComment(1).Font.Italic = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomCommentItalic", False)
        TBr(1).Buttons("Italic").value = IIf(TxtComment(1).Font.Italic = True, tbrPressed, tbrUnpressed)
    
        TxtComment(0).Font.Underline = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopCommentUnder", False)
        TBr(0).Buttons("Under").value = IIf(TxtComment(0).Font.Underline = True, tbrPressed, tbrUnpressed)
        TxtComment(1).Font.Underline = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomCommentUnder", False)
        TBr(1).Buttons("Under").value = IIf(TxtComment(1).Font.Underline = True, tbrPressed, tbrUnpressed)
    
        TxtComment(0).Font.name = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopCommentFontName", TxtComment(0).Font.name)
        TxtComment(0).Font.Size = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "TopCommentFontSize", TxtComment(0).Font.Size)
        TxtComment(1).Font.name = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomCommentFontName", TxtComment(1).Font.name)
        TxtComment(1).Font.Size = GetSetting(SystemOptions.SysRegsAppPath, "BarcodeSetting", "BottomCommentFontSize", TxtComment(1).Font.Size)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Tbr_ButtonClick(Index As Integer, _
                            ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrTrap

    Select Case Button.key

        Case "Under"
            TxtComment(Index).Font.Underline = Button.value

        Case "Bold"
            TxtComment(Index).Font.Bold = Button.value

        Case "Italic"
            TxtComment(Index).Font.Italic = Button.value

        Case "Right"
            TxtComment(Index).Alignment = vbRightJustify

        Case "Left"
            TxtComment(Index).Alignment = vbLeftJustify

        Case "Center"
            TxtComment(Index).Alignment = vbCenter
    End Select

    Exit Sub
ErrTrap:
End Sub

