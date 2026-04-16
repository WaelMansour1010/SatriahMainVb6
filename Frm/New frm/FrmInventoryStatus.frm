VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInventoryStatus 
   Caption         =   "«·„Êﬁ› «·Õ«·Ì ··«’‰«›"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   15780
   Icon            =   "FrmInventoryStatus.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   15780
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
      Height          =   9525
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15780
      _cx             =   27834
      _cy             =   16801
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
      BackColor       =   4210752
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
      GridRows        =   4
      GridCols        =   3
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmInventoryStatus.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9525
         Index           =   4
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   15780
         _cx             =   27834
         _cy             =   16801
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
         Align           =   5
         AutoSizeChildren=   8
         BorderWidth     =   2
         ChildSpacing    =   2
         Splitter        =   -1  'True
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
         GridRows        =   2
         GridCols        =   2
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmInventoryStatus.frx":03F6
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   9525
            Index           =   3
            Left            =   0
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   15780
            _cx             =   27834
            _cy             =   16801
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
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "»Ì«‰«  «·Õ”«»"
            Align           =   5
            AutoSizeChildren=   8
            BorderWidth     =   2
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
            GridRows        =   12
            GridCols        =   10
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmInventoryStatus.frx":0444
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   2355
               Index           =   1
               Left            =   30
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   30
               Width           =   15705
               _cx             =   27702
               _cy             =   4154
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
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
               ForeColor       =   192
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
               Begin VB.CheckBox Check4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ﬂ· «·„Œ«“‰"
                  Height          =   255
                  Left            =   12720
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   120
                  Width           =   1695
               End
               Begin VB.Frame Frame9 
                  Caption         =   "Õœœ «·› —Â"
                  Height          =   975
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   240
                  Width           =   9495
                  Begin VB.OptionButton Option12 
                     Alignment       =   1  'Right Justify
                     Caption         =   " «—ÌŒ „Õœœ"
                     Height          =   195
                     Left            =   4560
                     RightToLeft     =   -1  'True
                     TabIndex        =   89
                     Top             =   600
                     Width           =   1215
                  End
                  Begin VB.ComboBox CmbMonth 
                     Height          =   315
                     Left            =   6120
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   87
                     Top             =   600
                     Width           =   945
                  End
                  Begin VB.OptionButton Option5 
                     Alignment       =   1  'Right Justify
                     Caption         =   "‘Â—ÌÂ"
                     Height          =   195
                     Left            =   7920
                     RightToLeft     =   -1  'True
                     TabIndex        =   86
                     Top             =   600
                     Width           =   1215
                  End
                  Begin VB.OptionButton Option6 
                     Alignment       =   1  'Right Justify
                     Caption         =   "—»⁄ À«·À"
                     Height          =   195
                     Left            =   2520
                     RightToLeft     =   -1  'True
                     TabIndex        =   85
                     Top             =   240
                     Width           =   1095
                  End
                  Begin VB.OptionButton Option7 
                     Alignment       =   1  'Right Justify
                     Caption         =   "—»⁄ À«‰Ì"
                     Height          =   195
                     Left            =   3720
                     RightToLeft     =   -1  'True
                     TabIndex        =   84
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.OptionButton Option8 
                     Alignment       =   1  'Right Justify
                     Caption         =   "—»⁄ «Ê·"
                     Height          =   195
                     Left            =   4800
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.OptionButton Option9 
                     Alignment       =   1  'Right Justify
                     Caption         =   "—»⁄ —«»⁄"
                     Height          =   195
                     Left            =   1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.OptionButton Option3 
                     Alignment       =   1  'Right Justify
                     Caption         =   "”‰ÊÌ…"
                     Height          =   195
                     Left            =   8160
                     RightToLeft     =   -1  'True
                     TabIndex        =   81
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   975
                  End
                  Begin VB.OptionButton Option14 
                     Alignment       =   1  'Right Justify
                     Caption         =   "‰’› «Ê·"
                     Height          =   195
                     Left            =   7200
                     RightToLeft     =   -1  'True
                     TabIndex        =   80
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.OptionButton Option11 
                     Alignment       =   1  'Right Justify
                     Caption         =   "‰’› À«‰Ì "
                     Height          =   195
                     Left            =   6120
                     RightToLeft     =   -1  'True
                     TabIndex        =   79
                     Top             =   240
                     Width           =   1095
                  End
                  Begin MSComCtl2.DTPicker DTPickerAccFrom 
                     BeginProperty DataFormat 
                        Type            =   1
                        Format          =   "dd/MM/yyyy"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   11265
                        SubFormatType   =   3
                     EndProperty
                     Height          =   345
                     Left            =   2640
                     TabIndex        =   90
                     ToolTipText     =   "„‰  «—ÌŒ ﬁœÌ„"
                     Top             =   600
                     Width           =   1500
                     _ExtentX        =   2646
                     _ExtentY        =   609
                     _Version        =   393216
                     CalendarBackColor=   -2147483624
                     CalendarTitleBackColor=   10383715
                     CheckBox        =   -1  'True
                     CustomFormat    =   "yyyy/M/d"
                     Format          =   68812803
                     CurrentDate     =   37357
                  End
                  Begin MSComCtl2.DTPicker DTPickerAccTo 
                     BeginProperty DataFormat 
                        Type            =   1
                        Format          =   "dd/MM/yyyy"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   11265
                        SubFormatType   =   3
                     EndProperty
                     Height          =   345
                     Left            =   600
                     TabIndex        =   91
                     ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
                     Top             =   600
                     Width           =   1500
                     _ExtentX        =   2646
                     _ExtentY        =   609
                     _Version        =   393216
                     CalendarBackColor=   -2147483624
                     CalendarTitleBackColor=   10383715
                     CheckBox        =   -1  'True
                     CustomFormat    =   "yyyy/M/d"
                     Format          =   68812803
                     CurrentDate     =   40858
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„‰"
                     Height          =   285
                     Index           =   10
                     Left            =   4260
                     RightToLeft     =   -1  'True
                     TabIndex        =   93
                     Top             =   600
                     Width           =   315
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "≈·Ï"
                     Height          =   285
                     Index           =   11
                     Left            =   2100
                     RightToLeft     =   -1  'True
                     TabIndex        =   92
                     Top             =   600
                     Width           =   315
                  End
                  Begin VB.Label Label8 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "Õœœ «·‘Â—"
                     Height          =   255
                     Left            =   6960
                     RightToLeft     =   -1  'True
                     TabIndex        =   88
                     Top             =   600
                     Width           =   975
                  End
               End
               Begin VB.ComboBox CboYear1 
                  Height          =   315
                  Left            =   9600
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   76
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   1065
               End
               Begin VB.ComboBox CboYear 
                  Height          =   315
                  Left            =   12120
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   74
                  Top             =   840
                  Width           =   1065
               End
               Begin VB.CheckBox ChKBlock 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìﬁ«› «· ⁄«„·"
                  ForeColor       =   &H000000C0&
                  Height          =   195
                  Left            =   5160
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   3420
                  Width           =   1215
               End
               Begin VB.Frame Frame6 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ»Ì⁄Â «·—’Ìœ"
                  ForeColor       =   &H000000C0&
                  Height          =   1335
                  Left            =   5040
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   3600
                  Width           =   5055
                  Begin VB.Frame Frame5 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " "
                     Height          =   375
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   47
                     Top             =   720
                     Width           =   3495
                     Begin VB.OptionButton Differenttype 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   " Õ–Ì— ›ﬁÿ"
                        Height          =   195
                        Index           =   1
                        Left            =   120
                        RightToLeft     =   -1  'True
                        TabIndex        =   49
                        Top             =   120
                        Value           =   -1  'True
                        Width           =   1125
                     End
                     Begin VB.OptionButton Differenttype 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "„‰⁄ „‰ « „«„ «·⁄„·Ì…"
                        Height          =   195
                        Index           =   0
                        Left            =   1440
                        RightToLeft     =   -1  'True
                        TabIndex        =   48
                        Top             =   120
                        Width           =   1725
                     End
                  End
                  Begin VB.Frame Frame3 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " "
                     Height          =   375
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   44
                     Top             =   240
                     Width           =   3495
                     Begin VB.OptionButton DepitOrCredit 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "„œÌ‰"
                        Height          =   195
                        Index           =   0
                        Left            =   1800
                        RightToLeft     =   -1  'True
                        TabIndex        =   46
                        Top             =   120
                        Width           =   1365
                     End
                     Begin VB.OptionButton DepitOrCredit 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "œ«∆‰"
                        Height          =   195
                        Index           =   1
                        Left            =   240
                        RightToLeft     =   -1  'True
                        TabIndex        =   45
                        Top             =   120
                        Value           =   -1  'True
                        Width           =   1005
                     End
                  End
                  Begin VB.Label Label4 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "›Ì Õ«·… „Œ«·›… ÿ»Ì⁄… «·Õ”«»"
                     ForeColor       =   &H000000C0&
                     Height          =   375
                     Left            =   3720
                     RightToLeft     =   -1  'True
                     TabIndex        =   50
                     Top             =   720
                     Width           =   1215
                  End
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "·Â „Ê«“‰‹‹‹Â"
                  ForeColor       =   &H000000C0&
                  Height          =   195
                  Left            =   6480
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   3420
                  Width           =   1695
               End
               Begin VB.Frame Frame4 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "’·«ÕÌ… «· ⁄«„·"
                  ForeColor       =   &H000000C0&
                  Height          =   1215
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   3960
                  Width           =   4935
                  Begin VB.OptionButton Authority 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„” Œœ„"
                     Height          =   195
                     Index           =   2
                     Left            =   3120
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   840
                     Value           =   -1  'True
                     Width           =   885
                  End
                  Begin VB.OptionButton Authority 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Ã„Ê⁄Â"
                     Height          =   195
                     Index           =   1
                     Left            =   3120
                     RightToLeft     =   -1  'True
                     TabIndex        =   37
                     Top             =   480
                     Width           =   885
                  End
                  Begin VB.OptionButton Authority 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ· «·„” Œœ„Ì‰"
                     Height          =   195
                     Index           =   0
                     Left            =   2160
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   240
                     Width           =   1845
                  End
                  Begin MSDataListLib.DataCombo DataCombo1 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   39
                     Top             =   480
                     Width           =   3135
                     _ExtentX        =   5530
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DataCombo2 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   40
                     Top             =   840
                     Width           =   3135
                     _ExtentX        =   5530
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„—ﬂ“ «· ﬂ·›…"
                  ForeColor       =   &H000000C0&
                  Height          =   1095
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   3960
                  Width           =   4935
                  Begin VB.CheckBox Check2 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·Â „—ﬂ“  ﬂ·›Â"
                     Height          =   255
                     Left            =   3480
                     RightToLeft     =   -1  'True
                     TabIndex        =   33
                     Top             =   240
                     Width           =   1215
                  End
                  Begin VB.Frame Frame2 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰Ê⁄ «·„—ﬂ“"
                     Enabled         =   0   'False
                     Height          =   495
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   28
                     Top             =   120
                     Width           =   3135
                     Begin VB.OptionButton Option2 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "€Ì— „Õœœ"
                        Height          =   195
                        Left            =   1200
                        RightToLeft     =   -1  'True
                        TabIndex        =   30
                        Top             =   240
                        Width           =   975
                     End
                     Begin VB.OptionButton Option1 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "„Õœœ"
                        Height          =   195
                        Left            =   120
                        RightToLeft     =   -1  'True
                        TabIndex        =   29
                        Top             =   240
                        Width           =   975
                     End
                  End
                  Begin MSDataListLib.DataCombo DcCostCenter 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   32
                     Top             =   720
                     Width           =   3015
                     _ExtentX        =   5318
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label3 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ „—ﬂ“ «· ﬂ·›Â"
                     Height          =   255
                     Left            =   3360
                     RightToLeft     =   -1  'True
                     TabIndex        =   31
                     Top             =   720
                     Width           =   1215
                  End
               End
               Begin VB.TextBox TxtAccount_NameE 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   5160
                  RightToLeft     =   -1  'True
                  TabIndex        =   23
                  Top             =   3720
                  Width           =   3330
               End
               Begin MSDataListLib.DataCombo DboParentAccount 
                  Height          =   315
                  Left            =   5115
                  TabIndex        =   19
                  Top             =   3900
                  Width           =   3375
                  _ExtentX        =   5953
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.TextBox TxtAccount_Code 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   5820
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   -210
                  Visible         =   0   'False
                  Width           =   795
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   1095
                  Index           =   5
                  Left            =   120
                  TabIndex        =   10
                  TabStop         =   0   'False
                  Top             =   3480
                  Width           =   4905
                  _cx             =   8652
                  _cy             =   1931
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
                  ForeColor       =   192
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   "‰Ê⁄ «·Õ”«»"
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
                  Begin VB.ComboBox Combo1 
                     Height          =   315
                     ItemData        =   "FrmInventoryStatus.frx":0567
                     Left            =   2520
                     List            =   "FrmInventoryStatus.frx":0574
                     RightToLeft     =   -1  'True
                     TabIndex        =   52
                     Top             =   675
                     Width           =   1215
                  End
                  Begin VB.ComboBox Combo2 
                     Height          =   315
                     ItemData        =   "FrmInventoryStatus.frx":0591
                     Left            =   0
                     List            =   "FrmInventoryStatus.frx":05A4
                     RightToLeft     =   -1  'True
                     TabIndex        =   51
                     Top             =   675
                     Width           =   1215
                  End
                  Begin VB.CheckBox Check3 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«»  Ã„Ì⁄Ì"
                     Height          =   210
                     Left            =   30
                     RightToLeft     =   -1  'True
                     TabIndex        =   26
                     Top             =   225
                     Width           =   1335
                  End
                  Begin VB.OptionButton OptAccountType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«» ‰Â«∆Ï"
                     Height          =   210
                     Index           =   0
                     Left            =   1710
                     RightToLeft     =   -1  'True
                     TabIndex        =   12
                     Top             =   225
                     Value           =   -1  'True
                     Width           =   1215
                  End
                  Begin VB.OptionButton OptAccountType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«» —∆Ì”ÌÏ"
                     Height          =   195
                     Index           =   1
                     Left            =   3270
                     RightToLeft     =   -1  'True
                     TabIndex        =   11
                     Top             =   225
                     Width           =   1335
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿ»Ì⁄Â «·Õ”«»"
                     ForeColor       =   &H00000000&
                     Height          =   300
                     Index           =   3
                     Left            =   3600
                     RightToLeft     =   -1  'True
                     TabIndex        =   54
                     Top             =   705
                     Width           =   1230
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " »ÊÌ» «·Õ”«»"
                     ForeColor       =   &H00000000&
                     Height          =   300
                     Index           =   9
                     Left            =   1320
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   705
                     Width           =   990
                  End
                  Begin VB.Image Img 
                     Height          =   240
                     Index           =   1
                     Left            =   4620
                     Picture         =   "FrmInventoryStatus.frx":05D0
                     Top             =   225
                     Width           =   240
                  End
                  Begin VB.Image Img 
                     Height          =   240
                     Index           =   0
                     Left            =   2940
                     Picture         =   "FrmInventoryStatus.frx":095A
                     Top             =   225
                     Width           =   240
                  End
               End
               Begin VB.TextBox TxtAccount_Name 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   5190
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   4020
                  Visible         =   0   'False
                  Width           =   3330
               End
               Begin VB.TextBox TxtAccount_Serial 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6510
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   -1170
                  Width           =   1995
               End
               Begin VB.TextBox TxtAccount_ID 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   5
                  Top             =   -360
                  Visible         =   0   'False
                  Width           =   915
               End
               Begin MSDataListLib.DataCombo DCCURRENCY 
                  Height          =   315
                  Left            =   5160
                  TabIndex        =   22
                  Top             =   3630
                  Width           =   765
                  _ExtentX        =   1349
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcBalanceSheet 
                  Height          =   315
                  Left            =   10800
                  TabIndex        =   56
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   3525
                  _ExtentX        =   6218
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboStoreName 
                  Height          =   315
                  Left            =   9960
                  TabIndex        =   95
                  Top             =   360
                  Width           =   3525
                  _ExtentX        =   6218
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "”‰Â „ﬁ«—‰Â"
                  Height          =   255
                  Left            =   10680
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”‰Â "
                  Height          =   255
                  Left            =   13320
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   840
                  Width           =   975
               End
               Begin VB.Label Label5 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   " ﬁ—Ì— ⁄‰ «·«’‰«› «·„ÊÃÊœÂ ›Ì «·›—Ê⁄ Ê «·„” Êœ⁄"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   24
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   615
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   1320
                  Width           =   10215
               End
               Begin VB.Label lbllevel 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   255
                  Left            =   8280
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   1800
                  Width           =   1575
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄„·…"
                  Height          =   195
                  Index           =   8
                  Left            =   5880
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   3600
                  Width           =   510
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Name En"
                  Height          =   285
                  Index           =   7
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   -1170
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «·Õ”«» «·—∆Ì”Ì   "
                  Height          =   435
                  Index           =   6
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   3660
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Õ”«»"
                  Height          =   285
                  Index           =   2
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   -630
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Œ — «·„Œ“‰"
                  Height          =   225
                  Index           =   0
                  Left            =   13080
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   480
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·Õ”«»"
                  Height          =   345
                  Index           =   1
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   4
                  Top             =   -120
                  Visible         =   0   'False
                  Width           =   930
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   780
               Index           =   7
               Left            =   12945
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   8715
               Visible         =   0   'False
               Width           =   2805
               _cx             =   4948
               _cy             =   1376
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
               Begin ImpulseButton.ISButton CmdN 
                  Height          =   255
                  Index           =   0
                  Left            =   30
                  TabIndex        =   14
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   450
                  ButtonStyle     =   1
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
                  ButtonImage     =   "FrmInventoryStatus.frx":0CE4
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton CmdN 
                  Height          =   255
                  Index           =   1
                  Left            =   450
                  TabIndex        =   15
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   450
                  ButtonStyle     =   1
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
                  ButtonImage     =   "FrmInventoryStatus.frx":107E
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton CmdN 
                  Height          =   285
                  Index           =   2
                  Left            =   1320
                  TabIndex        =   16
                  Top             =   30
                  Width           =   585
                  _ExtentX        =   1032
                  _ExtentY        =   503
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
                  ButtonImage     =   "FrmInventoryStatus.frx":1418
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FgAccounts 
               Height          =   6300
               Left            =   30
               TabIndex        =   34
               Top             =   2400
               Width           =   15705
               _cx             =   27702
               _cy             =   11112
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
               Cols            =   18
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmInventoryStatus.frx":17B2
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   930
                  Index           =   2
                  Left            =   0
                  TabIndex        =   69
                  TabStop         =   0   'False
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   15090
                  _cx             =   26617
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
                  Begin VB.TextBox TxtModflg 
                     Alignment       =   1  'Right Justify
                     Height          =   120
                     Left            =   45
                     RightToLeft     =   -1  'True
                     TabIndex        =   70
                     Top             =   15
                     Visible         =   0   'False
                     Width           =   1260
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
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
                     Index           =   5
                     Left            =   45
                     RightToLeft     =   -1  'True
                     TabIndex        =   72
                     Top             =   1695
                     Width           =   16350
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "„·«ÕŸ… Â«„…:-"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000080&
                     Height          =   195
                     Index           =   4
                     Left            =   12195
                     RightToLeft     =   -1  'True
                     TabIndex        =   71
                     Top             =   45
                     Width           =   3015
                  End
                  Begin VB.Image Image1 
                     Height          =   240
                     Left            =   15435
                     Picture         =   "FrmInventoryStatus.frx":1AE0
                     Top             =   45
                     Width           =   240
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   780
               Index           =   0
               Left            =   30
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   8715
               Width           =   15705
               _cx             =   27702
               _cy             =   1376
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   600
                  Index           =   0
                  Left            =   18720
                  TabIndex        =   60
                  Top             =   135
                  Width           =   1950
                  _ExtentX        =   3440
                  _ExtentY        =   1058
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
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   4210752
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  ColorToggledText=   -2147483631
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   4210752
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   600
                  Index           =   1
                  Left            =   16365
                  TabIndex        =   61
                  Top             =   135
                  Width           =   2040
                  _ExtentX        =   3598
                  _ExtentY        =   1058
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
                  Height          =   600
                  Index           =   2
                  Left            =   14040
                  TabIndex        =   62
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   1320
                  _ExtentX        =   2328
                  _ExtentY        =   1058
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
                  Height          =   600
                  Index           =   3
                  Left            =   12150
                  TabIndex        =   63
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   1740
                  _ExtentX        =   3069
                  _ExtentY        =   1058
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   600
                  Index           =   4
                  Left            =   9045
                  TabIndex        =   64
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   2985
                  _ExtentX        =   5265
                  _ExtentY        =   1058
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
                  Height          =   600
                  Index           =   5
                  Left            =   7125
                  TabIndex        =   65
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   1058
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
                  Height          =   600
                  Index           =   6
                  Left            =   300
                  TabIndex        =   66
                  TabStop         =   0   'False
                  Top             =   135
                  Width           =   2100
                  _ExtentX        =   3704
                  _ExtentY        =   1058
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   600
                  Index           =   7
                  Left            =   4755
                  TabIndex        =   67
                  Top             =   135
                  Width           =   2055
                  _ExtentX        =   3625
                  _ExtentY        =   1058
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "ÿ»«⁄…"
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
                  Height          =   600
                  Left            =   2475
                  TabIndex        =   68
                  Top             =   135
                  Width           =   2175
                  _ExtentX        =   3836
                  _ExtentY        =   1058
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
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Currency"
               Height          =   2355
               Left            =   5580
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   825
               Width           =   7350
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄„·… «·Õ”«»"
               Height          =   780
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   30
               Width           =   1680
            End
         End
      End
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   10305
      Index           =   6
      Left            =   20520
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   0
      Width           =   6675
      _cx             =   11774
      _cy             =   18177
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
      Begin MSComctlLib.ImageList ImgLstChartTree 
         Left            =   600
         Top             =   1980
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryStatus.frx":1E6A
               Key             =   "Expanded_Node"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryStatus.frx":2CBC
               Key             =   "Root"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryStatus.frx":3056
               Key             =   "Open_Node"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryStatus.frx":33F0
               Key             =   "Closed_Node"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryStatus.frx":378A
               Key             =   "Item"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TrvAccounts 
         Height          =   10215
         HelpContextID   =   380
         Left            =   3330
         TabIndex        =   58
         Top             =   45
         Visible         =   0   'False
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   18018
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   706
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImgLstChartTree"
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "FrmInventoryStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xAccCol As Collection
Dim IntCurrentIndex As Integer
Dim Dcombos As ClsDataCombos
Dim StrTemp As String

Private Sub Check2_Click()

    If check2.value = vbChecked Then
        Frame2.Enabled = True
    Else
        Frame2.Enabled = False
    End If

End Sub

Private Sub CmbMonth_Click()
    Dim StartDate As Integer
    Dim EndDate As Integer
    StartDate = CmbMonth.ListIndex + 1
 
    DTPickerAccFrom.value = StartDate & "-1-" & CboYear.text
    DTPickerAccTo.value = StartDate & "-" & get_last_month_day(DTPickerAccFrom.value) & "-" & CboYear.text
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim cReport As ClsAccReports

    Select Case Index

        Case 0
            Me.TxtModFlg.text = "N"
       
            Dccurrency.BoundText = 1
            tXTAccount_Serial.text = ""
            Me.DboParentAccount.BoundText = StrTemp
            TxtAccount_Name.text = ""
            TxtAccount_NameE.text = ""

        Case 1
            Me.TxtModFlg.text = "E"

        Case 2
            SaveData

        Case 3
            Me.TxtModFlg.text = "R"

        Case 4
            DelAccount

        Case 5
            Account_search.show
            Account_search.case_id = 0
        
        Case 6
            Unload Me

        Case 7
            PrintReport
    End Select

End Sub

Function PrintReport()
    'save data
    Dim RsDev As ADODB.Recordset
    Dim sql As String
    sql = "Delete  From  BalanceSheetReport"
    Cn.Execute sql
        
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "BalanceSheetReport", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    With Me.FgAccounts

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Account_Code")) <> "" Then
     
                RsDev.AddNew
           
                RsDev("Account_Code").value = .TextMatrix(i, .ColIndex("Account_Code"))
                RsDev("DebitValue").value = val(.TextMatrix(i, .ColIndex("DepitValue")))
                RsDev("CreditValue").value = val(.TextMatrix(i, .ColIndex("CreditValue")))
                RsDev("Totals").value = val(.TextMatrix(i, .ColIndex("Totals")))
                RsDev("Totals").value = val(.TextMatrix(i, .ColIndex("Totals")))
                RsDev("Balance").value = val(.TextMatrix(i, .ColIndex("Balance")))
              
                RsDev("ComparisonYearTotal").value = val(.TextMatrix(i, .ColIndex("ComparisonYearTotal")))
                RsDev("ComparisonYearDepit").value = val(.TextMatrix(i, .ColIndex("ComparisonYearDepit")))
                RsDev("ComparisonYearCredit").value = val(.TextMatrix(i, .ColIndex("ComparisonYearCredit")))
               
                RsDev.update
            End If
       
        Next i

    End With

    Dim cAccountReport As ClsAccReports
    Set cAccountReport = New ClsAccReports
    cAccountReport.ShowBalanceSheetNew
    Set cAccountReport = Nothing

End Function

Private Sub CmdN_Click(Index As Integer)
    Dim StrTemp As String
    Dim xx As Object

    Select Case Index

        Case 0

            'Back Move
            If IntCurrentIndex = 0 Then
                IntCurrentIndex = xAccCol.count
            Else
                IntCurrentIndex = IntCurrentIndex - 1
            End If

            If IntCurrentIndex = 0 Then
                'Me.CmdN(0).Enabled = False
            Else

                If IntCurrentIndex <= xAccCol.count Then
                    StrTemp = xAccCol.Item("A" & IntCurrentIndex)
                    Me.Retrive StrTemp, False
                Else
                    IntCurrentIndex = 1
                End If
            End If

        Case 1

            'Forward Move
            If IntCurrentIndex = 0 Then
                IntCurrentIndex = xAccCol.count
            Else
                IntCurrentIndex = IntCurrentIndex + 1
            End If

            If IntCurrentIndex = 0 Then
                'Me.CmdN(0).Enabled = False
            Else

                If IntCurrentIndex <= xAccCol.count Then
                    StrTemp = xAccCol.Item("A" & IntCurrentIndex)
                    Me.Retrive StrTemp, False
                Else
                    IntCurrentIndex = 1
                End If
            End If

        Case 2
            GetUpLevel
    End Select

End Sub

Private Sub DcBalanceSheet_Change()
    Dim BalanceSheetAccounts As String
    Dim msgrslt As Integer

    If val(DcBalanceSheet.BoundText) = 0 Then Exit Sub
    BalanceSheetAccounts = BalanceSheetAccount(val(DcBalanceSheet.BoundText), "Account_Code")

    'LoadData Val(DcBalanceSheet.BoundText)
    X = MsgBox("Â·  —Ìœ «ŸÂ«—Õ”«»«  ’›—ÌÂ ‰⁄„ «„ ·« ", vbInformation + vbYesNo)

    If X = vbYes Then
        FillGridWithData BalanceSheetAccounts, True
    Else
        FillGridWithData BalanceSheetAccounts, False
    End If

End Sub

Public Function ExpandAllNodes()
    On Error Resume Next
    Dim expAll As Integer

    For expAll = 1 To TrvAccounts.Nodes.count

        If TrvAccounts.Nodes(expAll).Children Then
            TrvAccounts.Nodes(expAll).Expanded = True
        End If

    Next

End Function

Private Sub FillGridWithData(AccountsCodes As String, _
                             Optional showZeroAccounts As Boolean)
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    StrSQL = " select * from QryBalanceSheetGeneral( '" & SQLDate(DTPickerAccFrom.value) & "','" & SQLDate(DTPickerAccTo.value) & "',1) "

    StrSQL = StrSQL & "where account_code in (" + AccountsCodes & ")"
 
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.FgAccounts
        .Clear flexClearScrollable, flexClearEverything

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = .FixedRows + rs.RecordCount
 
            For i = 2 To .Rows - 1

                '            .TextMatrix(I, .ColIndex("Account_ID")) = IIf(IsNull(rs("Account_ID").value), "", rs("Account_ID").value)
                If val(rs("Balance").value) = 0 And val(rs("DebitValue").value) = 0 And val(rs("CreditValue").value) = 0 Then

                    If showZeroAccounts = False Then
                        GoTo ll
                    End If
                End If
         
                '.TextMatrix(I, .ColIndex("Balance")) = IIf(IsNull(rs("Balance").value), "", rs("Balance").value)
                .TextMatrix(i, .ColIndex("opening_balance")) = IIf(IsNull(rs("opening_balance").value), "", rs("opening_balance").value)
                .TextMatrix(i, .ColIndex("Parent_account_code")) = IIf(IsNull(rs("Parent_account_code").value), "", rs("Parent_account_code").value)

                If .TextMatrix(i, .ColIndex("Parent_account_code")) = "r" Then
                    .TextMatrix(i, .ColIndex("Totals")) = IIf(IsNull(rs("Balance").value), "", rs("Balance").value)
                    .Cell(flexcpForeColor, i, 4, i, 4) = vbRed
                ElseIf .TextMatrix(i, .ColIndex("Parent_account_code")) <> "r" Then
           
                    If rs("last_account").value = 0 Then
                        .TextMatrix(i, .ColIndex("Balance")) = IIf(IsNull(rs("Balance").value), "", rs("Balance").value)
                    Else
                        .TextMatrix(i, .ColIndex("DepitValue")) = IIf(IsNull(rs("DebitValue").value), "", rs("DebitValue").value)
                        .TextMatrix(i, .ColIndex("CreditValue")) = IIf(IsNull(rs("CreditValue").value), "", rs("CreditValue").value)
                        .Cell(flexcpForeColor, i, 4, i, 4) = vbBlue
                    End If
                End If
           
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)

                If SystemOptions.UserInterface = EnglishInterface Then
                    .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
               
                Else
                    .TextMatrix(i, .ColIndex("Account_Name")) = Space$(Len(.TextMatrix(i, .ColIndex("Account_Serial"))) * 2) & IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
            
                End If
             
                If rs("last_account").value = True Then
                    .Cell(flexcpPicture, i, .ColIndex("Account_ID")) = Me.ImgLstChartTree.ListImages("Item").ExtractIcon
                    .Cell(flexcpFontBold, i, .ColIndex("Account_Name")) = False
                
                Else
                    .Cell(flexcpPicture, i, .ColIndex("Account_ID")) = Me.ImgLstChartTree.ListImages("Closed_Node").ExtractIcon
                    .Cell(flexcpFontBold, i, .ColIndex("Account_Name")) = True
                    .Cell(flexcpFontName, i, .ColIndex("Account_Name")) = "Tahoma"
                End If

ll:
                rs.MoveNext
            Next i

        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        End If

        '    .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub FgAccounts_Click()
    Dim LngMouseRow As Long
    Dim LngMouseCol As Long
    Dim XNode As MSComctlLib.Node
    Dim StrAcountCode  As String
    Dim account_name As String

    With Me.FgAccounts
        
        LngMouseCol = .MouseCol
        LngMouseRow = .MouseRow

        If LngMouseRow <= 0 Then Exit Sub
        If LngMouseCol <> .ColIndex("Account_Name") Then
            Exit Sub
        End If

        If LngMouseCol = .ColIndex("Account_Name") Then
            If .Cell(flexcpFontBold, LngMouseRow, LngMouseCol) = True Then
                StrAcountCode = .TextMatrix(LngMouseRow, .ColIndex("Account_Code"))
                account_name = .TextMatrix(LngMouseRow, .ColIndex("Account_Name"))
                '      Me.Retrive StrAcountCode
                '      Set XNode = Me.TrvAccounts.Nodes(StrAcountCode & "G")
                '      Me.TrvAccounts.Nodes(XNode.key).EnsureVisible
                '      Me.TrvAccounts.Nodes(XNode.key).Expanded = True
                '      Me.TrvAccounts.Nodes(XNode.key).Selected = True
                viewReport StrAcountCode, account_name
            End If

        Else
            Exit Sub
        End If

    End With

End Sub

Function viewReport(StrAcountCode As String, StrAccountName As String)
    Dim StrAccountSerial As String
    Dim cAccountReport As ClsAccReports
    Dim ClsAcc As ClsAccounts
    Set ClsAcc = New ClsAccounts
    StrAccountSerial = ClsAcc.Get_Account_Serial(StrAcountCode)
    Set cAccountReport = New ClsAccReports
    cAccountReport.BegineDate = Me.DTPickerAccFrom.value
    cAccountReport.EndDate = Me.DTPickerAccTo.value
    cAccountReport.ShowLedger2 StrAcountCode, StrAccountName, , , , StrAccountSerial
    Set cAccountReport = Nothing

End Function

Private Sub FgAccounts_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    Dim LngMouseRow As Long
    Dim LngMouseCol As Long

    With Me.FgAccounts
        .ToolTipText = ""
        .MousePointer = flexDefault
        ' .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbBlack
        .Cell(flexcpFontUnderline, 0, 0, .Rows - 1, .Cols - 1) = False
        
        LngMouseCol = .MouseCol
        LngMouseRow = .MouseRow

        If LngMouseRow <= 0 Then Exit Sub
        If LngMouseCol <> .ColIndex("Account_Name") Then
            .MousePointer = flexDefault
            Exit Sub
        End If

        If LngMouseCol = .ColIndex("Account_Name") Then
            If .Cell(flexcpFontBold, LngMouseRow, LngMouseCol) = True Then
                .MousePointer = flexHand
                .ToolTipText = "≈÷€ÿ Â‰« Õ Ï Ì„ﬂ‰ﬂ „‘«Âœ… ﬂ‘›    «·Õ”«»"
                '  .Cell(flexcpForeColor, LngMouseRow, LngMouseCol) = vbBlue
                .Cell(flexcpFontUnderline, LngMouseRow, LngMouseCol) = True
            End If

        Else
            .MousePointer = flexDefault
            '.Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbBlack
            .Cell(flexcpFontUnderline, 0, 0, .Rows - 1, .Cols - 1) = False
        End If

    End With

End Sub

Private Sub ChangeLang()
    ChKBlock.Caption = "Block"
    Frame6.Caption = "Balance Type"
    DepitOrCredit(0).Caption = "Depit"
    DepitOrCredit(1).Caption = "Credit"
    Label4.Caption = "In Different Case"
    Differenttype(0).Caption = "Acess Deny"
    Differenttype(1).Caption = "Alarm Only"
    Lbl(3).Caption = "Acc. Type"
    Lbl(9).Caption = "Acc. Class."
    Frame4.Caption = "Authority"
    Authority(0).Caption = "All Users"
    Authority(1).Caption = "Group"
    Authority(2).Caption = "User"

    Frame1.Caption = "Cost Center"
    Frame2.Caption = "C.C Type"
    Option1.Caption = "Fixed"
    Option2.Caption = "Not Fixed"
    Label3.Caption = "CC Name"
    Me.Caption = "Account Chart"
    Ele(1).Caption = "Account Data"
    Lbl(1).Caption = "Account#"
    Lbl(0).Caption = "Acc. Code"
    Lbl(6).Caption = "Parent Account"
    Lbl(2).Caption = "Name A"
    Lbl(7).Caption = "Name E"
    'lbl(3).Caption = "Derived Account From This Acc"
    Lbl(4).Caption = "Note"
    Lbl(5).Visible = False
    check1.Caption = " have A Budget"
    check2.Caption = "Cost Center"
    Check3.Caption = "Sum account"
    Lbl(8).Caption = "Curr"

    Ele(5).Caption = "Account Type"

    With FgAccounts
        .TextMatrix(0, .ColIndex("Account_ID")) = "Account ID"
        .TextMatrix(0, .ColIndex("Account_Serial")) = " Account Code"
        .TextMatrix(0, .ColIndex("Account_name")) = "Account Name"
        .TextMatrix(0, .ColIndex("OpenAccount")) = "Opening Balance"
        .TextMatrix(0, .ColIndex("AccountState")) = "Account State"
        .TextMatrix(0, .ColIndex("DateCreated")) = "DateCreated"
        .TextMatrix(0, .ColIndex("CurrentAccount")) = "Current Account"
    End With

    OptAccountType(0).Caption = "Final"
    OptAccountType(1).Caption = "Master"

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

End Sub

Private Sub Form_Load()
    Dim Msg As String
    Dim My_SQL As String
    My_SQL = "  select id,code from currency"

    fill_combo Dccurrency, My_SQL

    Dim GrdBack As ClsBackGroundPic

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
  
    Me.Lbl(5).Caption = Msg
    '----------------------------
    Set xAccCol = New Collection
    '----------------------------
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set Dcombos = New ClsDataCombos

    If SystemOptions.UserInterface = EnglishInterface Then
        Dcombos.GetAccountingCodesENg Me.DboParentAccount
    Else
        Dcombos.GetAccountingCodes Me.DboParentAccount
    End If
  
    Dcombos.GetBalanceheeet DcBalanceSheet
    Me.Height = 8605
    Me.Width = 16000
    Resize_Form Me
    Set GrdBack = New ClsBackGroundPic

    With Me.FgAccounts
        Set .WallPaper = GrdBack.Picture
        .GridLines = flexGridNone

        '   .AutoSize 0, .Cols - 1, False
        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        End If

    End With
 
    Dcombos.GetStores Me.DCboStoreName

    With Me.TrvAccounts
        .Appearance = ccFlat
        .Checkboxes = False
        .BorderStyle = ccNone
        .LineStyle = tvwRootLines
        .SingleSel = False
    End With

    'LoadData
    Me.TxtModFlg.text = "R"
    'Me.Retrive "r"
    'Me.TrvAccounts.Nodes("r").EnsureVisible
    'Me.TrvAccounts.Nodes("r").Expanded = True
    'Me.TrvAccounts.Nodes("r").Selected = True
    Dim StrSQL As String
    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
    fill_combo Me.DcCostCenter, StrSQL

    'If OPEN_NEW_SCREEN = True Then
    'Cmd_Click (0)
    'End If
    CboYear.Clear

    For i = 2010 To 2050
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = 1
    CboYear1.Clear

    For i = 2010 To 2050
        CboYear1.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = 1
 
    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

End Sub

Private Sub LoadData(BalanceSheetId As Integer)
    'ModTree.LoadTreeAccountBalanceSheetPrint Me.TrvAccounts, True, BalanceSheetId
    'ExpandAllNodes
End Sub

Public Sub Retrive(StrAccountCode As String, _
                   Optional BolPutInCol As Boolean = True)

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Static IntIndexPut As Integer

    If BolPutInCol = True Then
        IntIndexPut = IntIndexPut + 1
        xAccCol.add StrAccountCode, "A" & IntIndexPut
        IntCurrentIndex = IntIndexPut
    Else
        'IntCurrentIndex = IntIndexPut
    End If

    'IntIndexPut = IntIndexPut + 1
    'xAccCol.Add StrAccountCode, "A" & IntIndexPut
    'IntCurrentIndex = IntIndexPut
    
    StrSQL = "Select * From Accounts Where Account_Code='" & StrAccountCode & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        Me.TxtAccount_ID.text = IIf(IsNull(rs("Account_ID").value), "", rs("Account_ID").value)
        Me.TxtAccount_Code.text = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
        Me.tXTAccount_Serial.text = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
        Me.TxtAccount_Name.text = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
        Me.TxtAccount_NameE.text = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)

        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbllevel.Caption = "«·„” ÊÏ : " & CountAs(Me.TxtAccount_Code.text)
        Else
            Me.lbllevel.Caption = "Level:" & CountAs(Me.TxtAccount_Code.text)
        End If

        Me.check1.value = IIf(rs("mowazna").value = True, vbChecked, Unchecked)
        Me.check2.value = IIf(rs("cost_center").value = True, vbChecked, Unchecked)
     
        If IsNull(rs("cost_center_type").value) Then
            Option2.value = True
        Else

            If rs("cost_center_type").value = 0 Then
                Option2.value = True
            ElseIf rs("cost_center_type").value = 1 Then
                Option1.value = True
            End If

            Me.DcCostCenter.BoundText = IIf(IsNull(rs("cost_center_id").value), "", rs("cost_center_id").value)
        End If
     
        Me.Check3.value = IIf(rs("Sum_account").value = True, vbChecked, Unchecked)
     
        Me.Dccurrency.BoundText = IIf(IsNull(rs("currenct_code").value), "", rs("currenct_code").value)
    
        If rs("last_account").value = True Then
            Me.OptAccountType(0).value = True
            Me.OptAccountType(1).value = False
        Else
            Me.OptAccountType(0).value = False
            Me.OptAccountType(1).value = True
        End If

        Me.DboParentAccount.BoundText = IIf(IsNull(rs("Parent_Account_Code").value), "", rs("Parent_Account_Code").value)
    
    End If

    StrSQL = "Select * From Accounts Where Parent_Account_Code='" & StrAccountCode & "'"
    StrSQL = StrSQL + " Order By Accounts.last_account, Account_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.FgAccounts
        .Clear flexClearScrollable, flexClearEverything

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = .FixedRows + rs.RecordCount

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Account_ID")) = IIf(IsNull(rs("Account_ID").value), "", rs("Account_ID").value)
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)

                If SystemOptions.UserInterface = EnglishInterface Then
                    .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                Else
                    .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                End If
            
                If Not IsNull(rs("DateCreated").value) Then
                    .TextMatrix(i, .ColIndex("DateCreated")) = DisplayDate(rs("DateCreated").value)
                End If

                '.TextMatrix(i, .ColIndex("")) = IIf(IsNull(Rs("").Value), "", Rs("").Value)
                '.TextMatrix(i, .ColIndex("")) = IIf(IsNull(Rs("").Value), "", Rs("").Value)
                If rs("last_account").value = True Then
                    .Cell(flexcpPicture, i, .ColIndex("Account_ID")) = Me.ImgLstChartTree.ListImages("Item").ExtractIcon
                    .Cell(flexcpFontBold, i, .ColIndex("Account_Name")) = False
                
                Else
                    .Cell(flexcpPicture, i, .ColIndex("Account_ID")) = Me.ImgLstChartTree.ListImages("Closed_Node").ExtractIcon
                    .Cell(flexcpFontBold, i, .ColIndex("Account_Name")) = True
                    .Cell(flexcpFontName, i, .ColIndex("Account_Name")) = "Tahoma"
                End If

                rs.MoveNext
            Next i

        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        End If

        .AutoSize 0, .Cols - 1, False
    End With

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
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                
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

Private Sub Form_Unload(Cancel As Integer)

    Do While xAccCol.count > 0
        xAccCol.Remove xAccCol.count
    Loop

    Set xAccCol = Nothing
    Set Dcombos = Nothing
End Sub

Private Sub Option1_Click()

    If Option1.value = True Then
        DcCostCenter.Enabled = True
    Else
        DcCostCenter.Enabled = False
    End If

End Sub

Private Sub Option11_Click()
    DTPickerAccFrom.value = "1-7-" & CboYear.text
    DTPickerAccTo.value = "31-12-" & CboYear.text
End Sub

Private Sub Option14_Click()
    DTPickerAccFrom.value = "1-1-" & CboYear.text
    DTPickerAccTo.value = "30-6-" & CboYear.text
End Sub

Private Sub Option2_Click()

    If Option1.value = True Then
        DcCostCenter.Enabled = True
    Else
        DcCostCenter.Enabled = False
    End If

End Sub

Private Sub Option3_Click()
    DTPickerAccFrom.value = "1-1-" & CboYear.text
    DTPickerAccTo.value = "31-12-" & CboYear.text
End Sub

Private Sub Option6_Click()
    DTPickerAccFrom.value = "7-1-" & CboYear.text
    DTPickerAccTo.value = "30-9-" & CboYear.text
End Sub

Private Sub Option7_Click()
    DTPickerAccFrom.value = "4-1-" & CboYear.text
    DTPickerAccTo.value = "30-6-" & CboYear.text
End Sub

Private Sub Option8_Click()
    DTPickerAccFrom.value = "1-1-" & CboYear.text
    DTPickerAccTo.value = "31-3-" & CboYear.text
End Sub

Private Sub Option9_Click()
    DTPickerAccFrom.value = "10-1-" & CboYear.text
    DTPickerAccTo.value = "31-12-" & CboYear.text
End Sub

Private Sub TrvAccounts_NodeClick(ByVal Node As MSComctlLib.Node)

    If Not Node Is Nothing Then
        If InStr(1, Node.key, "G", vbTextCompare) <> 0 Then
            StrTemp = Node.key
            StrTemp = Mid(StrTemp, 1, Len(StrTemp) - 1)
        Else
            StrTemp = Node.key
        End If

        If Me.TxtModFlg.text = "R" Then
            Me.Retrive StrTemp
        Else
            Me.DboParentAccount.BoundText = StrTemp
        End If
    End If

End Sub

Private Sub TxtAccount_Name_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub TxtAccount_NameE_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub TxtModFlg_Change()

    Select Case TxtModFlg.text

        Case "N"
            Me.TxtAccount_ID.Enabled = False
            Me.tXTAccount_Serial.Enabled = True
            Me.TxtAccount_Name.Enabled = True
            Cmd(0).Enabled = False
            Cmd(1).Enabled = False
            Cmd(2).Enabled = True
            Cmd(3).Enabled = True
            Cmd(4).Enabled = False
            Cmd(5).Enabled = False
            Cmd(7).Enabled = False
        
        Case "E"
            Me.TxtAccount_ID.Enabled = False
            Me.tXTAccount_Serial.Enabled = True
            Me.TxtAccount_Name.Enabled = True
            Cmd(0).Enabled = False
            Cmd(1).Enabled = False
            Cmd(2).Enabled = True
            Cmd(3).Enabled = True
            Cmd(4).Enabled = False
            Cmd(5).Enabled = False
            Cmd(7).Enabled = False

        Case "R"
            Me.TxtAccount_ID.Enabled = False
            Me.tXTAccount_Serial.Enabled = False
            Me.TxtAccount_Name.Enabled = False
        
            Cmd(0).Enabled = True
            Cmd(1).Enabled = True
            Cmd(2).Enabled = False
            Cmd(3).Enabled = False
            Cmd(4).Enabled = True
            Cmd(5).Enabled = True
            Cmd(7).Enabled = True
        
    End Select

End Sub

Private Sub MoveInCollection(IntDir As Integer)
    Dim StrTemp  As String

    If IntDir = 0 Then
        ' Õ—ﬂ ··√„«„
    
    ElseIf IntDir = 1 Then

        ' Õ—ﬂ ··Œ·›
        If IntCurrentIndex = 0 Then
            StrTemp = xAccCol.Item(1)
            IntCurrentIndex = 1
        Else
            StrTemp = xAccCol.Item(IntCurrentIndex - 1)
        End If

        Me.Retrive StrTemp, False
    End If

    'Set Buttons

End Sub

Private Sub GetUpLevel()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim StrTemp As String

    If val(Me.TxtAccount_ID.text) <> 0 Then
        StrSQL = "Select * From Accounts Where Account_ID=" & val(Me.TxtAccount_ID.text)
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            StrTemp = IIf(IsNull(rs("Parent_Account_Code").value), "", rs("Parent_Account_Code").value)

            If Trim$(StrTemp) <> "" Then
                Me.Retrive StrTemp, True
            End If
        End If
    End If

End Sub

Private Sub DelAccount()

    Dim Msg As String
    Dim RsAcccounts As ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrNodeKey As String
    Dim XNode As MSComctlLib.Node
    Dim IntRes As Integer

    'If Not DoPremis(Do_Delete, "Frm_General_Journal", "«·œ·Ì· «·„Õ«”»Ï") Then Exit Sub
    On Error GoTo ErrTrap
    StrAccountCode = Trim(Me.TxtAccount_Code.text)

    'If left(StrAccountCode, 6) = "a1a2a3" Or left(StrAccountCode, 6) = "a1a2a1" Or left(StrAccountCode, 6) = "a2a3a1" Or left(StrAccountCode, 6) = "a3a1a4" Then
    '    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·Õ”«»"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    Exit Sub
    'End If
    If CheckDelAccount1(StrAccountCode) = False Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·Õ”«» ·«‰Â Õ”«» —∆Ì”Ì ·œÌ… «»‰«¡"
        Else
            Msg = "Can't Delete this Account because it have Child Account"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If CheckDelAccount(StrAccountCode) = True Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "This Acccount will be Deleted " ' & Me.TxtAccount_ID.text
            Msg = Msg & Chr(13) & "Account Code :- " & Me.tXTAccount_Serial.text
            Msg = Msg & Chr(13) & "Account Name " & Me.TxtAccount_Name.text
            Msg = Msg & Chr(13) & ""
            Msg = Msg & Chr(13) & "Sure you want delte ??"
          
        Else
          
            Msg = "”Ê› Ì „ Õ–› «·Õ”«» —ﬁ„:- " '& Me.TxtAccount_ID.text
            Msg = Msg & Chr(13) & " ﬂÊœ «Ê „”·”· «·Õ”«»:- " & Me.tXTAccount_Serial.text
            Msg = Msg & Chr(13) & "«”„ «·Õ”«» :- " & Me.TxtAccount_Name.text
            Msg = Msg & Chr(13) & ""
            Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ ⁄„·Ì… «·Õ–› ..øø"
        
        End If

        IntRes = MsgBox(Msg, vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo, App.title)

        If IntRes = vbNo Then
            Exit Sub
        End If

        Set RsAcccounts = New ADODB.Recordset
        StrSQL = "select *  From  ACCOUNTS Where Account_Code = '" & StrAccountCode & "'"
        RsAcccounts.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If RsAcccounts("last_account").value = True Then
            StrNodeKey = StrAccountCode & ""
        Else
            StrNodeKey = StrAccountCode & "G"
        End If

        RsAcccounts.delete
        RsAcccounts.Close

        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Account was Deleted......."
        Else
            Msg = " „  ⁄„·Ì… «·Õ–› ...!"
        End If

        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Set XNode = Me.TrvAccounts.Nodes(StrNodeKey)
        Me.TrvAccounts.Nodes.Remove XNode.key
        Set RsAcccounts = Nothing
        'LoadData
    Else

        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Can't Delete this Account"
        Else
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·Õ”«» ...!!"
        End If
 
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim rs As ADODB.Recordset
    Dim StrSQL  As String
    Dim Msg As String
    Dim BolLastAccount As Boolean
    Dim StrNewAccountCode As String

    If Me.TxtAccount_Name.text = "" And Me.TxtAccount_NameE.text <> "" Then
        Me.TxtAccount_Name.text = Me.TxtAccount_NameE.text
    End If

    If Me.TxtAccount_Name.text <> "" And Me.TxtAccount_NameE.text = "" Then
        Me.TxtAccount_NameE.text = Me.TxtAccount_Name.text
    End If

    If Trim$(Me.TxtAccount_Name.text) = "" Then
        Msg = "ÌÃ» ﬂ «»… «”„ «·Õ”«»"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If SystemOptions.UserInterface = EnglishInterface Then

        If Trim$(Me.TxtAccount_NameE.text) = "" Then
            Msg = "Must Enter Account Name"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

    End If

    If Trim$(Me.DboParentAccount.BoundText) = "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Must Specify Parent Account"
        Else
            Msg = "ÌÃ»  ÕœÌœ «”„ «·Õ”«» «·—∆Ì”Ì «·–Ì ”Ê› Ì ›—⁄ „‰Â Â–« «·Õ”«»"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If Me.OptAccountType(0).value = True Then
        BolLastAccount = True
    Else
        BolLastAccount = False
    End If

    Dim cost_center_type As Integer
    Dim cost_center_id As String

    If Option1.value = True Then
        If DcCostCenter.BoundText = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Must Specify COST CENTER"
                       
            Else
                Msg = "ÌÃ»  ÕœÌœ «”„ „—ﬂ“ «· ﬂ·›…"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcCostCenter.SetFocus
            SendKeys "{F4}"
    
            Exit Sub
        Else
            cost_center_id = DcCostCenter.BoundText
        End If

    End If

    If Option2.value = True Then
        cost_center_type = 0
    Else
        cost_center_type = 1
    End If

    If Me.TxtModFlg.text = "N" Then
        Dim mowazna As Boolean
        Dim cost_center As Boolean
        Dim Sum_account As Boolean

        If check1.value = vbChecked Then
            mowazna = 1
        Else
            mowazna = 0
        End If

        If check2.value = vbChecked Then
            cost_center = 1
        Else
            cost_center = 0
        End If

        If Check3.value = vbChecked Then
            Sum_account = 1
        Else
            Sum_account = 0
        End If

        StrNewAccountCode = ModAccounts.AddNewAccount(Me.DboParentAccount.BoundText, Trim$(Me.TxtAccount_Name.text), BolLastAccount, False, Trim$(Me.TxtAccount_NameE.text), Dccurrency.BoundText, mowazna, cost_center, Sum_account, , tXTAccount_Serial)
    
        If StrNewAccountCode <> "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Saved"
            Else
                Msg = " „  ⁄„·Ì… «·Œ›Ÿ."
            End If
    
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.TxtModFlg.text = "R"
        End If

    ElseIf Me.TxtModFlg.text = "E" Then
        StrSQL = "Select * From Accounts Where Account_ID=" & val(Me.TxtAccount_ID.text) & ""
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

        If Not (rs.BOF Or rs.EOF) Then

            '·Ê «‰ «·Õ”«» Õ”«» —∆Ì”Ì ÊÌ—Ìœ «‰ ÌÃ⁄·Â Õ”«» ‰Â«∆
            If rs("last_account").value = False And OptAccountType(0).value = True Then
                If GetAccountChilds(Me.TxtAccount_Code.text) <> 0 Then
                    If SystemOptions.UserInterface = EnglishInterface Then
                        Msg = "This Account is Master Account  And have child "
                        Msg = Msg & Chr(13) & "Can't change to last Account"
            
                    Else
         
                        Msg = "Â–« «·Õ”«» Õ”«» —∆Ì”Ì ÊÌÕ ÊÏ ⁄·Ï Õ”«»«  „ ›—⁄… „‰Â"
                        Msg = Msg & Chr(13) & "Ê·« Ì„ﬂ‰  ⁄œÌ·Â ≈·Ï Õ”«» ‰Â«∆"
                    End If
            
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If

            ElseIf rs("last_account").value = True And Me.OptAccountType(1).value = True Then
                '«·Õ”«» «·„—«œ  ⁄œÌ·Â Õ”«» ‰Â«∆ ÊÌ—«œ  ÕÊÌ·Â ≈·Ï Õ”«» —∆Ì”Ï
            
            Else
                Dim currency_code As Integer

                If Not IsNull(TxtAccount_Code.text) Then
                    If Me.Dccurrency.BoundText <> "" Then
                        currency_code = Me.Dccurrency.BoundText
                    Else
                        currency_code = 1
                    End If

                    ModAccounts.EditAccount rs("Account_Code").value, Me.TxtAccount_Name.text, Me.TxtAccount_NameE.text, check1.value, check2.value, currency_code, Check3.value, tXTAccount_Serial

                    If SystemOptions.UserInterface = EnglishInterface Then
                        Msg = "Saved"
                    Else
                  
                        Msg = " „  ⁄„·Ì… «·Œ›Ÿ."
                    End If
                  
                    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Me.TxtModFlg.text = "R"
                End If
            End If
        End If
    End If

    Dcombos.GetAccountingCodes Me.DboParentAccount
    'LoadData
    Me.Retrive StrNewAccountCode
End Sub

Public Function GetAccountChilds(StrAccountCode As String) As Long
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    StrSQL = "Select * From Accounts Where Account_Code like '" & StrAccountCode & "%'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        GetAccountChilds = 0
    ElseIf rs.RecordCount = 0 Then
        GetAccountChilds = 0
    Else
        GetAccountChilds = rs.RecordCount
    End If

End Function
