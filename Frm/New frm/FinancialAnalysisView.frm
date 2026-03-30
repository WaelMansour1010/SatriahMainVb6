VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FinancialAnalysisView 
   Caption         =   "⁄—÷ «·„⁄«œ·«  «·„«·ÌÂ"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   14595
   Icon            =   "FinancialAnalysisView.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   14595
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9525
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14595
      _cx             =   25744
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
      _GridInfo       =   $"FinancialAnalysisView.frx":038A
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
         Width           =   14595
         _cx             =   25744
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
         _GridInfo       =   $"FinancialAnalysisView.frx":03F5
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
            Width           =   14595
            _cx             =   25744
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
            _GridInfo       =   $"FinancialAnalysisView.frx":0442
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.TextBox txtDown 
               Alignment       =   1  'Right Justify
               CausesValidation=   0   'False
               Height          =   2355
               Left            =   9870
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   6345
               Width           =   4680
            End
            Begin VB.TextBox txtUp 
               Alignment       =   1  'Right Justify
               CausesValidation=   0   'False
               Height          =   2355
               Left            =   9870
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   3195
               Width           =   4695
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   2355
               Index           =   1
               Left            =   30
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   30
               Width           =   14520
               _cx             =   25612
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
               Caption         =   "«Œ — ‘ﬂ· «·„⁄«œ·Â"
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
               Begin VB.TextBox TxtModflg 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Text            =   "Text1"
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.Frame Frame9 
                  Caption         =   "Õœœ «·› —Â"
                  Height          =   975
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   240
                  Width           =   9495
                  Begin VB.OptionButton Option12 
                     Alignment       =   1  'Right Justify
                     Caption         =   " «—ÌŒ „Õœœ"
                     Height          =   195
                     Left            =   4560
                     RightToLeft     =   -1  'True
                     TabIndex        =   84
                     Top             =   600
                     Width           =   1215
                  End
                  Begin VB.ComboBox CmbMonth 
                     Height          =   315
                     Left            =   6120
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   82
                     Top             =   600
                     Width           =   945
                  End
                  Begin VB.OptionButton Option5 
                     Alignment       =   1  'Right Justify
                     Caption         =   "‘Â—ÌÂ"
                     Height          =   195
                     Left            =   7920
                     RightToLeft     =   -1  'True
                     TabIndex        =   81
                     Top             =   600
                     Width           =   1215
                  End
                  Begin VB.OptionButton Option6 
                     Alignment       =   1  'Right Justify
                     Caption         =   "—»⁄ À«·À"
                     Height          =   195
                     Left            =   2520
                     RightToLeft     =   -1  'True
                     TabIndex        =   80
                     Top             =   240
                     Width           =   1095
                  End
                  Begin VB.OptionButton Option7 
                     Alignment       =   1  'Right Justify
                     Caption         =   "—»⁄ À«‰Ì"
                     Height          =   195
                     Left            =   3720
                     RightToLeft     =   -1  'True
                     TabIndex        =   79
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.OptionButton Option8 
                     Alignment       =   1  'Right Justify
                     Caption         =   "—»⁄ «Ê·"
                     Height          =   195
                     Left            =   4800
                     RightToLeft     =   -1  'True
                     TabIndex        =   78
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.OptionButton Option9 
                     Alignment       =   1  'Right Justify
                     Caption         =   "—»⁄ —«»⁄"
                     Height          =   195
                     Left            =   1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   77
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.OptionButton Option3 
                     Alignment       =   1  'Right Justify
                     Caption         =   "”‰ÊÌ…"
                     Height          =   195
                     Left            =   8160
                     RightToLeft     =   -1  'True
                     TabIndex        =   76
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
                     TabIndex        =   75
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.OptionButton Option11 
                     Alignment       =   1  'Right Justify
                     Caption         =   "‰’› À«‰Ì "
                     Height          =   195
                     Left            =   6120
                     RightToLeft     =   -1  'True
                     TabIndex        =   74
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
                     TabIndex        =   85
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
                     Format          =   52756483
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
                     TabIndex        =   86
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
                     Format          =   52756483
                     CurrentDate     =   40858
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„‰"
                     Height          =   285
                     Index           =   10
                     Left            =   4140
                     RightToLeft     =   -1  'True
                     TabIndex        =   88
                     Top             =   600
                     Width           =   435
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "≈·Ï"
                     Height          =   285
                     Index           =   11
                     Left            =   2100
                     RightToLeft     =   -1  'True
                     TabIndex        =   87
                     Top             =   600
                     Width           =   435
                  End
                  Begin VB.Label Label8 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "Õœœ «·‘Â—"
                     Height          =   255
                     Left            =   7080
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   600
                     Width           =   855
                  End
               End
               Begin VB.ComboBox CboYear1 
                  Height          =   315
                  Left            =   9600
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   71
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   1065
               End
               Begin VB.ComboBox CboYear 
                  Height          =   315
                  Left            =   12120
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   69
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
                  TabIndex        =   54
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
                  TabIndex        =   42
                  Top             =   3600
                  Width           =   5055
                  Begin VB.Frame Frame5 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " "
                     Height          =   375
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   46
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
                        TabIndex        =   48
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
                        TabIndex        =   47
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
                     TabIndex        =   43
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
                        TabIndex        =   45
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
                        TabIndex        =   44
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
                     TabIndex        =   49
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
                  TabIndex        =   40
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
                  TabIndex        =   34
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
                     TabIndex        =   37
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
                     TabIndex        =   36
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
                     TabIndex        =   35
                     Top             =   240
                     Width           =   1845
                  End
                  Begin MSDataListLib.DataCombo DataCombo1 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   38
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
                     TabIndex        =   39
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
                     ItemData        =   "FinancialAnalysisView.frx":0563
                     Left            =   2520
                     List            =   "FinancialAnalysisView.frx":0570
                     RightToLeft     =   -1  'True
                     TabIndex        =   51
                     Top             =   675
                     Width           =   1215
                  End
                  Begin VB.ComboBox Combo2 
                     Height          =   315
                     ItemData        =   "FinancialAnalysisView.frx":058D
                     Left            =   0
                     List            =   "FinancialAnalysisView.frx":05A0
                     RightToLeft     =   -1  'True
                     TabIndex        =   50
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
                     TabIndex        =   53
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
                     TabIndex        =   52
                     Top             =   705
                     Width           =   990
                  End
                  Begin VB.Image Img 
                     Height          =   240
                     Index           =   1
                     Left            =   4620
                     Picture         =   "FinancialAnalysisView.frx":05CC
                     Top             =   225
                     Width           =   240
                  End
                  Begin VB.Image Img 
                     Height          =   240
                     Index           =   0
                     Left            =   2940
                     Picture         =   "FinancialAnalysisView.frx":0956
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
                  Left            =   9600
                  TabIndex        =   55
                  Top             =   360
                  Width           =   3525
                  _ExtentX        =   6218
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label LblDoenValue 
                  Alignment       =   1  'Right Justify
                  Caption         =   "0"
                  Height          =   375
                  Left            =   11640
                  RightToLeft     =   -1  'True
                  TabIndex        =   112
                  Top             =   1440
                  Width           =   1695
               End
               Begin VB.Label Lblupvalue 
                  Alignment       =   1  'Right Justify
                  Caption         =   "0"
                  Height          =   375
                  Left            =   9840
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   1560
                  Width           =   1695
               End
               Begin VB.Label Label9 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Label9"
                  Height          =   15
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   2400
                  Width           =   13095
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "”‰Â „ﬁ«—‰Â"
                  Height          =   255
                  Left            =   10680
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   1215
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”‰Â «·„«·ÌÂ"
                  Height          =   255
                  Left            =   13320
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   840
                  Width           =   975
               End
               Begin VB.Label Label5 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄—÷ ‰ «∆Ã  «·„⁄«œ·«  «·„«·ÌÂ"
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
                  Left            =   2520
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   1440
                  Width           =   7455
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
                  TabIndex        =   41
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
                  Caption         =   "«Œ — «·„⁄«œ·…"
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
               Height          =   1560
               Index           =   7
               Left            =   30
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   7140
               Width           =   9810
               _cx             =   17304
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
               Begin VB.TextBox TxtFinal 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   960
                  Width           =   3120
               End
               Begin VB.TextBox TxtGeneral 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   960
                  Width           =   1920
               End
               Begin VB.TextBox TxtGeneralValue 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   120
                  Width           =   2040
               End
               Begin VB.TextBox TxtFinancialEquationsOpr 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   840
               End
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
                  ButtonImage     =   "FinancialAnalysisView.frx":0CE0
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
                  ButtonImage     =   "FinancialAnalysisView.frx":107A
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
                  ButtonImage     =   "FinancialAnalysisView.frx":1414
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   510
                  Index           =   8
                  Left            =   6360
                  TabIndex        =   99
                  TabStop         =   0   'False
                  Top             =   600
                  Width           =   510
                  _cx             =   900
                  _cy             =   900
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
                     Index           =   3
                     Left            =   30
                     TabIndex        =   100
                     Top             =   1350
                     Visible         =   0   'False
                     Width           =   525
                     _ExtentX        =   926
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
                     ButtonImage     =   "FinancialAnalysisView.frx":17AE
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton CmdN 
                     Height          =   255
                     Index           =   4
                     Left            =   -1470
                     TabIndex        =   101
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
                     ButtonImage     =   "FinancialAnalysisView.frx":1B48
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton CmdN 
                     Height          =   285
                     Index           =   5
                     Left            =   1320
                     TabIndex        =   102
                     Top             =   30
                     Visible         =   0   'False
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
                     ButtonImage     =   "FinancialAnalysisView.frx":1EE2
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VB.Image Image7 
                     Height          =   360
                     Left            =   120
                     Picture         =   "FinancialAnalysisView.frx":227C
                     Stretch         =   -1  'True
                     Top             =   0
                     Width           =   450
                  End
               End
               Begin VB.Image Image3 
                  Height          =   720
                  Left            =   2280
                  Picture         =   "FinancialAnalysisView.frx":285A
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   705
               End
               Begin VB.Image Image4 
                  Height          =   720
                  Left            =   1440
                  Picture         =   "FinancialAnalysisView.frx":2F47
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   810
               End
               Begin VB.Image Image5 
                  Height          =   720
                  Left            =   720
                  Picture         =   "FinancialAnalysisView.frx":3525
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   810
               End
               Begin VB.Image Image6 
                  Height          =   720
                  Left            =   0
                  Picture         =   "FinancialAnalysisView.frx":3DE5
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ﬁÌ„Â «·‰Â«∆Ì…"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   375
                  Index           =   12
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   960
                  Width           =   1440
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ﬁÌ„Â"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   375
                  Index           =   5
                  Left            =   7320
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   1080
                  Width           =   2400
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄„·ÌÂ"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   375
                  Index           =   4
                  Left            =   7320
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   600
                  Width           =   2400
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Õ«’· ﬁ”„Â «·»”ÿ / «·„ﬁ«„"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   375
                  Index           =   100
                  Left            =   7320
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   120
                  Width           =   2400
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   780
               Index           =   0
               Left            =   30
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   8715
               Width           =   14520
               _cx             =   25612
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
                  Left            =   17325
                  TabIndex        =   59
                  Top             =   135
                  Width           =   1785
                  _ExtentX        =   3149
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
                  Left            =   15120
                  TabIndex        =   60
                  Top             =   135
                  Width           =   1890
                  _ExtentX        =   3334
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
                  Index           =   20
                  Left            =   6630
                  TabIndex        =   61
                  Top             =   135
                  Width           =   1200
                  _ExtentX        =   2117
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
                  AcclimateGrayTones=   -1  'True
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   4210752
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   600
                  Index           =   3
                  Left            =   11235
                  TabIndex        =   62
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   1605
                  _ExtentX        =   2831
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
                  Left            =   8370
                  TabIndex        =   63
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   2745
                  _ExtentX        =   4842
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
                  Left            =   6600
                  TabIndex        =   64
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   1635
                  _ExtentX        =   2884
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
                  Left            =   285
                  TabIndex        =   65
                  TabStop         =   0   'False
                  Top             =   135
                  Width           =   1935
                  _ExtentX        =   3413
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
                  Left            =   4410
                  TabIndex        =   66
                  Top             =   135
                  Width           =   1890
                  _ExtentX        =   3334
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
                  Left            =   2280
                  TabIndex        =   67
                  Top             =   135
                  Width           =   2025
                  _ExtentX        =   3572
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
            Begin VSFlex8UCtl.VSFlexGrid FgAccounts 
               Height          =   2355
               Left            =   30
               TabIndex        =   91
               Top             =   2400
               Width           =   9810
               _cx             =   17304
               _cy             =   4154
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
               Rows            =   3
               Cols            =   11
               FixedRows       =   2
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FinancialAnalysisView.frx":4551
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
            Begin VSFlex8UCtl.VSFlexGrid FgAccounts1 
               Height          =   2355
               Left            =   30
               TabIndex        =   96
               Top             =   4770
               Width           =   9810
               _cx             =   17304
               _cy             =   4154
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
               Rows            =   3
               Cols            =   11
               FixedRows       =   2
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FinancialAnalysisView.frx":4734
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
            Begin VB.Label LblDoenValueview 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   765
               Left            =   9870
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   5565
               Width           =   2100
            End
            Begin VB.Label Lblupvalueview 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   780
               Left            =   9870
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   2400
               Width           =   2100
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„ﬁ«„"
               Height          =   765
               Left            =   11985
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   5565
               Visible         =   0   'False
               Width           =   2565
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               Caption         =   "«·»”ÿ"
               Height          =   780
               Left            =   11985
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   2400
               Visible         =   0   'False
               Width           =   2565
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Currency"
               Height          =   1560
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   825
               Width           =   6810
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄„·… «·Õ”«»"
               Height          =   780
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   30
               Width           =   1560
            End
         End
      End
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   10305
      Index           =   6
      Left            =   20520
      TabIndex        =   56
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
               Picture         =   "FinancialAnalysisView.frx":4917
               Key             =   "Expanded_Node"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FinancialAnalysisView.frx":5769
               Key             =   "Root"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FinancialAnalysisView.frx":5B03
               Key             =   "Open_Node"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FinancialAnalysisView.frx":5E9D
               Key             =   "Closed_Node"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FinancialAnalysisView.frx":6237
               Key             =   "Item"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TrvAccounts 
         Height          =   10215
         HelpContextID   =   380
         Left            =   3330
         TabIndex        =   57
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
   Begin VB.Image Image8 
      Height          =   720
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "FinancialAnalysisView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xAccCol As Collection
Dim IntCurrentIndex As Integer
Dim Dcombos As ClsDataCombos
Dim StrTemp As String

Private objScript As Object

Private Sub CboYear_Change()
On Error Resume Next
    DTPickerAccFrom.value = "1-1-" & CboYear.Text
    DTPickerAccTo.value = "31-12-" & CboYear.Text
    DcBalanceSheet_Change
End Sub

Private Sub CboYear_Click()
    CboYear_Change
End Sub

Private Sub Check2_Click()

    If Check2.value = vbChecked Then
        Frame2.Enabled = True
    Else
        Frame2.Enabled = False
    End If

End Sub

Private Sub CmbMonth_Click()
    Dim StartDate As Integer
    Dim EndDate As Integer
    StartDate = CmbMonth.ListIndex + 1
 
    DTPickerAccFrom.value = StartDate & "-1-" & CboYear.Text
    DTPickerAccTo.value = StartDate & "-" & get_last_month_day(DTPickerAccFrom.value) & "-" & CboYear.Text
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim cReport As ClsAccReports

    Select Case Index

        Case 0
            Me.TxtModFlg.Text = "N"
       
            Dccurrency.BoundText = 1
            tXTAccount_Serial.Text = ""
            Me.DboParentAccount.BoundText = StrTemp
            TxtAccount_Name.Text = ""
            TxtAccount_NameE.Text = ""

        Case 1
            Me.TxtModFlg.Text = "E"

        Case 20

            If val(DcBalanceSheet.BoundText) <> 0 Then
                AddRecord val(DcBalanceSheet.BoundText), DTPickerAccFrom.value, DTPickerAccTo.value, val(TxtFinal.Text)
            End If

            'SaveData
        Case 3
            Me.TxtModFlg.Text = "R"

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

Function AddRecord(FinancialEquationsId As Integer, Fromdate As Date, ToDate As Date, Finalvalue As Double)
    Dim RsDev As ADODB.Recordset
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "FinancialEquationsHistory", Cn, adOpenStatic, adLockOptimistic, adCmdTable
           
    RsDev.AddNew
    RsDev("FinancialEquationsId").value = FinancialEquationsId
    RsDev("fromdate").value = Fromdate
    RsDev("ToDate").value = ToDate
    RsDev("Finalvalue").value = Finalvalue
    RsDev.update

    MsgBox " „ «·«÷«›…", vbInformation
End Function

Function PrintReport()
    'save data
    Dim RsDev As ADODB.Recordset
    Dim sql As String
        
    Dim cAccountReport As ClsAccReports
    Set cAccountReport = New ClsAccReports
    cAccountReport.ShowFinancialEquationsHistory val(DcBalanceSheet.BoundText)
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
    Dim Operator As String
    Dim generalvalue As Double
    Dim msgrslt As Integer

    If val(DcBalanceSheet.BoundText) = 0 Then Exit Sub
    BalanceSheetAccounts = getFinancialEquationData(val(DcBalanceSheet.BoundText), Operator, generalvalue)

    TxtFinancialEquationsOpr.Text = Operator

    Select Case Operator

        Case ""
            Image7.Picture = Image8.Picture

        Case "+"
            Image7.Picture = Image3.Picture

        Case "-"
            Image7.Picture = Image4.Picture

        Case "*"
            Image7.Picture = Image5.Picture

        Case "/"
            Image7.Picture = Image6.Picture
    End Select

    TxtGeneral.Text = generalvalue

    'LoadData Val(DcBalanceSheet.BoundText)

    FillGridWithData BalanceSheetAccounts, True
    FillGridWithData1 BalanceSheetAccounts, True

    totals

End Sub

Function totals()
    Dim TotalValue As Double
    Dim Equation As String
    Set objScript = CreateObject("MSScriptControl.ScriptControl")
    objScript.Language = "VBScript"

    If val(Me.LblDoenValue.Caption) <> 0 Then
        TotalValue = val(Me.Lblupvalue.Caption) / val(Me.LblDoenValue.Caption)
        TxtGeneralValue = Round(TotalValue, 2)

        If Me.TxtFinancialEquationsOpr.Text <> "" Then
            Equation = TotalValue & Me.TxtFinancialEquationsOpr.Text & val(Me.TxtGeneral)
            TotalValue = objScript.Eval(Equation)
            TxtFinal = Round(TotalValue, 2)
        Else
            TxtFinal = Round(TxtGeneralValue, 2)
        End If
 
    Else
        'MsgBox "·« Ì„ﬂ‰ «·ﬁ”„Â ⁄·Ï ’›—", vbCritical
        TotalValue = 0
    End If

End Function

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
    Dim Balance As Double
    Dim balancestr As String
    Dim StrAccountCode As String
    Dim account_name As String
    Dim TotalValue As Double
    Dim Equation As String
    Dim equationtext As String
    Dim Opr As String
    StrSQL = "SELECT     dbo.ACCOUNTS.Account_Code, dbo.ACCOUNTS.Account_Name, dbo.FinancialEquationsData.FinancialEquationsId, dbo.FinancialEquationsData.UpOrDown, "
    StrSQL = StrSQL & " dbo.FinancialEquationsData.Opr"
    StrSQL = StrSQL & " FROM         dbo.ACCOUNTS INNER JOIN"
    StrSQL = StrSQL & " dbo.FinancialEquationsData ON dbo.ACCOUNTS.Account_Code = dbo.FinancialEquationsData.Account_Code"
    StrSQL = StrSQL & " where UpOrDown=0  and FinancialEquationsId=" & val(DcBalanceSheet.BoundText)

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.FgAccounts
        .Clear flexClearScrollable, flexClearEverything

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = .FixedRows + rs.RecordCount
       
            Equation = "0"
            equationtext = ""

            For i = 2 To .Rows - 1
                StrAccountCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                account_name = IIf(IsNull(rs("account_name").value), "", rs("account_name").value)
                '   get_balanceFromGlNew StrAccountCode, , , True, Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value, , , Balance
           
                'WriteCustomerBalPublic StrAccountCode, balancestr
                'Balance = val(balancestr)
                Balance = GetActualAccountBalance(StrAccountCode, 0, DTPickerAccFrom.value, DTPickerAccTo.value)
                Opr = IIf(IsNull(rs("Opr").value), "", rs("Opr").value)
           
                equationtext = equationtext & account_name & Opr
           
                .TextMatrix(i, .ColIndex("Account_Code")) = StrAccountCode
                .TextMatrix(i, .ColIndex("account_name")) = IIf(IsNull(rs("account_name").value), "", rs("account_name").value)
                .TextMatrix(i, .ColIndex("Balance")) = Round(Balance, SystemOptions.SysDefCurrencyForamt)
                .TextMatrix(i, .ColIndex("Operator")) = Opr
          
                Equation = Equation & Balance & Opr
            
                rs.MoveNext
            Next i

        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        End If

        '    .AutoSize 0, .Cols - 1, False
        Set objScript = CreateObject("MSScriptControl.ScriptControl")
        objScript.Language = "VBScript"
        Equation = mId(Equation, 1, Len(Equation) - 1)

        TotalValue = objScript.Eval(Equation)
        txtUp.Text = equationtext & Round(Abs(TotalValue), 2)
        Lblupvalue.Caption = Abs(TotalValue)
        Lblupvalueview.Caption = FormatNumber(Abs(TotalValue), SystemOptions.SysDefCurrencyForamt, True, True, True)

    End With

End Sub

Private Sub FillGridWithData1(AccountsCodes As String, _
                              Optional showZeroAccounts As Boolean)
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Balance As Double
    Dim balancestr As String
    Dim StrAccountCode As String
    Dim account_name As String
    Dim TotalValue As Double
    Dim Equation As String
    Dim equationtext As String
    Dim Opr As String
    StrSQL = "SELECT     dbo.ACCOUNTS.Account_Code, dbo.ACCOUNTS.Account_Name, dbo.FinancialEquationsData.FinancialEquationsId, dbo.FinancialEquationsData.UpOrDown, "
    StrSQL = StrSQL & " dbo.FinancialEquationsData.Opr"
    StrSQL = StrSQL & " FROM         dbo.ACCOUNTS INNER JOIN"
    StrSQL = StrSQL & " dbo.FinancialEquationsData ON dbo.ACCOUNTS.Account_Code = dbo.FinancialEquationsData.Account_Code"
    StrSQL = StrSQL & " where UpOrDown=1  and FinancialEquationsId=" & val(DcBalanceSheet.BoundText)

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.FgAccounts1
        .Clear flexClearScrollable, flexClearEverything

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = .FixedRows + rs.RecordCount
       
            Equation = " "
            equationtext = "1"

            For i = 2 To .Rows - 1
                StrAccountCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                account_name = IIf(IsNull(rs("account_name").value), "", rs("account_name").value)
                '    get_balanceFromGlNew StrAccountCode, , , True, Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value, , , Balance
                 
                'WriteCustomerBalPublic StrAccountCode, balancestr
                'Balance = val(balancestr)
                Balance = GetActualAccountBalance(StrAccountCode, 0, DTPickerAccFrom.value, DTPickerAccTo.value)
                Opr = IIf(IsNull(rs("Opr").value), "", rs("Opr").value)
           
                equationtext = equationtext & account_name & Opr
           
                .TextMatrix(i, .ColIndex("Account_Code")) = StrAccountCode
                .TextMatrix(i, .ColIndex("account_name")) = IIf(IsNull(rs("account_name").value), "", rs("account_name").value)
                .TextMatrix(i, .ColIndex("Balance")) = Round(Balance, SystemOptions.SysDefCurrencyForamt)
                .TextMatrix(i, .ColIndex("Operator")) = Opr
          
                Equation = Equation & Balance & Opr
            
                rs.MoveNext
            Next i

        Else
            Equation = 1
        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        End If

        '    .AutoSize 0, .Cols - 1, False
        Set objScript = CreateObject("MSScriptControl.ScriptControl")
        objScript.Language = "VBScript"

        If Equation <> "1" Then
            Equation = mId(Equation, 1, Len(Equation) - 1)
        End If

        TotalValue = objScript.Eval(Equation)
 
        txtDown.Text = equationtext & Round(Abs(TotalValue), 2)
        LblDoenValue.Caption = Abs(TotalValue)

        LblDoenValueview.Caption = FormatNumber(Abs(TotalValue), SystemOptions.SysDefCurrencyForamt, True, True, True)

    End With

End Sub

Private Sub DTPickerAccFrom_Change()
    DcBalanceSheet_Change
End Sub

Private Sub DTPickerAccTo_Change()
    DcBalanceSheet_Change
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
    lbl(3).Caption = "Acc. Type"
    lbl(9).Caption = "Acc. Class."
    Frame4.Caption = "Authority"
    Authority(0).Caption = "All Users"
    Authority(1).Caption = "Group"
    Authority(2).Caption = "User"

    Label5.Caption = "Financial Equations Results"
    ELe(1).Caption = "Select Equations"
    lbl(0).Caption = "Select Equations"
    Cmd(20).Caption = "Save"

    Frame9.Caption = "Select Period"
    Option14.Caption = "1st Half"
    Option11.Caption = "2nd Half"
    Option5.Caption = "Monthly"
    Option9.Caption = "Dates"
    Option12.Caption = "Dates"
    Option3.Caption = "Annuals"

    Label8.Caption = "Month"
    lbl(100).Caption = "Result"
    lbl(4).Caption = "Operator"
    lbl(5).Caption = "value"
    lbl(12).Caption = "Net"
    lbl(10).Caption = "From"
    lbl(11).Caption = "To"

    Label6.Caption = "Year"
    Label7.Caption = "Comp Year"
    Option8.Caption = "1st Qr"
    Option7.Caption = "2nd Qr"
    Option6.Caption = "3rd Qr"
    Option9.Caption = "4th Qr"

    Frame1.Caption = "Cost Center"
    Frame2.Caption = "C.C Type"
    Option1.Caption = "Fixed"
    Option2.Caption = "Not Fixed"
    Label3.Caption = "CC Name"
    Me.Caption = "Financial Equations Results"
    ELe(1).Caption = "Account Data"
    lbl(1).Caption = "Account#"
    lbl(0).Caption = "Acc. Code"
    lbl(6).Caption = "Parent Account"
    lbl(2).Caption = "Name A"
    lbl(7).Caption = "Name E"
    'lbl(3).Caption = "Derived Account From This Acc"
 
    lbl(5).Visible = False
    Check1.Caption = " have A Budget"
    Check2.Caption = "Cost Center"
    Check3.Caption = "Sum account"
    lbl(8).Caption = "Curr"

    ELe(5).Caption = "Account Type"

    With FgAccounts
 
        .TextMatrix(1, .ColIndex("Account_name")) = "Account Name"
        .TextMatrix(1, .ColIndex("Operator")) = "Operator"
        .TextMatrix(1, .ColIndex("Balance")) = "Balance"
    End With

    With FgAccounts1
 
        .TextMatrix(1, .ColIndex("Account_name")) = "Account Name"
        .TextMatrix(1, .ColIndex("Operator")) = "Operator"
        .TextMatrix(1, .ColIndex("Balance")) = "Balance"
    End With

    OptAccountType(0).Caption = "Final"
    OptAccountType(1).Caption = "Master"

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    'Cmd(2).Caption = "Save"
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

    CboYear.Clear

    For i = 2012 To 2050
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = 1
    CboYear1.Clear

    For i = 2012 To 2050
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

    DTPickerAccFrom.value = "1-1-" & CboYear.Text
    DTPickerAccTo.value = "31-12-" & CboYear.Text
    
    Dim GrdBack As ClsBackGroundPic

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
     'Me.lbl(5).Caption = Msg
    '----------------------------
    Set xAccCol = New Collection
    '----------------------------
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(20).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
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
  
    Dcombos.GetFinancialEquations DcBalanceSheet
   
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

    With FgAccounts
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
     
        .MergeCol(.ColIndex("AccountName")) = True
        .MergeCol(.ColIndex("Operator")) = True
        .Cell(flexcpText, 0, .ColIndex("Account_Name"), 0, .ColIndex("Operator")) = "     "
        .Cell(flexcpAlignment, 0, .ColIndex("Account_Name"), 0, .ColIndex("Operator")) = flexAlignCenterCenter
    End With

    With FgAccounts1
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
     
        .MergeCol(.ColIndex("AccountName")) = True
        .MergeCol(.ColIndex("Operator")) = True
        .Cell(flexcpText, 0, .ColIndex("Account_Name"), 0, .ColIndex("Operator")) = "   "
        .Cell(flexcpAlignment, 0, .ColIndex("Account_Name"), 0, .ColIndex("Operator")) = flexAlignCenterCenter
    End With

    With Me.TrvAccounts
        .Appearance = ccFlat
        .Checkboxes = False
        .BorderStyle = ccNone
        .LineStyle = tvwRootLines
        .SingleSel = False
    End With

    'LoadData
    Me.TxtModFlg.Text = "R"
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
        xAccCol.Add StrAccountCode, "A" & IntIndexPut
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
        Me.TxtAccount_ID.Text = IIf(IsNull(rs("Account_ID").value), "", rs("Account_ID").value)
        Me.TxtAccount_Code.Text = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
        Me.tXTAccount_Serial.Text = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
        Me.TxtAccount_Name.Text = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
        Me.TxtAccount_NameE.Text = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)

        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbllevel.Caption = "«·„” ÊÏ : " & CountAs(Me.TxtAccount_Code.Text)
        Else
            Me.lbllevel.Caption = "Level:" & CountAs(Me.TxtAccount_Code.Text)
        End If

        Me.Check1.value = IIf(rs("mowazna").value = True, vbChecked, Unchecked)
        Me.Check2.value = IIf(rs("cost_center").value = True, vbChecked, Unchecked)
     
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
    DTPickerAccFrom.value = "1-7-" & CboYear.Text
    DTPickerAccTo.value = "31-12-" & CboYear.Text
End Sub

Private Sub Option14_Click()
    DTPickerAccFrom.value = "1-1-" & CboYear.Text
    DTPickerAccTo.value = "30-6-" & CboYear.Text
End Sub

Private Sub Option2_Click()

    If Option1.value = True Then
        DcCostCenter.Enabled = True
    Else
        DcCostCenter.Enabled = False
    End If

End Sub

Private Sub Option3_Click()
    On Error Resume Next
    DTPickerAccFrom.value = "1-1-" & CboYear.Text
    DTPickerAccTo.value = "31-12-" & CboYear.Text
End Sub

Private Sub Option6_Click()
    On Error Resume Next
    DTPickerAccFrom.value = "7-1-" & CboYear.Text
    DTPickerAccTo.value = "30-9-" & CboYear.Text
End Sub

Private Sub Option7_Click()
    DTPickerAccFrom.value = "4-1-" & CboYear.Text
    DTPickerAccTo.value = "30-6-" & CboYear.Text
End Sub

Private Sub Option8_Click()
    DTPickerAccFrom.value = "1-1-" & CboYear.Text
    DTPickerAccTo.value = "31-3-" & CboYear.Text
End Sub

Private Sub Option9_Click()
    DTPickerAccFrom.value = "10-1-" & CboYear.Text
    DTPickerAccTo.value = "31-12-" & CboYear.Text
End Sub

Private Sub TrvAccounts_NodeClick(ByVal Node As MSComctlLib.Node)

    If Not Node Is Nothing Then
        If InStr(1, Node.Key, "G", vbTextCompare) <> 0 Then
            StrTemp = Node.Key
            StrTemp = mId(StrTemp, 1, Len(StrTemp) - 1)
        Else
            StrTemp = Node.Key
        End If

        If Me.TxtModFlg.Text = "R" Then
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

    Select Case TxtModFlg.Text

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
            '       Cmd(2).Enabled = False
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

    If val(Me.TxtAccount_ID.Text) <> 0 Then
        StrSQL = "Select * From Accounts Where Account_ID=" & val(Me.TxtAccount_ID.Text)
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
    StrAccountCode = Trim(Me.TxtAccount_Code.Text)

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
            Msg = Msg & CHR(13) & "Account Code :- " & Me.tXTAccount_Serial.Text
            Msg = Msg & CHR(13) & "Account Name " & Me.TxtAccount_Name.Text
            Msg = Msg & CHR(13) & ""
            Msg = Msg & CHR(13) & "Sure you want delte ??"
          
        Else
          
            Msg = "”Ê› Ì „ Õ–› «·Õ”«» —ﬁ„:- " '& Me.TxtAccount_ID.text
            Msg = Msg & CHR(13) & " ﬂÊœ «Ê „”·”· «·Õ”«»:- " & Me.tXTAccount_Serial.Text
            Msg = Msg & CHR(13) & "«”„ «·Õ”«» :- " & Me.TxtAccount_Name.Text
            Msg = Msg & CHR(13) & ""
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ ⁄„·Ì… «·Õ–› ..øø"
        
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
        Me.TrvAccounts.Nodes.Remove XNode.Key
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

    If Me.TxtAccount_Name.Text = "" And Me.TxtAccount_NameE.Text <> "" Then
        Me.TxtAccount_Name.Text = Me.TxtAccount_NameE.Text
    End If

    If Me.TxtAccount_Name.Text <> "" And Me.TxtAccount_NameE.Text = "" Then
        Me.TxtAccount_NameE.Text = Me.TxtAccount_Name.Text
    End If

    If Trim$(Me.TxtAccount_Name.Text) = "" Then
        Msg = "ÌÃ» ﬂ «»… «”„ «·Õ”«»"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If SystemOptions.UserInterface = EnglishInterface Then

        If Trim$(Me.TxtAccount_NameE.Text) = "" Then
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

    If Me.TxtModFlg.Text = "N" Then
        Dim mowazna As Boolean
        Dim cost_center As Boolean
        Dim Sum_account As Boolean

        If Check1.value = vbChecked Then
            mowazna = 1
        Else
            mowazna = 0
        End If

        If Check2.value = vbChecked Then
            cost_center = 1
        Else
            cost_center = 0
        End If

        If Check3.value = vbChecked Then
            Sum_account = 1
        Else
            Sum_account = 0
        End If

        StrNewAccountCode = ModAccounts.AddNewAccount(Me.DboParentAccount.BoundText, Trim$(Me.TxtAccount_Name.Text), BolLastAccount, False, Trim$(Me.TxtAccount_NameE.Text), Dccurrency.BoundText, mowazna, cost_center, Sum_account, , tXTAccount_Serial)
    
        If StrNewAccountCode <> "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Saved"
            Else
                Msg = " „  ⁄„·Ì… «·Œ›Ÿ."
            End If
    
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.TxtModFlg.Text = "R"
        End If

    ElseIf Me.TxtModFlg.Text = "E" Then
        StrSQL = "Select * From Accounts Where Account_ID=" & val(Me.TxtAccount_ID.Text) & ""
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

        If Not (rs.BOF Or rs.EOF) Then

            '·Ê «‰ «·Õ”«» Õ”«» —∆Ì”Ì ÊÌ—Ìœ «‰ ÌÃ⁄·Â Õ”«» ‰Â«∆
            If rs("last_account").value = False And OptAccountType(0).value = True Then
                If GetAccountChilds(Me.TxtAccount_Code.Text) <> 0 Then
                    If SystemOptions.UserInterface = EnglishInterface Then
                        Msg = "This Account is Master Account  And have child "
                        Msg = Msg & CHR(13) & "Can't change to last Account"
            
                    Else
         
                        Msg = "Â–« «·Õ”«» Õ”«» —∆Ì”Ì ÊÌÕ ÊÏ ⁄·Ï Õ”«»«  „ ›—⁄… „‰Â"
                        Msg = Msg & CHR(13) & "Ê·« Ì„ﬂ‰  ⁄œÌ·Â ≈·Ï Õ”«» ‰Â«∆"
                    End If
            
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If

            ElseIf rs("last_account").value = True And Me.OptAccountType(1).value = True Then
                '«·Õ”«» «·„—«œ  ⁄œÌ·Â Õ”«» ‰Â«∆ ÊÌ—«œ  ÕÊÌ·Â ≈·Ï Õ”«» —∆Ì”Ï
            
            Else
                Dim currency_code As Integer

                If Not IsNull(TxtAccount_Code.Text) Then
                    If Me.Dccurrency.BoundText <> "" Then
                        currency_code = Me.Dccurrency.BoundText
                    Else
                        currency_code = 1
                    End If

                    ModAccounts.EditAccount rs("Account_Code").value, Me.TxtAccount_Name.Text, Me.TxtAccount_NameE.Text, Check1.value, Check2.value, currency_code, Check3.value, tXTAccount_Serial

                    If SystemOptions.UserInterface = EnglishInterface Then
                        Msg = "Saved"
                    Else
                  
                        Msg = " „  ⁄„·Ì… «·Œ›Ÿ."
                    End If
                  
                    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Me.TxtModFlg.Text = "R"
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
