VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form System_alarms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   ‰»ÌÂ«  «·ÌÊ„    "
   ClientHeight    =   10905
   ClientLeft      =   195
   ClientTop       =   525
   ClientWidth     =   15330
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004000&
   Icon            =   "systeml_alarms.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10905
   ScaleWidth      =   15330
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic14 
      Height          =   10905
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   15330
      _cx             =   27040
      _cy             =   19235
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   10365
         Left            =   0
         TabIndex        =   3
         Top             =   600
         Width           =   15180
         _cx             =   26776
         _cy             =   18283
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
         Appearance      =   1
         MousePointer    =   0
         Version         =   801
         BackColor       =   14726431
         ForeColor       =   -2147483640
         FrontTabColor   =   8421631
         BackTabColor    =   14726431
         TabOutlineColor =   14726431
         FrontTabForeColor=   -2147483639
         Caption         =   $"systeml_alarms.frx":000C
         Align           =   0
         CurrTab         =   14
         FirstTab        =   12
         Style           =   2
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   0   'False
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   4
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   -1  'True
         MultiRow        =   -1  'True
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   1
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   8775
            Left            =   -19635
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin C1SizerLibCtl.C1Elastic CApprovedData 
               Height          =   855
               Left            =   9075
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   120
               Width           =   5400
               _cx             =   9525
               _cy             =   1508
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
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
               Begin ALLButtonS.ALLButton ALLButton18 
                  Height          =   375
                  Left            =   90
                  TabIndex        =   6
                  Top             =   360
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   661
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0120
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label30 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«   „” ‰œ«  ﬁÌœ «·«⁄ „«œ"
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
                  Height          =   375
                  Left            =   1500
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   0
                  Width           =   3150
               End
               Begin VB.Label Label19 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„” ‰œ«  ﬁÌœ «·«⁄ „«œ"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1350
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   360
                  Width           =   3870
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   8775
            Left            =   -19335
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin C1SizerLibCtl.C1Elastic CAccount 
               Height          =   2055
               Left            =   7200
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   0
               Width           =   7335
               _cx             =   12938
               _cy             =   3625
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
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
               Begin ALLButtonS.ALLButton ALLButton1 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   11
                  Top             =   225
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   635
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":013C
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton8 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   12
                  Top             =   675
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   635
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   2
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   16777152
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0158
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton10 
                  Height          =   345
                  Left            =   120
                  TabIndex        =   13
                  Top             =   1140
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   609
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0174
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton19 
                  Height          =   345
                  Left            =   120
                  TabIndex        =   14
                  Top             =   1590
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   609
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   2
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0190
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label21 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «·Õ”«»« "
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
                  Height          =   360
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   0
                  Width           =   2175
               End
               Begin VB.Label Label20 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«⁄ „«œ«  Ê «·÷„«‰«  «·»‰ﬂÌ…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   1590
                  Width           =   5055
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ŒÿÂ   Ê“Ì⁄ «·Õ”«»« "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   465
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   1140
                  Width           =   5055
               End
               Begin VB.Label Label8 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ÕœÊœ «∆ „«‰ «·⁄„·«¡"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   675
                  Width           =   5055
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«Ê—«ﬁ «·„«·Ì…  «·„” Õﬁ…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   15
                  Top             =   345
                  Width           =   5055
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   8775
            Left            =   -19035
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin C1SizerLibCtl.C1Elastic CFixed 
               Height          =   975
               Left            =   7200
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   0
               Width           =   7335
               _cx             =   12938
               _cy             =   1720
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
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
               Begin ALLButtonS.ALLButton ALLButton29 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   22
                  Top             =   345
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   661
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":01AC
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label43 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «·«’Ê· «·À«» …"
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
                  Height          =   375
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   0
                  Width           =   2175
               End
               Begin VB.Label Label44 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «ﬁ”«ÿ «·«’Ê· «·À«» …"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1800
                  RightToLeft     =   -1  'True
                  TabIndex        =   23
                  Top             =   345
                  Width           =   5175
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   8775
            Left            =   -18735
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin C1SizerLibCtl.C1Elastic CHR 
               Height          =   855
               Left            =   7200
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   0
               Width           =   7215
               _cx             =   12726
               _cy             =   1508
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
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
               Begin ALLButtonS.ALLButton ALLButton4 
                  Height          =   450
                  Left            =   120
                  TabIndex        =   27
                  Top             =   330
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   794
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":01C8
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label28 
                  Alignment       =   1  'Right Justify
                  Caption         =   "‘∆Ê‰ «·„ÊŸ›Ì‰"
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
                  Height          =   465
                  Left            =   1995
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   0
                  Width           =   3255
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  Caption         =   "‘∆Ê‰ «·„ÊŸ›Ì‰"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   315
                  Width           =   5295
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   8775
            Left            =   -18435
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin C1SizerLibCtl.C1Elastic CStores 
               Height          =   3975
               Left            =   7080
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   0
               Width           =   7335
               _cx             =   12938
               _cy             =   7011
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
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
               Begin ALLButtonS.ALLButton ALLButton2 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   32
                  Top             =   345
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   661
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":01E4
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton13 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   33
                  Top             =   705
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   635
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0200
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton14 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   34
                  Top             =   1050
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   635
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":021C
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton24 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   35
                  Top             =   2220
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   635
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0238
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton25 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   36
                  Top             =   2565
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   635
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0254
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton27 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   37
                  Top             =   2910
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   635
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0270
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton30 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   38
                  Top             =   3255
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   635
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":028C
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton34 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   39
                  Top             =   1395
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   635
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":02A8
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton9 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   116
                  Top             =   1800
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   635
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":02C4
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  ÕÃÊ“«  «·«’‰«›"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   1750
                  Width           =   4935
               End
               Begin VB.Label Label22 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «·„” Êœ⁄« "
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
                  Height          =   360
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   0
                  Width           =   2175
               End
               Begin VB.Label Label45 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «· ÕÊÌ·«  »Ì‰ «·„Œ«“‰"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   3135
                  Width           =   4935
               End
               Begin VB.Label Label41 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «·ﬂ„Ì«  /„ÕÃÊ“/„”·„ /„ »ﬁÌ    "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   2790
                  Width           =   4935
               End
               Begin VB.Label Label38 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  ”‰œ«  «·«” ·«„"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   2445
                  Width           =   4935
               End
               Begin VB.Label Label37 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «·ÿ·»«  «·œ«Œ·Ì…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   2100
                  Width           =   4935
               End
               Begin VB.Label Label15 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«’‰«› «· Ì ﬁ«—» «‰ Â«¡ ÷„«‰Â«"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   1035
                  Width           =   4935
               End
               Begin VB.Label Label14 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«’‰«› «·—«ﬂœ…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   690
                  Width           =   4935
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«’‰«› «· Ì »·€  Õœ «·ÿ·»"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   345
                  Width           =   4935
               End
               Begin VB.Label Label49 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«’‰«› «· Ì ﬁ«—» «‰ Â«¡ ’·«ÕÌ Â«"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   1395
                  Width           =   4935
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   8775
            Left            =   -18135
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin C1SizerLibCtl.C1Elastic CPurchase 
               Height          =   1575
               Left            =   7200
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   0
               Width           =   7335
               _cx             =   12938
               _cy             =   2778
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
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
               Begin ALLButtonS.ALLButton ALLButton26 
                  Height          =   555
                  Left            =   120
                  TabIndex        =   51
                  Top             =   840
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   979
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":02E0
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton32 
                  Height          =   555
                  Left            =   120
                  TabIndex        =   52
                  Top             =   375
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   979
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":02FC
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label39 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «·„‘ —Ì« "
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
                  Height          =   555
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   15
                  Width           =   2175
               End
               Begin VB.Label Label47 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  ÿ·»«  «·‘—«¡ ⁄‰ › —…     "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   555
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   495
                  Width           =   4935
               End
               Begin VB.Label Label40 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  ”‰œ«  «·«” ·«„ €Ì— „—»Êÿ… »›Ê« Ì—"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   555
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   960
                  Width           =   4935
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   8775
            Left            =   -17835
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin C1SizerLibCtl.C1Elastic CSAles 
               Height          =   1845
               Left            =   660
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   120
               Width           =   13755
               _cx             =   24262
               _cy             =   3254
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
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
               Begin ALLButtonS.ALLButton ALLButton11 
                  Height          =   555
                  Left            =   165
                  TabIndex        =   58
                  Top             =   1245
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   979
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0318
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton20 
                  Height          =   495
                  Left            =   165
                  TabIndex        =   59
                  Top             =   360
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   873
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0334
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton21 
                  Height          =   435
                  Left            =   165
                  TabIndex        =   60
                  Top             =   855
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   767
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0350
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin MSDataListLib.DataCombo DCboStore2Name 
                  Height          =   555
                  Left            =   4950
                  TabIndex        =   138
                  Top             =   420
                  Width           =   2955
                  _ExtentX        =   5212
                  _ExtentY        =   979
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo Dcbranch 
                  Height          =   555
                  Left            =   2070
                  TabIndex        =   140
                  Top             =   360
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   979
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  ForeColor       =   &H00000000&
                  Height          =   465
                  Index           =   36
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   360
                  Width           =   660
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·„Œ“‰"
                  Height          =   450
                  Index           =   47
                  Left            =   7800
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   405
                  Width           =   1215
               End
               Begin VB.Label Label23 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «·„»Ì⁄«  "
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
                  Height          =   510
                  Left            =   5865
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   0
                  Width           =   2805
               End
               Begin VB.Label Label24 
                  Alignment       =   1  'Right Justify
                  Caption         =   "›Ê« Ì— „»Ì⁄«  ·Ì” ·Â« ”‰œ ’—› „Œ“‰Ì"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   510
                  Left            =   8115
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   450
                  Width           =   5175
               End
               Begin VB.Label Label33 
                  Alignment       =   1  'Right Justify
                  Caption         =   "”‰œ«   ÕÊÌ· „Œ“‰Ì ·„  ’œ— »Â« ›Ê« Ì— „»Ì⁄«   Œ·«·"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   4605
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   855
                  Width           =   8685
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  Caption         =   " √Ê«„— «·»Ì⁄ «· Ì  „ «⁄ „«œÂ« Ê·„ Ì’œ— »Â« ›« Ê—…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   555
                  Left            =   4605
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   1245
                  Width           =   8685
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   8775
            Left            =   -17535
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin C1SizerLibCtl.C1Elastic CSalesInss 
               Height          =   1335
               Left            =   7200
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   0
               Width           =   7335
               _cx             =   12938
               _cy             =   2355
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
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
               Begin ALLButtonS.ALLButton ALLButton16 
                  Height          =   465
                  Left            =   120
                  TabIndex        =   67
                  Top             =   720
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   820
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":036C
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton3 
                  Height          =   435
                  Left            =   120
                  TabIndex        =   68
                  Top             =   300
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   767
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0388
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label310 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «·«ﬁ”«ÿ"
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
                  Height          =   465
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   0
                  Width           =   2175
               End
               Begin VB.Label Label17 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«ﬁ”«ÿ «·„ √Œ—…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   465
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   720
                  Width           =   5055
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·«ﬁ”«ÿ «· Ì Õ«‰ Êﬁ   ”œ«œÂ«"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   300
                  Width           =   5055
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   8775
            Left            =   -17235
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin C1SizerLibCtl.C1Elastic CProjects 
               Height          =   1815
               Left            =   7320
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   0
               Width           =   7215
               _cx             =   12726
               _cy             =   3201
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
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
               Begin ALLButtonS.ALLButton ALLButton6 
                  Height          =   390
                  Left            =   120
                  TabIndex        =   74
                  Top             =   360
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   688
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":03A4
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton7 
                  Height          =   405
                  Left            =   120
                  TabIndex        =   75
                  Top             =   900
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   714
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":03C0
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton31 
                  Height          =   390
                  Left            =   120
                  TabIndex        =   76
                  Top             =   1290
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   688
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":03DC
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label26 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «·„‘«—Ì⁄"
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
                  Height          =   405
                  Left            =   1995
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   0
                  Width           =   3255
               End
               Begin VB.Label Label46 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  ‰”» «· Õﬁﬁ ··⁄„·Ì« "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   1290
                  Width           =   5295
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„” Œ·’«  «·„‘«—Ì⁄ «· Ì ·„  ”œœ »«·ﬂ«„·"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   900
                  Width           =   5295
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«‰Õ—«›«  «·„‘«—Ì⁄"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   510
                  Width           =   5295
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic10 
            Height          =   8775
            Left            =   -16935
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin C1SizerLibCtl.C1Elastic CRS 
               Height          =   6255
               Left            =   6600
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   0
               Width           =   7815
               _cx             =   13785
               _cy             =   11033
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
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
               Begin ALLButtonS.ALLButton ALLButton5 
                  Height          =   510
                  Left            =   120
                  TabIndex        =   83
                  Top             =   2325
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   900
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":03F8
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton35 
                  Height          =   495
                  Left            =   120
                  TabIndex        =   84
                  Top             =   600
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   873
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0414
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton36 
                  Height          =   510
                  Left            =   120
                  TabIndex        =   85
                  Top             =   1095
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   900
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0430
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton37 
                  Height          =   510
                  Left            =   120
                  TabIndex        =   86
                  Top             =   1680
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   900
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":044C
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton38 
                  Height          =   495
                  Left            =   120
                  TabIndex        =   87
                  Top             =   2940
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   873
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0468
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label27 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  ≈œ«—… «·√„·«ﬂ"
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
                  Height          =   600
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   0
                  Width           =   3525
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «·’Ì«‰… «· Ì ·„  ﬁ›·"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   2940
                  Width           =   5295
               End
               Begin VB.Label Label32 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  œ›⁄«  «·„·«ﬂ «·„” Õﬁ…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   510
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   1590
                  Width           =   5295
               End
               Begin VB.Label Label50 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «‰ Â«¡ «·⁄ﬁÊœ"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   390
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   1095
                  Width           =   4935
               End
               Begin VB.Label Label51 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «·«ﬁ”«ÿ «·„” Õﬁ… ⁄·Ï «·„” √Ã—Ì‰"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   375
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   615
                  Width           =   5415
               End
               Begin VB.Label Label52 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «·⁄—«»Ì‰ «·„‰ ÂÌ… „œ… «·”„«Õ ·Â«"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   510
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   2325
                  Width           =   5415
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic11 
            Height          =   8775
            Left            =   -16635
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin C1SizerLibCtl.C1Elastic Ctransportation 
               Height          =   2895
               Left            =   6480
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   0
               Width           =   7935
               _cx             =   13996
               _cy             =   5106
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
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
               Begin ALLButtonS.ALLButton ALLButton1t 
                  Height          =   495
                  Left            =   0
                  TabIndex        =   118
                  Top             =   480
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   873
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0484
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton2t 
                  Height          =   495
                  Left            =   0
                  TabIndex        =   119
                  Top             =   1560
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   873
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":04A0
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton3t 
                  Height          =   495
                  Left            =   0
                  TabIndex        =   120
                  Top             =   960
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   873
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":04BC
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton12 
                  Height          =   495
                  Left            =   0
                  TabIndex        =   127
                  Top             =   2040
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   873
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":04D8
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label53 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  Œÿ… «·’Ì«‰…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   129
                  Top             =   2040
                  Width           =   5055
               End
               Begin VB.Label Label13 
                  Alignment       =   2  'Center
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   128
                  Top             =   2160
                  Visible         =   0   'False
                  Width           =   1215
               End
               Begin VB.Label Label5t 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«Ã„«·Ì  ⁄œœ «·„⁄œ« /«·”Ì«—«  «· Ì ”Ì‰ ÂÏ  √„Ì‰Â«"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   1080
                  Width           =   4935
               End
               Begin VB.Label dT3 
                  Alignment       =   2  'Center
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   1560
                  Width           =   1215
               End
               Begin VB.Label dT2 
                  Alignment       =   2  'Center
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   1080
                  Width           =   1215
               End
               Begin VB.Label Label4t 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«Ã„«·Ì  ⁄œœ «·„⁄œ« /«·”Ì«—«  «· Ì ”Ì‰ ÂÏ ›Õ’Â«"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   1560
                  Width           =   5055
               End
               Begin VB.Label dT1 
                  Alignment       =   2  'Center
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Label Label2t 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«Ã„«·Ì  ⁄œœ «·„⁄œ« /«·”Ì«—«  «· Ì ” ‰ ÂÌ «” „«— Â«"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   2880
                  RightToLeft     =   -1  'True
                  TabIndex        =   121
                  Top             =   480
                  Width           =   4815
               End
               Begin VB.Label Label29 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«   «·‰ﬁ·Ì« "
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
                  Height          =   375
                  Left            =   1995
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   0
                  Width           =   3255
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic12 
            Height          =   8775
            Left            =   -16335
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin C1SizerLibCtl.C1Elastic CShipments 
               Height          =   1335
               Left            =   7320
               TabIndex        =   98
               TabStop         =   0   'False
               Top             =   0
               Width           =   7215
               _cx             =   12726
               _cy             =   2355
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
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
               Begin ALLButtonS.ALLButton ALLButton22 
                  Height          =   405
                  Left            =   120
                  TabIndex        =   99
                  Top             =   270
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   714
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":04F4
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton23 
                  Height          =   405
                  Left            =   120
                  TabIndex        =   100
                  Top             =   795
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   714
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0510
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label35 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·›—ﬁ »Ì‰ «·ﬂ„Ì… «·„ÿ·Ê»… Ê «·„‘ÕÊ‰…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   390
                  Width           =   5295
               End
               Begin VB.Label Label36 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·›—ﬁ »Ì‰ «·ﬂ„Ì…   «·„‘ÕÊ‰…  Ê «·„” ·„…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   795
                  Width           =   5295
               End
               Begin VB.Label Label34 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«   «·‘Õ‰"
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
                  Height          =   405
                  Left            =   1995
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   0
                  Width           =   3255
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic13 
            Height          =   8775
            Left            =   -16035
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin C1SizerLibCtl.C1Elastic CMaintenance 
               Height          =   3015
               Left            =   7320
               TabIndex        =   105
               TabStop         =   0   'False
               Top             =   0
               Width           =   7215
               _cx             =   12726
               _cy             =   5318
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
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
               Begin ALLButtonS.ALLButton ALLButton15 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   106
                  Top             =   480
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   661
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":052C
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton17 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   107
                  Top             =   960
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   661
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0548
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton28 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   108
                  Top             =   1440
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   661
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0564
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton33 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   109
                  Top             =   1920
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   661
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":0580
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ALLButtonS.ALLButton ALLButton39 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   130
                  Top             =   2400
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   661
                  BTYPE           =   3
                  TX              =   "⁄—÷"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   -1  'True
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "systeml_alarms.frx":059C
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label54 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  Œÿ… «·’Ì«‰…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   131
                  Top             =   2400
                  Width           =   5055
               End
               Begin VB.Label Label25 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «·’Ì«‰…"
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
                  Height          =   375
                  Left            =   1995
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   0
                  Width           =   3255
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·’Ì«‰… «·Êﬁ«∆Ì…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   480
                  Width           =   5295
               End
               Begin VB.Label Label18 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«ÃÂ“… «·’Ì«‰… «·Ã«Â“… ·· ”·Ì„"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   112
                  Top             =   960
                  Width           =   5295
               End
               Begin VB.Label Label42 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  «·’Ì«‰… «·„› ÊÕ…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   1440
                  Width           =   5295
               End
               Begin VB.Label Label48 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ«  ⁄ﬁÊœ «·’Ì«‰… Ê«·÷„«‰"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   1920
                  Width           =   5295
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic15 
            Height          =   8775
            Left            =   -15735
            TabIndex        =   132
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin ALLButtonS.ALLButton ALLButton41 
               Height          =   375
               Left            =   9360
               TabIndex        =   134
               Top             =   600
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "⁄—÷"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "systeml_alarms.frx":05B8
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label56 
               Alignment       =   1  'Right Justify
               Caption         =   " ‰»ÌÂ«  «·„ﬁ«Ì”« "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   11160
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   600
               Width           =   2655
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic16 
            Height          =   8775
            Left            =   45
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   1545
            Width           =   14505
            _cx             =   25585
            _cy             =   15478
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Begin ALLButtonS.ALLButton ALLButton40 
               Height          =   375
               Left            =   9360
               TabIndex        =   136
               Top             =   600
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "⁄—÷"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "systeml_alarms.frx":05D4
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ALLButtonS.ALLButton ALLButton42 
               Height          =   360
               Left            =   9360
               TabIndex        =   142
               Top             =   1200
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   635
               BTYPE           =   3
               TX              =   "⁄—÷"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "systeml_alarms.frx":05F0
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label57 
               Alignment       =   1  'Right Justify
               Caption         =   " ‰»ÌÂ«  «·ÿ·»«  «·œ«Œ·Ì…"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   9240
               RightToLeft     =   -1  'True
               TabIndex        =   143
               Top             =   1080
               Width           =   4935
            End
            Begin VB.Label Label55 
               Alignment       =   1  'Right Justify
               Caption         =   "„ «»⁄Â «·«‰ «Ã"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   11160
               RightToLeft     =   -1  'True
               TabIndex        =   137
               Top             =   600
               Width           =   2655
            End
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "   ‰»ÌÂ«  «·ÌÊ„      "
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
         Height          =   855
         Index           =   2
         Left            =   6555
         RightToLeft     =   -1  'True
         TabIndex        =   115
         Top             =   -120
         Width           =   3000
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   960
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2040
      Top             =   4200
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì  ⁄œœ  «·„ÊŸ›Ì‰  «· Ì ” ‰ ÂÌ  √„Ì‰« Â„"
      Height          =   855
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   11160
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Label d5 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   615
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   9960
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "System_alarms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Askinterval As String
Dim Askcount As Integer
Dim BolShowRequest As Boolean

Private Sub ALLButton1_Click()

    If checkApility("FrmPaymentTime") = False Then
        Exit Sub
    End If

    FrmPaymentTime.show
    FrmPaymentTime.ZOrder 0
End Sub

Private Sub ALLButton10_Click()

    If checkApility("FrmAccountDestributionView") = False Then
        Exit Sub
    End If

    FrmAccountDestributionView.show
End Sub

Private Sub ALLButton11_Click()

    If checkApility("FrmAccreditOrder") = False Then
        Exit Sub
    End If

    frmaccreditOrder.show
End Sub

Private Sub ALLButton12_Click()

Unload FrmCarExpireLicens
FrmCarExpireLicens.Indx = 1
FrmCarExpireLicens.show


End Sub

Private Sub ALLButton13_Click()

    If checkApility("FrmStagnantItems") = False Then
        Exit Sub
    End If

    OpenScreen PopUpSowStagnantItems
End Sub

Private Sub ALLButton14_Click()

    If checkApility("FrmGuaranteeAlram") = False Then
        Exit Sub
    End If
FrmGuaranteeAlram.Ind = 0
    OpenScreen PopUpShowGuaranteeAlram
End Sub

Private Sub ALLButton15_Click()

    If checkApility("FrmPerfMantAlaram") = False Then
        Exit Sub
    End If

    FrmPerfMantAlaram.show
End Sub

Private Sub ALLButton16_Click()

    If checkApility("FrmCustomerBalances") = False Then
        Exit Sub
    End If

    FrmCustomerBalances.show
End Sub

Private Sub ALLButton17_Click()

    If checkApility("FrmManStore") = False Then
        Exit Sub
    End If

'    FrmManStore.show
'    FrmManStore.ZOrder 0
 
'    FrmManStore.TabMain.CurrTab = 3

End Sub

Private Sub ALLButton18_Click()
'    If checkApility("FrmApprovalTransactions") = False Then
''        Exit Sub
 '   End If
    
    
    FrmApprovalTransactions.ScreenName = ""
    FrmApprovalTransactions.show
    
    
    
End Sub

Private Sub ALLButton19_Click()
FrmEmpExpir20.show
End Sub

Private Sub ALLButton1t_Click()
Unload FrmCarExpireLicens
FrmCarExpireLicens.Indx = 0
FrmCarExpireLicens.show

End Sub

Private Sub ALLButton2_Click()

    If checkApility("FrmRequest") = False Then
        Exit Sub
    End If

    FrmRequest.show
    FrmRequest.ZOrder 0
End Sub

Private Sub ALLButton20_Click()
Dim sql As String
Dim path As String
sql = "SELECT     dbo.Transactions.Transaction_Date, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, dbo.TblBranchesData.branch_name, "
sql = sql & "  dbo.TblBranchesData.branch_namee, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
sql = sql & "  dbo.TblEmployee.Emp_Namee , dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Transactions.Nots"
sql = sql & "  FROM         dbo.Transactions LEFT OUTER JOIN"
sql = sql & "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
sql = sql & "  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "   dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
sql = sql & "    dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID"
sql = sql & "  WHERE      (dbo.Transactions.Transaction_Type = 21) AND (dbo.Transactions.Nots IS NULL or dbo.Transactions.Nots='' )"
If val(DCboStore2Name.BoundText) <> 0 And Me.DCboStore2Name.Text <> "" Then
    sql = sql + " and  dbo.Transactions.StoreID=" & val(DCboStore2Name.BoundText)
End If
If val(Dcbranch.BoundText) <> 0 And Me.Dcbranch.Text <> "" Then
    sql = sql + " and  dbo.Transactions.BranchId=" & val(Dcbranch.BoundText)
End If
  
path = App.path & "\Reports\REPORTS NEW\SalesWithNoIsuueVchr.rpt"
PrintSimpleReport sql, path
End Sub

Private Sub ALLButton21_Click()
Dim sql As String
Dim path As String
X = InputBox("Õœœ ⁄œœ «·«Ì«„ ·· ﬁ—Ì—")

sql = "SELECT     Transactions_1.order_no, dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Type, dbo.Transactions.Transaction_Date, dbo.Transactions.StoreID, "
sql = sql & "                        dbo.Transactions.CusID, dbo.Transactions.BranchId, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
sql = sql & "  dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, TblStore_1.StoreName AS tostorename, TblStore_1.StoreNamee AS tostorenamee,"
sql = sql & "  TblBranchesData_1.branch_name AS tobranch_name, TblBranchesData_1.branch_namee AS tobranch_namee"
sql = sql & "  FROM         dbo.TblCustemers INNER JOIN"
sql = sql & "  dbo.TblStore INNER JOIN"
sql = sql & "  dbo.Transactions ON dbo.TblStore.StoreID = dbo.Transactions.StoreID ON dbo.TblCustemers.CusID = dbo.Transactions.CusID INNER JOIN"
sql = sql & "  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id INNER JOIN"
sql = sql & "  dbo.Transactions Transactions_2 ON dbo.Transactions.Transaction_ID = Transactions_2.ReturnID INNER JOIN"
sql = sql & "  dbo.TblStore TblStore_1 ON Transactions_2.StoreID = TblStore_1.StoreID INNER JOIN"
sql = sql & "  dbo.TblBranchesData TblBranchesData_1 ON TblStore_1.linked = TblBranchesData_1.branch_id LEFT OUTER JOIN"
sql = sql & "  dbo.Transactions Transactions_1 ON dbo.Transactions.NoteSerial1 = Transactions_1.order_no"
sql = sql & "  WHERE     (dbo.Transactions.Transaction_Type = 10) AND (Transactions_1.order_no IS NULL) AND (Transactions_1.Transaction_Type IS NULL) AND"
sql = sql & "  (Transactions_2.Transaction_Type = 11)"
 If X > 0 Then
 sql = sql & " and  (dbo.Transactions.Transaction_Date >=" & SQLDate(DateAdd("d", -1 * val(X), Date), True) & ")"
 End If
  
 

path = App.path & "\Reports\REPORTS NEW\SalesWithNoIsuueVchr1.rpt"
PrintSimpleReport sql, path, , CStr(X)

End Sub

Private Sub ALLButton22_Click()
FrmDffrentChargRequairAlrm.show

End Sub

Private Sub ALLButton23_Click()
FrmDiffrentReceptChargAlrm.show
End Sub

Private Sub ALLButton24_Click()
FrmInternalRequesAlarm.show
End Sub

Private Sub ALLButton25_Click()
FrmReceptRawMatrialsAlarm.show
End Sub

Private Sub ALLButton26_Click()
FrmAlarmReceptNoInBillBuy.show
End Sub

Private Sub ALLButton27_Click()
FrmAlarmQauntety.show
End Sub

Private Sub ALLButton28_Click()
FrmAlarmRequiredMaintain.show
End Sub

Private Sub ALLButton29_Click()
   If checkApility("FrmInstallmentVendorAlarm") = False Then
        Exit Sub
    End If

FrmInstallmentVendorAlarm.show
FrmInstallmentVendorAlarm.TabMain.CurrTab = 1
End Sub

Private Sub ALLButton3_Click()

    If checkApility("FrmInstallmentMustPay") = False Then
        Exit Sub
    End If

    FrmInstallmentMustPay.show
    FrmInstallmentMustPay.ZOrder 0
End Sub

Private Sub ALLButton30_Click()
FrmMoveAlarm.show
End Sub

Private Sub ALLButton31_Click()
FrmProjectAlarm.show
End Sub

Private Sub ALLButton32_Click()
FrmAlarmPurchaseOrders.show
End Sub

Private Sub ALLButton33_Click()
FrmMaintainanceAlarm.show

End Sub

Private Sub ALLButton34_Click()
   If checkApility("FrmGuaranteeAlram") = False Then
        Exit Sub
    End If
FrmGuaranteeAlram.Ind = 1
    OpenScreen PopUpShowGuaranteeAlram

End Sub

Private Sub ALLButton35_Click()
  AskOption = GetSetting(StrAppRegPath, "View_Type", "RentInstallments", True)
    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_RentInstallments", "")
    Askcount = GetSetting(StrAppRegPath, "Setting", "Count_RentInstallments", 0)
    
    If AskOption = True And Askinterval <> "" Then
    rentInstallmentdate = DateAdd((Askinterval), 1 * Askcount, Date)
    
    End If
    RSRentAlarm.mIndex = 0
    RSRentAlarm.show
End Sub

Private Sub ALLButton36_Click()
FrmRsContractAlarm.show
End Sub

Private Sub ALLButton37_Click()
frmAqarInstallAlert.show
End Sub

Private Sub ALLButton38_Click()
RSMentnanceAlarm.show
End Sub

Private Sub ALLButton39_Click()

Unload FrmCarExpireLicens
FrmCarExpireLicens.Indx = 1
FrmCarExpireLicens.show

End Sub

Private Sub ALLButton4_Click()

    If checkApility("all_alarms") = False Then
        Exit Sub
    End If

    all_alarms.show
End Sub

Private Sub ChangeLang()
    Me.Caption = "Today Alarms"
    Label1(2).Caption = Me.Caption
    Label2t.Caption = "Total No of Expired Residence"
    Label4t.Caption = "Total No of Expired Passport"
    Label5t.Caption = "Total No of Expired License"
 '  Label7t.Caption = "Total No of Expired ID"
'    Label9t.Caption = "Total No of Expired Insurance"
    ALLButton1t.Caption = "View"
    ALLButton2t.Caption = "View"
    ALLButton3t.Caption = "View"
    Label56.Caption = "Contracting Alarms"
    Label56.Caption = "Production Alarms"
    Label55.Caption = "Production Alarm"
    Label57.Caption = "Internal Order Alarm"
    ALLButton40.Caption = "View"
    ALLButton42.Caption = "View"
    ALLButton41.Caption = "View"
    ALLButton40.Caption = "View"
    
'    ALLButton4t.Caption = "View"
'    ALLButton5t.Caption = "View"

C1Tab1.TabCaption(0) = "Doc To Approve"
C1Tab1.TabCaption(1) = "Accounts "
C1Tab1.TabCaption(2) = "Fixed Assets"
C1Tab1.TabCaption(3) = "HR"
C1Tab1.TabCaption(4) = "Stores"
C1Tab1.TabCaption(5) = "Purchases"
C1Tab1.TabCaption(6) = "Sales"
C1Tab1.TabCaption(7) = "Installments"
C1Tab1.TabCaption(8) = "Projects"
C1Tab1.TabCaption(9) = "Real Estates"
C1Tab1.TabCaption(10) = "Transportation"
C1Tab1.TabCaption(11) = "Shimpments"
C1Tab1.TabCaption(12) = "Maintenance"

C1Tab1.TabCaption(13) = "Project2"
C1Tab1.TabCaption(14) = "Production"

Label10.Caption = "Reservations Vchr. Alarms"

Label21.Caption = "Accounts Alarms"
Label22.Caption = "Stock Alarms"
ALLButton32.Caption = "View"
Label23.Caption = "Sales Alarms"
Label310.Caption = "Installments Alarms"

Label25.Caption = "Maintenance Alarms"
Label47.Caption = "PO Alarmas"
Label28.Caption = "HR Alarms"
Label46.Caption = "Operations Achivement  Alarms"
Label26.Caption = "Projects Alarms"
Label43.Caption = "Fixed Asset Alarms"
Label48.Caption = "Maintenance and Guarantee Contracts"
Label44.Caption = "Fixed Asset  Installments Alarms"
ALLButton21.Caption = "View"
ALLButton22.Caption = "View"
ALLButton23.Caption = "View"
ALLButton31.Caption = "View"
ALLButton33.Caption = "View"

ALLButton24.Caption = "View"
ALLButton25.Caption = "View"
ALLButton26.Caption = "View"
ALLButton27.Caption = "View"
ALLButton28.Caption = "View"
ALLButton29.Caption = "View"

ALLButton30.Caption = "View"

Label37.Caption = "Internal Order Alarm"
Label38.Caption = "Recive Voucher Alarm"
Label41.Caption = "Recerved Qty Alarm"
Label45.Caption = "Moving Alarm"

Label39.Caption = "Purchase Alarms"
Label40.Caption = "Recive Vchr Without Invoice"

Label33.Caption = "Moving Voucher  Without Invoice"
Label34.Caption = "Shipping"
Label42.Caption = "Opening Work order"

Label35.Caption = "Different Between order Qty and Shipped"
Label36.Caption = "Different Between Shipped Qty and Recived"


Label27.Caption = "RealState Alarms"
Label29.Caption = "Transportation Alarms"
Label30.Caption = "Pending Approval Doc Alarms"
ALLButton34.Caption = "View"
Label20.Caption = "LC Alarms"
   Label19.Caption = "Pending Doc To Approval"
    Label14.Caption = "Items"
    Label15.Caption = "Items that will end guaranteed"
    Label49.Caption = "Items that will Expire "
    
    Label16.Caption = "Preventive maintenance alerts"
    Label17.Caption = "Overdue installment"
    Me.Caption = "Today Alarms"
    Label1(2).Caption = Me.Caption
    Label3.Caption = "Projects Variances "
    Label6.Caption = "Payable Projects Invoices"
    ALLButton6.Caption = "View"
    ALLButton7.Caption = "View"
    ALLButton10.Caption = "View"
    ALLButton11.Caption = "View"
    ALLButton18.Caption = "View"
    
     ALLButton18.Caption = "View"
      ALLButton19.Caption = "View"
      Label24.Caption = "Sales Invoice  have't Issue Vchr"
            ALLButton20.Caption = "View"

    
'    Label13.Caption = "Transportation Alarms"
'    ALLButton12.Caption = "View"
    Label12.Caption = "Approved P.Os Have't Sales Inv"
    Label2.Caption = "Financial Outstanding"
    Label11.Caption = "Accounts Distribution"
    Label5.Caption = "Due installment"
    Label4.Caption = "Items Request"
    Label7.Caption = "HR"
    Label8.Caption = " Credit Limit Alarms"
'    Label10.Caption = " Real Estate Mangement"
    Label18.Caption = "Main. Equipment ready for delivery"
    ALLButton1.Caption = "View"
    ALLButton2.Caption = "View"
    ALLButton3.Caption = "View"
    ALLButton4.Caption = "View"
 
    ALLButton8.Caption = "View"
'    ALLButton9.Caption = "View"
    ALLButton10.Caption = "View"
    ALLButton11.Caption = "View"
'    ALLButton12.Caption = "View"
    ALLButton13.Caption = "View"
    ALLButton14.Caption = "View"
    ALLButton15.Caption = "View"
    ALLButton16.Caption = "View"
    ALLButton17.Caption = "View"

End Sub

Private Sub ALLButton40_Click()
'   If checkApility("FrmInstallmentVendorAlarm") = False Then
'        Exit Sub
'    End If

FrmInstallmentVendorAlarm.show
FrmInstallmentVendorAlarm.TabMain.CurrTab = 2
If SystemOptions.UserInterface = ArabicInterface Then
FrmInstallmentVendorAlarm.EleHeader.Caption = " ‰»ÌÂ«  «·«‰ «Ã"
Else
FrmInstallmentVendorAlarm.EleHeader.Caption = "Production Alarms"
End If
FrmInstallmentVendorAlarm.Caption = FrmInstallmentVendorAlarm.EleHeader.Caption


End Sub

Private Sub ALLButton41_Click()
   If checkApility("FrmInstallmentVendorAlarm") = False Then
        Exit Sub
    End If

FrmInstallmentVendorAlarm.show
FrmInstallmentVendorAlarm.TabMain.CurrTab = 0
If SystemOptions.UserInterface = ArabicInterface Then
FrmInstallmentVendorAlarm.EleHeader.Caption = " ‰»ÌÂ«  «·« ›«ﬁÌ« "
Else
FrmInstallmentVendorAlarm.EleHeader.Caption = "Contracting Alarms"
End If
FrmInstallmentVendorAlarm.Caption = FrmInstallmentVendorAlarm.EleHeader.Caption
End Sub

Private Sub ALLButton42_Click()
RSRentAlarm.mIndex = 1
RSRentAlarm.show

End Sub

Private Sub ALLButton5_Click()
RSArbonAlarm.show
End Sub

Private Sub ALLButton6_Click()

    If checkApility("ProjectsAlarm1") = False Then
        Exit Sub
    End If

    ProjectsAlarm1.show
End Sub

Private Sub ALLButton7_Click()

    If checkApility("ProjectsBillAlarm") = False Then
        Exit Sub
    End If

    ProjectsBillAlarm.show
End Sub

Private Sub ALLButton8_Click()

    If checkApility("ArrowsFollowAlarm") = False Then
        Exit Sub
    End If

    'ArrowsFollowAlarm.show
    
            Ageng_all.show
End Sub

Private Sub ALLButton9_Click()

    If checkApility("FrmGuaranteeAlram") = False Then
        Exit Sub
    End If
FrmGuaranteeAlram.Ind = 2
    OpenScreen PopUpShowGuaranteeAlram
End Sub

Private Sub Form_Load()

    Me.Height = 10000
    Me.Width = 17600
    
    'Me.left = (mdifrmmain.Width - Me.Width) / 2
    'Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
        Me.Left = (mdifrmmain.Width - Me.Width) / 2 - 1200
 Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    C1Tab1.CurrTab = 0
    
        Me.Left = 0 '(mdifrmmain.Width - Me.Width) / 2
    Me.Top = -100 '(mdifrmmain.Height - Me.Height) / 2 - 500

    Me.Width = (mdifrmmain.Width) - 500
    Me.Height = (mdifrmmain.Height) - 600
    
'   Label31.Caption = Format(Date, "YYYY-mm-DD")
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    'SkinFramework1.ApplyWindow Me.hWnd
    ' SkinFramework1.LoadSkin App.path & "\style\Vista.cjstyles", ""
 
    '  Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_Expirepas", "D")
    '  Askcount = GetSetting(StrAppRegPath, "Setting", "count_Expirepas", 0)
   
    
'         C1Tab1.TabVisible(0) = False
         C1Tab1.TabVisible(1) = False
         C1Tab1.TabVisible(2) = False
         C1Tab1.TabVisible(3) = False
         C1Tab1.TabVisible(4) = False
         C1Tab1.TabVisible(5) = False
         C1Tab1.TabVisible(6) = False
         C1Tab1.TabVisible(7) = False
         C1Tab1.TabVisible(8) = False
         C1Tab1.TabVisible(9) = False
         C1Tab1.TabVisible(10) = False
         C1Tab1.TabVisible(11) = False
         C1Tab1.TabVisible(12) = False
         
 
'If mdifrmmain.MnuAccounts.Visible = True Then
C1Tab1.TabVisible(1) = True
'CAccount.Visible = False
'End If


If mdifrmmain.StockControl.Visible = True Then
C1Tab1.TabVisible(4) = True
'CStores.Visible = False
End If


If mdifrmmain.Purchase.Visible = True Then
C1Tab1.TabVisible(5) = True
'CPurchase.Visible = False
End If

 

If mdifrmmain.Sales.Visible = True Then
C1Tab1.TabVisible(6) = True
'CSAles.Visible = False
End If

If mdifrmmain.SalesIns.Visible = True Then
C1Tab1.TabVisible(7) = True
End If

If mdifrmmain.MNUFixedAssets.Visible = True Then
C1Tab1.TabVisible(2) = True
'CFixed.Visible = False
End If

 

If mdifrmmain.mnuEmployee.Visible = True Then
C1Tab1.TabVisible(3) = True
'CHR.Visible = False
End If


 If mdifrmmain.MnuProjects.Visible = True Then
 C1Tab1.TabVisible(8) = True
'CProjects.Visible = False
End If



 If mdifrmmain.AssetsMngBase.Visible = True Then
 C1Tab1.TabVisible(9) = True
'CRS.Visible = False
End If

 If mdifrmmain.TransporterMain.Visible = True And mdifrmmain.MnuMaintnance.Visible = True Then
'Ctransportation.Visible = False
 C1Tab1.TabVisible(10) = True
  C1Tab1.TabVisible(12) = True
  
End If


 If mdifrmmain.shipmentMnu.Visible = True Then
'CShipments.Visible = False
 C1Tab1.TabVisible(11) = True
 
End If

 
 If mdifrmmain.MnuMaintnance.Visible = True Then
'CMaintenance.Visible = False
 C1Tab1.TabVisible(12) = True
End If



 If mdifrmmain.prdo.Visible = False Then '«‰ «Ã
 
End If

 If 1 = 0 Then '«⁄ „«œ
 C1Tab1.TabVisible(0) = True
End If
'****************************************************************************
'transportation
Set rs = New ADODB.Recordset
    Dim Dcombos As New ClsDataCombos
  Dcombos.GetStoresByUser Me.DCboStore2Name, , CInt(user_id)
  Dcombos.GetBranches Me.Dcbranch
  
'****************************************************************************
End Sub

Private Sub ALLButton2t_Click()
 
    FrmCarExpireTest.show
End Sub

Private Sub ALLButton3t_Click()
    FrmCarExpireInsurance.show
End Sub

 Private Sub Timer1_Timer()

    If Label1(2).ForeColor = &HFFFFFF Then
        Label1(2).ForeColor = &HFF&
    Else
        Label1(2).ForeColor = &HFFFFFF
    End If

End Sub

