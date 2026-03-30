VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDocApproval 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈⁄œ«œ«  «⁄ „«œ «·‘«‘«        "
   ClientHeight    =   8760
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   14745
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   580
   Icon            =   "frmDocApproval.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   14745
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8760
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14745
      _cx             =   26009
      _cy             =   15452
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmDocApproval.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   7725
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   14685
         _cx             =   25903
         _cy             =   13626
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   12648447
         ForeColor       =   128
         FrontTabColor   =   14871017
         BackTabColor    =   8454143
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "."
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7305
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   14595
            _cx             =   25744
            _cy             =   12885
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   7155
               Index           =   1
               Left            =   0
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   0
               Width           =   14745
               _cx             =   26009
               _cy             =   12621
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
               Begin VB.TextBox txtApprovNameE 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   2880
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   1800
                  Width           =   3885
               End
               Begin VB.TextBox txtApprovName 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   9045
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   1800
                  Width           =   3885
               End
               Begin VB.Frame Frame17 
                  Caption         =   "Êﬁ  «·«⁄ „«œ «·„ Êﬁ⁄"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   615
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   960
                  Width           =   2055
                  Begin VB.ComboBox CboTimeCateg 
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     ItemData        =   "frmDocApproval.frx":040F
                     Left            =   120
                     List            =   "frmDocApproval.frx":041C
                     TabIndex        =   65
                     Top             =   240
                     Width           =   855
                  End
                  Begin VB.TextBox TxTtimeCount 
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   64
                     Top             =   240
                     Width           =   615
                  End
               End
               Begin VB.Frame Frame2 
                  Caption         =   "ÿ—Ìﬁ… «·«—”«·"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   1245
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   -1320
                  Visible         =   0   'False
                  Width           =   5070
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     Caption         =   "ÌœÊÌ"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   0
                     Left            =   3720
                     RightToLeft     =   -1  'True
                     TabIndex        =   51
                     Top             =   240
                     Width           =   1215
                  End
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·Ì« ⁄‰œ «·Õ›Ÿ"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   1
                     Left            =   2280
                     RightToLeft     =   -1  'True
                     TabIndex        =   50
                     Top             =   240
                     Width           =   1215
                  End
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·Ì« ⁄‰œ «·ÿ»«⁄Â"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   2
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   49
                     Top             =   240
                     Width           =   1335
                  End
               End
               Begin VB.TextBox TxtRowNumber 
                  Alignment       =   1  'Right Justify
                  Height          =   495
                  Left            =   600
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Text            =   "0"
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   1020
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«ŸÂ«— ﬂ· «·„ÊŸ›Ì‰"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   2070
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   -1230
                  Visible         =   0   'False
                  Width           =   2700
               End
               Begin VB.Frame Frame1 
                  Caption         =   "«·„” ÊÌ« "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   1065
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   2100
                  Width           =   14355
                  Begin VB.TextBox TxtValue1 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   3120
                     RightToLeft     =   -1  'True
                     TabIndex        =   62
                     Text            =   "0"
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   495
                  End
                  Begin VB.TextBox TxtSearchCode 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   9480
                     TabIndex        =   55
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   1050
                  End
                  Begin VB.TextBox TxtValue 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   4200
                     RightToLeft     =   -1  'True
                     TabIndex        =   57
                     Text            =   "0"
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   855
                  End
                  Begin VB.TextBox TxtRemarks 
                     Alignment       =   1  'Right Justify
                     Height          =   615
                     Left            =   960
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   59
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1455
                  End
                  Begin VB.OptionButton Option1 
                     Alignment       =   1  'Right Justify
                     Caption         =   "ﬂ· «·„ÊŸ›Ì‰"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Left            =   11640
                     RightToLeft     =   -1  'True
                     TabIndex        =   26
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1575
                  End
                  Begin VB.OptionButton Option2 
                     Alignment       =   1  'Right Justify
                     Caption         =   "„ﬂÊ‰«  «·—”«·…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Left            =   10440
                     RightToLeft     =   -1  'True
                     TabIndex        =   25
                     Top             =   240
                     Value           =   -1  'True
                     Visible         =   0   'False
                     Width           =   1455
                  End
                  Begin MSDataListLib.DataCombo DCEmployee 
                     Height          =   315
                     Left            =   6960
                     TabIndex        =   56
                     Top             =   600
                     Width           =   5790
                     _ExtentX        =   10213
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
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   390
                     Index           =   20
                     Left            =   6120
                     TabIndex        =   58
                     Top             =   480
                     Width           =   720
                     _ExtentX        =   1270
                     _ExtentY        =   688
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "≈÷«›…"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "frmDocApproval.frx":0432
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " œﬁÌﬁ…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Index           =   10
                     Left            =   3000
                     RightToLeft     =   -1  'True
                     TabIndex        =   61
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   1155
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„·«ÕŸ« "
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Index           =   9
                     Left            =   2280
                     RightToLeft     =   -1  'True
                     TabIndex        =   47
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   795
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " «·ﬁÌ„…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   465
                     Index           =   4
                     Left            =   5160
                     RightToLeft     =   -1  'True
                     TabIndex        =   46
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   795
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Õœœ „” ÊÌ«  «·«⁄ „«œ"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   3
                     Left            =   12795
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   600
                     Width           =   1500
                  End
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   11460
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   1080
                  Width           =   1470
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  Height          =   465
                  Left            =   285
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   2850
                  Visible         =   0   'False
                  Width           =   375
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   3315
                  Left            =   135
                  TabIndex        =   7
                  Top             =   3240
                  Width           =   14415
                  _cx             =   25426
                  _cy             =   5847
                  Appearance      =   2
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   24
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmDocApproval.frx":07CC
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   915
                  Index           =   5
                  Left            =   -510
                  TabIndex        =   8
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   15165
                  _cx             =   26749
                  _cy             =   1614
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
                  Picture         =   "frmDocApproval.frx":0B6A
                  Caption         =   "≈⁄œ«œ«  «⁄ „«œ «·‘«‘«        "
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
                  PicturePos      =   0
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
                     Left            =   1695
                     TabIndex        =   9
                     Top             =   90
                     Width           =   495
                     _ExtentX        =   873
                     _ExtentY        =   661
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
                     ButtonImage     =   "frmDocApproval.frx":1844
                     ColorButton     =   16777215
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
                     Left            =   630
                     TabIndex        =   10
                     Top             =   90
                     Width           =   495
                     _ExtentX        =   873
                     _ExtentY        =   661
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
                     ButtonImage     =   "frmDocApproval.frx":1BDE
                     ColorButton     =   16777215
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
                     Left            =   2220
                     TabIndex        =   11
                     Top             =   90
                     Width           =   495
                     _ExtentX        =   873
                     _ExtentY        =   661
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
                     ButtonImage     =   "frmDocApproval.frx":1F78
                     ColorButton     =   16777215
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
                     Left            =   1155
                     TabIndex        =   12
                     Top             =   90
                     Width           =   495
                     _ExtentX        =   873
                     _ExtentY        =   661
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
                     ButtonImage     =   "frmDocApproval.frx":2312
                     ColorButton     =   16777215
                     ColorHighlight  =   4194304
                     ColorHoverText  =   16777215
                     ColorShadow     =   -2147483631
                     ColorOutline    =   -2147483631
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                     ColorToggledHoverText=   16777215
                     ColorTextShadow =   16777215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0FFFF&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "Â–… «·‘«‘…  ﬁÊ„ » ⁄—Ì› «·„” ÊÌ«  «·Œ«’… »«⁄ „«œ ﬂ· ‘«‘…"
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
                     Height          =   375
                     Index           =   6
                     Left            =   3840
                     TabIndex        =   70
                     Top             =   240
                     Width           =   4725
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   585
                  Index           =   3
                  Left            =   2505
                  TabIndex        =   13
                  TabStop         =   0   'False
                  Top             =   3030
                  Visible         =   0   'False
                  Width           =   5070
                  _cx             =   8943
                  _cy             =   1032
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
                  Caption         =   " Õœœ «·› —…"
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
                  Begin VB.ComboBox CboYear 
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   15
                     Top             =   165
                     Width           =   1005
                  End
                  Begin VB.ComboBox CmbMonth 
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   75
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   14
                     Top             =   180
                     Width           =   1485
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”‰…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Index           =   2
                     Left            =   2955
                     RightToLeft     =   -1  'True
                     TabIndex        =   17
                     Top             =   180
                     Width           =   1020
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Â—"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   0
                     Left            =   1425
                     RightToLeft     =   -1  'True
                     TabIndex        =   16
                     Top             =   270
                     Width           =   645
                  End
               End
               Begin MSComCtl2.DTPicker XPDtbTrans 
                  Height          =   345
                  Left            =   8535
                  TabIndex        =   22
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   1830
                  _ExtentX        =   3228
                  _ExtentY        =   609
                  _Version        =   393216
                  Format          =   236388353
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DCComponent 
                  Height          =   315
                  Left            =   9045
                  TabIndex        =   23
                  Top             =   1440
                  Width           =   3885
                  _ExtentX        =   6853
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   21
                  Left            =   12360
                  TabIndex        =   60
                  Top             =   6720
                  Width           =   1410
                  _ExtentX        =   2487
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " Õ–› ”ÿ—"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "frmDocApproval.frx":26AC
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo dcBranch 
                  Height          =   315
                  Left            =   6840
                  TabIndex        =   66
                  Top             =   1080
                  Width           =   3765
                  _ExtentX        =   6641
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
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
               Begin MSDataListLib.DataCombo DcboEmpDepartments 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   74
                  Top             =   1200
                  Width           =   3675
                  _ExtentX        =   6482
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«œ«—…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   14
                  Left            =   3360
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   1320
                  Width           =   1530
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«œ«—…"
                  Height          =   225
                  Index           =   7
                  Left            =   4680
                  TabIndex        =   75
                  Top             =   60
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·«⁄ „«œ «‰Ã·Ì“Ì"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   13
                  Left            =   6915
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   1800
                  Width           =   1890
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·«⁄ „«œ ⁄—»Ì"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   12
                  Left            =   12960
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   1800
                  Width           =   1530
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   11
                  Left            =   10680
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   1080
                  Width           =   705
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«Ã„«·Ì"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   6840
                  Visible         =   0   'False
                  Width           =   1215
               End
               Begin VB.Label LblSum 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
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
                  Height          =   255
                  Left            =   1200
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   6720
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.Label LBLWhereSTR 
                  Alignment       =   1  'Right Justify
                  Height          =   255
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   1740
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·Õ—ﬂ…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   5
                  Left            =   13035
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   1500
                  Width           =   1410
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ  "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Index           =   8
                  Left            =   10470
                  RightToLeft     =   -1  'True
                  TabIndex        =   21
                  Top             =   630
                  Visible         =   0   'False
                  Width           =   585
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   180
                  Index           =   7
                  Left            =   13440
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   1080
                  Width           =   750
               End
               Begin VB.Label Label5 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Height          =   285
                  Left            =   13305
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   1110
                  Width           =   495
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   90
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   960
         Left            =   30
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   7770
         Width           =   14685
         _cx             =   25903
         _cy             =   1693
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
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   11880
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   90
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
            ButtonImage     =   "frmDocApproval.frx":2C46
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   12765
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   225
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
            ButtonImage     =   "frmDocApproval.frx":2FE0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   13965
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   150
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   2
            Caption         =   ""
            BackColor       =   14871017
            FontSize        =   14.25
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "frmDocApproval.frx":337A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   11100
            TabIndex        =   36
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   873
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
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   1
            Left            =   10200
            TabIndex        =   37
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
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
            Height          =   495
            Index           =   2
            Left            =   9390
            TabIndex        =   38
            Top             =   480
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            CausesValidation=   0   'False
            Height          =   495
            Index           =   3
            Left            =   8235
            TabIndex        =   39
            Top             =   480
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            Height          =   495
            Index           =   4
            Left            =   7080
            TabIndex        =   40
            Top             =   480
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            CausesValidation=   0   'False
            Height          =   495
            Index           =   6
            Left            =   3360
            TabIndex        =   41
            Top             =   480
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            Height          =   495
            Index           =   5
            Left            =   5910
            TabIndex        =   42
            Top             =   480
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            Height          =   495
            Index           =   7
            Left            =   4560
            TabIndex        =   73
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "‰”Œ… „„«À·…"
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
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   225
            Width           =   1740
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   8640
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   3570
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   225
            Width           =   4695
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   10185
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   225
            Width           =   1455
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   4
      Top             =   6840
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "⁄—÷"
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
      ButtonImage     =   "frmDocApproval.frx":3714
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "frmDocApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double

Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long
Dim rs As ADODB.Recordset
Dim Msg  As String
Dim componentUnit As Integer

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
     On Error GoTo ErrTrap
    
    If txtid.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (txtid.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            StrSQL = "Delete From TblApprovalDefDetails Where lMessageDefID=" & val(Me.txtid.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
   
            If Not rs.RecordCount < 1 Then
                rs.delete
                rs.MoveFirst '

                If rs.RecordCount < 1 Then
                    clear_all Me
                    Grid.Clear flexClearScrollable, flexClearEverything
                    Grid.rows = 1
                    Grid.Enabled = False
                
                    TxtModFlg_Change
                    LabCurrRec.Caption = 0
                    LabCountRec.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub

Public Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2011 To 2050
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = IntDefIndex

End Sub

Private Sub ChkDetails_Click()
    FillGridWithData
End Sub

Private Sub ALLButton1_Click()
    FrmShowCol1.show
End Sub

Function check_previous_dev(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from notes where salary=" & year & Month
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev = False
    Else
        check_previous_dev = True
    End If
 
End Function

Function check_previous_dev1(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from salary_voucher where m_year='" & year & "' and m_month='" & Month & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev1 = False
    Else
        check_previous_dev1 = True
    End If
 
End Function

Function Create_dev()
  
End Function

Function Create_dev1()
    
End Function

Private Sub CboPayMentType_Click()

End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
 
Function create_report_data()

End Function

Private Sub CmdPrint_Click()
 
End Sub

Private Sub Combo1_Click()
 
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then
 
 
    If Trim(dcBranch.BoundText) = "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify Departement"
       Else
           Msg = " Õœœ ›—⁄ «Ê·« "
        End If

       MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        dcBranch.SetFocus
        Sendkeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
 
 
 
                    If Trim(Me.DCComponent.BoundText) = "" Then
                                        If SystemOptions.UserInterface = ArabicInterface Then
                                            Msg = "ÌÃ» ≈Œ Ì«— «·‘«‘…..!!"
                                        Else
                                            Msg = "Must Select Screen    ..!!"
                                        End If
            
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        DCComponent.SetFocus
                        Sendkeys "{F4}"
                        Exit Sub
                    End If
 
 
 
         If Trim(Me.CboTimeCateg.ListIndex) = -1 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ» ≈Œ Ì«—   «·„œ… ..!!"
                        Else
                            Msg = "Must Select period    ..!!"
                        End If
            
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        CboTimeCateg.SetFocus
                        Sendkeys "{F4}"
                        Exit Sub
        End If
        
        
        
                If val(TxTtimeCount.text) = 0 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»  ÕœÌœ   «·„œ… ..!!"
                        Else
                            Msg = "Must enter period    ..!!"
                        End If
            
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        TxTtimeCount.SetFocus
                       ' SendKeys "{F4}"
                        Exit Sub
        End If
        
        
    End If
 
 

 
 
    
    With Me.Grid

        If .rows = .FixedRows Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— „” ÊÏ  Ê«Õœ ⁄·Ï «·«ﬁ· ..!!"
            Else
                Msg = "Must Select one Level At Least    ..!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
                         
    End With
 
    '-------------------------------------------------------------------------------------------
    Dim EmployeeSalary As Double
    Dim NoOfHours As Double
    Dim NoOfMinutes As Double
    Dim cProgress As ClsProgress
    Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
    cProgress.StartProgress

    DoEvents
 
    Dim i As Long
    'Check

    Cn.BeginTrans
    BeginTrans = True
    
    If TxtModFlg.text = "N" Then
    Me.txtid.text = CStr(new_id("TblApprovalDef", "id", "", True))
        rs.AddNew
        rs("id").value = val(Me.txtid.text)
    ElseIf Me.TxtModFlg.text = "E" Then
        StrSQL = "Delete From TblApprovalDefDetails Where lMessageDefID=" & val(Me.txtid.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If
 
    rs("screenName").value = (DCComponent.BoundText)
 rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
 
     rs("timeCount").value = (TxTtimeCount.text)
     rs("TimeCateg").value = (CboTimeCateg.ListIndex)
   rs("ApprovName").value = (txtApprovName.text)
rs("ApprovNamee").value = (txtApprovNameE.text)

      If val(Me.DcboEmpDepartments.BoundText) = 0 Then
            rs("DepartmentID").value = Null
        Else
            rs("DepartmentID").value = val(Me.DcboEmpDepartments.BoundText)
        End If
        
    rs.update
    
    Dim IntDEV_Type As Integer
    Dim SngDEV_Value As Single
         
    Dim RsSerial As ADODB.Recordset
 
    Dim LngSerialCount As Long
 
    Set RsDev = New ADODB.Recordset
    RsDev.Open "TblApprovalDefDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    With Me.Grid

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("MessageDefID")) <> "" Then

                RsDev.AddNew
 
                RsDev("lMessageDefID").value = val(txtid.text)
  
                RsDev("PlainMessageID").value = .TextMatrix(i, .ColIndex("MessageDefID"))
     
                RsDev.update
                    
            End If
       
            '        End If
        Next i

    End With
     
    Cn.CommitTrans
    BeginTrans = False
    '    XPTxtCurrent.Caption = rs.AbsolutePosition
    '    XPTxtCount.Caption = rs.RecordCount

    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    
    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Operation Saved Successfully " & CHR(13)
                Msg = Msg + "Do You Want New Operation"
            End If
        
            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If

            '  Fg_Journal.Enabled = False
    End Select

    TxtModFlg.text = "R"
    'End If

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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub FocusGridOnCell(LngRow As Long, _
                            LngCol As Long)
    On Local Error GoTo ErrTrap

    With Me.Grid
        .row = LngRow
        .Col = LngCol
        .ShowCell LngRow, LngCol
        .SetFocus
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Cmd_Click(index As Integer)

    ' On Error GoTo ErrTrap
    Select Case index

        Case 0
            'If DoPremis(Do_New, Me.name, True) = False Then
            '    Exit Sub
            'End If
            TxtModFlg.text = "N"
            clear_all Me
           ' Me.txtid.Text = CStr(new_id("TblApprovalDef", "id", "", True))
        
            ' Me.DCboUserName.BoundText = user_id
            XPDtbTrans.value = Date
       
            '        XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.rows = 1
            Grid.Enabled = True
            Opt(0).value = True

            'Me.DcBranch.BoundText = branch_id
        Case 1
            '  If DoPremis(Do_Edit, Me.name, True) = False Then
            '      Exit Sub
            '  End If
            TxtModFlg.text = "E"
            '  Me.DCboUserName.BoundText = user_id
        
            '  Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True

        Case 2
     
            SaveData
           
        Case 3
            Undo

        Case 4
            'If DoPremis(Do_Delete, Me.name, True) = False Then
            '    Exit Sub
            'End If
            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            Load FrmNotesSearch
            FrmNotesSearch.SearchType = 3
            FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7
        txtid.text = ""
        TxtModFlg.text = "N"
            Grid.rows = Grid.rows + 1
            Grid.Enabled = True
            Cmd(2).Enabled = True
            '   ViewDataList
    
        Case 20
            addrow

        Case 21
            RemoveGridRow
        
        
    
    End Select

    Exit Sub
ErrTrap:

End Sub
 
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Function SHow_grig_col()
 
End Function

Function ChangeGridView(componentUnit As Integer)

    With Grid

        Select Case componentUnit

            Case 0 'ﬁÌ„…

                .ColHidden(.ColIndex("HourRate")) = True
                .ColHidden(.ColIndex("NoofDays")) = True
                .ColHidden(.ColIndex("NoOfMinutes")) = True
                .ColHidden(.ColIndex("NoOfHour")) = True
                lbl(4).Caption = "ﬁÌ„…"
                lbl(10).Visible = False
                TxtValue1.Visible = False

            Case 1 '«Ì«„
                .ColHidden(.ColIndex("NoofDays")) = False
                 
                .ColHidden(.ColIndex("HourRate")) = False
                
                .ColHidden(.ColIndex("NoOfMinutes")) = True
                .ColHidden(.ColIndex("NoOfHour")) = True
             
                lbl(4).Caption = "⁄œœ «·«Ì«„"
              
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(0, .ColIndex("HourRate")) = "„⁄œ· «·ÌÊ„ "
                Else
                    .TextMatrix(0, .ColIndex("HourRate")) = "Day Rate "
                End If

                lbl(10).Visible = False
                TxtValue1.Visible = False

            Case 2 '”«⁄« 
                .ColHidden(.ColIndex("NoofDays")) = True
                 
                .ColHidden(.ColIndex("HourRate")) = False
                
                .ColHidden(.ColIndex("NoOfMinutes")) = False
                .ColHidden(.ColIndex("NoOfHour")) = False
                lbl(4).Caption = "⁄œœ «·”«⁄« "
                lbl(10).Visible = True
                TxtValue1.Visible = True
              
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(0, .ColIndex("HourRate")) = "„⁄œ· «·”«⁄Â "
                Else
                    .TextMatrix(0, .ColIndex("HourRate")) = "Hour Rate "
                End If
 
        End Select

    End With

End Function

Private Sub DcEmployee_Click(Area As Integer)

    If val(DCEmployee.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DCEmployee.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
 
End Sub

Private Sub DCEmployee_KeyDown(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyReturn Then
   
        Sendkeys "{TAB}"
         
    End If

End Sub

Private Sub DCEmployee_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 1
        FrmEmployeeSearch.show
  
    End If

End Sub

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    With Grid

        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpPictureAlignment, 0, 0, .rows - 1, .Cols - 1) = flexPicAlignRightCenter
            .ColComboList(.ColIndex("Unit")) = "#1;ﬁÌ„…|#2;«Ì«„|#3;”«⁄« "
        Else
            .ColComboList(.ColIndex("Unit")) = "#1;Value|#2;Days|#3;Hours"
        End If
    
    End With



    Dim My_SQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "SELECT distinct  ScreenName,ScreenCaption From Screens where FlgShow=1  order by ScreenCaption "
    Else
        My_SQL = "SELECT  distinct  ScreenName,ScreenTitleEng From Screens where FlgShow=1  order by ScreenTitleEng "
    End If

    fill_combo DCComponent, My_SQL

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "SELECT LevelID,Name From TbLLevels order by Name "
    Else
        My_SQL = "SELECT LevelID,Name From TbLLevels  order by Name "
    End If

    fill_combo DCEmployee, My_SQL

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmpDepartments Me.DcboEmpDepartments
    Dcombos.GetBranches dcBranch
    Set cSearchDCombo = New clsDCboSearch

    Set BKGrndPic = New ClsBackGroundPic

    With Me.Grid
        .rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
        '    .WallPaper = BKGrndPic.Picture
        '    .AutoSize 0, .Cols - 1, False
    End With
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    YearMonth

    Dim StrSQL  As String
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblApprovalDef  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.text = "R"
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
 
End Sub

Private Sub ChangeLang()
 
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Same Copy"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"
    Opt(0).Caption = "Value"
    Opt(1).Caption = "Days"
    Opt(2).Caption = "Hours"
    Frame2.Caption = "Component Value"
    lbl(12).Caption = ""

    Frame1.Caption = "Select Employees"
    Option1.Caption = "All Employees"
    Option2.Caption = "Select Emp"
    lbl(3).Caption = "Select "
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "   Approval Settings"
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(8).Caption = "Date"
    Ele(3).Caption = "Select Interval"
    lbl(2).Caption = "Year"
    lbl(0).Caption = "Month"
 
    lbl(5).Caption = "Component Name"
    lbl(4).Caption = "Value"
    lbl(9).Caption = "Remark"

    Label2(0).Caption = "Current Record"
    Label2(2).Caption = "Total Record"
    Cmd(20).Caption = "Add"
    Cmd(21).Caption = "Remove"
    lbl(12).Caption = "Arabic Approved Name"
    lbl(13).Caption = "English Approved Name"
    lbl(11).Caption = "Branch"
    Frame17.Caption = "Approved Time"
    lbl(6).Caption = "This screen Define Approval for Transactions"

    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("MessageDefName")) = "Name"
        .TextMatrix(0, .ColIndex("MessageDefID")) = "Code"
 
    End With
With CboTimeCateg
         .Clear
        .AddItem "Minute"
        .AddItem "Hour"
        .AddItem "Day"

End With
Frame1.Caption = "Levels"
lbl(3).Caption = "Levels"

lbl(5).Caption = "Transaction"
lbl(12).Caption = "Ar. name"
lbl(13).Caption = "Eng. name"

lbl(7).Caption = "Id"
lbl(11).Caption = "Branch"
Frame17.Caption = "Exp. Time"



End Sub

Public Sub get_all_employee()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim j As Integer

    Dim sql As String
    Dim i As Integer

    sql = "Select * from TblEmployee "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With Grid

        .rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .rows = Rs3.RecordCount + 1
            Rs3.MoveFirst
         
            For i = 1 To Rs3.RecordCount
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(Rs3.Fields("Emp_ID").value), "", Rs3.Fields("Emp_ID").value)
                       
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
                Else
                    .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Namee").value), "", Rs3.Fields("Emp_Namee").value)

                End If
                       
                Rs3.MoveNext
            Next i
            
            '.Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
            ' .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
            ' .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 10
            ' .Cell(flexcpFontName, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Tahoma"
            .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

End Sub

Public Sub FillGridWithData()
    Exit Sub

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim LstDay As Date
    Dim FrstDay As Date
    Dim StrTxt As String
    Dim My_SQL As String
    Dim StrWhere As String
    Dim StrGrp As String
    Dim IntMonth As Integer
    Dim IntYear As Integer
    Dim Msg As String

    On Error GoTo ErrTrap
    'If DateDiff("d", Me.DtpFrom.Value, Me.DtpTO.Value, vbSaturday) < 0 Then
    '    Exit Sub
    'End If
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    If Me.CmbMonth.ListIndex = -1 Then Exit Sub
    If Me.CboYear.ListIndex = -1 Then Exit Sub

    'If Val(Me.TxtMonthHours.text) = 0 Then
    '    Msg = "ÌÃ» ≈œŒ«· ⁄œœ ”«⁄«  «·⁄„· ·Â–« «·‘Â—"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    Exit Sub
    'End If
    IntYear = val(Me.CboYear.text)
    IntMonth = Me.CmbMonth.ListIndex + 1

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        Dim ID As String
        My_SQL = " Select Emp_ID,Emp_Code,Emp_Name,DepartmentID,project_id "
        My_SQL = My_SQL + ",IsNUll(Emp_Salary,0)as Emp_Salary,IsNUll(Emp_Salary_sakn,0)as Emp_Salary_sakn,IsNUll(Emp_Salary_bus,0)as Emp_Salary_bus,IsNUll(Emp_Salary_food,0)as Emp_Salary_food,IsNUll(Emp_Salary_others,0)as Emp_Salary_others,IsNUll(Emp_Salary_mob,0)as Emp_Salary_mob,IsNUll(Emp_Salary_mang,0)as Emp_Salary_mang,  "
        My_SQL = My_SQL + "IsNUll( TotalDiscount,0)as TotalDiscount,"
        My_SQL = My_SQL + "IsNUll(TotalMokafea, 0) As TotalMokafea"
        My_SQL = My_SQL + ""
        My_SQL = My_SQL + ",(IsNUll(Emp_Salary,0)+IsNUll( TotalMokafea,0))-"
        My_SQL = My_SQL + "(IsNUll(TotalDiscount,0)) as EmpTotalNet "
    
        My_SQL = My_SQL + " From "
        My_SQL = My_SQL + "("
        My_SQL = My_SQL + "SELECT TOP 100 PERCENT  dbo.TblEmployee.project_id, dbo.TblEmployee.DepartmentID , dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang,"
        My_SQL = My_SQL + "dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary,"
        My_SQL = My_SQL + "SUM(QryAllDiscountWithMkafea.TotalDiscount) AS TotalDiscount,"
        My_SQL = My_SQL + "SUM(QryAllDiscountWithMkafea.Mokafea) AS TotalMokafea"
        My_SQL = My_SQL + ""
    
        My_SQL = My_SQL + " From dbo.QryAllDiscountWithMkafea(" & IntMonth & "," & IntYear & ")"
        My_SQL = My_SQL + " QryAllDiscountWithMkafea RIGHT OUTER JOIN"
        My_SQL = My_SQL + " dbo.TblEmployee ON QryAllDiscountWithMkafea.Emp_ID = dbo.TblEmployee.Emp_ID"
    
        'If Dcemp.text <> "" Then
        'My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and dbo.TblEmployee.emp_code='" & Dcemp.BoundText & "'"
        'Else
        'If Dcdep.text <> "" Then
        '
        '        If dcproject.BoundText = "" Then
        '        My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and dbo.TblEmployee.DepartmentID='" & Dcdep.BoundText & "'"
        '        Else
        '         My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and dbo.TblEmployee.DepartmentID='" & Dcdep.BoundText & "' and dbo.TblEmployee.project_id='" & Me.dcproject.BoundText & "'"
        '        End If
        'Else
        '    If Dcdep.text = "" Then
    
        '             If dcproject.BoundText <> "" Then
        '
        '              My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and  dbo.TblEmployee.project_id='" & Me.dcproject.BoundText & "'"
        '              Else
        '              My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1"
        '             End If
    
        ' Else
    
        ' My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1"
        ' End If
        ' End If
        ' End If
    
        My_SQL = My_SQL + " GROUP BY dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code,dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others,dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang,"
        My_SQL = My_SQL + " dbo.TblEmployee.Emp_Salary,dbo.TblEmployee.DepartmentID ,dbo.TblEmployee.project_id"
        My_SQL = My_SQL + " ORDER BY dbo.TblEmployee.Emp_ID"
    
        My_SQL = My_SQL + ")XTable"
    Else
        FrstDay = "1-" & CmbMonth.ListIndex + 1 & "-" & year(Date)
        LstDay = DateAdd("d", -1, "1-" & CmbMonth.ListIndex + 2 & "-" & year(Date))

        My_SQL = "select Emp_ID,Emp_Name,Emp_Salary ,sum(TotalDiscount) as TotalDiscount," & "sum(Mokafea) as Mokafea  From QryEmpAllValues where TransDate >=#" & Format(FrstDay, "mm/dd/yyyy") & "# and TransDate<=#" & Format(LstDay, "mm/dd/yyyy") & "# " & StrWhere & " GROUP BY Emp_ID, Emp_Name, " & "Emp_Salary  "
    End If

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Ser")) = i
                ',DepartmentID,project_id
            
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
            
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
               
                .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), "", rs.Fields("Emp_Salary").value)
            
                .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                
                .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Format(rs.Fields("TotalMokafea").value, SystemOptions.SysDefCurrencyForamt))
            
                '.TextMatrix(I, .ColIndex("TotalAdvance")) = IIf(IsNull(Rs.Fields("TotalAdvance").Value), _
                 "", Format(Rs.Fields("TotalAdvance").Value, SystemOptions.SysDefCurrencyForamt))
           
                '   .TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(Rs.Fields("EmpTotalNet").value), _
                '      "", Format(Rs.Fields("EmpTotalNet").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), "", Format(rs.Fields("Emp_Salary_sakn").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), "", Format(rs.Fields("Emp_Salary_bus").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), "", Format(rs.Fields("Emp_Salary_food").value))
                               
                .TextMatrix(i, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), "", Format(rs.Fields("Emp_Salary_mob").value))
                                    
                .TextMatrix(i, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), "", Format(rs.Fields("Emp_Salary_mang").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), "", Format(rs.Fields("Emp_Salary_others").value))
            
                rs.MoveNext
            
            Next

            rs.Close
        End If

        GetAdvanceValues IntMonth, IntYear
        GetWorkHours
        CalculateNets
        .rows = .rows + 1
        .TextMatrix(.rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        .IsSubtotal(.rows - 1) = True
        Dim SngTotal As Single
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .rows - 1, .ColIndex("Emp_Salary"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .rows - 1, .ColIndex("EmpTotalNet"))
        .TextMatrix(.rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
        net_value = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .rows - 1, .ColIndex("CorrectEmpTotalNet"))
        .TextMatrix(.rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .rows - 1, .ColIndex("Emp_Salary_sakn"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .rows - 1, .ColIndex("Emp_Salary_bus"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_food"), .rows - 1, .ColIndex("Emp_Salary_food"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_food")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_others"), .rows - 1, .ColIndex("Emp_Salary_others"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_others")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OverTimePrice"), .rows - 1, .ColIndex("OverTimePrice"))
        .TextMatrix(.rows - 1, .ColIndex("OverTimePrice")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.rows - 1, .ColIndex("Mokafea")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance"))
        .TextMatrix(.rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .rows - 1, .ColIndex("total1"))
        .TextMatrix(.rows - 1, .ColIndex("total1")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .rows - 1, .ColIndex("total2"))
        .TextMatrix(.rows - 1, .ColIndex("total2")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mang"), .rows - 1, .ColIndex("Emp_Salary_mang"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_mang")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mob"), .rows - 1, .ColIndex("Emp_Salary_mob"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_mob")) = SngTotal
    
        .Cell(flexcpBackColor, .rows - 1, 1, .rows - 1, .Cols - 1) = vbYellow
        .Cell(flexcpFontBold, .rows - 1, 1, .rows - 1, .Cols - 1) = True
        .Cell(flexcpFontSize, .rows - 1, 1, .rows - 1, .Cols - 1) = 10
        .Cell(flexcpFontName, .rows - 1, 1, .rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
    End With

ErrTrap:
End Sub

Public Sub FillGridWithData2()
    Exit Sub
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim LstDay As Date
    Dim FrstDay As Date
    Dim StrTxt As String
    Dim My_SQL As String
    Dim StrWhere As String
    Dim StrGrp As String
    Dim IntMonth As Integer
    Dim IntYear As Integer
    Dim Msg As String

    'On Error GoTo ErrTrap
    'If DateDiff("d", Me.DtpFrom.Value, Me.DtpTO.Value, vbSaturday) < 0 Then
    '    Exit Sub
    'End If
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    If Me.CmbMonth.ListIndex = -1 Then Exit Sub
    If Me.CboYear.ListIndex = -1 Then Exit Sub

    IntYear = val(Me.CboYear.text)
    IntMonth = Me.CmbMonth.ListIndex + 1

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        Dim ID As String
    
        My_SQL = "SELECT    id,project_id, DepartmentID,id, Emp_Code, Emp_Name, Emp_Salary, Emp_Salary_sakn, Emp_Salary_bus, Emp_Salary_food, Emp_Salary_mob, Emp_Salary_mang, Emp_Salary_others,"
        My_SQL = My_SQL + "OverTimePrice, Mokafea, SalesCom, total1, TotalAdvance, TotalDiscount, total2, EmpTotalNet, sgn, m_year, m_month, payed"
        My_SQL = My_SQL + " from dbo.emp_salary WHERE     (m_year = '" & Me.CboYear.text & "') AND (m_month = '" & Me.CmbMonth.text & "') AND (payed =0) "

        'If Dcemp.text <> "" Then
        'My_SQL = My_SQL + "  and  emp_code='" & Dcemp.BoundText & "'"
        'Else
        'If Dcdep.text <> "" Then
    
        '            If dcproject.BoundText = "" Then
        '            My_SQL = My_SQL + "  and  DepartmentID='" & Dcdep.BoundText & "'"
        '            Else
        '             My_SQL = My_SQL + "   and  DepartmentID='" & Dcdep.BoundText & "' and  project_id='" & Me.dcproject.BoundText & "'"
        '            End If
        ' Else
        '     If Dcdep.text = "" Then
        '
        '              If dcproject.BoundText <> "" Then
        '
        '               My_SQL = My_SQL + "  and  project_id='" & Me.dcproject.BoundText & "'"
        '              End If
    
        '  End If
        '  End If
        '  End If
    
        rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        'With Me.Grid1
        '    .Rows = 2
        '    .Clear flexClearScrollable
        '    If Rs.RecordCount > 0 Then
        '        .Rows = Rs.RecordCount + 1
        '        Rs.MoveFirst
        '        For I = 1 To .Rows - 1
        '
        '            .TextMatrix(I, .ColIndex("Ser")) = I
        '
        '          '  .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(Rs.Fields("ID").value), _
        '            "", Rs.Fields("ID").value)
        '
        '                        .TextMatrix(I, .ColIndex("id")) = IIf(IsNull(Rs.Fields("id").value), _
        '            "", Rs.Fields("id").value)
        '
        '            .TextMatrix(I, .ColIndex("Emp_Code")) = IIf(IsNull(Rs.Fields("Emp_Code").value), _
        '            "", Rs.Fields("Emp_Code").value)
        '
        '
        '                        .TextMatrix(I, .ColIndex("dep")) = IIf(IsNull(Rs.Fields("DepartmentID").value), _
        '            "", Rs.Fields("DepartmentID").value)
        '
        '
        '                        .TextMatrix(I, .ColIndex("project")) = IIf(IsNull(Rs.Fields("project_id").value), _
        '            "", Rs.Fields("project_id").value)
        '
        '
        '            .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(Rs.Fields("Emp_Name").value), _
        '            "", Rs.Fields("Emp_Name").value)
        '
        '            .TextMatrix(I, .ColIndex("Emp_Salary")) = IIf(IsNull(Rs.Fields("Emp_Salary").value), _
                     "", Rs.Fields("Emp_Salary").value)
        '
        '            .TextMatrix(I, .ColIndex("TotalDiscount")) = IIf(IsNull(Rs.Fields("TotalDiscount").value), _
        '            "", Format(Rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
        '
        '            .TextMatrix(I, .ColIndex("Mokafea")) = IIf(IsNull(Rs.Fields("Mokafea").value), _
        '            "", Format(Rs.Fields("Mokafea").value, SystemOptions.SysDefCurrencyForamt))
        '
        '
        '                        .TextMatrix(I, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(Rs.Fields("Emp_Salary_sakn").value), _
        '            "", Format(Rs.Fields("Emp_Salary_sakn").value))
        '
        '
        '                        .TextMatrix(I, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(Rs.Fields("Emp_Salary_bus").value), _
        '            "", Format(Rs.Fields("Emp_Salary_bus").value))
        '
        '
        '                        .TextMatrix(I, .ColIndex("Emp_Salary_food")) = IIf(IsNull(Rs.Fields("Emp_Salary_food").value), _
        '            "", Format(Rs.Fields("Emp_Salary_food").value))
        '
        '                               .TextMatrix(I, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(Rs.Fields("Emp_Salary_mob").value), _
        '            "", Format(Rs.Fields("Emp_Salary_mob").value))
        '
        ''                                    .TextMatrix(I, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(Rs.Fields("Emp_Salary_mang").value), _
        ''            "", Format(Rs.Fields("Emp_Salary_mang").value))
            
        ''
        '                       .TextMatrix(I, .ColIndex("Emp_Salary_others")) = IIf(IsNull(Rs.Fields("Emp_Salary_others").value), _
        '           "", Format(Rs.Fields("Emp_Salary_others").value))
        '
        '                             .TextMatrix(I, .ColIndex("OverTimePrice")) = IIf(IsNull(Rs.Fields("OverTimePrice").value), _
        '           "", Format(Rs.Fields("OverTimePrice").value))
        '
        '
        '                             .TextMatrix(I, .ColIndex("SalesCom")) = IIf(IsNull(Rs.Fields("SalesCom").value), _
        '           "", Format(Rs.Fields("SalesCom").value))
        '
        '
        '         .TextMatrix(I, .ColIndex("total1")) = IIf(IsNull(Rs.Fields("total1").value), _
        '           "", Format(Rs.Fields("total1").value))
        '
        '          .TextMatrix(I, .ColIndex("TotalAdvance")) = IIf(IsNull(Rs.Fields("TotalAdvance").value), _
        '           "", Format(Rs.Fields("TotalAdvance").value))
        '
        '              .TextMatrix(I, .ColIndex("total2")) = IIf(IsNull(Rs.Fields("total2").value), _
        '           "", Format(Rs.Fields("total2").value))
        '
        '                          .TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(Rs.Fields("EmpTotalNet").value), _
        '           "", Format(Rs.Fields("EmpTotalNet").value))
        '
        '
        '           Rs.MoveNext
        '
        '       Next
        '      Rs.Close
        '   End If
        '
        '   GetAdvanceValues IntMonth, IntYear
        '   GetWorkHours
        '   CalculateNets
        '   .Rows = .Rows + 1
        '   .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        '   .IsSubtotal(.Rows - 1) = True
        '   Dim SngTotal As Single
        '   SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .Rows - 1, .ColIndex("Emp_Salary"))
        '   .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary")) = SngTotal
        '
        '   SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .Rows - 1, .ColIndex("EmpTotalNet"))
        '   .TextMatrix(.Rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
        '   net_value1 = SngTotal
        '   SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .Rows - 1, .ColIndex("CorrectEmpTotalNet"))
        '   .TextMatrix(.Rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
        '
        '       SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .Rows - 1, .ColIndex("Emp_Salary_sakn"))
        '   .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
        '
        '
        '
        '       SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .Rows - 1, .ColIndex("Emp_Salary_bus"))
        '   .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
        
        '       SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_food"), .Rows - 1, .ColIndex("Emp_Salary_food"))
        '   .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_food")) = SngTotal
        '
        '
        '
        '       SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_others"), .Rows - 1, .ColIndex("Emp_Salary_others"))
        '   .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_others")) = SngTotal
        '
    
        '         SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OverTimePrice"), .Rows - 1, .ColIndex("OverTimePrice"))
        '   .TextMatrix(.Rows - 1, .ColIndex("OverTimePrice")) = SngTotal
        '
        '         SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .Rows - 1, .ColIndex("Mokafea"))
        '   .TextMatrix(.Rows - 1, .ColIndex("Mokafea")) = SngTotal
        '
        '
        '         SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .Rows - 1, .ColIndex("SalesCom"))
        '   .TextMatrix(.Rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        '
        '         SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .Rows - 1, .ColIndex("TotalAdvance"))
        '   .TextMatrix(.Rows - 1, .ColIndex("TotalAdvance")) = SngTotal
        '
        '             SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .Rows - 1, .ColIndex("TotalDiscount"))
        '   .TextMatrix(.Rows - 1, .ColIndex("TotalDiscount")) = SngTotal
        '
        '                 SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .Rows - 1, .ColIndex("total1"))
        '   .TextMatrix(.Rows - 1, .ColIndex("total1")) = SngTotal
        '
        '                 SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .Rows - 1, .ColIndex("total2"))
        '   .TextMatrix(.Rows - 1, .ColIndex("total2")) = SngTotal
        '
        '                     SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mang"), .Rows - 1, .ColIndex("Emp_Salary_mang"))
        '   .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mang")) = SngTotal
        '
        'SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mob"), .Rows - 1, .ColIndex("Emp_Salary_mob"))
        '   .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mob")) = SngTotal
        '
        '
        '   .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
        '   .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        '   .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 10
        '   .Cell(flexcpFontName, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Tahoma"
        '   .AutoSize 0, .Cols - 1, False
        'End With
    End If

ErrTrap:
End Sub

Private Sub GetWorkHours()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim LngFindRow As Long
    Dim i As Integer
    Dim X As Long
    Dim Y  As Long
    Dim Z As Long
    Dim IntYear As Integer, IntMonth As Integer
    Dim IntDefWorkHours As Integer

    IntYear = val(Me.CboYear.text)
    IntMonth = Me.CmbMonth.ListIndex + 1

    StrSQL = "SELECT dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
    StrSQL = StrSQL + " dbo.ConvertMintsToHours(sum(dbo.tblPresentTime.WorkHoursCount)) AS WorkHours,"
    StrSQL = StrSQL + " dbo.ConvertMintsToHours(SUM( dbo.tblPresentTime.WorkHoursCount - dbo.tblPresentTime.CurrentWorkMints))as OverTime"
    StrSQL = StrSQL + " FROM  dbo.TblEmployee LEFT OUTER JOIN"
    StrSQL = StrSQL + " dbo.tblPresentTime ON dbo.TblEmployee.Emp_ID = dbo.tblPresentTime.Emp_ID"
    'CONVERT (nvarchar(50),GenPresentTime ,111)
    'StrSQL = StrSQL + " Where CONVERT (nvarchar(50),GenPresentTime ,101) >=" & SQLDate(Me.DtpFrom.Value, True) & " AND " & _
     " CONVERT (nvarchar(50),GenPresentTime ,101) <=" & SQLDate(Me.DtpTO.Value, True)
    StrSQL = StrSQL + " Where Month(GenPresentTime)=" & IntMonth & " AND Year(GenPresentTime)=" & IntYear & ""
    StrSQL = StrSQL + " Group By dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Exit Sub
    End If

    'IntDefWorkHours = Val(Me.TxtMonthHours.text)
    If IntDefWorkHours = 0 Then Exit Sub

    Y = ConvertHoursToMints(IntDefWorkHours & ":00")

    With Me.Grid
        .Cell(flexcpText, .FixedRows, .ColIndex("DefWorkHours"), .rows - 1, .ColIndex("DefWorkHours")) = IntDefWorkHours

        For i = 1 To rs.RecordCount
            LngFindRow = .FindRow(rs("Emp_ID").value, .FixedRows, .ColIndex("Emp_ID"), False, True)

            If LngFindRow <> -1 Then
                If Not (IsNull(rs("WorkHours").value)) Then
                    .TextMatrix(LngFindRow, .ColIndex("WorkHours")) = rs("WorkHours").value
                    Z = ConvertHoursToMints(rs("WorkHours").value)
                    X = Z - Y

                    If X < 0 Then
                        .TextMatrix(LngFindRow, .ColIndex("OverTime")) = "-" & ConvertMintsToHours(Abs(X))
                    Else
                        .TextMatrix(LngFindRow, .ColIndex("OverTime")) = ConvertMintsToHours(Abs(X))
                    End If
                
                    If InStr(1, .TextMatrix(LngFindRow, .ColIndex("OverTime")), "-", vbTextCompare) <> 0 Then
                        .Cell(flexcpForeColor, LngFindRow, .ColIndex("OverTime")) = vbRed
                    End If

                Else
                    .TextMatrix(LngFindRow, .ColIndex("WorkHours")) = "00:00"
                    .TextMatrix(LngFindRow, .ColIndex("OverTime")) = "00:00"
                End If
            End If

            rs.MoveNext
        Next i

    End With

End Sub

Private Sub CalculateNets()
    Dim i As Integer
    Dim SngHourPrice As Single
    Dim SngOverTimePrice As Single

    Dim NetTotal As Single
    Dim SngTemp As Single
    'On Error GoTo ErrTrap
    On Error Resume Next

    With Me.Grid

        For i = .FixedRows To .rows - 1
            SngHourPrice = val(.TextMatrix(i, .ColIndex("Emp_Salary"))) / val(.TextMatrix(i, .ColIndex("DefWorkHours")))

            If .TextMatrix(i, .ColIndex("OverTime")) <> "" Then
                SngTemp = ConvertHoursToMints(.TextMatrix(i, .ColIndex("OverTime")))
                SngTemp = SngTemp * (1 / 60)
                SngOverTimePrice = SngTemp * SngHourPrice
                .TextMatrix(i, .ColIndex("OverTimePrice")) = SngOverTimePrice

                If SngOverTimePrice < 0 Then
                    .Cell(flexcpForeColor, i, .ColIndex("OverTimePrice")) = vbRed
                End If
            End If

            .TextMatrix(i, .ColIndex("total1")) = val(.TextMatrix(i, .ColIndex("Emp_Salary"))) + val(.TextMatrix(i, .ColIndex("Emp_Salary_sakn"))) + val(.TextMatrix(i, .ColIndex("Emp_Salary_bus"))) + val(.TextMatrix(i, .ColIndex("Emp_Salary_food"))) + val(.TextMatrix(i, .ColIndex("Emp_Salary_others"))) + val(.TextMatrix(i, .ColIndex("OverTimePrice"))) + val(.TextMatrix(i, .ColIndex("Mokafea"))) + val(.TextMatrix(i, .ColIndex("SalesCom"))) + val(.TextMatrix(i, .ColIndex("Emp_Salary_mob"))) + val(.TextMatrix(i, .ColIndex("Emp_Salary_mang")))
            .TextMatrix(i, .ColIndex("total2")) = val(.TextMatrix(i, .ColIndex("TotalAdvance"))) + val(.TextMatrix(i, .ColIndex("TotalDiscount")))
            .TextMatrix(i, .ColIndex("EmpTotalNet")) = val(.TextMatrix(i, .ColIndex("total1"))) - val(.TextMatrix(i, .ColIndex("total2")))
      
            '.TextMatrix(I, .ColIndex("EmpTotalNet")) = Val(.TextMatrix(I, .ColIndex("EmpTotalNet"))) - Val(.TextMatrix(I, .ColIndex("TotalAdvance")))
            '.TextMatrix(I, .ColIndex("EmpTotalNet")) = Val(.TextMatrix(I, .ColIndex("EmpTotalNet"))) + SngOverTimePrice
            '.TextMatrix(I, .ColIndex("EmpTotalNet")) = Format(Val(.TextMatrix(I, .ColIndex("EmpTotalNet"))), SystemOptions.SysDefCurrencyForamt)
            '.TextMatrix(I, .ColIndex("CorrectEmpTotalNet")) = CorrectCurrency(Val(.TextMatrix(I, .ColIndex("EmpTotalNet"))))
        Next i

    End With

    Exit Sub
ErrTrap:
    'Resume
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

                ' btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub Grid_AfterEdit(ByVal row As Long, _
                           ByVal Col As Long)
    Exit Sub
    Dim NoOfHours As Double
    Dim NoOfMinutes As Double

    With Grid

        Select Case .ColKey(Col)

            Case "Unit"
                .TextMatrix(row, .ColIndex("HourRate")) = 1

            Case "HourRate", "NoOfHour"
    
                If val(.TextMatrix(row, .ColIndex("Unit"))) <> 3 Then
                    .TextMatrix(row, .ColIndex("Value")) = val(.TextMatrix(row, .ColIndex("HourRate"))) * val(.TextMatrix(row, .ColIndex("NoOfHour")))
                Else
    
                    NoOfHours = Int(val(.TextMatrix(row, .ColIndex("NoOfHour"))))

                    If NoOfHours > 0 Then
                        NoOfMinutes = val(.TextMatrix(row, .ColIndex("NoOfHour"))) Mod NoOfHours
                        NoOfMinutes = (NoOfMinutes + NoOfHours + 60)
                  
                        .TextMatrix(row, .ColIndex("Value")) = (NoOfMinutes * val(.TextMatrix(row, .ColIndex("NoOfHour")))) / 60
                    Else
                        .TextMatrix(row, .ColIndex("Value")) = 0
                    End If
    
                End If
    
        End Select

    End With

End Sub

Private Sub Grid_BeforeEdit(ByVal row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid

        Select Case .ColKey(Col)

            Case "Emp_Code"
                .ComboList = ""
                Cancel = True
 
            Case "Emp_Name"
 
                Cancel = True
            
            Case "remarks"
     
                Cancel = True
            
            Case "HourRate"
                Cancel = True
             
        End Select

    End With

End Sub

Private Sub Grid_StartEdit(ByVal row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Exit Sub
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Me.Grid

        Select Case .ColKey(Col)

            Case "Emp_Name"
 
                'Full Path Display
                StrSQL = "SELECT TblEmployee.Emp_Code, TblEmployee.Emp_Name As FirstName " & " FROM TblEmployee "

                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Grid.BuildComboList(rs, "FirstName", "Emp_Code")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
            
        End Select

    End With

End Sub

Private Sub Grid_StartPage(ByVal hDC As Long, _
                           ByVal Page As Long, _
                           Cancel As Boolean)
    Dim s As String

    s = "„— »«  «·„ÊŸ›Ì‰ - Page " & Page & " - " & Now
    TextOut hDC, 100, 100, s, Len(s)
End Sub

Private Sub ISButton2_Click()
    FillGridWithData

    DoEvents
    Dim xApp As New CRAXDRT.Application
    Dim rs As New ADODB.Recordset
    Dim My_SQL As String
    Dim xReport As New CRAXDRT.Report

    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
    rs.Open My_SQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    Set xReport = xApp.OpenReport(App.path & "\reports\emp\REPORT10.rpt")
    xReport.Database.SetDataSource rs
    Dim FrmReport As New FrmReportViewer
    '   FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
  
    FrmReport.CRViewer.viewReport
    FrmReport.show
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
     
    Screen.MousePointer = vbDefault
    ' xReport.ReportTitle = X
    Sendkeys "{RIGHT}"

End Sub

Private Sub ISButton3_Click()

 '   Form3.show
 '   Form3.case_id = 11

End Sub

Private Sub TxtMonthHours_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtMonthHours.text, 1)
End Sub

Private Sub GetAdvanceValues(IntMonth As Integer, _
                             IntYear As Integer)
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim LngFindRow As Long
    On Error GoTo hErr
    StrSQL = "Select Emp_ID,Sum(TotalAdvance)as CCC From ( SELECT QryAllEmpAdvance.Emp_ID,QryA" & "llEmpAdvance.TotalAdvance FROM   dbo.QryAllEmpAdvance(" & IntMonth & "," & IntYear & ") QryAllEmpAdvance )" & "Xtable Group By Emp_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Exit Sub
    End If

    With Me.Grid
        rs.MoveFirst
        .Cell(flexcpText, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance")) = 0

        For i = 1 To rs.RecordCount
            LngFindRow = .FindRow(rs("Emp_ID").value, .FixedRows, .ColIndex("Emp_ID"), False, True)

            If LngFindRow <> -1 Then
                If Not (IsNull(rs("CCC").value)) Then
                    .TextMatrix(LngFindRow, .ColIndex("TotalAdvance")) = rs("CCC").value
                End If
            End If

            rs.MoveNext
        Next i

    End With

hErr:
    'Stop
End Sub

Private Sub Label3_Click()

End Sub

Sub addrow()
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
    Dim des As String

    If Me.TxtModFlg.text <> "R" Then
 
        If Trim(Me.DCComponent.BoundText) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— «·‘«‘…..!!"
            Else
                Msg = "Must Select Screen    ..!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCComponent.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
 
    End If
 
    If (Me.DCEmployee.BoundText) = "" Then
            
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ  " & des & "   «·—”«·…    ...!!!"
        Else
            Msg = "must select " & des & "  Message  ...!!!"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If val(Me.TxtRowNumber.text) <> 0 Then
        LngRow = val(Me.TxtRowNumber.text)
    Else
        Me.Grid.rows = Me.Grid.rows + 1
        LngRow = Me.Grid.rows - 1
    End If

    Dim EmployeeSalary As Double
 
    On Error Resume Next
 
    With Me.Grid
 
        .TextMatrix(LngRow, .ColIndex("MessageDefID")) = val(DCEmployee.BoundText)
    
        .TextMatrix(LngRow, .ColIndex("MessageDefName")) = Me.DCEmployee.text
   
        .AutoSize 0, .Cols - 1, False
    End With

    Me.DCEmployee.BoundText = ""
  
    ReLineGrid
 
End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
  
    With Me.Grid

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("MessageDefID")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
         
            End If

        Next i
   
    End With
 
    Coloring

End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    'Exit Sub
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.rows = 1
 
    If rs.RecordCount < 1 Then
        LabCurrRec.Caption = 0
        LabCountRec.Caption = 0
        Exit Sub
    End If
 
    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.txtid.text = IIf(IsNull(rs("id").value), "", (rs("id").value))
    DCComponent.BoundText = IIf(IsNull(rs("screenName").value), "", rs("screenName").value)
 Me.DcboEmpDepartments.BoundText = IIf(IsNull(rs("DepartmentID").value), "", rs("DepartmentID").value)


 Me.TxTtimeCount.text = IIf(IsNull(rs("timeCount").value), "", (rs("timeCount").value))
 
 Me.txtApprovNameE.text = IIf(IsNull(rs("ApprovNameE").value), "", (rs("ApprovNameE").value))
Me.txtApprovName.text = IIf(IsNull(rs("ApprovName").value), "", (rs("ApprovName").value))

 
   CboTimeCateg.ListIndex = IIf(IsNull(rs("TimeCateg").value), -1, rs("TimeCateg").value)

      dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
      
    StrSQL = "SELECT     dbo.TblApprovalDefDetails.PlainMessageID, dbo.TbLLevels.Name, dbo.TbLLevels.Namee"
    StrSQL = StrSQL & " FROM         dbo.TblApprovalDef INNER JOIN"
    StrSQL = StrSQL & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
    StrSQL = StrSQL & " dbo.TbLLevels ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbLLevels.LevelID"
    StrSQL = StrSQL & " WHERE     (dbo.TblApprovalDefDetails.lMessageDefID = " & val(Me.txtid.text) & ")"
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .rows - 1
  
                .TextMatrix(i, .ColIndex("MessageDefID")) = IIf(IsNull(RsDev("PlainMessageID").value), 0, val(RsDev("PlainMessageID").value))

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("MessageDefName")) = IIf(IsNull(RsDev("Name").value), "", (RsDev("Name").value))
                Else
                    .TextMatrix(i, .ColIndex("MessageDefName")) = IIf(IsNull(RsDev("Namee").value), "", (RsDev("Namee").value))
            
                End If
  
                RsDev.MoveNext
            Next i

            LblSum.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
 
            .AutoSize 0, .Cols - 1, False
        End With

    End If
 
    LabCurrRec.Caption = rs.AbsolutePosition
    LabCountRec.Caption = rs.RecordCount
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub

Private Sub RemoveGridRow()

    With Me.Grid

        If .row <= 0 Then Exit Sub
        .RemoveItem .row
    End With

    ReLineGrid
End Sub

Private Sub Option1_Click()
Exit Sub
    If Me.TxtModFlg.text <> "R" Then
 
        If Trim(Me.DCComponent.BoundText) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
      '          Msg = "ÌÃ» ≈Œ Ì«— «·„›—œ..!!"
            Else
      '          Msg = "Must Select Component    ..!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCComponent.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
 
    End If
 
    If Option1.value = True Then
        get_all_employee
    Else

        With Me.Grid
            .rows = 2
            .Clear flexClearScrollable
        End With

    End If

End Sub
Private Sub Option2_Click()
Exit Sub
    If Me.TxtModFlg.text <> "R" Or Me.TxtModFlg.text <> "E" Then
 
        If Trim(Me.DCComponent.BoundText) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— «·„›—œ..!!"
            Else
                Msg = "Must Select Component    ..!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            DCComponent.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
 
    End If
 
    If Option2.value = True Then
        '     With Me.Grid
        '       .Rows = 1
        '     .Clear flexClearScrollable
        '     End With
    End If

End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.text = "N" Then
        'CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False
        Cmd(7).Enabled = False
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.text = "E" Then
        'CmdRemove.Enabled = True
        
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True
        Cmd(7).Enabled = False
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        'Ele(1).Enabled = False

        'CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True
        Cmd(7).Enabled = True
        Cmd(5).Enabled = True

    End If

End Sub

Private Sub TxtSearchCode_KeyDown(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyReturn Then
   
        Sendkeys "{TAB}"
        Sendkeys "{TAB}"
    End If

End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DCEmployee.BoundText = EmpID
    End If

End Sub

Private Sub TxtValue_KeyDown(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyReturn Then
   
        Cmd_Click (20)
        TxtSearchCode.SetFocus
        TxtSearchCode.text = ""
         
    End If

End Sub

Private Sub Coloring()
    Dim i As Integer
    Dim IntCounter As Integer

    With Grid

        For i = .FixedRows To .rows - 1
        
            If i Mod 2 = 0 Then
                .Cell(flexcpBackColor, i, 1, i, 21) = &HFFFFC0
            Else
                .Cell(flexcpBackColor, i, 1, i, 21) = vbWhite
            End If

        Next i

    End With

    'line_no1 = IntCounter

End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtValue.text, 0)
End Sub

Private Sub XPBtnMove_Click(index As Integer)

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    On Error GoTo ErrTrap

    Select Case index

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

Private Sub XPDtbTrans_Change()
    On Error Resume Next
    CboYear.text = year(XPDtbTrans.value)
    CmbMonth.ListIndex = Month(XPDtbTrans.value) - 1
End Sub
