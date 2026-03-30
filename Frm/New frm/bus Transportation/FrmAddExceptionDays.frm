VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmAddExceptionDays 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " «÷«›«  «·«Ì«„ "
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   Icon            =   "FrmAddExceptionDays.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   6045
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   6105
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6045
      _cx             =   10663
      _cy             =   10769
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   756
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   4452
         Width           =   5856
         _cx             =   10319
         _cy             =   1323
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
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   288
            Width           =   768
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   288
            Width           =   600
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   2
            Left            =   4416
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   288
            Width           =   1128
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   4
            Left            =   900
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   288
            Width           =   1008
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   3732
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   636
         Width           =   5856
         _cx             =   10319
         _cy             =   6588
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
         Begin VB.OptionButton opt_Vendor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ ⁄Âœ"
            Height          =   252
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   1680
            Width           =   732
         End
         Begin VB.OptionButton opt_MA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—… «· ⁄·Ì„Ì…"
            Height          =   252
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   1680
            Width           =   1212
         End
         Begin VB.OptionButton opt_City 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Õ«›Ÿ…"
            Height          =   252
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   1680
            Width           =   852
         End
         Begin VB.OptionButton optAll 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ·"
            Height          =   252
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   1680
            Width           =   732
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Enabled         =   0   'False
            Height          =   1692
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   1920
            Width           =   5532
            Begin VB.TextBox txtRecordno 
               Alignment       =   1  'Right Justify
               Height          =   312
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   240
               Width           =   1752
            End
            Begin VB.TextBox txtfullcode 
               Alignment       =   1  'Right Justify
               Height          =   312
               Left            =   2688
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   240
               Width           =   1380
            End
            Begin MSDataListLib.DataCombo dcCustomer 
               Height          =   288
               Left            =   120
               TabIndex        =   31
               Top             =   600
               Width           =   3936
               _ExtentX        =   6932
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcMangerialAreaID 
               Height          =   288
               Left            =   120
               TabIndex        =   32
               Top             =   1320
               Width           =   3936
               _ExtentX        =   6932
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcCity 
               Height          =   288
               Left            =   120
               TabIndex        =   33
               Top             =   960
               Width           =   3936
               _ExtentX        =   6932
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·”Ã· "
               Height          =   312
               Index           =   21
               Left            =   1668
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   240
               Width           =   888
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂÊœ"
               Height          =   312
               Index           =   22
               Left            =   4488
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   240
               Width           =   612
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ ⁄Âœ"
               Height          =   312
               Index           =   6
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   600
               Width           =   468
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Õ«›Ÿ… "
               Height          =   312
               Index           =   10
               Left            =   4548
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   960
               Width           =   600
               WordWrap        =   -1  'True
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«œ«—… «· ⁄·Ì„Ì…"
               Height          =   312
               Index           =   1
               Left            =   3792
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   1320
               Width           =   1356
               WordWrap        =   -1  'True
            End
         End
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   288
            Left            =   1992
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   2208
         End
         Begin VB.TextBox txtDC 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   240
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   1272
            Width           =   3936
         End
         Begin VB.TextBox txtdayvalue 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   7272
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   3708
            Visible         =   0   'False
            Width           =   768
         End
         Begin MSDataListLib.DataCombo dcDuration 
            Height          =   288
            Left            =   240
            TabIndex        =   28
            Top             =   636
            Width           =   3936
            _ExtentX        =   6932
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcMonth 
            Height          =   288
            Left            =   240
            TabIndex        =   29
            Top             =   996
            Width           =   3936
            _ExtentX        =   6932
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·ÌÊ„"
            Height          =   288
            Index           =   9
            Left            =   7872
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   3708
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«Ì«„"
            Height          =   348
            Index           =   7
            Left            =   4368
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   1272
            Width           =   900
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·› —…"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4200
            TabIndex        =   22
            Top             =   996
            Width           =   1068
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”‰… «·œ—«”Ì…"
            ForeColor       =   &H00000000&
            Height          =   372
            Left            =   4200
            TabIndex        =   16
            Top             =   636
            Width           =   1068
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   288
            Index           =   0
            Left            =   4248
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   1020
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   585
         Left            =   0
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   6180
         _cx             =   10901
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   " «÷«›«  «·«Ì«„ "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
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
         CaptionStyle    =   1
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
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   10
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAddExceptionDays.frx":038A
            ColorButton     =   -2147483634
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
            Left            =   90
            TabIndex        =   11
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAddExceptionDays.frx":0724
            ColorButton     =   -2147483634
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
            Left            =   1680
            TabIndex        =   12
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAddExceptionDays.frx":0ABE
            ColorButton     =   -2147483634
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
            Left            =   615
            TabIndex        =   13
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAddExceptionDays.frx":0E58
            ColorButton     =   -2147483634
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   756
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   5280
         Width           =   5880
         _cx             =   10372
         _cy             =   1323
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
            Height          =   516
            Index           =   0
            Left            =   4908
            TabIndex        =   2
            Top             =   120
            Width           =   924
            _ExtentX        =   1640
            _ExtentY        =   900
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
            ButtonImage     =   "FrmAddExceptionDays.frx":11F2
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
            Height          =   516
            Index           =   1
            Left            =   3840
            TabIndex        =   3
            Top             =   120
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   900
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
            ButtonImage     =   "FrmAddExceptionDays.frx":7A54
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
            Height          =   516
            Index           =   2
            Left            =   2904
            TabIndex        =   4
            Top             =   120
            Width           =   936
            _ExtentX        =   1640
            _ExtentY        =   900
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
            ButtonImage     =   "FrmAddExceptionDays.frx":E2B6
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
            Height          =   516
            Index           =   3
            Left            =   1980
            TabIndex        =   5
            Top             =   120
            Width           =   924
            _ExtentX        =   1640
            _ExtentY        =   900
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
            ButtonImage     =   "FrmAddExceptionDays.frx":14B18
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
            Height          =   516
            Index           =   4
            Left            =   948
            TabIndex        =   6
            Top             =   120
            Width           =   1008
            _ExtentX        =   1773
            _ExtentY        =   900
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
            ButtonImage     =   "FrmAddExceptionDays.frx":1B37A
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
            Height          =   516
            Index           =   6
            Left            =   12
            TabIndex        =   7
            Top             =   120
            Width           =   924
            _ExtentX        =   1640
            _ExtentY        =   900
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
            ButtonImage     =   "FrmAddExceptionDays.frx":21BDC
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
   End
End
Attribute VB_Name = "FrmAddExceptionDays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim rsVendor As ADODB.Recordset
Dim TTP As clstooltip

Dim FromDate_ As Date
Dim ToDate_ As Date
Dim FromDateH_ As String
Dim ToDateH_ As String


Private Sub Cmd_Click(Index As Integer)
 '    On Error GoTo ErrTrap

    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "N"
            clear_all Me
            txtid.Text = CStr(new_id("TblAddExceptionDays", "ID", "", True))
         '   txtName.SetFocus
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"

        Case 2

          
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Company

        Case 5

        Case 6
            Unload Me
         Case 7
   '      print_report2
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

 

Private Sub dcContract_Click(Area As Integer)

 


End Sub

Private Sub dcViolation_Click(Area As Integer)

End Sub



Private Sub dtpDate_Change()
        VBA.Calendar = vbCalGreg
       ' dtpDateH.value = ToHijriDate(dtpDate.value)
End Sub

Private Sub dtpDateH_LostFocus()
'dtpDate.value = ToGregorianDate(dtpDateH.value)
End Sub

Private Sub Dtp_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub dcCity_Change()
Dim str As String
Set Rs_Temp = New ADODB.Recordset
Set dcMangerialAreaID.RowSource = Rs_Temp
If SystemOptions.UserInterface = ArabicInterface Then
    str = " Select ID , Name   from TblManagerialArea  where cityid = " & val(dcCity.BoundText)
Else
    str = " Select ID , NameE   from TblManagerialArea  where cityid = " & val(dcCity.BoundText)
End If
fill_combo dcMangerialAreaID, str
dcMangerialAreaID.Refresh

End Sub



Private Sub dcCustomer_Click(Area As Integer)
Dim val1, val2, recordno As String, Fullcode As String
If dcCustomer.BoundText = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and cusid = " & dcCustomer.BoundText
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     End If
     TxtRecordNo.Text = recordno
     TxtFullcode.Text = Fullcode

End Sub

Private Sub dcCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
        Unload FrmCompanySearch
        FrmCompanySearch.lblSearchtype = "20202020"
        FrmCompanySearch.show vbModal
End If
End Sub

Private Sub dcDuration_Change()
Dim i  As Integer, str As String, typ As Integer
    i = val(dcDuration.BoundText)
    If i > 0 Then
        str = "  select id , Name  from TblDurations_Details where did =   " & i
        fill_combo dcMonth, str
    Else
        str = "  select id , Name  from TblDurations_Details where did =   " & -1
        fill_combo dcMonth, str
    End If
    
 
End Sub

Private Sub dcDuration_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
     '   Unload FrmSearch_Duration
     '   FrmSearch_Duration.SendForm = "ConfirmVacation"
     '   FrmSearch_Duration.show
End If
End Sub

Private Sub Form_Activate()
'    XPTxtBoxID.SetFocus
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

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim Dcombos As ClsDataCombos
    Dim str As String

    
    Set Dcombos = New ClsDataCombos
    'Dcombos.GetCustomersSuppliers 2, dcVendor

    Dcombos.GetCustomersSuppliers 2, dcCustomer
    Dcombos.getCountriesGovernments dcCity
   str = "  select id , name  from TblDurations  "
   fill_combo dcDuration, str
   
 
   str = "   "
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " «À»«  «· ⁄ÿ·  "
    LogTextE = " Open Window " & " Confirm  Violation "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "O", "", ""

    Dim My_SQL As String
       
    
    Resize_Form Me
    
    AddTip
    Set rs = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblAddExceptionDays "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    Me.TxtModFlg.Text = "R"
    
    XPBtnMove_Click 2
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

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

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

 
 
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    'Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  «À»«  «· ⁄ÿ·  "
    LogTextE = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "O", "", ""

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub


Private Sub FromDate_GotFocus()
If dcDuration.BoundText = "" Then
    MsgBox ("«Œ — «·”‰… «·œ—«”Ì… «Ê·« ")
    dcDuration.SetFocus
    Exit Sub
End If

If dcMonth.BoundText = "" Then
    MsgBox ("«Œ — «·› —… «Ê·« ")
    dcMonth.SetFocus
    Exit Sub
End If
End Sub

Private Sub FromDate_LostFocus()
 
End Sub

Private Sub FromDateH_GotFocus()
If dcDuration.BoundText = "" Then
    MsgBox ("«Œ — «·”‰… «·œ—«”Ì… «Ê·« ")
    dcDuration.SetFocus
    Exit Sub
End If

If dcMonth.BoundText = "" Then
    MsgBox ("«Œ — «·› —… «Ê·« ")
    dcMonth.SetFocus
    Exit Sub
End If
End Sub

Private Sub Fromdateh_LostFocus()

 

End Sub


Private Sub ToDate_GotFocus()
If dcDuration.BoundText = "" Then
    MsgBox ("«Œ — «·”‰… «·œ—«”Ì… «Ê·« ")
    dcDuration.SetFocus
    Exit Sub
End If

If dcMonth.BoundText = "" Then
    MsgBox ("«Œ — «·› —… «Ê·« ")
    dcMonth.SetFocus
    Exit Sub
End If
End Sub

 

Private Sub TODateH_GotFocus()
If dcDuration.BoundText = "" Then
    MsgBox ("«Œ — «·”‰… «·œ—«”Ì… «Ê·« ")
    dcDuration.SetFocus
    Exit Sub
End If

If dcMonth.BoundText = "" Then
    dcMonth.SetFocus
    MsgBox ("«Œ — «·› —… «Ê·« ")
    Exit Sub
End If
End Sub

Private Sub todateH_LostFocus()
 
End Sub



Private Sub opt_City_Click()


TxtFullcode.Enabled = False
TxtRecordNo.Enabled = False
dcCustomer.Enabled = False
dcCity.Enabled = True
dcMangerialAreaID.Enabled = False
End Sub

Private Sub opt_MA_Click()


TxtFullcode.Enabled = False
TxtRecordNo.Enabled = False
dcCustomer.Enabled = False
dcCity.Enabled = True
dcMangerialAreaID.Enabled = True
End Sub

Private Sub opt_Vendor_Click()


Frame1.Enabled = True

TxtFullcode.Enabled = False
TxtRecordNo.Enabled = False
dcCustomer.Enabled = True
dcCity.Enabled = False
dcMangerialAreaID.Enabled = False

End Sub

Private Sub optAll_Click()
Frame1.Enabled = False
End Sub



Private Sub optSpes_Click()
Frame1.Enabled = True
End Sub

Private Sub txtfullcode_Change()
Dim val1, val2
If TxtFullcode.Text = "" Then Exit Sub
Dim str As String, recordno As String, CusID As String
recordno = ""
CusID = ""

    str = " select * From TblCustemers where Type=2  and fullcode = '" & TxtFullcode & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
     Else
        TxtRecordNo.Text = ""
        dcCustomer.BoundText = ""
    End If
    
    TxtRecordNo.Text = recordno
    dcCustomer.BoundText = CusID
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = " «÷«›… «·«Ì«„"
            Else
                Me.Caption = "Violation Types"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            Me.txtid.locked = True
            'Me.txtName.locked = True
          '  Me.XPMTxtRemark.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
C1Elastic2.Enabled = False
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «À»«  «· ⁄ÿ·( ÃœÌœ )"
            Else
                Me.Caption = "Violation Types (New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «À»«  «· ⁄ÿ·( ÃœÌœ )"
            Else
                Me.Caption = "Violation Types(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.txtid.locked = True
          '  Me.txtName.locked = False
       C1Elastic2.Enabled = True
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «À»«  «· ⁄ÿ· (  ⁄œÌ· )"
            Else
                Me.Caption = "Violation Types(Edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            Me.txtid.locked = True
           ' Me.txtName.locked = False
       '     Me.XPMTxtRemark.locked = False
       C1Elastic2.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    

    txtid.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    dcDuration.BoundText = IIf(IsNull(rs("DurationID").value), "", Trim(rs("DurationID").value))
    dcMonth.BoundText = IIf(IsNull(rs("MonthID").value), "", rs("MonthID").value)
    txtDC.Text = IIf(IsNull(rs("Days").value), "", rs("Days").value)
    
    dcCustomer.BoundText = IIf(IsNull(rs("VendorID").value), "", rs("VendorID").value)
    dcCity.BoundText = IIf(IsNull(rs("CityID").value), "", rs("CityID").value)
    dcMangerialAreaID.BoundText = IIf(IsNull(rs("ManagerialAreaID").value), "", rs("ManagerialAreaID").value)
    
    optAll.value = IIf(IsNull(rs("Alls").value), False, rs("Alls").value)
    opt_Vendor.value = IIf(IsNull(rs("vendor").value), False, rs("vendor").value)
    opt_City.value = IIf(IsNull(rs("city").value), False, rs("city").value)
    opt_MA.value = IIf(IsNull(rs("MA").value), False, rs("MA").value)
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub




Private Sub TxtName_GotFocus()
On Error Resume Next
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub TxtNameE_GotFocus()
 SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub txtRecordNo_Change()
Dim val1, val2, CusID As String, Fullcode As String
If TxtRecordNo.Text = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and recordno = '" & TxtRecordNo.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
         CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     Else
        dcCustomer.BoundText = ""
        TxtFullcode.Text = ""
    End If
    
   dcCustomer.BoundText = CusID
   TxtFullcode.Text = Fullcode

End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
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

Function CuurentLogdata(Optional Currentmode As String)
   
  

End Function
 
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
    
        If dcDuration.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ «Œ — «·”‰… «·œ—«”Ì… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcDuration.SetFocus
            SendKeys ("{F4}")
            Exit Sub
        End If

      
        
          If dcMonth.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ  «Œ — «·› —… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcMonth.SetFocus
            'SendKeys ("{F4}")
            Exit Sub
        End If
        
        If Not val(txtDC.Text) > 0 Then
                MsgBox ("«œŒ· «·«Ì«„ «Ê·«")
                Exit Sub
        End If
            
        
        Select Case Me.TxtModFlg.Text
            Case "N"
            rs.AddNew
            txtid.Text = CStr(new_id("TblAddExceptionDays", "ID", "", True))
            Case "E"
              '  StrSQL = "select * From  TblViolationTypes where Name='" & Trim(txtName.text) & "'"
           
        End Select

        Cn.BeginTrans
        BeginTrans = True
          
        rs("ID").value = val(txtid.Text)
        rs("DurationID").value = IIf(dcDuration.BoundText = "", Null, dcDuration.BoundText)
        rs("MonthID") = IIf(dcMonth.BoundText = "", Null, dcMonth.BoundText)
        rs("CreationDate") = Date
        rs("Days") = val(txtDC.Text)
        rs("UserID") = user_id
        
        rs("VendorID").value = IIf(dcCustomer.BoundText = "", Null, dcCustomer.BoundText)
        rs("CityID").value = IIf(dcCity.BoundText = "", Null, dcCity.BoundText)
        rs("ManagerialAreaID").value = IIf(dcMangerialAreaID.BoundText = "", Null, dcMangerialAreaID.BoundText)
        
        
        rs("Alls").value = optAll.value
        rs("vendor").value = opt_Vendor.value
         rs("city").value = opt_City.value
        rs("MA").value = opt_MA.value
        
        rs.update
        
        
    
        
        
        Cn.CommitTrans
        
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
       'CuurentLogdata

        Select Case Me.TxtModFlg.Text

            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ «·»Ì«‰«    " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & CHR(13)
                    Msg = Msg + "Do you want enter another One"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
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
            rs.find "ID='" & val(txtid.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
            
        If txtid.Text <> "" Then

    
        Msg = "”Ì „ Õ–› »Ì«‰«  «À»«  «· ⁄ÿ· —ﬁ„ " & CHR(13)
        Msg = Msg + (txtid.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs.RecordCount < 1 Then
                StrSQL = "delete From TblAddExceptionDays where  ID =" & val(txtid.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                                 
                   StrSQL = "SELECT  *  From TblAddExceptionDays  "
                   Set rs = New ADODB.Recordset
                   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
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
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «· ⁄ÿ· "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub


Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  «À»«  «· ⁄ÿ·  ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ·  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «À»«  «· ⁄ÿ· " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ·  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «À»«  «· ⁄ÿ·  «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «À»«  «· ⁄ÿ· " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰« «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ «À»«  «· ⁄ÿ· " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «· ⁄ÿ·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«   «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
    '    .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtBoxName_GotFocus()

    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtBoxNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub


