VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSocialInsurance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈⁄œ«œ «· √„Ì‰«  «·«Ã „«⁄Ì…"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15345
   Icon            =   "FrmSocialInsurance.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   15345
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   23640
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   960
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmSocialInsurance.frx":6852
      Left            =   23640
      List            =   "FrmSocialInsurance.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   2040
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   23640
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Text            =   "modflag"
      Top             =   1320
      Visible         =   0   'False
      Width           =   465
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8100
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   15315
      _cx             =   27014
      _cy             =   14288
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
      Begin VB.TextBox TxtRemarks 
         Alignment       =   2  'Center
         Height          =   3135
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3240
         Width           =   6690
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic8 
         Height          =   870
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   7275
         Width           =   15300
         _cx             =   26988
         _cy             =   1535
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
         Begin ImpulseButton.ISButton btnNew 
            Height          =   435
            Left            =   13740
            TabIndex        =   14
            Top             =   285
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   767
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
            ButtonImage     =   "FrmSocialInsurance.frx":687B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   435
            Left            =   9945
            TabIndex        =   15
            Top             =   285
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   767
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
            ButtonImage     =   "FrmSocialInsurance.frx":D0DD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   435
            Left            =   12000
            TabIndex        =   16
            Top             =   285
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   767
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
            ButtonImage     =   "FrmSocialInsurance.frx":D477
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   435
            Left            =   8040
            TabIndex        =   17
            Top             =   285
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   767
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
            ButtonImage     =   "FrmSocialInsurance.frx":13CD9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   435
            Left            =   6120
            TabIndex        =   18
            Top             =   285
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   767
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
            ButtonImage     =   "FrmSocialInsurance.frx":14073
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   435
            Left            =   4215
            TabIndex        =   19
            Top             =   285
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   767
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
            ButtonImage     =   "FrmSocialInsurance.frx":1460D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   510
            Left            =   2595
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   285
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   900
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
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
            ButtonImage     =   "FrmSocialInsurance.frx":149A7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   435
            Left            =   2520
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   285
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   767
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
            ButtonImage     =   "FrmSocialInsurance.frx":1B209
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   780
         Left            =   0
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   6465
         Width           =   15300
         _cx             =   26988
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   570
            Left            =   120
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   6045
            _cx             =   10663
            _cy             =   1005
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
            Begin VB.Label LabCountRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00C00000&
               Height          =   240
               Left            =   630
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   120
               Width           =   795
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00800000&
               Height          =   240
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   120
               Width           =   705
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   240
               Index           =   1
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   120
               Width           =   1560
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   240
               Index           =   0
               Left            =   4455
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   120
               Width           =   1320
            End
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   10830
            TabIndex        =   28
            Top             =   225
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   270
            Left            =   8745
            TabIndex        =   29
            ToolTipText     =   "Õ–› «·’› «·Õ«·Ì"
            Top             =   225
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·’› «·Õ«·Ì"
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
            ButtonImage     =   "FrmSocialInsurance.frx":1B5A3
            ButtonImageDisabled=   "FrmSocialInsurance.frx":21E05
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   270
            Left            =   7125
            TabIndex        =   30
            ToolTipText     =   "Õ–› «·ﬂ·"
            Top             =   225
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·ﬂ· "
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
            ButtonImage     =   "FrmSocialInsurance.frx":40FEF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   270
            Index           =   8
            Left            =   13965
            TabIndex        =   31
            Top             =   225
            Width           =   885
         End
      End
      Begin C1SizerLibCtl.C1Elastic Frm2 
         Height          =   2025
         Left            =   0
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   750
         Width           =   15300
         _cx             =   26988
         _cy             =   3572
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
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Height          =   1455
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   600
            Width           =   15135
            Begin VB.TextBox TxtResidentVal1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   10920
               TabIndex        =   69
               Top             =   600
               Width           =   810
            End
            Begin VB.TextBox TxtCitizenVal1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   12360
               TabIndex        =   68
               Top             =   600
               Width           =   810
            End
            Begin VB.TextBox TxtCitizenVal2 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   12360
               TabIndex        =   4
               Top             =   960
               Width           =   810
            End
            Begin VB.TextBox TxtAccount3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   960
               Width           =   1185
            End
            Begin VB.TextBox TxtAccount1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   0
               Top             =   480
               Width           =   1185
            End
            Begin VB.TextBox TxtAccount4 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   9240
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   960
               Width           =   1185
            End
            Begin VB.TextBox TxtAccount2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10320
               RightToLeft     =   -1  'True
               TabIndex        =   2
               Top             =   0
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.TextBox TxtResidentVal2 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   10920
               TabIndex        =   5
               Top             =   960
               Width           =   810
            End
            Begin MSDataListLib.DataCombo DcbAccount2 
               Height          =   315
               Left            =   6480
               TabIndex        =   3
               Top             =   0
               Visible         =   0   'False
               Width           =   3765
               _ExtentX        =   6641
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbAccount1 
               Height          =   315
               Left            =   120
               TabIndex        =   1
               Top             =   480
               Width           =   3765
               _ExtentX        =   6641
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbAccount4 
               Height          =   315
               Left            =   5400
               TabIndex        =   9
               Top             =   960
               Width           =   3765
               _ExtentX        =   6641
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbAccount3 
               Height          =   315
               Left            =   120
               TabIndex        =   7
               Top             =   960
               Width           =   3765
               _ExtentX        =   6641
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» «·«ÃÊ— «·„” Õﬁ… ··„ÊŸ›"
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   13
               Left            =   6600
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   600
               Width           =   2700
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "‰”»…  Õ„· «·‘—ﬂ…"
               Height          =   315
               Index           =   2
               Left            =   13080
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   960
               Width           =   2340
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "‰”»…  Õ„· «·„ÊŸ›"
               Height          =   315
               Index           =   0
               Left            =   13080
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   600
               Width           =   2340
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   70
               Left            =   13005
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   630
               Width           =   1365
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   10
               Left            =   10560
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   960
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   9
               Left            =   12000
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   960
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» „œÌ‰"
               Height          =   315
               Index           =   6
               Left            =   6960
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» œ«∆‰"
               Height          =   315
               Index           =   7
               Left            =   2160
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " „Ê«ÿ‰"
               Height          =   315
               Index           =   3
               Left            =   11955
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   240
               Width           =   1500
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " „ﬁÌ„"
               Height          =   315
               Index           =   1
               Left            =   10515
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   240
               Width           =   1500
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   4
               Left            =   12000
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   600
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   5
               Left            =   10560
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   600
               Width           =   300
            End
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   12345
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   135
            Width           =   1365
         End
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
            Height          =   360
            Left            =   6015
            TabIndex        =   34
            Top             =   135
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   635
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   360
            Left            =   9600
            TabIndex        =   35
            Top             =   120
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   635
            _Version        =   393216
            Format          =   175374337
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmSocialInsurance.frx":47851
            Height          =   315
            Left            =   225
            TabIndex        =   36
            Top             =   135
            Visible         =   0   'False
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
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
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   240
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   135
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblcode 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„"
            Height          =   270
            Left            =   14040
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   135
            Width           =   900
         End
         Begin VB.Label lblhjdate 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ «·ÂÃ—Ì"
            Height          =   270
            Left            =   8205
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   135
            Width           =   1305
         End
         Begin VB.Label lbldate 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   360
            Left            =   11310
            TabIndex        =   37
            Top             =   150
            Width           =   780
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   3135
         Left            =   7080
         TabIndex        =   10
         Top             =   3240
         Width           =   8100
         _cx             =   14287
         _cy             =   5530
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSocialInsurance.frx":47866
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
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   615
            Left            =   0
            TabIndex        =   41
            Top             =   1800
            Visible         =   0   'False
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   1085
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   780
         Left            =   0
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   0
         Width           =   15315
         _cx             =   27014
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
         BackColor       =   16777215
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
         Begin ImpulseButton.ISButton btnLast 
            Height          =   285
            Left            =   660
            TabIndex        =   43
            Top             =   225
            Visible         =   0   'False
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   503
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
            ButtonImage     =   "FrmSocialInsurance.frx":47909
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   285
            Left            =   2130
            TabIndex        =   44
            Top             =   225
            Visible         =   0   'False
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
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
            ButtonImage     =   "FrmSocialInsurance.frx":47CA3
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   285
            Left            =   1110
            TabIndex        =   45
            Top             =   225
            Visible         =   0   'False
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
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
            ButtonImage     =   "FrmSocialInsurance.frx":4803D
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   285
            Left            =   1665
            TabIndex        =   46
            Top             =   225
            Visible         =   0   'False
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
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
            ButtonImage     =   "FrmSocialInsurance.frx":483D7
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   555
            Left            =   13830
            Picture         =   "FrmSocialInsurance.frx":48771
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "≈⁄œ«œ «· √„Ì‰«  «·«Ã „«⁄Ì…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Index           =   2
            Left            =   9735
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   225
            Width           =   3750
         End
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„›—œ« "
         Height          =   315
         Index           =   12
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   2880
         Width           =   3300
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„·«ÕŸ« "
         Height          =   315
         Index           =   11
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   2880
         Width           =   3300
      End
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   20280
      TabIndex        =   51
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   -360
      Width           =   2100
      _ExtentX        =   3704
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
      Left            =   23640
      TabIndex        =   52
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   23640
      Top             =   2280
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
            Picture         =   "FrmSocialInsurance.frx":4DF43
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSocialInsurance.frx":4E2DD
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSocialInsurance.frx":4E677
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSocialInsurance.frx":4EA11
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSocialInsurance.frx":4EDAB
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSocialInsurance.frx":4F145
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSocialInsurance.frx":4F4DF
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSocialInsurance.frx":4FA79
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
      Left            =   19920
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   -960
      Width           =   855
   End
End
Attribute VB_Name = "FrmSocialInsurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecID As String
 Dim II As Long

Sub UpdtaMofr(Optional ID As Double = 0)
Dim Sql As String
If ID <> 0 Then
 Sql = "update mofrad  set SoInsID=0,Insurances=0 where id= " & ID & ""
                   Cn.Execute Sql
End If
End Sub

Private Sub DcbAccount1_Change()
DcbAccount1_Click (0)
End Sub

Private Sub DcbAccount1_Click(Area As Integer)
Me.TxtAccount1.Text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount1.BoundText)
End Sub

Private Sub DcbAccount1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 251161
    End If

End Sub

Private Sub DcbAccount2_Change()
DcbAccount2_Click (0)
End Sub

Private Sub DcbAccount2_Click(Area As Integer)
Me.TxtAccount2.Text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount2.BoundText)
End Sub

Private Sub DcbAccount2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 251162
    End If
End Sub

Private Sub DcbAccount3_Change()
DcbAccount3_Click (0)
End Sub

Private Sub DcbAccount3_Click(Area As Integer)
Me.TxtAccount3.Text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount3.BoundText)
End Sub

Private Sub DcbAccount3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 251163
    End If
End Sub

Private Sub DcbAccount4_Change()
DcbAccount4_Click (0)
End Sub

Private Sub DcbAccount4_Click(Area As Integer)
Me.TxtAccount4.Text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount4.BoundText)
End Sub

Private Sub DcbAccount4_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 251164
    End If
End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from TblSocialInsurance order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
   'load tblUsers -----------------------------------------------
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
    '''''''''''''''''''''''''''''''''''
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetAccountingCodes Me.DcbAccount1
    Dcombos.GetAccountingCodes Me.DcbAccount2
    Dcombos.GetAccountingCodes Me.DcbAccount3
    Dcombos.GetAccountingCodes Me.DcbAccount4
   
   ' BtnLast_Click
       If RsSavRec.RecordCount > 0 Then
    Else
     btnNew_Click
    End If
    FiLLTXT
    ShowTip
     If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If

   
       Me.Refresh
ErrTrap:
End Sub
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
    On Error GoTo ErrTrap
    If TxtModFlg = "E" Then
    StrSQL = "Delete From TblSocialInsuranceDet Where SoInsID='" & val(TxtSerial1.Text) & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If
    
    RsSavRec.Fields("RecorDate").value = XPDtbTrans.value
    RsSavRec.Fields("RecorDateH").value = Me.Txt_DateHigri.value
   ' RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("Acount_Code1").value = Me.DcbAccount1.BoundText
    RsSavRec.Fields("Acount_Code2").value = Me.DcbAccount2.BoundText
    RsSavRec.Fields("Acount_Code3").value = Me.DcbAccount3.BoundText
    RsSavRec.Fields("Acount_Code4").value = Me.DcbAccount4.BoundText
    RsSavRec.Fields("Remarks").value = IIf(TxtRemarks.Text <> "", Trim(TxtRemarks.Text), Null)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    RsSavRec.Fields("ResidentVal1").value = val(Me.TxtResidentVal1.Text)
    RsSavRec.Fields("ResidentVal2").value = val(Me.TxtResidentVal2.Text)
    RsSavRec.Fields("CitizenVal1").value = val(Me.TxtCitizenVal1.Text)
    RsSavRec.Fields("CitizenVal2").value = val(Me.TxtCitizenVal2.Text)
    
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' save grid
    Dim Sql As String
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblSocialInsuranceDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim I As Integer
    With Grid
     Sql = "update mofrad  set Insurances=0 where  (SoInsID IS NULL) "
                   Cn.Execute Sql
       For I = .FixedRows To .Rows - 1
         If .TextMatrix(I, .ColIndex("MofrdID")) <> "" Then
                RsDevsub.AddNew
                RsDevsub("SoInsID").value = val(Me.TxtSerial1.Text)
                RsDevsub("MofrdID").value = IIf((.TextMatrix(I, .ColIndex("MofrdID"))) = "", Null, val(.TextMatrix(I, .ColIndex("MofrdID"))))
                RsDevsub("mofrad_type").value = IIf((.TextMatrix(I, .ColIndex("mofrad_type"))) = "", Null, val(.TextMatrix(I, .ColIndex("mofrad_type"))))
                RsDevsub.update
                   Sql = "update mofrad  set SoInsID=1,Insurances=1 where id= " & val(.TextMatrix(I, .ColIndex("mofrad_type"))) & " "
                   Cn.Execute Sql
        End If
      Next I
     End With
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & Chr(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                FullGridData
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
               Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                FullGridData
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                FullGridData
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                FullGridData
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
    Dim I As Integer
    ProgressBar1.Visible = True
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value): ProgressBar1.value = 10
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecorDate").value), Date, RsSavRec.Fields("RecorDate").value): ProgressBar1.value = 20
    Txt_DateHigri.value = IIf(IsNull(RsSavRec.Fields("RecorDateH").value), "", RsSavRec.Fields("RecorDateH").value): ProgressBar1.value = 30
   ' Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value): ProgressBar1.value = 40
    DcbAccount1.BoundText = IIf(IsNull(RsSavRec.Fields("Acount_Code1").value), "", RsSavRec.Fields("Acount_Code1").value): ProgressBar1.value = 50
    DcbAccount2.BoundText = IIf(IsNull(RsSavRec.Fields("Acount_Code2").value), "", RsSavRec.Fields("Acount_Code2").value): ProgressBar1.value = 60
    DcbAccount3.BoundText = IIf(IsNull(RsSavRec.Fields("Acount_Code3").value), "", RsSavRec.Fields("Acount_Code3").value): ProgressBar1.value = 70
    DcbAccount4.BoundText = IIf(IsNull(RsSavRec.Fields("Acount_Code4").value), "", RsSavRec.Fields("Acount_Code4").value): ProgressBar1.value = 80
    TxtResidentVal1.Text = IIf(IsNull(RsSavRec.Fields("ResidentVal1").value), 0, RsSavRec.Fields("ResidentVal1").value): ProgressBar1.value = 90
    TxtResidentVal2.Text = IIf(IsNull(RsSavRec.Fields("ResidentVal2").value), 0, RsSavRec.Fields("ResidentVal2").value): ProgressBar1.value = 100
    TxtCitizenVal1.Text = IIf(IsNull(RsSavRec.Fields("CitizenVal1").value), 0, RsSavRec.Fields("CitizenVal1").value): ProgressBar1.value = 10
    TxtCitizenVal2.Text = IIf(IsNull(RsSavRec.Fields("CitizenVal2").value), 0, RsSavRec.Fields("CitizenVal2").value): ProgressBar1.value = 20
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value): ProgressBar1.value = 30
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value): ProgressBar1.value = 40
    LabCurrRec.Caption = RsSavRec.AbsolutePosition: ProgressBar1.value = 50
    LabCountRec.Caption = RsSavRec.RecordCount: ProgressBar1.value = 60
    ProgressBar1.Visible = False
    ProgressBar1.value = 0
    FullGridData
ErrTrap:
  ProgressBar1.Visible = False
 ProgressBar1.value = 0
End Sub
 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim Sql As String
  Sql = "SELECT     dbo.TblSocialInsuranceDet.ID, dbo.TblSocialInsuranceDet.SoInsID, dbo.TblSocialInsuranceDet.MofrdID, dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.mofrdat.mofrad_type"
  Sql = Sql + "   FROM         dbo.TblSocialInsuranceDet LEFT OUTER JOIN"
  Sql = Sql + "                    dbo.mofrdat ON dbo.TblSocialInsuranceDet.MofrdID = dbo.mofrdat.mofrad_code"
  Sql = Sql + "  Where (dbo.TblSocialInsuranceDet.SoInsID = " & val(TxtSerial1.Text) & ") "
  Rs1.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim I As Integer
     With Me.Grid
                    For I = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(I, .ColIndex("Ser")) = I
                   .TextMatrix(I, .ColIndex("MofrdID")) = IIf(IsNull(Rs1("MofrdID").value), "", Rs1("MofrdID").value)
                   .TextMatrix(I, .ColIndex("mofrad_type")) = IIf(IsNull(Rs1("mofrad_type").value), "", Rs1("mofrad_type").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs1("mofrad_name").value), "", Rs1("mofrad_name").value)
                    Else
                   .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs1("mofrad_namee").value), "", Rs1("mofrad_namee").value)
                    End If
                   Rs1.MoveNext
             Next I
        End With
        Exit Sub
ErrTrap:
    End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrTrap
    Dim StrAccountCode As String
    Dim LngRow As Long
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Sql As String
    With Grid

        Select Case .ColKey(Col)
 
            Case "Name"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("MofrdID"), False, True)
                .TextMatrix(Row, .ColIndex("MofrdID")) = StrAccountCode
                 Sql = "select * from mofrdat where mofrad_code =" & val(.TextMatrix(Row, .ColIndex("MofrdID"))) & " "
                 Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
                 If Rs3.RecordCount > 0 Then
                 .TextMatrix(Row, .ColIndex("mofrad_type")) = IIf(IsNull(Rs3("mofrad_type").value), 0, Rs3("mofrad_type").value)
                End If
            Case "MofrdID"
            Sql = "select * from mofrdat where mofrad_code =" & val(.TextMatrix(Row, .ColIndex("MofrdID"))) & " "
            Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Rs3.RecordCount > 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(Row, .ColIndex("Name")) = IIf(IsNull(Rs3("mofrad_name").value), "", Rs3("mofrad_name").value)
            Else
            .TextMatrix(Row, .ColIndex("Name")) = IIf(IsNull(Rs3("mofrad_namee").value), "", Rs3("mofrad_namee").value)
            End If
            Else
            .TextMatrix(Row, .ColIndex("Name")) = ""
            .TextMatrix(Row, .ColIndex("mofrad_type")) = IIf(IsNull(Rs3("mofrad_type").value), 0, Rs3("mofrad_type").value)
            End If
       End Select
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If
   End With
ErrTrap:
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Dim StrSQL As String
Dim StrComboList As String
With Grid
Select Case .ColKey(Col)
        Case "Name"
                StrSQL = " select * from mofrdat "
                Rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Grid.BuildComboList(Rs2, "mofrad_name", "mofrad_code")
                Else
                    StrComboList = Grid.BuildComboList(Rs2, "mofrad_namee", "mofrad_code")
                End If
                
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With
End Sub



 Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 On Error GoTo ErrTrap
      With Grid
        Select Case .ColKey(Col)
               Case "MofrdID"
               .ComboList = ""
        End Select
       End With
ErrTrap:
End Sub




Private Sub ISButton3_Click()
On Error Resume Next
If Me.TxtModFlg.Text <> "R" Then
    With Me.Grid
        If .Row <= 0 Then Exit Sub
        UpdtaMofr val(.TextMatrix(.Row, .ColIndex("mofrad_type")))
        .RemoveItem .Row
    End With
End If
 End Sub
 Private Sub ISButton4_Click()
 On Error Resume Next
 Dim I As Integer
If Me.TxtModFlg.Text <> "R" Then
For I = 1 To Grid.Rows - 1
UpdtaMofr val(Grid.TextMatrix(I, Grid.ColIndex("mofrad_type")))
Next I
 Me.Grid.Clear flexClearScrollable, flexClearEverything
 cleargriid
End If
 End Sub


Private Sub Txt_DateHigri_LostFocus()
 If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            XPDtbTrans.value = ToGregorianDate(Txt_DateHigri.value)
   End If
End Sub

Private Sub TxtAccount1_KeyPress(KeyAscii As Integer)
Me.DcbAccount1.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount1.Text)
End Sub



Private Sub TxtAccount2_KeyPress(KeyAscii As Integer)
Me.DcbAccount2.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount2.Text)
End Sub

Private Sub TxtAccount3_KeyPress(KeyAscii As Integer)
Me.DcbAccount3.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount3.Text)
End Sub

Private Sub TxtAccount4_KeyPress(KeyAscii As Integer)
Me.DcbAccount4.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount4.Text)
End Sub

Private Sub TxtCitizenVal1_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtCitizenVal1.Text, 0)
End Sub

Private Sub TxtCitizenVal2_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtCitizenVal2.Text, 0)
End Sub

Private Sub TxtResidentVal1_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtResidentVal1.Text, 0)
End Sub

Private Sub TxtResidentVal2_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtResidentVal2.Text, 0)
End Sub

' change date to hj
  Private Sub XPDtbTrans_Change()
  If Me.TxtModFlg.Text <> "R" Then
              Txt_DateHigri.value = ToHijriDate(XPDtbTrans.value)
   End If
   End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text
            '------------------------------ new record ----------------------------
        Case "N"
                  '------------------------- save record -----------------------------
          AddNewRecored
          AddNewRec
           
        '  BtnLast_Click
        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
Else
    MsgBox "Something went wrong while inserting data", vbOKOnly + vbMsgBoxRight, App.title
End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblSocialInsurance", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
' change id search
Private Sub TxtSerial1_Change()
    On Error GoTo ErrTrap
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
ErrTrap:
End Sub
' search for select id
Public Function FindRec(ByVal RecID As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "ID=" & RecID, , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
        FullGridData
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
    Unload Me
End Sub
' undo sub
 Private Sub BtnUndo_Click()
    On Error GoTo ErrTrap
    FindRec val(TxtSerial1.Text)
    Me.TxtModFlg.Text = "R"
    FiLLTXT
     BtnLast_Click
ErrTrap:
End Sub
' delet sub
Private Sub btnDelete_Click()
    On Error GoTo ErrTrap
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If X = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
               Dim I As Integer
               For I = 1 To Grid.Rows - 1
               UpdtaMofr val(Grid.TextMatrix(I, Grid.ColIndex("mofrad_type")))
                Next I
                RsSavRec.Find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.Delete
               '''''''''''''''''''''''''''''''
                 StrSQL = "Delete From TblSocialInsuranceDet Where SoInsID='" & val(TxtSerial1.Text) & "'"
                 Cn.Execute StrSQL, , adExecuteNoRecords
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
               cleargriid
     End If
                            '------------------------------ Move Next ---------------------------.
        Me.Refresh
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
           Cn.Errors.Clear
    End Select

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
                   RecID As String)
     FiLLRec
End Sub

Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        ISButton1.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
              
       
    ElseIf TxtModFlg.Text = "R" Then
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        ISButton1.Enabled = False
        If TxtSerial1.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
            ISButton1.Enabled = True
    End If
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        ISButton1.Enabled = True
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.Text = "E" Then
       Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        ISButton1.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    '    Grid.Enabled = False
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
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    cleargriid
    FiLLTXT
    FullGridData
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
            Msg = Msg & "Date will be updated now" & Chr(13)
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast
    cleargriid
    FiLLTXT
    FullGridData
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
                Msg = Msg & "Date will be updated now" & Chr(13)
            End If
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
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
      Grid.Rows = Grid.Rows + 1
        Frm2.Enabled = True
        Me.Dcbranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & Chr(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry" & Chr(13)
                Msg = Msg & "This recored can't be edited now" & Chr(13)
                Msg = Msg & "it's under modification by other user on the network"
            End If
            
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
                    If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    cleargriid

    TxtModFlg.Text = "N"
    CmbType.ListIndex = 0
    Me.DCboUserName.BoundText = user_id
    'Me.Dcbranch.BoundText = branch_id
    CmbType.ListIndex = 0
    'Dcbranch.SetFocus
    Me.Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
      clear_all Me
      cleargriid
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
    cleargriid
    FiLLTXT
    FullGridData
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
            Msg = Msg & "Date will be updated now" & Chr(13)
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
     cleargriid
    FiLLTXT
    FullGridData
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
            Msg = Msg & "Date will be updated now" & Chr(13)
        End If
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
    Wrap = Chr(13) + Chr(10)
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
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

Private Sub ISButton1_Click()
On Error GoTo ErrTrap
   If val(Me.TxtSerial1.Text) <> 0 Then
       print_report
   End If
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String)
On Error GoTo ErrTrap
    Dim Sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    Sql = "SELECT      dbo.TBLInsurances.IDINS, dbo.TBLInsurances.DateM, dbo.TBLInsurances.DateH, dbo.TBLInsurances.BranchID, dbo.TBLInsurances.SignalID, dbo.TBLInsurances.Monthe,"
    Sql = Sql & "      dbo.TBLInsurances.SubYear, dbo.TBLInsurances.SudePerce, dbo.TBLInsurances.UnSudePerce, dbo.TBLInsurances.Totall, dbo.TBLInsurancesJoin.IDINS AS IDINSJOIN,"
    Sql = Sql & "      dbo.TBLInsurancesJoin.EmpInsurances, dbo.TBLInsurancesJoin.InsValue, dbo.TBLInsurancesJoin.InsTotal, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
    Sql = Sql & "      dbo.mofrad.name, dbo.mofrad.nameE, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng, ACCOUNTS_1.Account_Name AS Account_Name2,"
    Sql = Sql & "      ACCOUNTS_1.Account_NameEng AS Account_NameEng2, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Fullcode"
    Sql = Sql & "      FROM         dbo.TblEmployee RIGHT OUTER JOIN"
    Sql = Sql & "       dbo.TBLInsurancesJoin ON dbo.TblEmployee.Emp_ID = dbo.TBLInsurancesJoin.EmpCode RIGHT OUTER JOIN"
    Sql = Sql & "       dbo.ACCOUNTS ACCOUNTS_1 INNER JOIN"
    Sql = Sql & "      dbo.TBLInsurances INNER JOIN"
    Sql = Sql & "     dbo.TblBranchesData ON dbo.TBLInsurances.BranchID = dbo.TblBranchesData.branch_id ON ACCOUNTS_1.Account_Code = dbo.TBLInsurances.Acount2 LEFT OUTER JOIN"
    Sql = Sql & "      dbo.ACCOUNTS ON dbo.TBLInsurances.Acount1 = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
    Sql = Sql & "     dbo.mofrad ON dbo.TBLInsurances.SignalID = dbo.mofrad.id ON dbo.TBLInsurancesJoin.IDINS = dbo.TBLInsurances.IDINS"
    Sql = Sql & " Where (dbo.TBLInsurances.IDINS = " & val(TxtSerial1.Text) & ")"
                       
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "InsurancesRPT.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "InsurancesRPTENN.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
     If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
    Else
        Msg = "There's no data to show"
    End If
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
      Exit Function
   End If
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData
    Dim cCompanyInfo As New ClsCompanyInfo
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
          xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:
  End Function
Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
    Me.Caption = "GOSI"
    ' labell name
    Me.Label1(2).Caption = Me.Caption
    Me.lblcode.Caption = "Code"
    Me.lbldate.Caption = "Date"
    Me.lblhjdate.Caption = "HJ Date"
    Me.Label3.Caption = "Branch"
    lbl(3).Caption = "Citizens "
    lbl(1).Caption = "Residents "
    lbl(6).Caption = "Debit Account"
    lbl(7).Caption = "Creditor Account"
    lbl(12).Caption = "Componenets"
    lbl(11).Caption = "Remarks"
    '''''''''''''''''''''''' next
   lbl(2).Caption = "Percentage Company"
   lbl(0).Caption = "Percentage Employee"
 
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
   
    ISButton3.Caption = "Delet Select"
    ISButton4.Caption = "Delet All"
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("Name")) = "Name"
        .TextMatrix(0, .ColIndex("MofrdID")) = "Code"
    End With
    
    lbl(13).Caption = "Calculation of wages due to the employee"
    
ErrTrap:
End Sub
Private Sub Dcbranch_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
'  DcboBox.SetFocus
  End If
ErrTrap:
End Sub

Private Sub cleargriid()
Me.Grid.Rows = 1
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
   On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    My_SQL = "TblSocialInsurance"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub




