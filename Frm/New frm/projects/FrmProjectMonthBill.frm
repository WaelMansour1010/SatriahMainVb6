VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmProjectMonthBill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«’œ«— «·„” Œ·’«  «·‘Â—Ì… ··„‘«—Ì⁄ "
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15300
   Icon            =   "FrmProjectMonthBill.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   15300
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   23640
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   960
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmProjectMonthBill.frx":6852
      Left            =   23640
      List            =   "FrmProjectMonthBill.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   36
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
      TabIndex        =   35
      Text            =   "modflag"
      Top             =   1320
      Visible         =   0   'False
      Width           =   465
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8100
      Left            =   0
      TabIndex        =   0
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
      Begin VB.CheckBox Check22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ÕœÌœ «·ﬂ·"
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   1680
         Width           =   1290
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic8 
         Height          =   870
         Left            =   0
         TabIndex        =   1
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
            TabIndex        =   2
            Top             =   285
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
            ButtonImage     =   "FrmProjectMonthBill.frx":687B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   435
            Left            =   9945
            TabIndex        =   3
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
            ButtonImage     =   "FrmProjectMonthBill.frx":D0DD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   435
            Left            =   12000
            TabIndex        =   4
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
            ButtonImage     =   "FrmProjectMonthBill.frx":D477
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   435
            Left            =   8040
            TabIndex        =   5
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
            ButtonImage     =   "FrmProjectMonthBill.frx":13CD9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   435
            Left            =   2160
            TabIndex        =   6
            Top             =   285
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
            ButtonImage     =   "FrmProjectMonthBill.frx":14073
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   435
            Left            =   255
            TabIndex        =   7
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
            ButtonImage     =   "FrmProjectMonthBill.frx":1460D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   510
            Left            =   6075
            TabIndex        =   8
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
            ButtonImage     =   "FrmProjectMonthBill.frx":149A7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   435
            Left            =   4200
            TabIndex        =   9
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
            ButtonImage     =   "FrmProjectMonthBill.frx":1B209
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   780
         Left            =   0
         TabIndex        =   10
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
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   120
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
               TabIndex        =   15
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
               TabIndex        =   14
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
               TabIndex        =   13
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
               TabIndex        =   12
               Top             =   120
               Width           =   1320
            End
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   10830
            TabIndex        =   16
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
            TabIndex        =   17
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
            ButtonImage     =   "FrmProjectMonthBill.frx":1B5A3
            ButtonImageDisabled=   "FrmProjectMonthBill.frx":21E05
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   270
            Left            =   7125
            TabIndex        =   18
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
            ButtonImage     =   "FrmProjectMonthBill.frx":40FEF
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
            TabIndex        =   19
            Top             =   225
            Width           =   885
         End
      End
      Begin C1SizerLibCtl.C1Elastic Frm2 
         Height          =   1305
         Left            =   0
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   750
         Width           =   15300
         _cx             =   26988
         _cy             =   2302
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
            BackColor       =   &H00FFFFC0&
            Height          =   525
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   74
            Top             =   600
            Width           =   4125
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«»«  «·‘—ﬂ…"
            Height          =   975
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   1920
            Visible         =   0   'False
            Width           =   3375
            Begin MSDataListLib.DataCombo DcbAccount4 
               Height          =   315
               Left            =   120
               TabIndex        =   67
               Top             =   240
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbAccount3 
               Height          =   315
               Left            =   120
               TabIndex        =   68
               Top             =   600
               Width           =   2805
               _ExtentX        =   4948
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
               Caption         =   "Õ”«» œ«∆‰"
               Height          =   315
               Index           =   13
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   600
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» „œÌ‰"
               Height          =   315
               Index           =   14
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   240
               Width           =   1380
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… «·‘—ﬂ…"
            Height          =   975
            Left            =   10320
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   1920
            Visible         =   0   'False
            Width           =   1935
            Begin VB.TextBox TxtStay1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   61
               Top             =   600
               Width           =   810
            End
            Begin VB.TextBox TxtCivilin1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   60
               Top             =   240
               Width           =   810
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
               Index           =   12
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   65
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
               Index           =   11
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   240
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„ﬁÌ„Ì‰"
               Height          =   315
               Index           =   10
               Left            =   1035
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   600
               Width           =   1500
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„Ê«ÿ‰Ì‰"
               Height          =   315
               Index           =   9
               Left            =   1035
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   240
               Width           =   1500
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«»«  «·„ÊŸ›"
            Height          =   975
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   1920
            Visible         =   0   'False
            Width           =   3495
            Begin MSDataListLib.DataCombo DcbAccount2 
               Height          =   315
               Left            =   120
               TabIndex        =   54
               Top             =   240
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbAccount1 
               Height          =   315
               Left            =   120
               TabIndex        =   55
               Top             =   600
               Width           =   2805
               _ExtentX        =   4948
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
               Caption         =   "Õ”«» œ«∆‰"
               Height          =   315
               Index           =   7
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   600
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» „œÌ‰"
               Height          =   315
               Index           =   6
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   240
               Width           =   1380
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… «·„ÊŸ›"
            Height          =   975
            Left            =   8160
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   1800
            Visible         =   0   'False
            Width           =   2175
            Begin VB.TextBox TxtCivilin 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   49
               Top             =   240
               Width           =   810
            End
            Begin VB.TextBox TxtStay 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   48
               Top             =   600
               Width           =   810
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„Ê«ÿ‰Ì‰"
               Height          =   315
               Index           =   3
               Left            =   1275
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   240
               Width           =   1500
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„ﬁÌ„Ì‰"
               Height          =   315
               Index           =   1
               Left            =   1275
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   600
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
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   240
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
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   50
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
            TabIndex        =   21
            Top             =   135
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   360
            Left            =   9600
            TabIndex        =   22
            Top             =   120
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   94437377
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   10920
            TabIndex        =   23
            Top             =   1800
            Visible         =   0   'False
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   705
            Index           =   3
            Left            =   8880
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _cx             =   11033
            _cy             =   1244
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
            Begin VB.ComboBox CmbMonth 
               Height          =   315
               Left            =   3315
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   225
               Width           =   1485
            End
            Begin VB.ComboBox CboYear 
               Height          =   315
               Left            =   75
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   42
               Top             =   225
               Width           =   1830
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—"
               Height          =   195
               Index           =   0
               Left            =   4905
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   240
               Width           =   870
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰…"
               Height          =   240
               Index           =   2
               Left            =   2160
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   240
               Width           =   900
            End
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   570
            Left            =   6600
            TabIndex        =   71
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   600
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   1005
            ButtonPositionImage=   2
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
            ButtonImage     =   "FrmProjectMonthBill.frx":47851
            ColorButton     =   14871017
            ButtonToggles   =   2
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmProjectMonthBill.frx":4E0B3
            Height          =   315
            Left            =   1065
            TabIndex        =   72
            Top             =   120
            Width           =   7380
            _ExtentX        =   13018
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   240
            Index           =   15
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   240
            Left            =   8040
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   135
            Width           =   1575
         End
         Begin VB.Label Labelbank 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„›—œ"
            Height          =   255
            Left            =   12480
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   1800
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label lblcode 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„"
            Height          =   270
            Left            =   14040
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   135
            Width           =   900
         End
         Begin VB.Label lbldate 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   360
            Left            =   11310
            TabIndex        =   24
            Top             =   150
            Width           =   780
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   4335
         Left            =   0
         TabIndex        =   27
         Top             =   2040
         Width           =   15300
         _cx             =   26987
         _cy             =   7646
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmProjectMonthBill.frx":4E0C8
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
            Left            =   1560
            TabIndex        =   28
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
         TabIndex        =   29
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
            TabIndex        =   30
            Top             =   225
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
            ButtonImage     =   "FrmProjectMonthBill.frx":4E20A
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   285
            Left            =   2130
            TabIndex        =   31
            Top             =   225
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
            ButtonImage     =   "FrmProjectMonthBill.frx":4E5A4
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   285
            Left            =   1110
            TabIndex        =   32
            Top             =   225
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
            ButtonImage     =   "FrmProjectMonthBill.frx":4E93E
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   285
            Left            =   1665
            TabIndex        =   33
            Top             =   225
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
            ButtonImage     =   "FrmProjectMonthBill.frx":4ECD8
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   4080
            Picture         =   "FrmProjectMonthBill.frx":4F072
            Stretch         =   -1  'True
            Top             =   120
            Width           =   525
         End
         Begin VB.Image Image1 
            Height          =   555
            Left            =   13830
            Picture         =   "FrmProjectMonthBill.frx":52CDA
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«’œ«— «·„” Œ·’«  «·‘Â—Ì… ··„‘«—Ì⁄ "
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
            Left            =   8895
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   225
            Width           =   4590
         End
      End
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   20280
      TabIndex        =   38
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
      TabIndex        =   39
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
            Picture         =   "FrmProjectMonthBill.frx":584AC
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProjectMonthBill.frx":58846
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProjectMonthBill.frx":58BE0
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProjectMonthBill.frx":58F7A
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProjectMonthBill.frx":59314
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProjectMonthBill.frx":596AE
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProjectMonthBill.frx":59A48
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProjectMonthBill.frx":59FE2
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
      TabIndex        =   40
      Top             =   -960
      Width           =   855
   End
End
Attribute VB_Name = "FrmProjectMonthBill"
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


Private Sub CboYear_Change()
CboYear_Click
End Sub

Private Sub CboYear_Click()
 On Error Resume Next
    Dim str As String
    str = "01/" & CmbMonth.ListIndex + 1 & "/" & CboYear.Text

    XPDtbTrans.value = MonthLastDay(CDate(str))
End Sub

Private Sub Check22_Click()
  Dim i As Integer

    If Check22.value = vbChecked Then

        With Me.Grid
 
            For i = 1 To .Rows - 2
        
                .TextMatrix(i, .ColIndex("ch")) = True
            Next i

        End With

    Else

        With Me.Grid

            For i = 1 To .Rows - 2
        
                .TextMatrix(i, .ColIndex("ch")) = False
            Next i

        End With

    End If
End Sub

Private Sub CmbMonth_Change()
CmbMonth_Click
End Sub

Private Sub CmbMonth_Click()
  On Error Resume Next
    Dim str As String
    str = "01/" & CmbMonth.ListIndex + 1 & "/" & CboYear.Text

    XPDtbTrans.value = MonthLastDay(CDate(str))
End Sub
    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from TblProjectMonthBill order by  ID "
    StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
    StrSQL = StrSQL & " order by  ID  "
    
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
    Dcombos.GetBranches Me.dcBranch
 
    
    YearMonth
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
       Me.Refresh
ErrTrap:
End Sub
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++

Sub GetEnd_user_id(Optional ID As Double = 0, Optional ByRef End_user_id As Long, Optional ByRef expanses_account As String)
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "select * from projects where id =" & ID & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
End_user_id = IIf(IsNull(Rs3("End_user_id").value), 0, Rs3("End_user_id").value)
expanses_account = IIf(IsNull(Rs3("expanses_account").value), "", Rs3("expanses_account").value)
Else
expanses_account = ""
End_user_id = 0
End If
End Sub

Sub CreatVoucher(Optional bill_date As Date, Optional ID As Long = 0, Optional Row As Integer, Optional ProjectIDNEw As Double)
    Dim des As String
    Dim Msg As String
    Dim Project_name As String
    Dim note_id As Long
    Dim ManualNO As String
    Dim revenue_account As String
    Dim Remarks As String
    Dim expanses_account As String
    Dim project_no As String
    Dim bill_to As Integer
    Dim MyBranch As String
    Dim UserID As Long
    Dim discount1ID As Integer
    Dim discount2ID As Integer
    Dim subContractorId As Long
    Dim total As Double
    Dim discount As Double
    Dim advancedPayment As Double
    Dim Results As Double
    Dim discount1value As Double
    Dim discount2value As Double
    Dim Account_Code_dynamic1 As String
    Dim StrSQL As String
    Dim RsDev As ADODB.Recordset
    Dim accountdep As String
    Dim End_user_id As Long
    Dim branch_no As Integer
    On Error GoTo ErrTrap
    GetProjectsBillInformation ID, project_no, bill_to, branch_no, total, Project_name, ManualNO, , UserID, discount, advancedPayment, revenue_account, Results, Remarks, discount1value, discount2value, discount1ID, discount2ID, subContractorId
    GetEnd_user_id val(project_no), End_user_id, expanses_account
    '''////////////////////////
    Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
  des = "„” Œ·’ —ﬁ„  :  " & ProjectIDNEw & CHR(13) & "  ··„‘—Ê⁄ " & Project_name & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & ManualNO
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer

tablename = "project_billl"
Filedname = "id"
NoteSerial1 = ProjectIDNEw
Notevalue = 0
 notytype = 5000
 Notevalue = total
 BranchID = branch_no
NoteDate = bill_date
Dim sql As String
 With Me.Grid
If Notevalue > 0 Then
                            
                                     CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des, , , "note_id"
                                              .TextMatrix(Row, .ColIndex("NoteID")) = NoteID
                                                     .TextMatrix(Row, .ColIndex("NoteSerial")) = NoteSerial
                                                        StrSQL = "UPDATE TblProjectMonthBillDet SET NoteID=" & NoteID & " ,NoteSerial='" & NoteSerial & "' WHERE ProjectID=" & ID & " and RecorDate =" & SQLDate(bill_date, True) & " and PrJMonBID =" & val(TxtSerial1.Text) & ""
                           Cn.Execute StrSQL
                    '                 Else
                    '                             If .TextMatrix(Row, .ColIndex("NoteID")) = "" Or .TextMatrix(Row, .ColIndex("NoteSerial")) = "" Then
                    '                        CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des, , , "note_id"
                    '                                              .TextMatrix(Row, .ColIndex("NoteID")) = NoteID
                    '                                              .TextMatrix(Row, .ColIndex("NoteSerial")) = NoteSerial
                    '       StrSQL = "UPDATE TblProjectMonthBillDet SET NoteID=" & NoteID & " ,NoteSerial='" & NoteSerial & "' WHERE ProjectID=" & ID & " and RecorDate =" & SQLDate(bill_date, True) & " and PrJMonBID =" & val(TxtSerial1.Text) & ""
                    '       Cn.Execute StrSQL
                    '                               Else
                    '                                             Sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                    '                                            Sql = Sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                    '                                               Sql = Sql & " where NoteID=" & val(.TextMatrix(Row, .ColIndex("NoteID")))
                    '                                               Cn.Execute Sql
                    '
                    '                             End If
                    '
                    '            End If
                         End If
                          note_id = val(.TextMatrix(Row, .ColIndex("NoteID")))
    End With
    
    ''''''''''''''''/////////////////////
    MyBranch = branch_no
    Set RsDev = New ADODB.Recordset
               StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   Dim LngDevID As Long
  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
  accountdep = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", End_user_id, "Account_code")
If bill_to = 0 Then
  Dim lineno As Integer
  lineno = 1
'    If accountdep = "" Then GoTo ll
    '«·ÿ—› «·„œÌ‰
    RsDev.AddNew
    
    RsDev("branch_id").value = branch_no

    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
    RsDev("Account_Code").value = accountdep '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
    RsDev("Value").value = total
    RsDev("Credit_Or_Debit").value = 0

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & ProjectIDNEw & CHR(13) & "  ··„‘—Ê⁄ " & Project_name & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & ManualNO
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & ProjectIDNEw & CHR(13) & "  To Project " & Project_name & " Manual No. " & ManualNO
    End If

    RsDev("Notes_ID").value = note_id
    RsDev("project_bill_no").value = ProjectIDNEw
  
   RsDev("RecordDate").value = NoteDate
    RsDev("UserID").value = UserID
    RsDev("branch_id").value = branch_no
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev.update
'll:
lineno = lineno + 1

'«·Õ”„Ì« 
'Account_Code_dynamic1
If discount > 0 Then
    RsDev.AddNew
    
    RsDev("branch_id").value = branch_no
    Account_Code_dynamic1 = get_account_code_branch(103, MyBranch)
    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
    RsDev("Account_Code").value = Account_Code_dynamic1 '⁄„Ì· ‰Â«∆Ì «Ê „ﬁ«Ê· »«ÿ‰
    RsDev("Value").value = discount
    RsDev("Credit_Or_Debit").value = 0

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & ProjectIDNEw & CHR(13) & "  ··„‘—Ê⁄ " & Project_name & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & ManualNO
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & ProjectIDNEw & CHR(13) & "  To Project " & Project_name & " Manual No  " & ManualNO
    End If

    RsDev("Notes_ID").value = note_id
    RsDev("project_bill_no").value = ProjectIDNEw
    RsDev("RecordDate").value = NoteDate ' DateValue(Now)
    RsDev("UserID").value = UserID
    RsDev("branch_id").value = branch_no
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev.update
'll:
lineno = lineno + 1
End If
Dim Account_Code_dynamic2 As String
'«·œ›⁄«  «·„ﬁœ„…
'Account_Code_dynamic2
If advancedPayment > 0 Then
    RsDev.AddNew
     Account_Code_dynamic2 = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", End_user_id, "Account_code2")
    RsDev("branch_id").value = branch_no
    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
    RsDev("Account_Code").value = Account_Code_dynamic2 '    Õ”«» œ›⁄«  „ﬁœ„…
    RsDev("Value").value = advancedPayment
    RsDev("Credit_Or_Debit").value = 0

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & ProjectIDNEw & CHR(13) & "  ··„‘—Ê⁄ " & Project_name & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & ManualNO & "œ›⁄«  „ﬁœ„…"
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & ProjectIDNEw & CHR(13) & "  To Project " & Project_name & " Manual No.  " & ManualNO
    End If

    RsDev("Notes_ID").value = note_id
    RsDev("project_bill_no").value = ProjectIDNEw
  
    RsDev("RecordDate").value = NoteDate
    RsDev("UserID").value = UserID
    RsDev("branch_id").value = branch_no
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev.update
'll:
lineno = lineno + 1

  RsDev.AddNew
    
    RsDev("branch_id").value = branch_no

    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
    RsDev("Account_Code").value = accountdep '    Õ”«» œ›⁄«  „ﬁœ„…
    RsDev("Value").value = advancedPayment
    RsDev("Credit_Or_Debit").value = 1

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & ProjectIDNEw & CHR(13) & "  ··„‘—Ê⁄ " & Project_name & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & ManualNO & "œ›⁄«  „ﬁœ„…"
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & ProjectIDNEw & CHR(13) & "  To Project " & Project_name & " Manual No   " & ManualNO
    End If

    RsDev("Notes_ID").value = note_id
    RsDev("project_bill_no").value = ProjectIDNEw
    RsDev("RecordDate").value = NoteDate
    RsDev("UserID").value = UserID
    RsDev("branch_id").value = branch_no
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev.update
'll:
lineno = lineno + 1

End If


'«·«Ì—œ« 

    '«·ÿ—› «·œ«∆‰
    If revenue_account = "" Then Exit Sub
    RsDev.AddNew
    RsDev("branch_id").value = branch_no
    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = lineno
 
    RsDev("Account_Code").value = revenue_account ' Account_Code_dynamic1

    
    RsDev("Value").value = Results '«·«Ì—«œ« 
    RsDev("Credit_Or_Debit").value = 1

    If SystemOptions.UserInterface = ArabicInterface Then
        RsDev("Double_Entry_Vouchers_Description").value = "„” Œ·’ —ﬁ„  :  " & ProjectIDNEw & CHR(13) & "  ··„‘—Ê⁄ " & Project_name & "   " & Remarks & " —ﬁ„ «·”‰œ " & ProjectIDNEw & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & ManualNO
    Else
        RsDev("Double_Entry_Vouchers_Description").value = "  Project Invoice No  :  " & ProjectIDNEw & CHR(13) & "  To Project " & Project_name
    End If

    RsDev("Notes_ID").value = note_id
    RsDev("project_bill_no").value = ProjectIDNEw
 
    RsDev("RecordDate").value = NoteDate
    RsDev("UserID").value = UserID
    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    RsDev.update
lineno = lineno + 1
Else
'
        'If SystemOptions.SubContactorHave3Account = True Then
                Dim Discount1 As Double
                Dim Discount2 As Double
                Dim netvalue As Double
                Dim TotalValue As Double
                Dim AdvancedAccount As String
                Dim GuranteeAccount As String
                Dim line_no As Integer
                
                            If discount1ID = 0 Then
                                Discount1 = 0
                            ElseIf discount1ID = 1 Then
                                Discount1 = discount1value * total / 100
                            ElseIf discount1ID = 2 Then
                                Discount1 = discount1value
                            End If
        
                            If discount2ID = 0 Then
                                Discount2 = 0
                            ElseIf discount2ID = 1 Then
                                Discount2 = discount2value * total / 100
                            ElseIf discount2ID = 2 Then
                                Discount2 = discount2value
                            End If
               AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", subContractorId, "Account_code2")
               GuranteeAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", subContractorId, "Account_code1")
               accountdep = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", subContractorId, "Account_code")
               line_no = 1
               Discount1 = Round(Discount1, 2)
                Discount2 = Round(Discount2, 2)
               netvalue = Round(total - Discount1 - Discount2, 2)
               TotalValue = Round(val(total), 2)
               
                              des = "„’—Ê›«  «·„‘«—Ì⁄ " & "   " & Remarks & " —ﬁ„ «·”‰œ " & ProjectIDNEw & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & ManualNO
           If TotalValue > 0 Then '
    
                
            
               If ModAccounts.AddNewDev(LngDevID, line_no, expanses_account, TotalValue, 0, Msg & des & "  " & "··„‘—Ê⁄   " & Project_name, note_id, , , , SQLDate(bill_date, True), UserID, , , , , , , , , setfoxy_Line, , , , , , , , , branch_no) = False Then
                                        GoTo ErrTrap
                                    
                                    End If
                    
                                    line_no = line_no + 1
              
  '////////////////////////////////////////////////////
       '  End If
         
         
               des = "Œ’„ ÷„«‰ «⁄„«· " & "   " & Remarks & " —ﬁ„ «·”‰œ " & ProjectIDNEw & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & ManualNO
           If Discount1 > 0 Then '÷„«‰ «·«⁄„«·
    
                
            
               If ModAccounts.AddNewDev(LngDevID, line_no, GuranteeAccount, Discount1, 1, Msg & des & "  " & "··„‘—Ê⁄   " & Project_name, note_id, , , , SQLDate(bill_date, True), UserID, , , , , , , , , setfoxy_Line, , , , , , , , , branch_no) = False Then
                                          GoTo ErrTrap
                                    End If
                    
                                    line_no = line_no + 1
            
  
         End If
         
         des = "Œ’„ œ›⁄«  „ﬁœ„…   " & "   " & Remarks & " —ﬁ„ «·”‰œ " & ProjectIDNEw & " —ﬁ„ «·„” Œ·’ «·ÌœÊÌ   " & ManualNO
           If Discount2 > 0 Then '
    
                
            
               If ModAccounts.AddNewDev(LngDevID, line_no, AdvancedAccount, Discount2, 1, Msg & des & "  " & "··„‘—Ê⁄   " & Project_name, note_id, , , , SQLDate(bill_date, True), UserID, , , , , , , , , setfoxy_Line, , , , , , , , , branch_no) = False Then
                                          GoTo ErrTrap
                                    End If
                    
                                    line_no = line_no + 1
              
  
         End If
         
         des = " «⁄„«·"
           If netvalue > 0 Then '
    
                
            
               If ModAccounts.AddNewDev(LngDevID, line_no, accountdep, netvalue, 1, Msg & des & "  " & "··„‘—Ê⁄   " & Project_name, note_id, , , , SQLDate(bill_date, True), UserID, , , , , , , , , setfoxy_Line, , , , , , , , , branch_no) = False Then
                                          GoTo ErrTrap
                                    End If
                    
                                    line_no = line_no + 1
               
  
         End If
         


End If
End If
ErrTrap:
End Sub
Public Sub FiLLRec()
    On Error GoTo ErrTrap
    If TxtModFlg = "E" Then
    Dim i As Integer
    StrSQL = "Delete From TblProjectMonthBillDet Where PrJMonBID=" & val(TxtSerial1.Text) & ""
    Cn.Execute StrSQL, , adExecuteNoRecords
            With Grid
             For i = 1 To .Rows - 1
             Cn.Execute " Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
             Cn.Execute " Delete From notes  Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
             Cn.Execute " Update  project_billl_Month set NewBill_ID=null,FlgBill=null where Bill_ID=" & val(.TextMatrix(i, .ColIndex("ProjectID"))) & ""
             Cn.Execute " Delete from project_billl  where id= " & val(.TextMatrix(i, .ColIndex("NewProjectID"))) & ""
             Cn.Execute " Delete from project_bill_details  where bill_id= " & val(.TextMatrix(i, .ColIndex("NewProjectID"))) & ""
   
             Next i
             End With
    End If
    
    RsSavRec.Fields("RecorDate").value = XPDtbTrans.value
    RsSavRec.Fields("BranchID").value = val(Me.dcBranch.BoundText)
    RsSavRec.Fields("Remarks").value = Me.TxtRemarks.Text
    RsSavRec.Fields("MonthID").value = IIf(val(CmbMonth.ListIndex) <> -1, val((CmbMonth.ListIndex)), Null)
    RsSavRec.Fields("YearID").value = IIf(val(CboYear.ListIndex) <> -1, val(CboYear.ListIndex), Null)
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' save grid
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblProjectMonthBillDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    Dim ProjectIDNEw As Double
    With Grid
       For i = .FixedRows To .Rows - 1
       ProjectIDNEw = 0
       If .Cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
         If val(.TextMatrix(i, .ColIndex("ProjectID"))) <> 0 Then
                RsDevsub.AddNew
                RsDevsub("PrJMonBID").value = Me.TxtSerial1.Text
                RsDevsub("ProjectID").value = IIf((.TextMatrix(i, .ColIndex("ProjectID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ProjectID"))))
                RsDevsub("NewProjectID").value = IIf((.TextMatrix(i, .ColIndex("NewProjectID"))) = "", Null, val(.TextMatrix(i, .ColIndex("NewProjectID"))))
               RsDevsub("RecorDate").value = IIf((Not IsDate(.TextMatrix(i, .ColIndex("RecorDate")))), Null, (.TextMatrix(i, .ColIndex("RecorDate"))))
               RsDevsub.update
               CreateBillMonthLy IIf((Not IsDate(.TextMatrix(i, .ColIndex("RecorDate")))), Null, (.TextMatrix(i, .ColIndex("RecorDate")))), IIf((.TextMatrix(i, .ColIndex("ProjectID"))) = "", 0, val(.TextMatrix(i, .ColIndex("ProjectID")))), ProjectIDNEw
               CreatVoucher IIf((Not IsDate(.TextMatrix(i, .ColIndex("RecorDate")))), Null, (.TextMatrix(i, .ColIndex("RecorDate")))), IIf((.TextMatrix(i, .ColIndex("ProjectID"))) = "", 0, val(.TextMatrix(i, .ColIndex("ProjectID")))), i, ProjectIDNEw
        End If
        End If
      Next i
     End With
     
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
                FiLLTXT
                FullGridData
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                Me.Refresh
                FullGridData
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                
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
    Dim i As Integer
    ProgressBar1.Visible = True
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value): ProgressBar1.value = 10
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecorDate").value), Date, RsSavRec.Fields("RecorDate").value): ProgressBar1.value = 20
    dcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value): ProgressBar1.value = 30
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value): ProgressBar1.value = 40
    CmbMonth.ListIndex = IIf(IsNull(RsSavRec.Fields("MonthID").value), -1, RsSavRec.Fields("MonthID").value): ProgressBar1.value = 50
    CboYear.ListIndex = IIf(IsNull(RsSavRec.Fields("YearID").value), -1, RsSavRec.Fields("YearID").value): ProgressBar1.value = 60
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)

    LabCurrRec.Caption = RsSavRec.AbsolutePosition: ProgressBar1.value = 70
    LabCountRec.Caption = RsSavRec.RecordCount: ProgressBar1.value = 80
    ProgressBar1.Visible = False
    ProgressBar1.value = 0
ErrTrap:
  ProgressBar1.Visible = False
 ProgressBar1.value = 0
End Sub
 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
  sql = "SELECT    * from TblProjectMonthBillDet "
  sql = sql + "  Where (PrJMonBID = " & val(TxtSerial1.Text) & ") "
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.Grid
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("RecorDate")) = IIf(IsNull(Rs1("RecorDate").value), "", Rs1("RecorDate").value)
                   .TextMatrix(i, .ColIndex("ProjectID")) = IIf(IsNull(Rs1("ProjectID").value), 0, Rs1("ProjectID").value)
                   .TextMatrix(i, .ColIndex("NewProjectID")) = IIf(IsNull(Rs1("NewProjectID").value), 0, Rs1("NewProjectID").value)
                   .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(Rs1("NoteID").value), 0, Rs1("NoteID").value)
                   .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(Rs1("NoteSerial").value), "", Rs1("NoteSerial").value)
                   .TextMatrix(i, .ColIndex("ch")) = 1
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub

Private Sub CreateBillMonthLy(Optional RecordDae As Date, Optional ProjectID As Double = 0, Optional ByRef ProjectIDNE As Double)
    Dim StrSQL As String
    Dim sql As String
        Dim ID As Long
        Dim RsNotesGeneral As ADODB.Recordset

        ID = CStr(new_id("project_billl", "ID", "", True))
        
        sql = "INSERT INTO  project_billl (ID ,bill_date,project_no,project_name ,End_user_name,Sub_user_name,End_user_account,Sub_user_account,bill_to,bill_type,revenue_account,note_id,total,NoteSerial,Branch_NO,dueDate,subContractorId,discount1ID,discount2ID,discount1value,discount2value,Remarks,ManualNO,duedate1,discount,Results,advancedPayment,UserID)SELECT " & ID & "," & SQLDate(RecordDae, True) & ",project_no,project_name,End_user_name,Sub_user_name,End_user_account,Sub_user_account,bill_to,bill_type ,revenue_account,note_id,total,NoteSerial,Branch_NO,dueDate,subContractorId,discount1ID ,discount2ID,discount1value,discount2value,Remarks,ManualNO,duedate1,discount,Results,advancedPayment,UserID From project_billl Where  ID =" & ProjectID
        Cn.Execute sql
        '
        sql = "INSERT INTO  dbo.project_bill_details(bill_id,project_no,item,cost,exe,percentage,exedate,line_no,item_id,Unit_id,Quantity,Price,Pre_Quantity,Pre_Value,Pre_Percent,Curr_Quantity,Curr_value,curr_Percent,tot_quantity,tot_value,tot_percent,item_unit,qty,total,discount,net,quntExc,totEx,oprid,discountEXE,NetExe,percentage1,Pre_Percent1,tot_percent1)SELECT  " & ID & ",project_no,item,cost,exe,percentage , exedate, line_no ,item_id,Unit_id,Quantity, Price, Pre_Quantity, Pre_Value,Pre_Percent,Curr_Quantity,Curr_value,curr_Percent,tot_quantity ,tot_value,tot_percent,item_unit,qty,total,discount,net,quntExc,totEx,oprid,discountEXE,NetExe,percentage1,Pre_Percent1,tot_percent1 From dbo.project_bill_details Where   bill_id =" & ProjectID
 Cn.Execute sql
        StrSQL = "UPDATE project_billl_Month SET NewBill_ID=" & ID & ",FlgBill=1 WHERE Bill_ID=" & ProjectID & " and RecordDate =" & SQLDate(RecordDae, True) & ""
        Cn.Execute StrSQL
StrSQL = "UPDATE project_billl_Month SET NewBill_ID=" & ID & ",FlgBill=1 WHERE Bill_ID=" & ProjectID & " and RecordDate =" & SQLDate(RecordDae, True) & ""
        Cn.Execute StrSQL

StrSQL = "UPDATE TblProjectMonthBillDet SET NewProjectID=" & ID & " WHERE ProjectID=" & ProjectID & " and RecorDate =" & SQLDate(RecordDae, True) & " and PrJMonBID =" & val(TxtSerial1.Text) & ""
        Cn.Execute StrSQL
 ProjectIDNE = ID
ErrTrap:

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid
Select Case .ColKey(Col)
Case "Show"
.ComboList = ""
Case "ProjectID"
Cancel = True
Case "RecorDate"
Cancel = True
Case "NewProjectID"
Cancel = True
Case "NoteSerial"
Cancel = True
End Select
End With
End Sub

Private Sub Grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With Grid
Select Case .ColKey(Col)
Case "Show"
Unload projectsbill
Load projectsbill
projectsbill.show
projectsbill.Retrive val(.TextMatrix(Row, .ColIndex("NewProjectID")))
End Select
End With
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid
Select Case .ColKey(Col)
Case "Show"
.ColComboList(.ColIndex("Show")) = "..."
End Select
End With
End Sub

Private Sub ImgFavorites_Click()
    AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton2_Click()
If Me.TxtModFlg.Text <> "R" Then
  On Error GoTo ErrTrap

    '+++++++++++++++++++++++++++++++++++++++++++++++

      If val(CmbMonth.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·‘Â— ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
               Else
            MsgBox "Please Select Month ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
           End If
             CmbMonth.SetFocus
            Exit Sub
      End If
     If val(CboYear.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·”‰… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
             Else
            MsgBox "Please Select Year ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
             CboYear.SetFocus
            Exit Sub
     End If
     FillTextGridData
ErrTrap:
End If
   End Sub
  Sub FillTextGridData()
  On Error GoTo ErrTrap
   Dim Rs1 As ADODB.Recordset
   Set Rs1 = New ADODB.Recordset
   Dim sql As String
       
 sql = "SELECT     Bill_ID, FlgBill, NewBill_ID, RecordDate, ID, MONTH(RecordDate) AS Mnth, YEAR(RecordDate) AS yar "
 sql = sql & " From dbo.project_billl_Month"
 sql = sql & " Where (Month(recorddate) = " & val(CmbMonth.ListIndex + 1) & ") And (year(recorddate) = " & val(CboYear.Text) & ")"
 sql = sql & " and (FlgBill IS NULL) "
      Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.Grid
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("NewProjectID")) = IIf(IsNull(Rs1("NewBill_ID").value), 0, Rs1("NewBill_ID").value)
                   .TextMatrix(i, .ColIndex("RecorDate")) = IIf(IsNull(Rs1("RecordDate").value), "", Rs1("RecordDate").value)
                   .TextMatrix(i, .ColIndex("ProjectID")) = IIf(IsNull(Rs1("Bill_ID").value), 0, Rs1("Bill_ID").value)
                   
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
 End Sub


Private Sub ISButton3_Click()
On Error Resume Next
If Me.TxtModFlg.Text <> "R" Then
       With Grid
       If .Cell(flexcpChecked, .Row, .ColIndex("ch")) = flexChecked Then
             Cn.Execute " Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(.TextMatrix(.Row, .ColIndex("NoteID")))
             Cn.Execute " Delete From notes  Where NoteID=" & val(.TextMatrix(.Row, .ColIndex("NoteID")))
             Cn.Execute " Update  project_billl_Month set NewBill_ID=null,FlgBill=null where Bill_ID=" & val(.TextMatrix(.Row, .ColIndex("ProjectID"))) & ""
             Cn.Execute " Delete from project_billl  where id= " & val(.TextMatrix(.Row, .ColIndex("NewProjectID"))) & ""
             Cn.Execute " Delete from project_bill_details  where bill_id= " & val(.TextMatrix(.Row, .ColIndex("NewProjectID"))) & ""
        End If
             End With
    With Me.Grid
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End If
 End Sub
 Private Sub ISButton4_Click()
 Dim i As Integer
 On Error Resume Next
 If Me.TxtModFlg.Text <> "R" Then
       With Grid
             For i = 1 To .Rows - 1
             Cn.Execute " Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
             Cn.Execute " Delete From notes  Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
             Cn.Execute " Update  project_billl_Month set NewBill_ID=null,FlgBill=null where Bill_ID=" & val(.TextMatrix(i, .ColIndex("ProjectID"))) & ""
             Cn.Execute " Delete from project_billl  where id= " & val(.TextMatrix(i, .ColIndex("NewProjectID"))) & ""
             Cn.Execute " Delete from project_bill_details  where bill_id= " & val(.TextMatrix(i, .ColIndex("NewProjectID"))) & ""
   
             Next i
             End With
 Me.Grid.Clear flexClearScrollable, flexClearEverything
 cleargriid
End If
 End Sub

' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Function CheckPeriod() As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "  SELECT     ID, YearID, MonthID"
sql = sql & " From dbo.TblProjectMonthBill"
sql = sql & " Where (YearID = " & val(CboYear.ListIndex) + 2010 & ") And (MonthID = " & val(CmbMonth.ListIndex) + 1 & ") And (ID <> " & val(TxtSerial1.Text) & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckPeriod = True
Else
CheckPeriod = False
End If
End Function
Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
  '  If CheckPeriod() = True Then
  '  If SystemOptions.UserInterface = ArabicInterface Then
  '  MsgBox "·«Ì„ﬂ‰  ﬂ—«— «·› —…"
  '  Else
  '  MsgBox "Can not repeat period"
  '  End If
  '  Exit Sub
  '  End If
    '---------------------- check if data Vaclete -----------------------
       If dcBranch.Text = "" Or (dcBranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄ «·ﬁ«∆„ »«·Õ—ﬂ…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please select Branch Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
          dcBranch.SetFocus
            Exit Sub
         End If



      If CmbMonth.Text = "" Or val(CmbMonth.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·‘Â— ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
             Else
            MsgBox "Please select Month ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
           End If
             CmbMonth.SetFocus
             Exit Sub
      End If
     If CboYear.Text = "" Or val(CboYear.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·”‰… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
             Else
            MsgBox "Please Select Year ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            CboYear.SetFocus
            Exit Sub
     End If


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
  MsgBox "Sorry error douring enter data", vbOKOnly + vbMsgBoxRight, App.title
  End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblProjectMonthBill", "ID", "")
    RsSavRec.AddNew
    TxtSerial1.Text = StrRecID
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
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "ID=" & RecId, , adSearchForward, 1
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
    Dim i As Integer
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
             With Grid
             For i = 1 To .Rows - 1
             Cn.Execute " Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
             Cn.Execute " Delete From notes  Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
             Cn.Execute " Update  project_billl_Month set NewBill_ID=null,FlgBill=null where Bill_ID=" & val(.TextMatrix(i, .ColIndex("ProjectID"))) & ""
             Cn.Execute " Delete from project_billl  where id= " & val(.TextMatrix(i, .ColIndex("NewProjectID"))) & ""
             Cn.Execute " Delete from project_bill_details  where bill_id= " & val(.TextMatrix(i, .ColIndex("NewProjectID"))) & ""
   
             Next i
             End With
                StrSQL = "Delete From TblProjectMonthBillDet Where PrJMonBID=" & val(TxtSerial1.Text) & ""
                 Cn.Execute StrSQL, , adExecuteNoRecords
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
  RsSavRec.delete
  
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
               LabCurrRec.Caption = 0
               LabCountRec.Caption = 0
               cleargriid
     End If
                            '------------------------------ Move Next ---------------------------.
        Me.Refresh
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "You can not delete this record because of its connection with other data"
            End If
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
              Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
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
        Me.dcBranch.BoundText = branch_id
        Frm2.Enabled = True
        Me.dcBranch.SetFocus
    End If
    Exit Sub
ErrTrap:

    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
          Else
           Msg = "Sorry ..." & CHR(13)
            Msg = Msg & "You can not edit this record now" & CHR(13)
            Msg = Msg & "It is in use by another user on the network"
          End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
                    If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
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
  '  Me.VSFlexGrid2.Rows = 1
    TxtModFlg.Text = "N"

    Me.DCboUserName.BoundText = user_id
    Me.dcBranch.BoundText = branch_id
    dcBranch.SetFocus
    Me.Grid.Clear flexClearScrollable, flexClearEverything
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
       Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
       Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
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
    Wrap = CHR(13) + CHR(10)
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
    Me.Caption = "Monthly Billing"
    Me.Label1(2).Caption = Me.Caption
    lbldate.Caption = "Date"
    Label3.Caption = "Branch"
    lbl(15).Caption = "Remarks"
    lblcode.Caption = "No"
    Ele(3).Caption = "Select Period"
    lbl(2).Caption = "Year"
    lbl(0).Caption = "Month"
   ' Label1(35).Caption = "GL"
   Check22.RightToLeft = False
   Check22.Caption = "Select All"
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    ISButton2.Caption = "ADD"
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
        .TextMatrix(0, .ColIndex("ProjectID")) = "No. Original Invoice"
        .TextMatrix(0, .ColIndex("RecorDate")) = "Date"
        .TextMatrix(0, .ColIndex("Show")) = "Show"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "GL"
        .TextMatrix(0, .ColIndex("NewProjectID")) = "No. New Invoice"
         .TextMatrix(0, .ColIndex("ch")) = "Select"
        
    End With
ErrTrap:
End Sub

Private Sub CmbMonth_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  CboYear.SetFocus
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
   My_SQL = "TblProjectMonthBill"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

Private Sub YearMonth()
    Dim i As Integer
    Dim IntDefIndex As Integer
    CmbMonth.Clear
    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next
    CmbMonth.ListIndex = Month(Date) - 1
    ''''''''''
    CboYear.Clear
    For i = 2010 To 2050
        CboYear.AddItem i
        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If
    Next
    CboYear.ListIndex = IntDefIndex
End Sub
'+++++++++++++++++++++++++++++++++ en
