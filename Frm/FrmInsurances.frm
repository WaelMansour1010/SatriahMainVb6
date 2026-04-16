VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmInsurances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· «À»«  «· √„Ì‰« "
   ClientHeight    =   8070
   ClientLeft      =   2760
   ClientTop       =   3660
   ClientWidth     =   15300
   Icon            =   "FrmInsurances.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   15300
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   23640
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   960
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmInsurances.frx":6852
      Left            =   23640
      List            =   "FrmInsurances.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   37
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
      TabIndex        =   36
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
            ButtonImage     =   "FrmInsurances.frx":687B
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
            ButtonImage     =   "FrmInsurances.frx":D0DD
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
            ButtonImage     =   "FrmInsurances.frx":D477
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
            ButtonImage     =   "FrmInsurances.frx":13CD9
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
            ButtonImage     =   "FrmInsurances.frx":14073
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
            ButtonImage     =   "FrmInsurances.frx":1460D
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
            ButtonImage     =   "FrmInsurances.frx":149A7
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
            ButtonImage     =   "FrmInsurances.frx":1B209
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
            ButtonImage     =   "FrmInsurances.frx":1B5A3
            ButtonImageDisabled=   "FrmInsurances.frx":21E05
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
            ButtonImage     =   "FrmInsurances.frx":40FEF
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
         Height          =   1785
         Left            =   0
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   750
         Width           =   15300
         _cx             =   26988
         _cy             =   3149
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
         Begin VB.Frame Frame9 
            Caption         =   "»Ì«‰«  „Õ«”»Ì…"
            Height          =   744
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   960
            Width           =   7290
            Begin VB.TextBox TxtNoteID 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   360
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   240
               Width           =   2415
            End
            Begin VB.CommandButton Command9 
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   375
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton Command8 
               Caption         =   "ﬂ‘› Õ”«»"
               Height          =   375
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   195
               Index           =   35
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   240
               Width           =   990
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«»«  «·‘—ﬂ…"
            Height          =   975
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   1920
            Visible         =   0   'False
            Width           =   3375
            Begin MSDataListLib.DataCombo DcbAccount4 
               Height          =   315
               Left            =   120
               TabIndex        =   73
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
               TabIndex        =   74
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
               TabIndex        =   76
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
               TabIndex        =   75
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
            TabIndex        =   65
            Top             =   1920
            Visible         =   0   'False
            Width           =   1935
            Begin VB.TextBox TxtStay1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   67
               Top             =   600
               Width           =   810
            End
            Begin VB.TextBox TxtCivilin1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   66
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
               TabIndex        =   71
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
               TabIndex        =   70
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
               TabIndex        =   69
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
               TabIndex        =   68
               Top             =   240
               Width           =   1500
            End
         End
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   255
            Index           =   0
            Left            =   6600
            TabIndex        =   62
            Top             =   480
            Width           =   975
            _Version        =   786432
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "ﬂ· «·›—Ê⁄"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«»«  «·„ÊŸ›"
            Height          =   975
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   1920
            Visible         =   0   'False
            Width           =   3495
            Begin MSDataListLib.DataCombo DcbAccount2 
               Height          =   315
               Left            =   120
               TabIndex        =   55
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
               TabIndex        =   56
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
               TabIndex        =   58
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
               TabIndex        =   57
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
            TabIndex        =   47
            Top             =   1800
            Visible         =   0   'False
            Width           =   2175
            Begin VB.TextBox TxtCivilin 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   50
               Top             =   240
               Width           =   810
            End
            Begin VB.TextBox TxtStay 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   49
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
               TabIndex        =   54
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
               TabIndex        =   53
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
               TabIndex        =   52
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
               TabIndex        =   51
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
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
            Height          =   360
            Left            =   7695
            TabIndex        =   22
            Top             =   135
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   635
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   360
            Left            =   9600
            TabIndex        =   23
            Top             =   120
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   93650945
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   10920
            TabIndex        =   24
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
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   720
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
               TabIndex        =   44
               Top             =   225
               Width           =   1485
            End
            Begin VB.ComboBox CboYear 
               Height          =   315
               Left            =   75
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   43
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
               TabIndex        =   46
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
               TabIndex        =   45
               Top             =   240
               Width           =   900
            End
         End
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   255
            Index           =   1
            Left            =   4920
            TabIndex        =   63
            Top             =   480
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "›—⁄ „Õœœ"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmInsurances.frx":47851
            Height          =   315
            Left            =   225
            TabIndex        =   77
            Top             =   480
            Width           =   4620
            _ExtentX        =   8149
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
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   930
            Left            =   7545
            TabIndex        =   78
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   720
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   1640
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
            ButtonImage     =   "FrmInsurances.frx":47866
            ColorButton     =   14871017
            ButtonToggles   =   2
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSDataListLib.DataCombo Dcbranch1 
            Bindings        =   "FrmInsurances.frx":4E0C8
            Height          =   315
            Left            =   225
            TabIndex        =   85
            Top             =   120
            Width           =   4620
            _ExtentX        =   8149
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
         Begin VB.Label lblhjdate 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ «·ÂÃ—Ì"
            Height          =   270
            Left            =   8205
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   735
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄ «·ﬁ«∆„ »«·Õ—ﬂ…"
            Height          =   240
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   64
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
            TabIndex        =   27
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
            TabIndex        =   26
            Top             =   135
            Width           =   900
         End
         Begin VB.Label lbldate 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   360
            Left            =   11310
            TabIndex        =   25
            Top             =   150
            Width           =   780
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   3390
         Left            =   120
         TabIndex        =   28
         Top             =   2520
         Width           =   15300
         _cx             =   26987
         _cy             =   5980
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
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmInsurances.frx":4E0DD
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
         ExplorerBar     =   3
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
            TabIndex        =   29
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
         TabIndex        =   30
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
            TabIndex        =   31
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
            ButtonImage     =   "FrmInsurances.frx":4E372
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   285
            Left            =   2130
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
            ButtonImage     =   "FrmInsurances.frx":4E70C
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   285
            Left            =   1110
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
            ButtonImage     =   "FrmInsurances.frx":4EAA6
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   285
            Left            =   1665
            TabIndex        =   34
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
            ButtonImage     =   "FrmInsurances.frx":4EE40
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   555
            Left            =   13830
            Picture         =   "FrmInsurances.frx":4F1DA
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ”ÃÌ· «À»«  «· √„Ì‰« "
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
            TabIndex        =   35
            Top             =   225
            Width           =   3750
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   570
         Left            =   120
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   5880
         Width           =   4605
         _cx             =   8123
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
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ã„Ê⁄"
            Height          =   240
            Index           =   3
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label TotalTXT 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   360
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   120
            Width           =   2775
         End
      End
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   20280
      TabIndex        =   39
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
      TabIndex        =   40
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
            Picture         =   "FrmInsurances.frx":549AC
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInsurances.frx":54D46
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInsurances.frx":550E0
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInsurances.frx":5547A
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInsurances.frx":55814
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInsurances.frx":55BAE
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInsurances.frx":55F48
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInsurances.frx":564E2
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
      TabIndex        =   41
      Top             =   -960
      Width           =   855
   End
End
Attribute VB_Name = "FrmInsurances"
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
Private Sub btnQuery_Click()
    On Error GoTo ErrTrap
    FrmInsurancesSearch.SendForm = 1
    Load FrmInsurancesSearch
    FrmInsurancesSearch.show vbModal
ErrTrap:
End Sub
Function ChekExistRecord() As Boolean
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim sql As String
sql = "Select * from TBLInsurances where  Monthe =" & val(CmbMonth.ListIndex) & " and SubYear =" & val(CboYear.Text) & " "
If Rd(0).value = True Then
sql = sql & " and (BranchID <> 0 or AllBranch=0 ) "
Else
sql = sql & " and (BranchID =" & val(Dcbranch.BoundText) & " or AllBranch=0) "
End If
sql = sql & " and IDINS<>" & val(TxtSerial1.Text) & ""
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
ChekExistRecord = True
Else
ChekExistRecord = False
End If

End Function
Private Sub CboYear_Change()
CboYear_Click
End Sub

Private Sub CboYear_Click()
 On Error Resume Next
    Dim str As String
    str = "01/" & CmbMonth.ListIndex + 1 & "/" & CboYear.Text

    XPDtbTrans.value = MonthLastDay(CDate(str))
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

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub

Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
Dim BasicSalaryAccount As String
Dim StrSQL As String
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords



    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
   
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Dim j As Integer
    Dim ColumnName As String
    Dim SalaryAccount As String
    Dim BonusAccount As String
    Dim DiscountAccount As String
    
Msg = "«À»«  «· √„Ì‰« -Õ’… «·‘—ﬂ… ⁄‰ «·› —… " & CmbMonth.Text & " ·”‰…" & CboYear.Text
 
        
 
 notes_id = general_noteid

  
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
    'C???? C??I?? C?C?C?CE
     
    Dim CValue As Double
    Dim Branch As Integer
    Dim ProjectID As Integer
    Dim mofradAccount As String
    Dim mofradAccount1 As String
    Dim Nationality As String
    BranchID = 1
   '********************************************************************************************************
         GetInsuranceAccount , , , , mofradAccount, mofradAccount1

 Dim Emp_id  As Double


   '*********************************************************************************************************
    With Grid


line_no = 1
        For i = .FixedRows To .Rows - 1
    BranchID = .TextMatrix(i, .ColIndex("BranchId"))
         Nationality = .TextMatrix(i, .ColIndex("Citirent"))
       Emp_id = val(.TextMatrix(i, .ColIndex("Empid")))
       
            If val(.TextMatrix(i, .ColIndex("InsTotal2"))) > 0 And .TextMatrix(i, .ColIndex("Empid")) <> "" Then    'C?C??? C???E??E IC??
          
        
        
                If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, val(.TextMatrix(i, .ColIndex("InsTotal2"))), 0, Msg & "··„ÊŸ›" & .TextMatrix(i, .ColIndex("EmpCode")) & " - " & .TextMatrix(i, .ColIndex("EmpName")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , Emp_id) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
                
                If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount1, val(.TextMatrix(i, .ColIndex("InsTotal2"))), 1, Msg & "··„ÊŸ›" & .TextMatrix(i, .ColIndex("EmpCode")) & " - " & .TextMatrix(i, .ColIndex("EmpName")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , Emp_id) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
            
                
                
                
            End If
     
     
     Next i
     
     End With
           
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
   End Function
Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "«À»«  «· √„Ì‰« -Õ’… «·‘—ﬂ… ⁄‰ «·› —… " & CmbMonth.Text & " ·”‰…" & CboYear.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
 

Dim sql As String
tablename = "TBLInsurances"
Filedname = "IDINS"
NoteSerial1 = val(TxtSerial1)
Notevalue = 0

 notytype = 8070
Notevalue = val(TotalTXT.Caption)
 

 BranchID = val(dcBranch1.BoundText)
NoteDate = (XPDtbTrans.value)
 
If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TxtNoteID.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial
                                     Else
                                                 If TxtNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                            CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                                 TxtNoteID.Text = NoteID
                                                                TxtNoteSerial.Text = NoteSerial
                                                   Else
                                                                 sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   sql = sql & " where NoteID=" & val(TxtNoteID.Text)
                                                                   Cn.Execute sql
                                                               
                                                 End If
                                       
                                End If

CREATE_VOUCHER_GE val(TxtNoteID.Text), BranchID, user_id, NoteDate
RsSavRec.Resync adAffectCurrent
 

     End If

End Function

Private Sub DcBranch1_Click(Area As Integer)
 If Me.TxtModFlg.Text <> "R" Then
          
              TxtNoteSerial.Text = ""
   End If
End Sub

Private Sub Dcbranch1_GotFocus()
  If Me.TxtModFlg.Text <> "R" Then
          
              TxtNoteSerial.Text = ""
   End If
 
End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from TBLInsurances order by  IDINS "
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
    '''''''''''''''''''''''''''''''''''
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetBranches Me.dcBranch1
    'Dcombos.GetInsurancesCode Me.DcboBox
    Dcombos.GetAccountingCodes Me.DcbAccount1
    Dcombos.GetAccountingCodes Me.DcbAccount2
    Dcombos.GetAccountingCodes Me.DcbAccount3
    Dcombos.GetAccountingCodes Me.DcbAccount4
    
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
Public Sub FiLLRec()
    On Error GoTo ErrTrap
    If TxtModFlg = "E" Then
    StrSQL = "Delete From TBLInsurancesJoin Where IDINS='" & val(TxtSerial1.Text) & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
    
              StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords


    End If
    
    RsSavRec.Fields("DateM").value = XPDtbTrans.value
    RsSavRec.Fields("DateH").value = Me.Txt_DateHigri.value
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("BranchID1").value = val(Me.dcBranch1.BoundText)
    RsSavRec.Fields("SignalID").value = val(Me.DcboBox.BoundText)
    RsSavRec.Fields("Monthe").value = IIf(val(CmbMonth.ListIndex) <> -1, val((CmbMonth.ListIndex)), Null)
    RsSavRec.Fields("SubYear").value = IIf(val(CboYear.ListIndex) <> -1, val(CboYear.Text), Null)
    RsSavRec.Fields("SudePerce").value = IIf(TxtCivilin.Text <> "", val(TxtCivilin.Text), Null)
    RsSavRec.Fields("UnSudePerce").value = IIf(TxtStay.Text <> "", val(TxtStay.Text), Null)
    RsSavRec.Fields("SudePerce1").value = IIf(TxtCivilin1.Text <> "", val(TxtCivilin1.Text), Null)
    RsSavRec.Fields("UnSudePerce1").value = IIf(TxtStay1.Text <> "", val(TxtStay1.Text), Null)
    RsSavRec.Fields("Acount1").value = Me.DcbAccount1.BoundText
    RsSavRec.Fields("Acount2").value = Me.DcbAccount2.BoundText
    RsSavRec.Fields("Acount3").value = Me.DcbAccount3.BoundText
    RsSavRec.Fields("Acount4").value = Me.DcbAccount4.BoundText
    RsSavRec.Fields("Totall").value = IIf(TotalTXT.Caption <> "", Trim(TotalTXT.Caption), Null)
    If Me.Rd(0).value = True Then
    RsSavRec.Fields("AllBranch").value = 0
    Else
    RsSavRec.Fields("AllBranch").value = 1
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' save grid
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TBLInsurancesJoin Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Grid
       For i = .FixedRows To .Rows - 1
         If .TextMatrix(i, .ColIndex("Ser")) <> "" Then
                RsDevsub.AddNew
                RsDevsub("IDINS").value = Me.TxtSerial1.Text
                RsDevsub("EmpCode").value = IIf((.TextMatrix(i, .ColIndex("Empid"))) = "", Null, .TextMatrix(i, .ColIndex("Empid")))
                RsDevsub("EmpName").value = IIf((.TextMatrix(i, .ColIndex("EmpName"))) = "", Null, .TextMatrix(i, .ColIndex("EmpName")))
                RsDevsub("EmpInsurances").value = IIf((.TextMatrix(i, .ColIndex("EmpInsurances"))) = "", Null, .TextMatrix(i, .ColIndex("EmpInsurances")))
                RsDevsub("InsValue").value = IIf((.TextMatrix(i, .ColIndex("InsValue"))) = "", Null, .TextMatrix(i, .ColIndex("InsValue")))
                RsDevsub("InsTotal").value = IIf((.TextMatrix(i, .ColIndex("InsTotal"))) = "", Null, .TextMatrix(i, .ColIndex("InsTotal")))
                RsDevsub("BranchId").value = IIf((.TextMatrix(i, .ColIndex("BranchId"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchId"))))
                RsDevsub("InsTotal2").value = IIf((.TextMatrix(i, .ColIndex("InsTotal2"))) = "", Null, .TextMatrix(i, .ColIndex("InsTotal2")))
                RsDevsub("CompRate").value = IIf((.TextMatrix(i, .ColIndex("CompRate"))) = "", Null, .TextMatrix(i, .ColIndex("CompRate")))
                RsDevsub("Citirent").value = IIf((.TextMatrix(i, .ColIndex("Citirent"))) = "", Null, .TextMatrix(i, .ColIndex("Citirent")))
                RsDevsub("WorkDays").value = IIf((.TextMatrix(i, .ColIndex("WorkDays"))) = "", Null, .TextMatrix(i, .ColIndex("WorkDays")))
                RsDevsub("BignDateWork").value = IIf((.TextMatrix(i, .ColIndex("BignDateWork"))) = "", Null, .TextMatrix(i, .ColIndex("BignDateWork")))
                
                
                RsDevsub.update
        End If
      Next i
     End With
     createVoucher
     
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
    Dim i As Integer
    ProgressBar1.Visible = True
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("IDINS").value), "", RsSavRec.Fields("IDINS").value): ProgressBar1.value = 10
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("DateM").value), Date, RsSavRec.Fields("DateM").value): ProgressBar1.value = 20
    Txt_DateHigri.value = IIf(IsNull(RsSavRec.Fields("DateH").value), "", RsSavRec.Fields("DateH").value): ProgressBar1.value = 30
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value): ProgressBar1.value = 40
    DcboBox.BoundText = IIf(IsNull(RsSavRec.Fields("SignalID").value), "", RsSavRec.Fields("SignalID").value): ProgressBar1.value = 50
    CmbMonth.ListIndex = IIf(IsNull(RsSavRec.Fields("Monthe").value), "", RsSavRec.Fields("Monthe").value): ProgressBar1.value = 60
    CboYear.Text = IIf(IsNull(RsSavRec.Fields("SubYear").value), "", RsSavRec.Fields("SubYear").value): ProgressBar1.value = 70
    TxtCivilin.Text = IIf(IsNull(RsSavRec.Fields("SudePerce").value), "", RsSavRec.Fields("SudePerce").value): ProgressBar1.value = 80
    TxtStay.Text = IIf(IsNull(RsSavRec.Fields("UnSudePerce").value), "", RsSavRec.Fields("UnSudePerce").value): ProgressBar1.value = 90
    Me.DcbAccount1.BoundText = IIf(IsNull(RsSavRec.Fields("Acount1").value), "", RsSavRec.Fields("Acount1").value): ProgressBar1.value = 100
    Me.DcbAccount2.BoundText = IIf(IsNull(RsSavRec.Fields("Acount2").value), "", RsSavRec.Fields("Acount2").value): ProgressBar1.value = 10
    TotalTXT.Caption = IIf(IsNull(RsSavRec.Fields("Totall").value), "", RsSavRec.Fields("Totall").value): ProgressBar1.value = 20
    Me.DcbAccount3.BoundText = IIf(IsNull(RsSavRec.Fields("Acount3").value), "", RsSavRec.Fields("Acount3").value): ProgressBar1.value = 30
    Me.DcbAccount4.BoundText = IIf(IsNull(RsSavRec.Fields("Acount4").value), "", RsSavRec.Fields("Acount4").value): ProgressBar1.value = 40
    TxtCivilin1.Text = IIf(IsNull(RsSavRec.Fields("SudePerce1").value), "", RsSavRec.Fields("SudePerce1").value): ProgressBar1.value = 50
    TxtStay1.Text = IIf(IsNull(RsSavRec.Fields("UnSudePerce1").value), "", RsSavRec.Fields("UnSudePerce1").value): ProgressBar1.value = 60
    dcBranch1.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID1").value), "", RsSavRec.Fields("BranchID1").value): ProgressBar1.value = 70
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
        Me.TxtNoteID.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)

Me.TxtNoteSerial.Text = IIf(IsNull(RsSavRec("NoteSerial").value), "", RsSavRec("NoteSerial").value)


    If Not (IsNull(RsSavRec.Fields("AllBranch").value)) Then
    If val(RsSavRec.Fields("AllBranch").value) = 0 Then
    Rd(0).value = True
    Else
    Rd(1).value = True
    End If
    End If
    LabCurrRec.Caption = RsSavRec.AbsolutePosition: ProgressBar1.value = 50
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
  sql = "SELECT     dbo.TBLInsurancesJoin.IDINS AS IDINSJOIN,TBLInsurancesJoin.BignDateWork,TBLInsurancesJoin.WorkDays, dbo.TBLInsurancesJoin.EmpCode, dbo.TBLInsurancesJoin.EmpInsurances, dbo.TBLInsurancesJoin.InsValue, "
  sql = sql + "                     dbo.TBLInsurancesJoin.InsTotal, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_ID,"
  sql = sql + "                       dbo.TBLInsurancesJoin.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmployee.Nationality,"
  sql = sql + "                     dbo.TblEmployee.NationalityE, dbo.TBLInsurancesJoin.payed, dbo.TBLInsurancesJoin.Citirent, dbo.TBLInsurancesJoin.InsTotal2,"
  sql = sql + "                       dbo.TBLInsurancesJoin.CompRate"
  sql = sql + "  FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
  sql = sql + "                      dbo.TBLInsurancesJoin ON dbo.TblBranchesData.branch_id = dbo.TBLInsurancesJoin.BranchId LEFT OUTER JOIN"
  sql = sql + "                     dbo.TblEmployee ON dbo.TBLInsurancesJoin.EmpCode = dbo.TblEmployee.Emp_ID"
  sql = sql + "  Where (dbo.TBLInsurancesJoin.IDINS = " & val(TxtSerial1.Text) & ") "
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.Grid
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(Rs1("BranchId").value), "", Rs1("BranchId").value)
                   .TextMatrix(i, .ColIndex("IDINS")) = IIf(IsNull(Rs1("IDINSJOIN").value), "", Rs1("IDINSJOIN").value)
                   .TextMatrix(i, .ColIndex("Empid")) = IIf(IsNull(Rs1("EmpCode").value), 0, Rs1("EmpCode").value)
                   .TextMatrix(i, .ColIndex("EmpCode")) = IIf(IsNull(Rs1("fullcode").value), "", Rs1("fullcode").value)
                   .TextMatrix(i, .ColIndex("Citirent")) = IIf(IsNull(Rs1("Citirent").value), "", Rs1("Citirent").value)
                   .TextMatrix(i, .ColIndex("InsTotal2")) = IIf(IsNull(Rs1("InsTotal2").value), 0, Rs1("InsTotal2").value)
                   .TextMatrix(i, .ColIndex("CompRate")) = IIf(IsNull(Rs1("CompRate").value), 0, Rs1("CompRate").value)
                   
                   
                    .TextMatrix(i, .ColIndex("BignDateWork")) = IIf(IsNull(Rs1("BignDateWork").value), "", Rs1("BignDateWork").value)
                    .TextMatrix(i, .ColIndex("WorkDays")) = IIf(IsNull(Rs1("WorkDays").value), "", Rs1("WorkDays").value)
                   
                                      
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Nationality")) = IIf(IsNull(Rs1("Nationality").value), "", Rs1("Nationality").value)
                   .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                    Else
                    .TextMatrix(i, .ColIndex("Nationality")) = IIf(IsNull(Rs1("NationalityE").value), "", Rs1("NationalityE").value)
                   .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(Rs1("Emp_Namee").value), "", Rs1("Emp_Namee").value)
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
                    End If
                   .TextMatrix(i, .ColIndex("EmpInsurances")) = IIf(IsNull(Rs1("EmpInsurances").value), "", Rs1("EmpInsurances").value)
                   .TextMatrix(i, .ColIndex("InsValue")) = IIf(IsNull(Rs1("InsValue").value), "", Rs1("InsValue").value)
                   .TextMatrix(i, .ColIndex("InsTotal")) = IIf(IsNull(Rs1("InsTotal").value), "", Rs1("InsTotal").value)
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub
Private Sub ISButton2_Click()
ISButton4_Click
If Me.TxtModFlg.Text <> "R" Then
  On Error GoTo ErrTrap
  If Rd(1).value = True Then
       If Dcbranch.Text = "" And val(Dcbranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Dcbranch.SetFocus
            Exit Sub
            Else
            MsgBox "Write Branch Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            Dcbranch.SetFocus
         End If
     End If
   End If
  If GetProInsurance() = False Then
  If SystemOptions.UserInterface = ArabicInterface Then
  MsgBox "Ì—ÃÏ ÷»ÿ ≈⁄œ«œ  «· √„Ì‰"
  Else
  MsgBox "Please GOSI Settings"
  End If
  Exit Sub
  End If
    '+++++++++++++++++++++++++++++++++++++++++++++++

     '  If CmbMonth.text = "" Then
     '   If SystemOptions.UserInterface = ArabicInterface Then
     '       MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·‘Â— ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
     '       CmbMonth.SetFocus
     '        Exit Sub
     '        Else
     '       MsgBox "Write Month ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
     '       CmbMonth.SetFocus
     '       Exit Sub
     '       End If
     ' End If
     'If CboYear.text = "" Then
     '   If SystemOptions.UserInterface = ArabicInterface Then
     '       MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·”‰… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
     '       CboYear.SetFocus
     '        Exit Sub
     '        Else
     '       MsgBox "Write Year ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
     '       CboYear.SetFocus
     '       Exit Sub
     '       End If
     'End If
    '  If TxtCivilin.text = "" Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ ‰”»… «·„Ê«ÿ‰Ì‰ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
    '        TxtCivilin.SetFocus
    '         Exit Sub
    '         Else
    '        MsgBox "Write Citizens Percentage ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    '        TxtCivilin.SetFocus
    '        Exit Sub
    '        End If
    ' End If
    '     If TxtStay.text = "" Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ ‰”»… «·„ﬁÌ„Ì‰ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
    '        TxtStay.SetFocus
    '         Exit Sub
    '         Else
    '        MsgBox "Write Residents Percentage ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    '        TxtStay.SetFocus
    '        Exit Sub
    ''        End If
    ' End If
    '     If DataCombo1.text = "" Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·Õ”«» «·„œÌ‰ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
    '        DataCombo1.SetFocus
    '         Exit Sub
    '         Else
    '        MsgBox "Write Debit Account ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    '        DataCombo1.SetFocus
    '        Exit Sub
    '        End If
    ' End If
    '     If DataCombo2.text = "" Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·Õ”«» «·œ«∆‰ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
    '        DataCombo2.SetFocus
    '         Exit Sub
    '         Else
    '        MsgBox "Write Creditor Account ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    '        DataCombo2.SetFocus
    '        Exit Sub
         '   End If
    ' End If
     FillTextGridData
ErrTrap:
End If
   End Sub
  Sub FillTextGridData()
  On Error GoTo ErrTrap
   Dim Rs1 As ADODB.Recordset
   Set Rs1 = New ADODB.Recordset
   Dim sql As String
       

  ' sql = " SELECT     TOP 100 PERCENT dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
  ' sql = sql & "                   SUM(dbo.EmpSalaryComponent.[Value]) AS Salary, dbo.mofrad.Insurances, dbo.jopstatus.Insurances AS InsurancesJob, dbo.jopstatus.resignationInt,"
  ' sql = sql & "                   dbo.TblEmployee.InsuranceState, dbo.TblEmployee.NationalityE, dbo.TblEmployee.Nationality, dbo.TblEmployee.InstanceDateM, dbo.TblEmployee.BranchId,"
  ' sql = sql & "                   dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_nameE"
  ' sql = sql & "  FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
  ' sql = sql & "                   dbo.TblEmployee ON dbo.TblBranchesData.branch_id = dbo.TblEmployee.BranchId LEFT OUTER JOIN"
  ' sql = sql & "                   dbo.jopstatus ON dbo.TblEmployee.jopstatusid = dbo.jopstatus.id LEFT OUTER JOIN"
  ' sql = sql & "                   dbo.mofrad RIGHT OUTER JOIN"
  ' sql = sql & "                  dbo.EmpSalaryComponent ON dbo.mofrad.id = dbo.EmpSalaryComponent.AccountCode ON dbo.TblEmployee.Emp_ID = dbo.EmpSalaryComponent.emp_ID"
  ' sql = sql & " WHERE     (DATEPART(year, dbo.EmpSalaryComponent.EntIncresDataM) <" & year(XPDtbTrans) & ") OR"
  ' sql = sql & "                     (DATEPART(year, dbo.EmpSalaryComponent.EntIncresDataM) IS NULL)"
  ' If Rd(1).value = True Then
  '  sql = sql & " and dbo.TblEmployee.BranchId =" & val(Me.Dcbranch.BoundText) & ""
  ' End If
  ' sql = sql & "GROUP BY dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.mofrad.Insurances,"
  ' sql = sql & "                     dbo.jopstatus.Insurances, dbo.jopstatus.resignationInt, dbo.TblEmployee.InsuranceState, dbo.TblEmployee.NationalityE, dbo.TblEmployee.Nationality,"
  ' sql = sql & "                    dbo.TblEmployee.InstanceDateM , dbo.TblEmployee.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE"
  ' sql = sql & "    HAVING      (dbo.mofrad.Insurances = 1) AND (dbo.jopstatus.Insurances = 1) AND (dbo.jopstatus.resignationInt IS NULL) AND (dbo.TblEmployee.InsuranceState = 1) AND"
  ' sql = sql & "                     (dbo.TblEmployee.InstanceDateM <=" & SQLDate(Me.XPDtbTrans.value, True) & " OR"
  ' sql = sql & "                     dbo.TblEmployee.InstanceDateM IS NULL) OR"
  ' sql = sql & "                     (dbo.jopstatus.resignationInt <> 2) AND (dbo.jopstatus.resignationInt <> 1)"
  ' sql = sql & " ORDER BY dbo.TblEmployee.Emp_ID"
 sql = "  SELECT     TOP 100 PERCENT dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
 sql = sql & "                     SUM(dbo.EmpSalaryComponent.[Value]) AS Salary, dbo.mofrad.Insurances, dbo.jopstatus.Insurances AS InsurancesJob, dbo.jopstatus.resignationInt,"
 sql = sql & "                     dbo.TblEmployee.InsuranceState, dbo.TblEmployee.NationalityE, dbo.TblEmployee.Nationality, dbo.TblEmployee.InstanceDateM, dbo.TblEmployee.BranchId,"
 sql = sql & "                      dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_nameE,dbo.TblEmployee.BignDateWork"
 sql = sql & " FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
 sql = sql & "                     dbo.TblEmployee ON dbo.TblBranchesData.branch_id = dbo.TblEmployee.BranchId LEFT OUTER JOIN"
 sql = sql & "                     dbo.jopstatus ON dbo.TblEmployee.jopstatusid = dbo.jopstatus.id LEFT OUTER JOIN"
 sql = sql & "                     dbo.mofrad RIGHT OUTER JOIN"
 sql = sql & "                     dbo.EmpSalaryComponent ON dbo.mofrad.id = dbo.EmpSalaryComponent.mofrad_type ON dbo.TblEmployee.Emp_ID = dbo.EmpSalaryComponent.emp_ID"
 
 
 sql = sql & "  WHERE     (DATEPART(year, dbo.EmpSalaryComponent.EntIncresDataM) < " & year(XPDtbTrans) & ") OR"
 sql = sql & "                     (DATEPART(year, dbo.EmpSalaryComponent.EntIncresDataM) IS NULL)"
  
  If Rd(1).value = True Then
   sql = sql & " and dbo.TblEmployee.BranchId =" & val(Me.Dcbranch.BoundText) & ""
   End If
  
 sql = sql & "  AND (dbo.TblEmployee.InsuranceState = 1)"
  
 sql = sql & " GROUP BY dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name,dbo.TblEmployee.BignDateWork, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.mofrad.Insurances,"
 sql = sql & "                     dbo.jopstatus.Insurances, dbo.jopstatus.resignationInt, dbo.TblEmployee.InsuranceState, dbo.TblEmployee.NationalityE, dbo.TblEmployee.Nationality,"
 sql = sql & "                     dbo.TblEmployee.InstanceDateM , dbo.TblEmployee.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE"
 sql = sql & " HAVING      (dbo.mofrad.Insurances = 1) AND (dbo.jopstatus.Insurances = 1) AND (dbo.jopstatus.resignationInt IS NULL) AND (dbo.TblEmployee.InsuranceState = 1) AND"
 sql = sql & "                     (dbo.TblEmployee.InstanceDateM <= " & SQLDate(Me.XPDtbTrans.value, True) & " OR"
 sql = sql & "                     dbo.TblEmployee.InstanceDateM IS NULL) OR"
 sql = sql & "                     (dbo.jopstatus.resignationInt <> 2) AND (dbo.jopstatus.resignationInt <> 1)"
 sql = sql & " ORDER BY dbo.TblEmployee.Emp_ID"
   
      Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.Grid
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("IDINS")) = TxtSerial1.Text
                   .TextMatrix(i, .ColIndex("EmpCode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(i, .ColIndex("BignDateWork")) = IIf(IsNull(Rs1("BignDateWork").value), "", Rs1("BignDateWork").value)
                    Dim MDate As Date
                    MDate = "01-" & CmbMonth.ListIndex + 1 & "-" & val(CboYear.Text)
                     
                    Dim ss As Date
                    .TextMatrix(i, .ColIndex("WorkDays")) = 30
                    ss = MonthLastDay(MDate)
                   If .TextMatrix(i, .ColIndex("BignDateWork")) <> "" Then
                        If CmbMonth.ListIndex + 1 = Month(.TextMatrix(i, .ColIndex("BignDateWork"))) And val(CboYear.Text) = year(.TextMatrix(i, .ColIndex("BignDateWork"))) Then
                            .TextMatrix(i, .ColIndex("WorkDays")) = DateDiff("d", .TextMatrix(i, .ColIndex("BignDateWork")), ss)
                        End If
                    End If
                   
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                   .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   .TextMatrix(i, .ColIndex("Nationality")) = IIf(IsNull(Rs1("Nationality").value), "", Rs1("Nationality").value)
                   
                   
                   If Not (IsNull(Rs1("Nationality").value)) Then
                    If Rs1("Nationality").value = "”⁄ÊœÌ" Or Rs1("Nationality").value = "”⁄ÊœÏ" Or Rs1("Nationality").value = "Saudi" Then
                   .TextMatrix(i, .ColIndex("InsValue")) = val(TxtCivilin.Text)
                   .TextMatrix(i, .ColIndex("CompRate")) = val(TxtCivilin1.Text)
                   .TextMatrix(i, .ColIndex("Citirent")) = "„Ê«ÿ‰"
                    Else
                   .TextMatrix(i, .ColIndex("InsValue")) = val(TxtStay.Text)
                   .TextMatrix(i, .ColIndex("CompRate")) = val(TxtStay1.Text)
                   .TextMatrix(i, .ColIndex("Citirent")) = "„ﬁÌ„"
                    End If
                      Else
                    .TextMatrix(i, .ColIndex("CompRate")) = 0
                    .TextMatrix(i, .ColIndex("InsValue")) = 0
                    .TextMatrix(i, .ColIndex("Citirent")) = 0
                   End If
                    Else
                    .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
                    .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(Rs1("Emp_Namee").value), "", Rs1("Emp_Namee").value)
                    .TextMatrix(i, .ColIndex("Nationality")) = IIf(IsNull(Rs1("NationalityE").value), "", Rs1("NationalityE").value)
                    If Not (IsNull(Rs1("NationalityE").value)) Then
                    If Rs1("NationalityE").value = UCase$("SAUDI") Then
                    .TextMatrix(i, .ColIndex("CompRate")) = val(TxtCivilin1.Text)
                    .TextMatrix(i, .ColIndex("InsValue")) = val(TxtCivilin.Text)
                    .TextMatrix(i, .ColIndex("Citirent")) = "Citize"
                     Else
                     .TextMatrix(i, .ColIndex("CompRate")) = val(TxtStay1.Text)
                    .TextMatrix(i, .ColIndex("InsValue")) = val(TxtStay.Text)
                    .TextMatrix(i, .ColIndex("Citirent")) = "Resident"
                     End If
                     
                   Else
                    .TextMatrix(i, .ColIndex("CompRate")) = 0
                    .TextMatrix(i, .ColIndex("InsValue")) = 0
                    .TextMatrix(i, .ColIndex("Citirent")) = 0
                   End If
                   End If
                   .TextMatrix(i, .ColIndex("EmpInsurances")) = IIf(IsNull(Rs1("Salary").value), "", Rs1("Salary").value)
                   .TextMatrix(i, .ColIndex("InsTotal")) = val(val(.TextMatrix(i, .ColIndex("EmpInsurances"))) * val(.TextMatrix(i, .ColIndex("InsValue"))) / 100)
                   .TextMatrix(i, .ColIndex("InsTotal2")) = val(val(.TextMatrix(i, .ColIndex("EmpInsurances"))) * val(.TextMatrix(i, .ColIndex("CompRate"))) / 100)
                   
                  .TextMatrix(i, .ColIndex("InsTotal")) = val(.TextMatrix(i, .ColIndex("InsTotal"))) / 30 * val(.TextMatrix(i, .ColIndex("WorkDays")))
                  .TextMatrix(i, .ColIndex("InsTotal2")) = val(.TextMatrix(i, .ColIndex("InsTotal2"))) / 30 * val(.TextMatrix(i, .ColIndex("WorkDays")))
                  
                   
                   .TextMatrix(i, .ColIndex("Empid")) = IIf(IsNull(Rs1("Emp_id").value), "", Rs1("Emp_id").value)
                   .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(Rs1("BranchId").value), "", Rs1("BranchId").value)
                   Rs1.MoveNext
             Next i
             TotalCelll
             
        End With
        Exit Sub
ErrTrap:
 End Sub
 Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 On Error GoTo ErrTrap
      With Grid
                If Col = .ColIndex("EmpCode") Then
                Cancel = True
                End If
                If Col = .ColIndex("EmpName") Then
                Cancel = True
                End If
       End With
ErrTrap:
End Sub
 Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 On Error GoTo ErrTrap
      TotalCelll
ErrTrap:
  End Sub
 Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
  On Error GoTo ErrTrap
     TotalCelll
ErrTrap:
 End Sub
  Sub TotalCelll()
  On Error GoTo ErrTrap
     Dim i As Integer
     With Me.Grid
     For i = 1 To .Rows - 1
        .TextMatrix(i, .ColIndex("InsTotal")) = val(val(.TextMatrix(i, .ColIndex("EmpInsurances"))) * val(.TextMatrix(i, .ColIndex("InsValue"))) / 100)
      Next i
     End With
     TotalGrid
ErrTrap:
 End Sub
  Sub TotalGrid()
  On Error GoTo ErrTrap
    With Me.Grid
      Me.TotalTXT.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("InsTotal2"), .Rows - 1, .ColIndex("InsTotal2"))
    End With
ErrTrap:
  End Sub
Private Sub ISButton3_Click()
On Error Resume Next
    With Me.Grid
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    TotalCelll
 End Sub
 Private Sub ISButton4_Click()
 On Error Resume Next
 Me.Grid.Clear flexClearScrollable, flexClearEverything
 cleargriid
 TotalTXT.Caption = 0
 End Sub

Private Sub Rd_Click(Index As Integer)
If Index = 0 Then
Dcbranch.BoundText = 0
Dcbranch.Enabled = False
Else
Dcbranch.Enabled = True
End If
End Sub



Private Sub Txt_DateHigri_LostFocus()
 If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            XPDtbTrans.value = ToGregorianDate(Txt_DateHigri.value)
   End If
End Sub

' change date to hj
  Private Sub XPDtbTrans_Change()
  If Me.TxtModFlg.Text <> "R" Then
              Txt_DateHigri.value = ToHijriDate(XPDtbTrans.value)
              TxtNoteSerial.Text = ""
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
       If dcBranch1.Text = "" Or (dcBranch1.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄ «·ﬁ«∆„ »«·Õ—ﬂ…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Dcbranch.SetFocus
            Exit Sub
            Else
            MsgBox "Please select Branch Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            dcBranch1.SetFocus
         End If
         End If
    If ChekDolpe() = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ Õ›Ÿ Â–« «·«ÀÌ«  ·«‰Â „ÊÃÊœ „”»ﬁ«"
    Else
    MsgBox "Can Not Saved this Process because  it is Already exist"
    End If
    Exit Sub
    End If
  If ChekExistRecord() = True Then
  If SystemOptions.UserInterface = ArabicInterface Then
  MsgBox " ·«Ì„ﬂ‰ «·Õ›Ÿ Â–« «·”‰œ „ÊÃÊœ „”»ﬁ« "
  Else
  MsgBox "You can not save this authority already exists"
  End If
  Exit Sub
  End If
    If Rd(1).value = True Then
      If Dcbranch.Text = "" Or (Dcbranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Dcbranch.SetFocus
            Exit Sub
            Else
            MsgBox "Write Branch Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            Dcbranch.SetFocus
         End If
     End If
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++
    ' If DcboBox.text = "" And val(DcboBox.BoundText) = 0 Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·„›—œ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
    '        DcboBox.SetFocus
    '         Exit Sub
    '         Else
    '        MsgBox "Write Single ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    '        DcboBox.SetFocus
    '        Exit Sub
    '        End If
    ' End If
      If CmbMonth.Text = "" And val(CmbMonth.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·‘Â— ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            CmbMonth.SetFocus
             Exit Sub
             Else
            MsgBox "Write Month ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CmbMonth.SetFocus
            Exit Sub
            End If
      End If
     If CboYear.Text = "" And val(CboYear.ListIndex) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·”‰… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            CboYear.SetFocus
             Exit Sub
             Else
            MsgBox "Write Year ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CboYear.SetFocus
            Exit Sub
            End If
     End If
     If ChekPayedSalary(val(CboYear.Text), val(CmbMonth.ListIndex) + 1, val(dcBranch1.BoundText)) = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ Õ–› ﬁÌœ «·—Ê« »  ··‘Â—  "
            Else
            MsgBox "Delete Salary Allocation JL"
            End If
            Exit Sub
    End If
     ' If TxtCivilin.text = "" Then
     '   If SystemOptions.UserInterface = ArabicInterface Then
     '       MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ ‰”»… «·„Ê«ÿ‰Ì‰ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
     '       TxtCivilin.SetFocus
     '        Exit Sub
     '        Else
     '       MsgBox "Write Citizens Percentage ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
     '       TxtCivilin.SetFocus
     '       Exit Sub
     '       End If
     'End If
     '    If TxtStay.text = "" Then
     '   If SystemOptions.UserInterface = ArabicInterface Then
     '       MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ ‰”»… «·„ﬁÌ„Ì‰ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
     '       TxtStay.SetFocus
     '        Exit Sub
     '        Else
     '       MsgBox "Write Residents Percentage ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
     '       TxtStay.SetFocus
     '       Exit Sub
     '       End If
     'End If
      '   If DataCombo1.text = "" Then
      '  If SystemOptions.UserInterface = ArabicInterface Then
      '      MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·Õ”«» «·„œÌ‰ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
      '      DataCombo1.SetFocus
      '       Exit Sub
      '       Else
      '      MsgBox "Write Debit Account ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
      '      DataCombo1.SetFocus
      '      Exit Sub
    '        End If
    ' End If
      '   If DataCombo2.text = "" Then
      '  If SystemOptions.UserInterface = ArabicInterface Then
      '      MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·Õ”«» «·œ«∆‰ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
      '      DataCombo2.SetFocus
      '       Exit Sub
      '       Else
      '      MsgBox "Write Creditor Account ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
      '      DataCombo2.SetFocus
      '      Exit Sub
      '      End If
    ' End If
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
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TBLInsurances", "IDINS", "")
    Me.TxtSerial1.Text = StrRecID
    
    RsSavRec.AddNew
    RsSavRec.Fields("IDINS").value = IIf(StrRecID <> "", StrRecID, Null)
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
Public Function FindRec(ByVal RecId As Long, Optional NoteID As Long = 0)
    On Error GoTo ErrTrap
    If NoteID = 0 Then
    RsSavRec.find "IDINS=" & RecId, , adSearchForward, 1
      
    Else
      RsSavRec.find "NoteID=" & NoteID, , adSearchForward, 1
    End If
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
                If ChekPaye() = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ Õ–› Â–Â «·⁄„·Ì… ·«‰Â« „— »ÿÂ »ﬁÌœ «·«” Õﬁ«ﬁ"
    Else
    MsgBox "Can not Delete This Process"
    End If
    Exit Sub
    End If
                RsSavRec.find "IDINS=" & val(TxtSerial1.Text), , adSearchForward, 1
              
               '''''''''''''''''''''''''''''''
                 StrSQL = "Delete From TBLInsurancesJoin Where IDINS='" & val(TxtSerial1.Text) & "'"
                 Cn.Execute StrSQL, , adExecuteNoRecords
                  
              StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords

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
       ' FillGridWithData
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
'Private Sub Grid_EnterCell()
 '   On Error GoTo ErrTrap
  '  FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("Ser")))
'ErrTrap:
'End Sub
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
    If ChekPaye() = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰  ⁄œÌ· Â–Â «·⁄„·Ì… ·«‰Â« „— »ÿÂ »ﬁÌœ «·«” Õﬁ«ﬁ"
    Else
    MsgBox "Can not Update This Process"
    End If
    Exit Sub
    End If
        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
        Me.dcBranch1.BoundText = branch_id
        Frm2.Enabled = True
        Me.dcBranch1.SetFocus
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
Function GetProInsurance() As Boolean
Dim sql As String
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
sql = "select * from TblSocialInsurance order by  ID "
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
GetProInsurance = True
   DcbAccount1.BoundText = IIf(IsNull(Rs6("Acount_Code1").value), "", Rs6("Acount_Code1").value)
    DcbAccount2.BoundText = IIf(IsNull(Rs6("Acount_Code2").value), "", Rs6("Acount_Code2").value)
    DcbAccount3.BoundText = IIf(IsNull(Rs6("Acount_Code3").value), "", Rs6("Acount_Code3").value)
    DcbAccount4.BoundText = IIf(IsNull(Rs6("Acount_Code4").value), "", Rs6("Acount_Code4").value)
    TxtStay.Text = IIf(IsNull(Rs6("ResidentVal1").value), 0, Rs6("ResidentVal1").value)
    TxtStay1.Text = IIf(IsNull(Rs6("ResidentVal2").value), 0, Rs6("ResidentVal2").value)
    TxtCivilin.Text = IIf(IsNull(Rs6("CitizenVal1").value), 0, Rs6("CitizenVal1").value)
    TxtCivilin1.Text = IIf(IsNull(Rs6("CitizenVal2").value), 0, Rs6("CitizenVal2").value)
Else
GetProInsurance = False
   DcbAccount1.BoundText = 0
    DcbAccount2.BoundText = 0
    DcbAccount3.BoundText = 0
    DcbAccount4.BoundText = 0
    TxtStay.Text = 0
    TxtStay1.Text = 0
    TxtCivilin.Text = 0
    TxtCivilin1.Text = 0
End If
End Function
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
    Rd(0).value = True
    
    Rd_Click (0)
    CmbType.ListIndex = 0
    Me.DCboUserName.BoundText = user_id
    Me.dcBranch1.BoundText = branch_id
    CmbType.ListIndex = 0
    dcBranch1.SetFocus
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
Private Sub BtnPrint_Click()
On Error GoTo ErrTrap
  If val(Me.TxtSerial1.Text) <> 0 Then
      print_report
  End If
ErrTrap:
End Sub
Private Sub ISButton1_Click()
On Error GoTo ErrTrap
   If val(Me.TxtSerial1.Text) <> 0 Then
       print_report
   End If
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String)
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    sql = "SELECT     dbo.TBLInsurances.IDINS, dbo.TBLInsurances.DateM, dbo.TBLInsurances.DateH, dbo.TBLInsurances.SignalID, dbo.TBLInsurances.Monthe, "
    sql = sql & "                  dbo.TBLInsurances.SubYear, dbo.TBLInsurances.SudePerce, dbo.TBLInsurances.UnSudePerce, dbo.TBLInsurances.Totall,"
    sql = sql & "                  dbo.TBLInsurancesJoin.IDINS AS IDINSJOIN, dbo.TBLInsurancesJoin.EmpInsurances, dbo.TBLInsurancesJoin.InsValue, dbo.TBLInsurancesJoin.InsTotal,"
    sql = sql & "                  dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Fullcode, dbo.TblEmployee.NationalityE,"
    sql = sql & "                  dbo.TBLInsurancesJoin.payed, dbo.TBLInsurancesJoin.Citirent, dbo.TBLInsurancesJoin.InsTotal2, dbo.TBLInsurancesJoin.CompRate,"
    sql = sql & "                  dbo.TBLInsurancesJoin.BranchId, TblBranchesData_1.branch_name, TblBranchesData_1.branch_namee, dbo.TBLInsurances.BranchID AS HBranchID,"
    sql = sql & "                  TblBranchesData_1.branch_name AS Hbranch_name, TblBranchesData_1.branch_namee AS Hbranch_namee, dbo.TBLInsurances.AllBranch,"
    sql = sql & "                  dbo.TBLInsurances.SudePerce1 , dbo.TBLInsurances.UnSudePerce1"
    sql = sql & "   FROM         dbo.TblBranchesData TblBranchesData_1 LEFT OUTER JOIN"
    sql = sql & "                  dbo.TBLInsurances ON TblBranchesData_1.branch_id = dbo.TBLInsurances.BranchID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData TblBranchesData_2 RIGHT OUTER JOIN"
    sql = sql & "                  dbo.TBLInsurancesJoin ON TblBranchesData_2.branch_id = dbo.TBLInsurancesJoin.BranchId LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblEmployee ON dbo.TBLInsurancesJoin.EmpCode = dbo.TblEmployee.Emp_ID ON dbo.TBLInsurances.IDINS = dbo.TBLInsurancesJoin.IDINS"
    sql = sql & " Where (dbo.TBLInsurances.IDINS = " & val(TxtSerial1.Text) & ")"
                       
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
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
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

    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic

   ' form name
    Me.Caption = "Proof of  Insurances"
    ' labell name
    Label1(35).Caption = "GL"
    Command9.Caption = "Print  GL"
    Command8.Caption = "Account"
    Frame9.Caption = "Data"
    Me.Label1(2).Caption = Me.Caption
    Me.lblcode.Caption = "Code"
    Me.lbldate.Caption = "Date"
    Me.lblhjdate.Caption = "HJ Date"
    Me.Label3.Caption = "Branch"
    Me.Labelbank.Caption = "Select Single "
    '''''''''''''' next
    ELe(3).Caption = "Select period"
    lbl(0).Caption = "Month"
    lbl(2).Caption = "Year"
    Rd(0).RightToLeft = False
    Rd(1).RightToLeft = False
    Rd(0).Caption = "All Branch"
    Rd(1).Caption = "Select Branch"
    '''''''''''''''''''''''' next
    Frame1.Caption = "Select Percentage"
    lbl(3).Caption = "citizens Percentage"
    lbl(1).Caption = "Residents Percentage"
    Frame2.Caption = "Select Account"
    lbl(6).Caption = "Debit account"
    lbl(7).Caption = "Creditor account"
    Label2(3).Caption = "Total"
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
        .TextMatrix(0, .ColIndex("EmpCode")) = "Emp Code"
        .TextMatrix(0, .ColIndex("EmpName")) = "Emp Name"
        .TextMatrix(0, .ColIndex("EmpInsurances")) = "Total Salary With Insurances"
        .TextMatrix(0, .ColIndex("InsValue")) = "Employee Percentage"
        .TextMatrix(0, .ColIndex("InsTotal")) = "Employee Value"
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
        .TextMatrix(0, .ColIndex("Nationality")) = "Nationality Name"
         .TextMatrix(0, .ColIndex("Citirent")) = "Citizen/Resident"
        .TextMatrix(0, .ColIndex("CompRate")) = "Company Percentage"
        .TextMatrix(0, .ColIndex("InsTotal2")) = "Company Value"
        
    End With
ErrTrap:
End Sub
Function ChekDolpe() As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
ChekDolpe = False
sql = "select * from TBLInsurances where SubYear=" & val(CboYear.ListIndex) & " and Monthe=" & val(CmbMonth.ListIndex) & "and ((AllBranch = 0) or (BranchID =" & val(Me.Dcbranch.BoundText) & "))and IDINS<>" & val(TxtSerial1.Text) & " "
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
ChekDolpe = True
Else
ChekDolpe = False
End If
End Function
Private Sub Dcbranch_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  DcboBox.SetFocus
  End If
ErrTrap:
End Sub
Private Sub DcboBox_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  CmbMonth.SetFocus
  End If
ErrTrap:
End Sub
Private Sub CmbMonth_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  CboYear.SetFocus
  End If
ErrTrap:
End Sub
Private Sub CboYear_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  TxtCivilin.SetFocus
  End If
ErrTrap:
End Sub
Private Sub TxtCivilin_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  TxtStay.SetFocus
  End If
ErrTrap:
End Sub
Private Sub TxtStay_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
'  DataCombo1.SetFocus
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
   My_SQL = "TBLInsurances"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub
Function ChekPaye() As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
ChekPaye = False
sql = "SELECT     IDINS, payed"
sql = sql & " From dbo.TBLInsurancesJoin"
sql = sql & " Where (payed = 1) And (IDINS = " & val(TxtSerial1.Text) & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
ChekPaye = True
Else
ChekPaye = False
End If
End Function


'Public Function MonthLastDay(ByVal dCurrDate As Date) As Date
'    Dim dFirstDayNextMonth As Date
'
'    MonthLastDay = Empty
'    dCurrDate = Format(dCurrDate, "DD/MM/YYYY")
'    dFirstDayNextMonth = DateSerial(CInt(Format(dCurrDate, "yyyy")), CInt(Format(dCurrDate, "mm")) + 1, 1)
'    MonthLastDay = DateAdd("d", -1, dFirstDayNextMonth)
'
'    Exit Function
'
'End Function

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
    For i = 2006 To 2050
        CboYear.AddItem i
        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If
    Next
    CboYear.ListIndex = IntDefIndex
End Sub
'+++++++++++++++++++++++++++++++++ en
