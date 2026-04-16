VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{AA91FA8F-BC1E-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseWizard.ocx"
Begin VB.Form FrmRegisteration 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ”ÃÌ· «·»—‰«„Ã"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "FrmRegisteration.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleExpire 
      Height          =   375
      Left            =   90
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   4560
      Width           =   6645
      _cx             =   11721
      _cy             =   661
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
      BackColor       =   14737632
      ForeColor       =   -2147483630
      FloodColor      =   128
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   2
      FloodPercent    =   1
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
   End
   Begin ImpulseWizard.ISWizard WzrdMain 
      Height          =   4425
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   7805
      BackColor       =   16777215
      FillStyle       =   1
      FinishEnabled   =   0   'False
      ShowStepNumber  =   -1  'True
      NumberOfSteps   =   5
      ScaleWidth      =   6795
      ScaleHeight     =   4425
      ActiveButtons   =   -1  'True
      ColorButton     =   16777215
      ControlCount    =   6
      Control(1).Name =   "Ele"
      Control(1).Index=   0
      Control(1).WizardStep=   1
      Control(1).Visible=   -1  'True
      Control(1).InternalID=   "A0E8DA14C4"
      Control(2).Name =   "Ele"
      Control(2).Index=   1
      Control(2).WizardStep=   2
      Control(2).Visible=   0   'False
      Control(2).InternalID=   "1FD623AB95"
      Control(3).Name =   "Lbl"
      Control(3).Index=   6
      Control(3).WizardStep=   5
      Control(3).Visible=   0   'False
      Control(3).InternalID=   "7AF0D3EB3E"
      Control(4).Name =   "Ele"
      Control(4).Index=   2
      Control(4).WizardStep=   4
      Control(4).Visible=   0   'False
      Control(4).InternalID=   "DA902CB31B"
      Control(5).Name =   "C1Elastic1"
      Control(5).Index=   -1
      Control(5).WizardStep=   3
      Control(5).Visible=   0   'False
      Control(5).InternalID=   "F0B7ABF330"
      Control(6).Name =   "Ele"
      Control(6).Index=   3
      Control(6).WizardStep=   5
      Control(6).Visible=   0   'False
      Control(6).InternalID=   "CF3C5A5069"
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3915
         Index           =   0
         Left            =   960
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   0
         Width           =   5835
         _cx             =   10292
         _cy             =   6906
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
         BackColor       =   12634304
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
         Begin VB.Image Image3 
            Height          =   1080
            Left            =   3720
            Picture         =   "FrmRegisteration.frx":058A
            Top             =   2400
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "‘ﬂ—« ·≈” Œœ«„ﬂ„ »—‰«„Ã  œÌ‰«„Ìﬂ »«Ì  ··Õ”«»«  ... «·»—‰«„Ã «·√ﬁÊÏ Ê«·√”Â·"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   765
            Index           =   1
            Left            =   240
            TabIndex        =   55
            Top             =   840
            Width           =   5445
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C8C0&
            Caption         =   "√Â·« »ﬂ„ ›Ï „⁄«·Ã  ”ÃÌ· «·»—‰«„Ã"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   525
            Index           =   27
            Left            =   30
            TabIndex        =   52
            Top             =   240
            Width           =   5775
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "Â–« «·„⁄«·Ã ”Ê› Ì’Õ»ﬂ„ ŒÿÊ… »ŒÿÊ… · ”ÃÌ· «·»—‰«„Ã Ê–·ﬂ Õ„«Ì… ·ÕﬁÊﬁﬂ„ ›Ï «·Õ’Ê· ⁄·Ï «·œ⁄„ «·›‰Ï «·„»«‘— „‰ «·‘—ﬂ… «Ê „‰ «·„Ê“⁄ "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   885
            Index           =   6
            Left            =   150
            TabIndex        =   51
            Top             =   1710
            Width           =   5565
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   1245
            Index           =   2
            Left            =   60
            TabIndex        =   50
            Top             =   2640
            Width           =   3375
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   0
            Left            =   3450
            Picture         =   "FrmRegisteration.frx":45F4
            Top             =   2610
            Width           =   240
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "„‰ ›÷·ﬂ ≈÷€ÿ Next ··„ «»⁄…"
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   3
            Left            =   3480
            TabIndex        =   49
            Top             =   3570
            Width           =   2265
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   5
            Left            =   4890
            Picture         =   "FrmRegisteration.frx":497E
            Top             =   300
            Width           =   240
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3915
         Index           =   1
         Left            =   1.00960e5
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   0
         Width           =   5835
         _cx             =   10292
         _cy             =   6906
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
         BackColor       =   12634304
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
         Begin VB.OptionButton OptRegType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "√—Ìœ  ”ÃÌ· «·»—‰«„Ã „‰ Œ·«· «·√ ’«· »«·Â« ›"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   2
            Left            =   1710
            TabIndex        =   42
            Top             =   1650
            Width           =   3615
         End
         Begin VB.OptionButton OptRegType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "·œÏ „·› «·Õ„«Ì… Ê«—Ìœ «· ”ÃÌ· „‰ Œ·«·Â"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   0
            Left            =   1710
            TabIndex        =   41
            Top             =   2370
            Width           =   3615
         End
         Begin VB.OptionButton OptRegType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "√—Ìœ «· ”ÃÌ· „‰ Œ·«· «·√‰ —‰ "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   1
            Left            =   1710
            TabIndex        =   40
            Top             =   1020
            Value           =   -1  'True
            Width           =   3615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "„‰ ›÷·ﬂ ≈÷€ÿ Next ··„ «»⁄…"
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   4
            Left            =   3480
            TabIndex        =   47
            Top             =   3570
            Width           =   2265
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C8C0&
            Caption         =   "√Œ — ÿ—Ìﬁ… «· ”ÃÌ· «·„·«∆„… ·ﬂ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   525
            Index           =   7
            Left            =   30
            TabIndex        =   46
            Top             =   240
            Width           =   5775
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "„·ÕÊŸ…:-"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   705
            Index           =   14
            Left            =   30
            TabIndex        =   45
            Top             =   2730
            Width           =   5295
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "„·ÕÊŸ…:-"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   30
            TabIndex        =   44
            Top             =   1350
            Width           =   5295
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "„·ÕÊŸ…:-"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   11
            Left            =   30
            TabIndex        =   43
            Top             =   2010
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   2
            Left            =   5370
            Picture         =   "FrmRegisteration.frx":4D08
            Top             =   990
            Width           =   240
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   3
            Left            =   5370
            Picture         =   "FrmRegisteration.frx":5292
            Top             =   1620
            Width           =   240
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   4
            Left            =   5400
            Picture         =   "FrmRegisteration.frx":561C
            Top             =   2340
            Width           =   240
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   6
            Left            =   4710
            Picture         =   "FrmRegisteration.frx":59A6
            Top             =   270
            Width           =   240
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3915
         Index           =   2
         Left            =   1.00960e5
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   0
         Width           =   5835
         _cx             =   10292
         _cy             =   6906
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
         BackColor       =   12634304
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
         Begin VB.Frame FraCustomerInfo 
            BackColor       =   &H00C0C8C0&
            Caption         =   "»Ì«‰«  «·⁄„Ì·"
            Height          =   1965
            Left            =   60
            TabIndex        =   26
            Top             =   690
            Width           =   5715
            Begin VB.TextBox TxtAddress 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   120
               TabIndex        =   31
               Top             =   1560
               Width           =   4125
            End
            Begin VB.TextBox TxtEmial 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   1260
               TabIndex        =   30
               Top             =   1230
               Width           =   2985
            End
            Begin VB.TextBox TxtMobile 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   1260
               TabIndex        =   29
               Top             =   900
               Width           =   2985
            End
            Begin VB.TextBox TxtPhone 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   1260
               TabIndex        =   28
               Top             =   570
               Width           =   2985
            End
            Begin VB.TextBox TxtCustomerName 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   120
               TabIndex        =   27
               Top             =   240
               Width           =   4125
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C8C0&
               Caption         =   "«·⁄‰Ê«‰"
               Height          =   315
               Index           =   17
               Left            =   4290
               TabIndex        =   36
               Top             =   1560
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C8C0&
               Caption         =   "«·»—Ìœ «·√·ﬂ —Ê‰Ï"
               Height          =   315
               Index           =   16
               Left            =   4290
               TabIndex        =   35
               Top             =   1230
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C8C0&
               Caption         =   "—ﬁ„ «·ÃÊ«·"
               Height          =   315
               Index           =   10
               Left            =   4290
               TabIndex        =   34
               Top             =   900
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C8C0&
               Caption         =   "—ﬁ„ «·Â« ›"
               Height          =   315
               Index           =   9
               Left            =   4290
               TabIndex        =   33
               Top             =   570
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C8C0&
               Caption         =   "«”„ «·⁄„Ì·"
               Height          =   315
               Index           =   8
               Left            =   4290
               TabIndex        =   32
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Image Image1 
            Height          =   1080
            Left            =   60
            Picture         =   "FrmRegisteration.frx":5D30
            Top             =   2760
            Width           =   1080
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   7
            Left            =   4230
            Picture         =   "FrmRegisteration.frx":9D9A
            Top             =   300
            Width           =   240
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "„‰ ›÷·ﬂ ≈÷€ÿ Next ··„ «»⁄…"
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   15
            Left            =   3480
            TabIndex        =   38
            Top             =   3570
            Width           =   2265
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C8C0&
            Caption         =   "„‰ ›÷·ﬂ «œŒ· »Ì«‰« ﬂ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   525
            Index           =   12
            Left            =   30
            TabIndex        =   37
            Top             =   240
            Width           =   5775
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   3915
         Left            =   1.00960e5
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   0
         Width           =   5835
         _cx             =   10292
         _cy             =   6906
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
         BackColor       =   12634304
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
         Begin VB.Frame FraActiveFile 
            BackColor       =   &H00C0C8C0&
            Caption         =   "„”«— „·› «· ‰‘Ìÿ ⁄·Ï «·ÃÂ«“"
            Height          =   765
            Left            =   60
            TabIndex        =   19
            Top             =   1140
            Width           =   5715
            Begin VB.CommandButton CmdBrows 
               Caption         =   "..."
               Height          =   345
               Left            =   90
               TabIndex        =   21
               Top             =   330
               Width           =   465
            End
            Begin VB.TextBox TxtFilePath 
               Height          =   345
               Left            =   570
               TabIndex        =   20
               Top             =   330
               Width           =   5055
            End
            Begin MSComDlg.CommonDialog Cdg 
               Left            =   240
               Top             =   180
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
         End
         Begin VB.Image Image2 
            Height          =   1080
            Left            =   4680
            Picture         =   "FrmRegisteration.frx":A124
            Top             =   1920
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C8C0&
            Caption         =   "≈Œ Ì«— „·› «· ‰‘Ìÿ «Ê «· ”ÃÌ· «·„ÊÃÊœ „”»ﬁ«"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   525
            Index           =   18
            Left            =   60
            TabIndex        =   24
            Top             =   240
            Width           =   5715
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   8
            Left            =   5430
            Picture         =   "FrmRegisteration.frx":E18E
            Top             =   300
            Width           =   240
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   1245
            Index           =   5
            Left            =   90
            TabIndex        =   23
            Top             =   1950
            Width           =   3375
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   1
            Left            =   3510
            Picture         =   "FrmRegisteration.frx":E518
            Top             =   1950
            Width           =   240
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "„‰ ›÷·ﬂ ≈÷€ÿ Next ··„ «»⁄…"
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   26
            Left            =   3480
            TabIndex        =   22
            Top             =   3570
            Width           =   2265
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3915
         Index           =   3
         Left            =   1.00960e5
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   5835
         _cx             =   10292
         _cy             =   6906
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
         BackColor       =   12634304
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
         Begin VB.Frame FraRegInfo 
            BackColor       =   &H00C0C8C0&
            Height          =   1575
            Left            =   60
            TabIndex        =   9
            Top             =   660
            Width           =   5745
            Begin VB.CommandButton Cmd 
               Caption         =   "‰”Œ"
               Height          =   285
               Left            =   4800
               TabIndex        =   60
               Top             =   570
               Width           =   825
            End
            Begin VB.TextBox TxtSerialNumber 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   30
               TabIndex        =   10
               Top             =   870
               Width           =   4755
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C8C0&
               Caption         =   " —ﬁ„ «·”Ì—Ì«· - «œŒ· «·—ﬁ„ «·–Ï «Œ– Â „‰ „Ê“⁄ «·»—‰«„Ã «Ê „‰ «·‘—ﬂ…"
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   20
               Left            =   30
               TabIndex        =   15
               Top             =   1260
               Width           =   4725
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C8C0&
               Caption         =   " «·—ﬁ„ «·Œ«’ - Â–« «·—ﬁ„ Œ«’ »ﬂ· ÃÂ«“ Ê·« Ì ﬂ—— „⁄ «Ï ÃÂ«“ «Œ—"
               ForeColor       =   &H00000040&
               Height          =   225
               Index           =   21
               Left            =   30
               TabIndex        =   14
               Top             =   600
               Width           =   4665
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C8C0&
               Caption         =   "«·—ﬁ„ «·Œ«’"
               Height          =   345
               Index           =   22
               Left            =   4890
               TabIndex        =   13
               Top             =   240
               Width           =   795
            End
            Begin VB.Label LblComputerID 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblComputerID"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   30
               TabIndex        =   12
               Top             =   240
               Width           =   4755
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C8C0&
               Caption         =   "—ﬁ„ «·”Ì—Ì«·"
               Height          =   315
               Index           =   23
               Left            =   4800
               TabIndex        =   11
               Top             =   900
               Width           =   885
            End
         End
         Begin VB.Frame FraActivate 
            BackColor       =   &H00000040&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   210
            TabIndex        =   6
            Top             =   2370
            Width           =   5265
            Begin VB.TextBox TxtActivateNumber 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   405
               Left            =   180
               TabIndex        =   7
               Top             =   120
               Width           =   3825
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00777777&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «· ”ÃÌ·"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000CCFF&
               Height          =   375
               Index           =   24
               Left            =   4020
               TabIndex        =   8
               Top             =   150
               Width           =   1005
            End
         End
         Begin ImpulseButton.ISButton CmdActivate 
            Height          =   495
            Left            =   210
            TabIndex        =   5
            Top             =   3210
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   873
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ”ÃÌ·"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmRegisteration.frx":E8A2
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "≈÷€ÿ ⁄·Ï “—  ”ÃÌ· ·»œ¡ ⁄„·Ì… «· ”ÃÌ·"
            ForeColor       =   &H00000040&
            Height          =   285
            Index           =   0
            Left            =   1470
            TabIndex        =   54
            Top             =   3300
            Width           =   2835
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C8C0&
            Caption         =   "„‰ ›÷·ﬂ «œŒ· —ﬁ„ «·”Ì—Ì«· «·Œ«’ »«·‰”Œ…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   525
            Index           =   19
            Left            =   60
            TabIndex        =   17
            Top             =   180
            Width           =   5715
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "„‰ ›÷·ﬂ ≈÷€ÿ Next ··„ «»⁄…"
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   25
            Left            =   3480
            TabIndex        =   16
            Top             =   3570
            Visible         =   0   'False
            Width           =   2265
         End
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C8C0&
         Caption         =   "Â–« «·„⁄«·Ã ”Ê› Ì’Õ»ﬂ„ ŒÿÊ… »ŒÿÊ… · ”ÃÌ· «·»—‰«„Ã Ê–·ﬂ Õ„«Ì… ·ÕﬁÊﬁﬂ„ ›Ï «·Õ’Ê· ⁄·Ï «·œ⁄„ «·›‰Ï «·„»«‘— „‰ «·‘—ﬂ… «Ê „‰ «·„Ê“⁄ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   945
         Index           =   28
         Left            =   1.00090e5
         TabIndex        =   53
         Top             =   2820
         Width           =   645
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   30
      TabIndex        =   2
      Top             =   3930
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton CmdConnect 
      Caption         =   "Connect"
      Height          =   435
      Left            =   30
      TabIndex        =   1
      Top             =   3180
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "SEND"
      Height          =   435
      Left            =   30
      TabIndex        =   0
      Top             =   3630
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label LblExpireCount 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      Height          =   255
      Left            =   750
      TabIndex        =   59
      Top             =   5070
      Width           =   495
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·„œ… «·„ »ﬁÌ… ·ﬂ"
      Height          =   255
      Index           =   30
      Left            =   1290
      TabIndex        =   58
      Top             =   5070
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·‰”Œ… «· Ã—Ì»Ì… ·„œ… 50 „—…  ‘€Ì·"
      Height          =   255
      Index           =   29
      Left            =   4140
      TabIndex        =   57
      Top             =   5070
      Width           =   2625
   End
End
Attribute VB_Name = "FrmRegisteration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BolStop As Boolean

Public WithEvents Conn As ADODB.Connection
Attribute Conn.VB_VarHelpID = -1
Dim m_UserCancelReg As Boolean

Private Sub Cmd_Click()
    Clipboard.SetText Me.LblComputerID.Caption
End Sub

Private Sub CmdActivate_Click()
    Dim Msg As String

    If Me.TxtActivateNumber.text = CreateKey Then 'registered
        save_confoguration "D11002D19y84", 0, TxtActivateNumber.text + "10111982"
        Msg = " „ ﬁ»Ê· —ﬁ„ «· ”ÃÌ·...  „  ‰‘Ìÿ «·»—‰«„Ã"
        Msg = Msg & Chr(13) & "‘ﬂ—« ·≈” Œœ«„ﬂ„ »—‰«„Ã œÌ‰«„Ìﬂ »«Ì  «·„ ﬂ«„·"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        SystemOptions.SysRegisterState = Registered
        SystemOptions.SysVersion = RegisterVersion
        Unload Me
    Else
        Msg = " „ —›÷ —ﬁ„ «· ”ÃÌ·...      "
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    
    End If

    Exit Sub

    If Trim(Me.TxtActivateNumber.text) = "ayman-18121996-251979" Then
    
    ElseIf OptRegType(0).value = True Then

        If RegViaActivateFile = True Then
            Msg = " „ ﬁ»Ê· „·› «· ‰‘Ìÿ"
            Msg = Msg & Chr(13) & " „  ‰‘Ìÿ «·»—‰«„Ã »‰Ã«Õ"
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            SystemOptions.SysRegisterState = Registered
            SystemOptions.SysVersion = RegisterVersion
            Unload Me
        End If

    ElseIf OptRegType(1).value = True Then

        If RegViaNet = True Then
            Msg = " „ ‰ ‘Ìÿ Ê ”ÃÌ· «·»—‰«„Ã »‰Ã«Õ"
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            SystemOptions.SysRegisterState = Registered
            SystemOptions.SysVersion = RegisterVersion
            Unload Me
        End If

    ElseIf OptRegType(2).value = True Then

        If RegViaPhone = True Then
            Unload Me
            SystemOptions.SysRegisterState = Registered
            SystemOptions.SysVersion = RegisterVersion
        End If
    End If

End Sub

Private Sub CmdBrows_Click()
    Dim sp As ImpulseFilesAndFolders
    Dim DD As String

    On Error GoTo ErrTrap
    Cdg.CancelError = False
    Set sp = New ImpulseFilesAndFolders
    DD = sp.SpecialFolders.Desktop
    Me.Cdg.InitDir = DD
    Cdg.ShowOpen

    If Cdg.FileName <> "" Then
        TxtFilePath.text = Cdg.FileName

        If Trim(TxtFilePath.text) <> "" Then
            MeLoadReg
        End If
    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub Command1_Click()

    Dim rs As New ADODB.Recordset
    Dim SS
    'With Conn
    '    If SystemOptions.SysConnectionType = ConnectRemote Then
    '        .ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & _
    '        StrServerIP & ";PORT=3306;DATABASE=" & _
    '        StrDataBaseName & ";USER=" & StrUserConnect & "" _
    '        & ";PASSWORD=" & StrPasswordConnect & ";OPTION=16427;"
    '    ElseIf SystemOptions.SysConnectionType = ConnecLocal Then
    '        .ConnectionString = "Provider=MSDASQL.1;Password=nour1234nour;" & _
    '        "Persist Security Info=True;User ID=root;Data Source=register"
    '    End If
    '    .Open
    'End With

    Set rs = New ADODB.Recordset
    open_my_connection
    rs.Open "registers", Conn, adOpenKeyset, adLockOptimistic, adCmdTable
    Open App.path & "\SmallAccountSerialNumbers.txt" For Input As #1

    Do While Not EOF(1)
        Line Input #1, SS
        rs.AddNew
        rs("SerialNumber").value = SS
        rs.update

        DoEvents
        Me.Caption = rs.RecordCount
    Loop

    Close #1
    MsgBox "Done"
End Sub

Private Sub Form_Load()
    Dim Encrp As New ImpulseEncryption
    Dim Msg As String
    SystemOptions.SysServerIP = "BYTE"

    Msg = " ”ÃÌ· »—‰«„Ã " & App.title
    'Lbl(0).Caption = Msg
    Msg = "»—Ã«¡ ≈ »«⁄ ŒÿÊ«   ”ÃÌ· «·»—‰«„Ã ·ﬂÏ Ì „  ”ÃÌ· «·»—‰«„Ã »‰Ã«Õ"
    'Lbl(1).Caption = Msg
    Msg = ""
    Msg = Msg & ""
    lbl(11).Caption = Msg
    lbl(13).Caption = " √ﬂœ „‰ ÊÃÊœ ≈ ’«· »«·√‰ —‰ "
    Msg = "„·ÕÊŸ…:- ≈–« ﬂ‰  ﬁ„  »«· ”ÃÌ· ›Ï «Ï „—… ”«»ﬁ… "
    Msg = Msg & "Ê·œÌﬂ „·› «· ‰‘Ìÿ ⁄·Ï ÃÂ«“ﬂ ›«‰Â Ì„ﬂ‰ﬂ «· ”ÃÌ· "
    Msg = Msg & "„—… «Œ—Ï »‰›” «·„·›"
    lbl(14).Caption = Msg

    Msg = "„·ÕÊŸ…:-"
    Msg = Msg & Chr(13) & "Â–« «·»—‰«„Ã „Õ„Ï »ÕﬁÊﬁ «·„·ﬂÌ… «·›ﬂ—Ì…"
    Msg = Msg & Chr(13) & "Ê«Ï ‰”Œ €Ì— ﬁ«‰Ê‰Ï ··»—‰«„Ã Ì⁄—÷ ··„”√·…"
    Msg = Msg & Chr(13) & "«·ﬁ«‰Ê‰Ì…"
    lbl(2).Caption = Msg
    Msg = "„·ÕÊŸ…:-"
    Msg = Msg & Chr(13) & "ﬁ„ » ÕœÌœ „·› «· ‰‘Ìÿ „‰ ⁄·Ï «·ÃÂ«“ "
    Msg = Msg & "ÊÃÊœ „·› «· ‰‘Ìÿ ÌÊ›— ⁄·Ìﬂ «·Êﬁ  ›Ï «·„—… "
    Msg = Msg & "«· «·… ⁄‰œ ≈⁄œ«œ «·»—‰«„Ã Ê ”ÃÌ·Â ·«‰ „·› «· ‰‘Ìÿ "
    Msg = Msg & "·«‰ «·»—‰«„Ã ÌﬁÊ„ »Õ›ÿ ≈⁄œ«œ«  «· ”ÃÌ· ›Ï Â–« «·„·›"
    lbl(5).Caption = Msg

    Set Conn = New ADODB.Connection
    CmdSend.Enabled = False
    'CmdStop.Enabled = False
    CenterForm Me
    'Msg = GetBoardData(False)
    Msg = GetHardDiskData(False)
    Msg = Msg & "**" '& GetProcessorData(False)
    Me.LblComputerID.Caption = Msg
    OptRegType(1).value = True
    OptRegType_Click 1
    'If 2 > 1 Then
    '    WzrdMain.CancelEnabled = False
    'End If
    FrmRegisteration.LblExpireCount = 50 - run_count ' 50 - SystemOptions.SysRunNumber
    FrmRegisteration.EleExpire.FloodPercent = SystemOptions.SysRunNumber * 2
    WriteInLogFile "Compelet Load FrmRegisteration"
End Sub

Public Function open_my_connection() As Boolean
    Dim Msg As String
    open_my_connection = False

    If IsInternetConnected = False Then
        Msg = "·«ÌÊÃœ ≈ ’«· »«·√‰ —‰  "
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        open_my_connection = False
        Exit Function
    End If

    On Error GoTo Open_my_connection_ErrTrap

    If Not Conn Is Nothing Then
    Else
        Set Conn = New ADODB.Connection
    End If

    If Conn.State = adStateClosed Then

        With Conn
            '        'StrServerIP = "65.98.113.18"
            '        StrServerIP = SystemOptions.SysServerName
            '        StrDataBaseName = SystemOptions.SysDataBaseName
            '        StrDataBaseName = "noursyst_moneyshare"
            '        StrPasswordConnect = SystemOptions.SysPasswordConnect
            '        StrUserConnect = SystemOptions.SysUserConnect
            .CursorLocation = adUseClient
            .ConnectionTimeout = 20

            If SystemOptions.SysConnectionType = ConnectRemote Then
                .ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & "65.98.113.18" & ";PORT=3306;DATABASE=" & "noursyst_SmallAccountRegister" & ";USER=" & "noursyst_account" & "" & ";PASSWORD=" & "account" & ";OPTION=16427;"
                Conn.Open
            ElseIf SystemOptions.SysConnectionType = ConnecLocal Then
                .ConnectionString = "Provider=MSDASQL.1;Password=nour1234nour;" & "Persist Security Info=True;User ID=root;Data Source=register"
                Conn.Open
            End If

            'cOpenConnection.OpenConnection Conn, Conn.ConnectionString
        End With

    End If

    'Dim I As Integer
    'For I = 0 To Conn.Properties.Count - 1
    '    Debug.Print Conn.Properties(I).Name, Conn.Properties(I).Value
    'Next I

    Do While Conn.State = adStateClosed
        DoEvents
    Loop

    If Conn.State = adStateOpen Then
        open_my_connection = True
    Else
        open_my_connection = False
    End If

    Exit Function
Open_my_connection_ErrTrap:
    open_my_connection = False

    If Err.Number = -2147467259 Then
        Msg = "«·»—‰«„Ã €Ì— ﬁ«œ— ⁄·Ï «·√ ’«· »«·”Ì—›—"
        Msg = Msg & Chr(13) & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ «·√ ’«· »«·√‰ —‰ "
    Else
        Msg = "«·»—‰«„Ã €Ì— ﬁ«œ— ⁄·Ï «·√ ’«· »«·”Ì—›—"
    End If

    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
End Function

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    If UnloadMode = VBRUN.QueryUnloadConstants.vbFormControlMenu Or UnloadMode = VBRUN.QueryUnloadConstants.vbAppTaskManager Then
        UserCancelReg = True
    End If

End Sub

Private Sub OptRegType_Click(Index As Integer)
    Dim BolStatus As Boolean

    If OptRegType(0).value = True Then
        CmdBrows.Enabled = True
        TxtFilePath.Enabled = True
        lbl(14).Enabled = True
    
        lbl(2).Enabled = False
        lbl(3).Enabled = False
        'Lbl(4).Enabled = False
    
        lbl(6).Enabled = False
        'Lbl(7).Enabled = False
        'Lbl(8).Enabled = False
        'Lbl(9).Enabled = False
        'Lbl(10).Enabled = False
        lbl(11).Enabled = False
        LblComputerID.Enabled = False
        TxtActivateNumber.Enabled = False
        'Lbl(12).Enabled = False
        lbl(13).Enabled = False
    
        lbl(11).Enabled = False
   
        TxtActivateNumber.Enabled = False
        FraActivate.Enabled = False
        TxtSerialNumber.Enabled = False
        lbl(2).Enabled = False
    
    ElseIf OptRegType(1).value = True Then

        CmdBrows.Enabled = False
        TxtFilePath.Enabled = False
        lbl(14).Enabled = False
    
        lbl(2).Enabled = True
        lbl(3).Enabled = True
        'Lbl(4).Enabled = True
    
        lbl(6).Enabled = True
        'Lbl(7).Enabled = True
        'Lbl(8).Enabled = True
        'Lbl(9).Enabled = True
        'Lbl(10).Enabled = True
        lbl(11).Enabled = True
        LblComputerID.Enabled = True
        TxtActivateNumber.Enabled = True
        'Lbl(12).Enabled = True
        lbl(13).Enabled = True
    
        lbl(11).Enabled = False
    
        TxtActivateNumber.Enabled = False
        FraActivate.Enabled = False
        TxtSerialNumber.Enabled = True
        lbl(2).Enabled = True
    
    ElseIf OptRegType(2).value = True Then
        CmdBrows.Enabled = False
        TxtFilePath.Enabled = False
        lbl(14).Enabled = False
    
        lbl(2).Enabled = False
        lbl(3).Enabled = False
        'Lbl(4).Enabled = False
    
        lbl(6).Enabled = False
        'Lbl(7).Enabled = False
        'Lbl(8).Enabled = False
        'Lbl(9).Enabled = False
        'Lbl(10).Enabled = False
        lbl(11).Enabled = False
        LblComputerID.Enabled = False
        TxtActivateNumber.Enabled = False
        'Lbl(12).Enabled = False
        lbl(13).Enabled = False
    
        lbl(11).Enabled = True
        TxtActivateNumber.Enabled = True
        FraActivate.Enabled = True
        TxtSerialNumber.Enabled = True
        lbl(2).Enabled = True
    End If

    OptRegType(0).ForeColor = IIf(OptRegType(0).value = True, vbRed, vbBlack)
    OptRegType(1).ForeColor = IIf(OptRegType(1).value = True, vbRed, vbBlack)
    OptRegType(2).ForeColor = IIf(OptRegType(2).value = True, vbRed, vbBlack)

    lbl(13).BackStyle = IIf(OptRegType(1).value = True, 1, 0)
    lbl(11).BackStyle = IIf(OptRegType(2).value = True, 1, 0)
    lbl(14).BackStyle = IIf(OptRegType(0).value = True, 1, 0)
End Sub

Private Function GetProcessorData(BolWithCaption As Boolean) As String
    Dim objNameSpace As SWbemServices, ObjCPUSet As SWbemObjectSet
    Dim objCpu As SWbemObject

    Set objNameSpace = GetObject("winmgmts:")
    Set ObjCPUSet = objNameSpace.InstancesOf("Win32_Processor")

    For Each objCpu In ObjCPUSet

        If BolWithCaption = True Then
            GetProcessorData = "»Ì«‰«  «·»—Ê””Ê—" & Chr(13) & Chr(10) & objCpu.GetObjectText_
        Else
            'GetProcessorData = objCpu.SerialNumber
            GetProcessorData = objCpu.ProcessorId
        End If

        Exit For
    Next

    Set objCpu = Nothing
    Set ObjCPUSet = Nothing
    Set objNameSpace = Nothing
End Function

'Private Function GetHardDiskData(BolWithCaption As Boolean) As String
'Dim objNameSpace As SWbemServices, ObjCPUSet As SWbemObjectSet
'Dim objCpu As SWbemObject
'Set objNameSpace = GetObject("winmgmts:")
'Set ObjCPUSet = objNameSpace.InstancesOf("Win32_PhysicalMedia")
'For Each objCpu In ObjCPUSet
'    Debug.Print objCpu.Name
'    If BolWithCaption = True Then
'        GetHardDiskData = "»Ì‰«‰«  «·Â«—œ œÌ”ﬂ" & Chr(10) & Chr(13) & objCpu.GetObjectText_
'    Else
'        'GetHardDiskData = objCpu.GetObjectText_
'        GetHardDiskData = objCpu.SerialNumber
'    End If
'    Debug.Print objCpu.SerialNumber
'Next
'Set objCpu = Nothing
'Set ObjCPUSet = Nothing
'Set objNameSpace = Nothing
'
'End Function

Public Function GetBoardData(BolWithCaption As Boolean) As String
    Dim objNameSpace As SWbemServices, ObjBiosSet As SWbemObjectSet
    Dim objBios As SWbemObject
    Set objNameSpace = GetObject("winmgmts:")
    Set ObjBiosSet = objNameSpace.InstancesOf("Win32_MotherboardDevice")

    For Each objBios In ObjBiosSet
        Debug.Print objBios.name

        If BolWithCaption = True Then
            GetBoardData = "»Ì«‰«  «··ÊÕ… «·√„" & Chr(13) & Chr(10) & objBios.GetObjectText_
        Else
            GetBoardData = objBios.GetObjectText_
        End If

    Next

    Set objBios = Nothing
    Set ObjBiosSet = Nothing
    Set objNameSpace = Nothing
End Function

Public Function OpenConnection()
    Dim StrConn As String
    StrConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\EmailData.mdb;Persist Security Info=False"
    Conn.Open StrConn
End Function

Private Function RegViaNet() As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim VarTemp As Variant
    Dim StrComputerID As String
    Dim StrSQL As String
    Dim RsTimeStamp As ADODB.Recordset

    On Error GoTo ErrTrap

    If Trim(TxtSerialNumber.text) = "" Then
        Msg = "„‰ ›÷·ﬂ ÌÃ» ﬂ «»… «·—ﬁ„ «·„”·”· «·Œ«’ »«·‰”Œ…"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtSerialNumber.SetFocus
        Exit Function
    End If

    If Trim(TxtCustomerName.text) = "" Then
        Msg = "„‰ ›÷·ﬂ ÌÃ» ﬂ «»… √”„ «·⁄„Ì·"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtCustomerName.SetFocus
        Exit Function
    End If

    If Trim(LblComputerID.Caption) = "" Then
        Msg = "·«Ì„ﬂ‰  ”ÃÌ· «·»—‰«„Ã «·√‰"
        Msg = Msg & Chr(13) & "«·»—‰«„Ã €Ì— ﬁ«œ— ⁄·Ï  ÕœÌœ «·—ﬁ„ «·Œ«’ ··ÃÂ«“"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        StrComputerID = Trim(LblComputerID.Caption)
    End If

    If Trim(StrComputerID) <> "" Then
        VarTemp = Split(StrComputerID, "**", , vbTextCompare)
    
    End If

    Me.MousePointer = vbArrowHourglass
    Set Conn = Nothing
    open_my_connection
    StrSQL = "SELECT CURRENT_TIMESTAMP()"
    Set RsTimeStamp = New ADODB.Recordset
    RsTimeStamp.Open StrSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

    Set rs = New ADODB.Recordset
    rs.Open "registers", Conn, adOpenStatic, adLockOptimistic, adCmdTable

    If rs.BOF Or rs.EOF Then
        Msg = "·«Ì„ﬂ‰ «· ”ÃÌ· «·√‰ "
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.MousePointer = vbDefault
        Exit Function
    End If

    rs.find "SerialNumber='" & Trim(Me.TxtSerialNumber.text) & "'", , adSearchForward, 1

    If rs.BOF Or rs.EOF Then
        Msg = "«·—ﬁ„ «·„”·”· «·–Ï «œŒ· Â €Ì— ’ÕÌÕ"
        Msg = Msg & Chr(13) & "»—Ã«¡ «· «ﬂœ „‰ «·—ﬁ„ «·„œŒ·"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.MousePointer = vbDefault
        Exit Function
    End If

    If rs("RegisterNum").value = 0 Or IsNull(rs("RegisterNum").value) = True Then
        'Not Register Before
        '«·—ﬁ„ «·„”·”· „ÊÃÊœ ›Ï ﬁ«⁄œ… «·»Ì«‰« 
        'Ê”Ê› Ì „ —»ÿ Â–« «·”Ì—Ì«· „⁄ Â–« «·⁄„Ì·
        rs("CustomerName").value = IIf(Trim(Me.TxtCustomerName.text) = "", Null, Trim(Me.TxtCustomerName.text))
        rs("CustomerPhone").value = IIf(Trim(Me.TxtPhone.text) = "", Null, Trim(Me.TxtPhone.text))
        rs("CustomerMobile").value = IIf(Trim(Me.TxtMobile.text) = "", Null, Trim(Me.TxtMobile.text))
        rs("CustomerAddress").value = IIf(Trim(Me.TxtAddress.text) = "", Null, Trim(Me.TxtAddress.text))
        rs("CustomerEmail").value = IIf(Trim(Me.TxtEmial.text) = "", Null, Trim(Me.TxtEmial.text))
        rs("ActivateMethod").value = 1
        rs("RegisterNum").value = rs("RegisterNum").value + 1
        rs("HardDiskID").value = VarTemp(0)
        rs("ProcessorID").value = VarTemp(1)

        If Not RsTimeStamp.BOF Or RsTimeStamp.EOF Then
            rs("RegisterDate").value = IIf(IsNull(RsTimeStamp("CURRENT_TIMESTAMP()").value) = True, Null, RsTimeStamp("CURRENT_TIMESTAMP()").value)
        Else
            rs("RegisterDate").value = Null
        End If

        Me.TxtActivateNumber.text = CreateKey
        rs("ActivateCode").value = Me.TxtActivateNumber.text
        rs.update

        RecRead.customername = Trim(Me.TxtCustomerName.text)
        RecRead.phone = Trim(Me.TxtPhone.text)
        RecRead.mobile = Trim(Me.TxtMobile.text)
        RecRead.Emial = Trim(Me.TxtEmial.text)
        RecRead.Address = Trim(Me.TxtAddress.text)
        RecRead.ComputerID = Trim(Me.LblComputerID.Caption)
        RecRead.SerialNumber = Trim(Me.TxtSerialNumber.text)
        RecRead.ActivateNumber = Me.TxtActivateNumber.text
        RecRead.HardDisk_ID = VarTemp(0)
        RecRead.Processor_ID = VarTemp(1)
        RecRead.MaxNumToRun = 50
        RecRead.CurRumNumber = SystemOptions.SysRunNumber
        RecRead.VersionType = 1
        RecRead.FristRunDate = RecRead.FristRunDate
        RecRead.LastRunDate = Date
    
        CreateNewRegFile
        SaveInRegFile RecRead
    
        Msg = " „  ”ÃÌ· «·»—‰«„Ã »‰Ã«Õ"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    ElseIf rs("RegisterNum").value > 0 Then

        'Â–« «·—ﬁ„ «·„”·”· „”Ã· „”»ﬁ«
        'ÊÃ«—Ï «· ”ÃÌ· „—… «Œ—Ï
        If (rs("HardDiskID").value & "**" & rs("ProcessorID").value) = Trim(Me.LblComputerID.Caption) Then
            'this is the same computer
            Me.TxtActivateNumber.text = rs("ActivateCode").value
            rs("RegisterNum").value = rs("RegisterNum").value + 1
            rs.update
            Me.TxtActivateNumber.text = rs("ActivateCode").value
        ElseIf rs("HardDiskID").value = VarTemp(0) Or rs("ProcessorID").value = VarTemp(1) Then
            'this is the same computer
            'but some hard ware may b changed
            Me.TxtActivateNumber.text = rs("ActivateCode").value
            rs("RegisterNum").value = rs("RegisterNum").value + 1
            rs.update
        ElseIf rs("HardDiskID").value <> VarTemp(0) And rs("ProcessorID").value <> VarTemp(1) Then
            'ﬂ„»ÌÊ —  «‰Ï
            '«·—ﬁ„ «·„”·”· ··Â«—œ €Ì— „ Ê«›ﬁ Êﬂ–·ﬂ «·—ﬁ„ «·„”·”· ··»—Ê””Ê—
            '—»„« ÌﬂÊ‰ Â–« ÃÂ«“ «Œ— €Ì— «·„”Ã· ⁄·ÌÂ «·‰”Œ… „‰ «·„—… «·”«»ﬁ…
            'Ê„„ﬂ‰ «‰ ÌﬂÊ‰ «·⁄„Ì· ﬁœ ﬁ«„ » €Ì— ÃÂ«“Â «Ê «‰Â «⁄ÿÏ «·‰”Œ…
            '≈·Ï ‘Œ’ «Œ— ... ÊÂ‰« ÌÃ»  ‰»ÌÂ «·„” Œœ„
            Msg = "⁄›Ê« .. Â–Â «·‰”Œ… ”Ã·  „”»ﬁ« „‰ ⁄·Ï ÃÂ«“ «Œ—"
            Msg = Msg & Chr(13) & "›≈–« ﬂ‰   „·ﬂ ‰”Œ… «’·Ì… „‰ «·»—‰«„Ã Ê„”Ã·… „”»ﬁ«"
            Msg = Msg & Chr(13) & "Õ«Ê· «·√ ’«· »«·œ⁄„ «·›‰Ï "
            MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.MousePointer = vbDefault
            Exit Function
        End If

    ElseIf rs("RegisterNum").value >= rs("MaxRegisterNum") Then ''RegisterBefore
        Msg = "⁄›Ê« .. Â–Â «·‰”Œ… ”Ã·  „”»ﬁ« «ﬂÀ— „‰ „—…"
        Msg = Msg & Chr(13) & "⁄œœ „—«  «· ”ÃÌ· «·”«»ﬁ… :-" & rs("RegisterNum").value
        Msg = Msg & Chr(13) & "Õ«Ê· «·√ ’«· »«·œ⁄„ «·›‰Ï "
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.MousePointer = vbDefault
        Exit Function
    End If

    RegViaNet = True
    Me.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Me.MousePointer = vbDefault
End Function

Private Function RegViaActivateFile() As Boolean
    Dim Msg As String

    If CreateKey <> val(TxtActivateNumber.text) Then
        Msg = "⁄›Ê« .. „·› «· ‰‘Ìÿ €Ì— „‰«”»"
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RegViaActivateFile = False
    Else
        RegViaActivateFile = True
    End If

    Exit Function
ErrTrap:
    RegViaActivateFile = False
End Function

Private Sub TxtSerialNumber_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Function CreateKey() As String
    Dim i As Integer
    Dim StrOutKey As String
    Dim StrChar As String
    Dim LngOutKey As Long
    Dim VarTemp As Variant
    VarTemp = Split(LblComputerID.Caption, "**", , vbBinaryCompare)

    For i = 1 To Len(LblComputerID.Caption)
        StrChar = Mid(LblComputerID.Caption, i, 1)

        If StrChar <> "" Then
            LngOutKey = LngOutKey + (Asc(StrChar) * 12)
        End If

    Next i

    LngOutKey = (LngOutKey + LngOutKey)
    LngOutKey = LngOutKey * 3
    CreateKey = CStr(LngOutKey)
End Function

Public Function CreateRegFile() As Boolean
    Dim IntFreeFile As Integer
    Dim StrFilePath As String
    Dim sp As ImpulseFilesAndFolders
    Dim DD As String
    Dim StrNewPath As String
    Dim StrNewEncrPath As String
    Dim Msg As String

    Dim Encryptor As ImpulseEncryption

    On Error GoTo ErrTrap
    IntFreeFile = FreeFile
    StrFilePath = App.path & "\TempReg.txt"

    If Dir(StrFilePath, vbNormal) <> "" Then
        Kill StrFilePath
    End If

    RecSave.customername = TxtCustomerName.text
    RecSave.phone = TxtPhone.text
    RecSave.mobile = TxtMobile.text
    RecSave.Emial = TxtEmial.text
    RecSave.Address = TxtAddress.text
    RecSave.ComputerID = LblComputerID.Caption
    RecSave.SerialNumber = TxtSerialNumber.text
    RecSave.ActivateNumber = TxtActivateNumber.text

    Open StrFilePath For Random As #IntFreeFile Len = Len(RecSave)
    Put #IntFreeFile, 1, RecSave
    Close #IntFreeFile

    Set sp = New ImpulseFilesAndFolders
    DD = sp.SpecialFolders.Desktop
    StrNewPath = DD & "\ProgrameRegTemp.txt"
    FileCopy StrFilePath, StrNewPath

    If Dir(StrFilePath, vbNormal) <> "" Then
        Kill StrFilePath
    End If

    Set Encryptor = New ImpulseEncryption
    Encryptor.EncryptionKey = EnKey
    StrNewEncrPath = DD & "\ProgrameReg.txt"

    If Dir(StrNewEncrPath) <> "" Then
        Kill StrNewEncrPath
    End If

    Encryptor.EncryptFile StrNewPath, StrNewEncrPath

    If Dir(StrNewPath) <> "" Then
        Kill StrNewPath
    End If

    Msg = "„·ÕÊŸ… Â«„…:-"
    Msg = Msg & Chr(13) & " „ ≈‰‘«¡ „·› «·Õ„«Ì… «·Œ«’… »«·»—‰«„Ã ⁄·Ï «·„”«— "
    Msg = Msg & Chr(13) & ""
    Msg = Msg & Chr(13) & StrNewEncrPath
    Msg = Msg & Chr(13) & ""
    Msg = Msg & Chr(13) & "≈” Œœ„ Â–« «·„·› ·«Õﬁ« ›Ï  ‰‘Ìÿ «·»—‰«„Ã œÊ‰"
    Msg = Msg & Chr(13) & "«·Õ«Ã… ≈·Ï «·√ ’«· »«·√‰ —‰  «Ê «·Â« ›"
    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    CreateRegFile = True
    Exit Function
ErrTrap:
    CreateRegFile = False
End Function

Public Sub EncrypitFile()

End Sub

Private Function RegViaPhone() As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim VarTemp As Variant
    Dim StrComputerID As String
    Dim StrSQL As String
    Dim RsTimeStamp As ADODB.Recordset

    On Error GoTo ErrTrap

    If Trim(TxtSerialNumber.text) = "" Then
        Msg = "„‰ ›÷·ﬂ ÌÃ» ﬂ «»… «·—ﬁ„ «·„”·”· «·Œ«’ »«·‰”Œ…"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtSerialNumber.SetFocus
        Exit Function
    End If

    If Trim(TxtCustomerName.text) = "" Then
        Msg = "„‰ ›÷·ﬂ ÌÃ» ﬂ «»… √”„ «·⁄„Ì·"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtCustomerName.SetFocus
        Exit Function
    End If

    If Trim(LblComputerID.Caption) = "" Then
        Msg = "·«Ì„ﬂ‰  ”ÃÌ· «·»—‰«„Ã «·√‰"
        Msg = Msg & Chr(13) & "«·»—‰«„Ã €Ì— ﬁ«œ— ⁄·Ï  ÕœÌœ «·—ﬁ„ «·Œ«’ ··ÃÂ«“"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        StrComputerID = Trim(LblComputerID.Caption)
    End If

    If Trim(StrComputerID) <> "" Then
        VarTemp = Split(StrComputerID, "**", , vbTextCompare)
    End If

    If Trim(TxtActivateNumber.text) = "" Then
        Msg = "ÌÃ» ﬂ «»… —ﬁ„ «· ”ÃÌ· «Ê «· ‰‘Ìÿ..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtActivateNumber.SetFocus
        Exit Function
    End If

    If CreateKey = Trim(TxtActivateNumber.text) Then

        RecRead.customername = Trim(Me.TxtCustomerName.text)
        RecRead.phone = Trim(Me.TxtPhone.text)
        RecRead.mobile = Trim(Me.TxtMobile.text)
        RecRead.Emial = Trim(Me.TxtEmial.text)
        RecRead.Address = Trim(Me.TxtAddress.text)
        RecRead.ComputerID = Trim(Me.LblComputerID.Caption)
        RecRead.SerialNumber = Trim(Me.TxtSerialNumber.text)
        RecRead.ActivateNumber = Me.TxtActivateNumber.text
        RecRead.HardDisk_ID = VarTemp(0)
        RecRead.Processor_ID = VarTemp(1)
        RecRead.MaxNumToRun = 50
        RecRead.CurRumNumber = SystemOptions.SysRunNumber
        RecRead.VersionType = 1
        RecRead.FristRunDate = RecRead.FristRunDate
        RecRead.LastRunDate = Date
    
        CreateNewRegFile 'SALIM TEST
        SaveInRegFile RecRead
    
        Msg = " „ ﬁ»Ê· —ﬁ„ «· ”ÃÌ·...  „  ‰‘Ìÿ «·»—‰«„Ã"
        Msg = Msg & Chr(13) & "‘ﬂ—« ·≈” Œœ«„ﬂ„ »—‰«„Ã œÌ‰«„Ìﬂ »«Ì  «·„ ﬂ«„·"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RegViaPhone = True
        Exit Function
    ElseIf CreateAontherKey = Trim(TxtActivateNumber.text) Then

        RecRead.customername = Trim(Me.TxtCustomerName.text)
        RecRead.phone = Trim(Me.TxtPhone.text)
        RecRead.mobile = Trim(Me.TxtMobile.text)
        RecRead.Emial = Trim(Me.TxtEmial.text)
        RecRead.Address = Trim(Me.TxtAddress.text)
        RecRead.ComputerID = Trim(Me.LblComputerID.Caption)
        RecRead.SerialNumber = Trim(Me.TxtSerialNumber.text)
        RecRead.ActivateNumber = Me.TxtActivateNumber.text
        RecRead.HardDisk_ID = VarTemp(0)
        RecRead.Processor_ID = VarTemp(1)
        RecRead.MaxNumToRun = 50
        RecRead.CurRumNumber = SystemOptions.SysRunNumber
        RecRead.VersionType = 1
        RecRead.FristRunDate = RecRead.FristRunDate
        RecRead.LastRunDate = Date
    
        CreateNewRegFile
        SaveInRegFile RecRead
    
        Msg = " „ ﬁ»Ê· —ﬁ„ «· ”ÃÌ·...  „  ‰‘Ìÿ «·»—‰«„Ã"
        Msg = Msg & Chr(13) & "‘ﬂ—« ·≈” Œœ«„ﬂ„ »—‰«„Ã œÌ‰«„Ìﬂ »«Ì  «·„ ﬂ«„·"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RegViaPhone = True
        Exit Function
    Else
        Msg = "—ﬁ„ «· ”ÃÌ· €Ì— „ﬁ»Ê· "
        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RegViaPhone = False
        Exit Function
    End If

    Exit Function
ErrTrap:
    RegViaPhone = False
End Function

Private Function CreateAontherKey() As String
    Dim VarTemp As Variant
    Dim StrNewKey As String
    Dim Part1 As String
    Dim Part2 As String
    Dim Part3 As String
    Dim Part4 As String

    On Error GoTo ErrTrap

    If Trim(Me.TxtSerialNumber.text) = "" Then
        Exit Function
    End If

    VarTemp = Split(TxtSerialNumber.text, "-", , vbTextCompare)

    Part1 = val(VarTemp(1)) * 2
    Part2 = val(VarTemp(2)) * 3
    Part3 = val(VarTemp(3)) * 4
    Part4 = val(VarTemp(4)) * 5

    Part1 = val(Part4) + val(Part1)
    Part2 = val(Part2) + val(Part3)
    Part3 = val(Part1) + val(Part3)
    Part4 = val(Part2) + val(Part4)

    StrNewKey = VarTemp(0) & "-" & Part1 & "-" & Part2 & "-" & Part3 & "-" & Part4
    CreateAontherKey = StrNewKey
    Exit Function
ErrTrap:
    CreateAontherKey = ""
End Function

Private Sub WzrdMain_CancelWizardRequest(Step As Integer)
    Dim Msg As String
    Msg = "«·⁄„Ì· «·⁄“Ì“ :-"
    Msg = Msg & Chr(13) & "ÌÃ» ⁄·Ìﬂ  ”ÃÌ· ‰”Œ ﬂ „‰ «·»—‰«„Ã"
    Msg = Msg & Chr(13) & "Õ Ï  ÷„‰ ≈” „—«— ⁄„· «·»—‰«„Ã"
    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    UserCancelReg = True
    Me.Hide
    End
End Sub

Private Sub WzrdMain_StepComplete(Step As Integer, _
                                  ImpendingStep As Integer)
    Dim Msg As String

    If Step = 1 And ImpendingStep = 2 Then

    ElseIf Step = 2 And ImpendingStep = 3 Then

        If OptRegType(0).value = False Then
            WzrdMain.MoveToStep 3
        End If

    ElseIf Step = 3 And ImpendingStep = 4 Then

        If OptRegType(0).value = True Then
            If Trim(Me.TxtFilePath.text) = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— „·› «· ‰‘Ìÿ"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                WzrdMain.MoveToStep 2
            End If

            If Dir(Trim(Me.TxtFilePath.text)) = "" Then
                Msg = "„·› «· ‰‘Ìÿ €Ì— „ÊÃÊœ"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                WzrdMain.MoveToStep 2
            End If
        End If

    ElseIf Step = 4 And ImpendingStep = 3 Then

        If OptRegType(0).value = False Then
            WzrdMain.MoveToStep 3
        End If

    ElseIf Step = 4 And ImpendingStep = 5 Then

        If Trim(TxtCustomerName.text) = "" Then
            Msg = "„‰ ›÷·ﬂ ÌÃ» ﬂ «»… √”„ «·⁄„Ì·"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            WzrdMain.MoveToStep 3
            TxtCustomerName.SetFocus
        
        End If
    End If

End Sub

Private Function MeLoadReg() As Boolean

    Dim Encryptor As ImpulseEncryption
    Dim StrDecrPath As String
    Dim IntFreeFile As Integer
    Dim StrAllFile  As String
    Dim VarTemp As Variant
    Dim StrFilePath As String
    Dim Msg As String
    On Error GoTo ErrTrap

    TxtFilePath.text = Trim(TxtFilePath.text)

    If Trim(TxtFilePath.text) = "" Then
        Exit Function
    End If

    If Dir(Trim(TxtFilePath.text)) = "" Then
        Exit Function
    End If

    If Dir(SystemOptions.SysRegFilePath) <> "" Then
        Kill SystemOptions.SysRegFilePath
    End If

    FileCopy TxtFilePath.text, SystemOptions.SysRegFilePath
    LoadRegFile
    TxtCustomerName.text = Trim(RecRead.customername)
    TxtPhone.text = Trim(RecRead.phone)
    TxtMobile.text = Trim(RecRead.mobile)
    TxtEmial.text = Trim(RecRead.Emial)
    TxtAddress.text = Trim(RecRead.Address)
    LblComputerID.Caption = Trim(RecRead.ComputerID)
    TxtSerialNumber.text = RecRead.SerialNumber
    TxtActivateNumber.text = RecRead.ActivateNumber
    MeLoadReg = True
    Exit Function
ErrTrap:
    MeLoadReg = False
End Function

Public Property Get UserCancelReg() As Boolean
    UserCancelReg = m_UserCancelReg
End Property

Public Property Let UserCancelReg(ByVal vNewValue As Boolean)
    m_UserCancelReg = vNewValue
End Property

Private Sub PutDataInRec()

End Sub
