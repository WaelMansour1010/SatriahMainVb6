VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{AA91FA8F-BC1E-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseWizard.ocx"
Begin VB.Form FrmClose 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«ﬁ›«· «·”‰Â «·„«·ÌÂ"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "FrmClose.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6915
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
            Picture         =   "FrmClose.frx":058A
            Top             =   2400
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "‘ﬂ—« ·≈” Œœ«„ﬂ„ »—‰«„Ã   œÌ‰«„Ìﬂ »«Ì  ··Õ”«»«  ... «·»—‰«„Ã «·√ﬁÊÏ Ê«·√”Â·"
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
            Caption         =   "√Â·« »ﬂ„ ›Ï „⁄«·Ã  «ﬁ›«· «·”‰Â «·„«·ÌÂ"
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
            Caption         =   "Â–« «·„⁄«·Ã ”Ê› Ì’Õ»ﬂ„ ŒÿÊ… »ŒÿÊ… ·«ﬁ›«· «·”‰Â «·„«·ÌÂ"
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
            Picture         =   "FrmClose.frx":45F4
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
            Picture         =   "FrmClose.frx":497E
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
         Begin VB.CheckBox Check9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "«· √ﬂœ „‰ «‰ Â«¡ «‰‘«¡ ﬂ«›Â   ”‰œ«  «·—« » ⁄‰ «·⁄«„ «·„‰’—›"
            Height          =   195
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   2640
            Width           =   4575
         End
         Begin VB.CheckBox Check8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "«· √ﬂœ „‰ ⁄œ„ ÊÃÊœ «’‰«› „ﬂ‘ÊﬁÂ »œÊ‰ —’Ìœ"
            Height          =   195
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   2400
            Width           =   3495
         End
         Begin VB.CheckBox Check7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "«· √ﬂœ „‰ « „«„ ﬂ«›Â ⁄„·Ì«  «· Ê“Ì⁄ ··Õ”«»« "
            Height          =   195
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   2160
            Width           =   3375
         End
         Begin VB.CheckBox Check6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "«· √ﬂœ „‰ ⁄œ„ ÊÃÊœ „»Ì⁄«  ·Ì” ·Â« «–‰ ’—›"
            Height          =   195
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   1920
            Width           =   3375
         End
         Begin VB.CheckBox Check5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "«· √ﬂœ „‰ ⁄œ„ ÊÃÊœ „‘ —Ì«  ·„  ”·„ »”‰œ «” ·«„"
            Height          =   195
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   1680
            Width           =   3735
         End
         Begin VB.CheckBox Check4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "«· √ﬂœ „‰ «‰ Â«¡ «· ”ÊÌ«  «·Ã—œÌÂ ·‘Â—  12 "
            Height          =   195
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   1440
            Width           =   3375
         End
         Begin VB.CheckBox Check3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "«· √ﬂœ „‰ «‰ Â«¡ «‰‘«¡ ﬂ«›Â «·ﬁÌÊœ «· ﬂ—«—«Ì…"
            Height          =   195
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   1200
            Width           =   3375
         End
         Begin VB.CheckBox Check2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "«· √ﬂœ „‰ «‰ Â«¡  «’œ«— ﬂ«›Â «ﬁ”«ÿ «·«Â·«ﬂ"
            Height          =   195
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   960
            Width           =   3255
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C8C0&
            Caption         =   "«ﬁ›«· Õ”«»«  «·”‰Â «·Õ«·ÌÂ"
            Height          =   195
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   720
            Width           =   2655
         End
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
            Left            =   6870
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
            Left            =   6510
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
            Left            =   7350
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
            Caption         =   "”ÌﬁÊ„ «·‰Ÿ«„ »«· «ﬂœ „‰ «· «·Ì"
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
            Left            =   6750
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
            Left            =   7230
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
            Left            =   6510
            TabIndex        =   43
            Top             =   2010
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   2
            Left            =   6810
            Picture         =   "FrmClose.frx":4D08
            Top             =   990
            Width           =   240
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   3
            Left            =   6690
            Picture         =   "FrmClose.frx":5292
            Top             =   1620
            Width           =   240
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   4
            Left            =   7200
            Picture         =   "FrmClose.frx":561C
            Top             =   2340
            Width           =   240
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   6
            Left            =   4710
            Picture         =   "FrmClose.frx":59A6
            Top             =   270
            Width           =   240
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3915
         Index           =   2
         Left            =   1.09840e5
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
            Picture         =   "FrmClose.frx":5D30
            Top             =   2760
            Width           =   1080
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   7
            Left            =   4230
            Picture         =   "FrmClose.frx":9D9A
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
         Begin VB.CommandButton CMDDO 
            Caption         =   "‰›–"
            Height          =   255
            Left            =   2400
            TabIndex        =   71
            Top             =   840
            Width           =   1215
         End
         Begin VB.Frame FraActiveFile 
            BackColor       =   &H00C0C8C0&
            Caption         =   "„”«— „·› «· ‰‘Ìÿ ⁄·Ï «·ÃÂ«“"
            Height          =   765
            Left            =   6420
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   " „ ⁄„· «·ﬁÌœ «·«›  «ÕÌ"
            Height          =   375
            Left            =   1560
            TabIndex        =   70
            Top             =   1320
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Image Image2 
            Height          =   1080
            Left            =   4680
            Picture         =   "FrmClose.frx":A124
            Top             =   1920
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C8C0&
            Caption         =   "Ì „ «·«‰ ⁄„· «·ﬁÌœ «·«›  «ÕÌ ··”‰… «·ÃœÌœ…"
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
            Picture         =   "FrmClose.frx":E18E
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
            Left            =   7050
            TabIndex        =   23
            Top             =   1950
            Width           =   3375
         End
         Begin VB.Image Img 
            Height          =   240
            Index           =   1
            Left            =   3510
            Picture         =   "FrmClose.frx":E518
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
         Left            =   1.07440e5
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
            ButtonImage     =   "FrmClose.frx":E8A2
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
      Top             =   6150
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
      Top             =   6030
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
      Top             =   6270
      Width           =   2625
   End
End
Attribute VB_Name = "FrmClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGard As New ADODB.Recordset
Dim LngNoteID As Long
Dim TxtNoteID As String
 
Dim TxtDEVID As String
Dim TxtSerial1 As String

Private Sub Command2_Click()

End Sub

Private Sub CmdDo_Click()
    'CreateOpeningBalnceJLVoucherHeader , CDate("01/01/2012")

    'If CreateStoreOpeningbalanceVoucher(1, 1, 1) = True Then

    'End If
 
    Dim StrSQL As String
    Dim rs As New ADODB.Recordset
    Dim I As Integer
    Dim j As Integer

    Dim RsStorses As New ADODB.Recordset
 
    StrSQL = "SELECT branch_id From TblBranchesData"
 
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    StrSQL = "SELECT StoreID From TblStore"
 
    RsStorses.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        MsgBox "·« ÌÊÃœ ›—Ê⁄ „”Ã·Â »«·‰Ÿ«„"
        Exit Sub
    End If

    If RsStorses.RecordCount = 0 Then
        MsgBox "·« ÌÊÃœ „Œ«“‰ „”Ã·Â »«·‰Ÿ«„"
        GoTo ll:
    End If

    '                             If CreateStoreOpeningbalanceVoucher(6, 1, 1) = True Then
    '                             Exit Sub
    '                             End If
    If rs.RecordCount > 0 Then

        For I = 1 To rs.RecordCount
            RsStorses.MoveFirst

            For j = 1 To RsStorses.RecordCount

                If CreateStoreOpeningbalanceVoucher(val(rs!branch_id), val(RsStorses!StoreId), 1) = True Then

                End If

                '   MsgBox "Branch" & I & "Store" & J
                RsStorses.MoveNext
            Next j
                            
            rs.MoveNext
        Next I

    End If

ll:
    MsgBox " „ «·«ﬁ›«·"
End Sub

Function CreateStoreOpeningbalanceVoucher(Optional BranchID As Integer, Optional StoreId As Integer, Optional IntervalID As Integer) As Boolean
    CreateStoreOpeningbalanceVoucher = False
    Dim StrSQL As String
    Dim rs As New ADODB.Recordset
    Dim I As Integer
    Dim j As Integer

    Dim FromDate As Variant
    Dim ToDate As Variant
    'getIntervalIDDates
    FromDate = "01/01/2011"
    ToDate = "31/12/2011"
    StrSQL = StrSQL & "  SELECT     SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.Transaction_Details.Item_ID AS ItemID, dbo.Transactions.StoreID, "
    StrSQL = StrSQL & "  dbo.TblStore.StoreName , dbo.Transaction_Details.order_no, dbo.Transaction_Details.BranchId, dbo.Transaction_Details.unitid"
    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL & "  WHERE     (dbo.Transactions.Transaction_Date >=   " & SQLDate(CDate(FromDate), True) & "  and  dbo.Transactions.Transaction_Date  <=  " & SQLDate(CDate(ToDate), True) & " )  "
 
    StrSQL = StrSQL & "  AND (dbo.Transactions.StoreID = " & StoreId & ") AND (dbo.Transaction_Details.BranchId = " & BranchID & ")"
    StrSQL = StrSQL & "  GROUP BY dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.Transaction_Details.order_no, dbo.Transaction_Details.BranchId,"
    StrSQL = StrSQL & "  dbo.Transaction_Details.unitid "
    StrSQL = StrSQL & "  Having (SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) <> 0)"
 
    RsGard.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    If RsGard.RecordCount > 0 Then

        CreateOpeningBalnceVoucher BranchID, StoreId, ToDate

    End If

    RsGard.Close
 
End Function

Function CreateOpeningBalnceJLVoucherHeader(Optional my_branch As Integer = 1, Optional StartNewIntervalDate As Date) As Long

    If TxtSerial1 = "" Then
        If Voucher_coding(val(my_branch), StartNewIntervalDate, 3, 101) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ﬁÌœ «›  «ÕÌ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Function
        Else
                   
            If Voucher_coding(val(my_branch), StartNewIntervalDate, 3, 101) = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Function
            Else
                TxtSerial1 = Voucher_coding(val(my_branch), StartNewIntervalDate, 3, 101)
            End If
        End If
    End If

    Dim RsNetes As ADODB.Recordset

    TxtNoteID = CStr(new_id("notes1", "NoteID", ""))
    TxtDEVID = CStr(new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", ""))

    '---------------------------Begine of Saving------------
 
    Set RsNetes = New ADODB.Recordset
    RsNetes.Open "NOTES1", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsNetes.AddNew
    RsNetes("branch_no").Value = my_branch
    RsNetes("NoteID").Value = TxtNoteID
    RsNetes("NoteType").Value = 101
    RsNetes("NoteSerial").Value = TxtSerial1
    RsNetes("NoteSerial1").Value = TxtSerial1
    
    RsNetes("numbering_type").Value = sand_numbering_type(0) ' „”·”· «·ﬁÌœ
    RsNetes("numbering_type1").Value = sand_numbering_type(3) ' „”·”· «·”‰œ
    
    RsNetes("sanad_year").Value = year(StartNewIntervalDate)
    RsNetes("sanad_month").Value = Month(StartNewIntervalDate)
    
    RsNetes("NoteDate").Value = StartNewIntervalDate
 
    RsNetes("Double_Entry_Vouchers_ID").Value = TxtDEVID
    
    RsNetes("Remark").Value = "«ﬁ›«· › —… „«·ÌÂ"
    RsNetes("UserID").Value = user_id
    
    RsNetes.update

    LngNoteID = TxtNoteID
End Function

Function CreateOpeningBalnceVoucher(Optional BranchID As Integer, Optional StoreId As Integer, Optional ToDate As Variant)

    Dim rs As New ADODB.Recordset
    Dim RSTransDetails As New ADODB.Recordset
    Dim XPTxtBillID As String
    Dim txtopening_balance_voucher_id As String
    Dim TxtTransSerial As String
    Dim NewIntervaldate As Date
    NewIntervaldate = DateAdd("D", 1, CDate(ToDate))

    rs.Open "[Transactions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 '   StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  
    XPTxtBillID = CStr(new_id("Transactions", "Transaction_ID", "", True))
    txtopening_balance_voucher_id = get_opening_balance_voucher_id
    TxtTransSerial = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=3"))

    Cn.BeginTrans
    BegineTrans = True
 
    rs.AddNew
    rs("Transaction_ID").Value = val(XPTxtBillID)
    rs("BranchId").Value = BranchID
    rs("opening_balance_voucher_id").Value = txtopening_balance_voucher_id
    rs("Transaction_Serial").Value = TxtTransSerial
    rs("Transaction_Date").Value = NewIntervaldate
    rs("Transaction_Type").Value = 3
    rs("UserID").Value = user_id
    rs("StoreID").Value = StoreId
    rs.update
 
    For RowNum = 1 To RsGard.RecordCount

        If val(RsGard!ItemID) <> 0 Then
            RSTransDetails.AddNew
            RSTransDetails("Transaction_ID").Value = XPTxtBillID
            RSTransDetails("Item_ID").Value = RsGard!ItemID
            RSTransDetails("Quantity").Value = RsGard!SumQty
 
            RSTransDetails("ItemCase").Value = 1
            RSTransDetails("Price").Value = ModItemCostPrice.GetCostItemPrice(RsGard!ItemID, 0, , , SystemOptions.SysMainStockCostMethod)
            
            RSTransDetails("ColorID").Value = 1
             
            RSTransDetails("BranchId").Value = BranchID
            ' IIf((FG.TextMatrix(RowNum, FG.ColIndex("BranchId")) = ""), 1, Val(FG.TextMatrix(RowNum, FG.ColIndex("BranchId"))))
               
            RSTransDetails("ItemSize").Value = ""
            RSTransDetails("UnitID").Value = RsGard!unitid
          
            RSTransDetails("ShowQty").Value = RsGard!SumQty

            Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double
            Dim Valu As Double
        
            LngCurItemID = RsGard!ItemID
            LngUnitID = IIf(IsNull(RsGard!unitid), 1, RsGard!unitid)
            DblQty = RsGard!SumQty

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
             
            RSTransDetails("QtyBySmalltUnit").Value = RsUnitData("UnitFactor").Value
            RSTransDetails("Quantity").Value = RSTransDetails("QtyBySmalltUnit").Value * RSTransDetails("showqty").Value
            
            RSTransDetails("price").Value = Round(RSTransDetails("price").Value * RSTransDetails("QtyBySmalltUnit").Value, 2)
            'RSTransDetails("price").value = Round(FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / RSTransDetails("Quantity").value, 2)
 
            RSTransDetails("order_no").Value = RsGard!order_no
 
            RSTransDetails("OpeningBurcahseQty").Value = RSTransDetails("Quantity").Value
            RSTransDetails("OpeningBurcahseValue").Value = RSTransDetails("Price").Value * RSTransDetails("Quantity").Value
            'RSTransDetails("OpeningSalesQty").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesQty")) = "", Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesQty"))))
            'RSTransDetails("OpeningSalesValue").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesValue")) = "", Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesValue"))))
            
            RSTransDetails.update
        End If

        RsGard.MoveNext
    Next RowNum
  
    Cn.CommitTrans
    BegineTrans = False
  
End Function

Private Sub Form_Load()
    CenterForm Me
End Sub
