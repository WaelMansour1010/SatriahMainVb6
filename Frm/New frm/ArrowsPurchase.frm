VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form ArrowsPurchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘—«¡ √”Â„"
   ClientHeight    =   9420
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   13110
   Icon            =   "ArrowsPurchase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   13110
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   660
      Left            =   15
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   8745
      Width           =   13320
      _cx             =   23495
      _cy             =   1164
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
      Begin ImpulseButton.ISButton btnNew 
         Height          =   330
         Left            =   8895
         TabIndex        =   58
         Top             =   195
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
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
         ButtonImage     =   "ArrowsPurchase.frx":000C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   330
         Left            =   7350
         TabIndex        =   59
         Top             =   195
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
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
         ButtonImage     =   "ArrowsPurchase.frx":03A6
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   330
         Left            =   8115
         TabIndex        =   60
         Top             =   195
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
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
         ButtonImage     =   "ArrowsPurchase.frx":0740
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   330
         Left            =   6585
         TabIndex        =   61
         Top             =   195
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
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
         ButtonImage     =   "ArrowsPurchase.frx":0ADA
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   330
         Left            =   5820
         TabIndex        =   62
         Top             =   195
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
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
         ButtonImage     =   "ArrowsPurchase.frx":0E74
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   6960
         TabIndex        =   63
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   1170
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
         ButtonImage     =   "ArrowsPurchase.frx":140E
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   6045
         TabIndex        =   64
         TabStop         =   0   'False
         ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   945
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
         ButtonImage     =   "ArrowsPurchase.frx":17A8
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnPrint 
         Height          =   285
         Left            =   5805
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1230
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
         ButtonImage     =   "ArrowsPurchase.frx":1B42
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   330
         Left            =   5025
         TabIndex        =   66
         Top             =   195
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
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
         ButtonImage     =   "ArrowsPurchase.frx":1EDC
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   225
         Width           =   540
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   210
         Index           =   2
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   225
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   210
         Index           =   0
         Left            =   3225
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   225
         Width           =   975
      End
   End
   Begin VB.TextBox TxtNoteID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   13080
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Text            =   "Text1"
      Top             =   7440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   0
      Width           =   13275
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   6840
         Top             =   240
      End
      Begin VB.Frame Frmo2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   50
            Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
            Top             =   15
            Width           =   2340
            _ExtentX        =   4128
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
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   45
            Width           =   855
         End
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2580
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Text            =   "modflag"
         Top             =   90
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   510
         Visible         =   0   'False
         Width           =   945
      End
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   3120
         Top             =   0
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
               Picture         =   "ArrowsPurchase.frx":2276
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ArrowsPurchase.frx":2610
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ArrowsPurchase.frx":29AA
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ArrowsPurchase.frx":2D44
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ArrowsPurchase.frx":30DE
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ArrowsPurchase.frx":3478
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ArrowsPurchase.frx":3812
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ArrowsPurchase.frx":3DAC
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   90
         TabIndex        =   52
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "ArrowsPurchase.frx":4146
         ColorButton     =   14871017
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   555
         TabIndex        =   53
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "ArrowsPurchase.frx":44E0
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1155
         TabIndex        =   54
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "ArrowsPurchase.frx":487A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1620
         TabIndex        =   55
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "ArrowsPurchase.frx":4C14
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‘—«¡ √”Â„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   9735
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   90
         Width           =   2670
      End
   End
   Begin VB.CommandButton Cmd 
      Caption         =   " Õ„Ì· «·«”⁄«— „‰ «·«‰ —‰ "
      Height          =   315
      Index           =   0
      Left            =   8400
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   9240
      Width           =   10215
      ExtentX         =   18018
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   9480
      Width           =   10815
      ExtentX         =   19076
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   8055
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   13290
      _cx             =   23442
      _cy             =   14208
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
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   12648447
      ForeColor       =   0
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   16711680
      Caption         =   "‘—«¡ «”Â„|»Ì«‰«  «·‘—«¡|„”«⁄œÂ"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   0   'False
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   1
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Picture(0)      =   "ArrowsPurchase.frx":4FAE
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   7590
         Index           =   0
         Left            =   45
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   420
         Width           =   13200
         _cx             =   23283
         _cy             =   13388
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
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Height          =   7575
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   0
            Visible         =   0   'False
            Width           =   13095
            Begin SHDocVwCtl.WebBrowser WebBrowser3 
               Height          =   6855
               Left            =   120
               TabIndex        =   79
               Top             =   480
               Width           =   12855
               ExtentX         =   22675
               ExtentY         =   12091
               ViewMode        =   0
               Offline         =   0
               Silent          =   0
               RegisterAsBrowser=   0
               RegisterAsDropTarget=   1
               AutoArrange     =   0   'False
               NoClientEdge    =   0   'False
               AlignLeft       =   0   'False
               NoWebView       =   0   'False
               HideFileNames   =   0   'False
               SingleClick     =   0   'False
               SingleSelection =   0   'False
               NoFolders       =   0   'False
               Transparent     =   0   'False
               ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
               Location        =   "http:///"
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«€·«ﬁ"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   11160
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame FraNote 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·Õ«›Ÿ…"
            Height          =   1725
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   3480
            Width           =   4755
            Begin VB.ComboBox CboPaymentType 
               Height          =   315
               Left            =   6240
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   96
               Top             =   240
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1270
               Locked          =   -1  'True
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   1320
               Width           =   1815
            End
            Begin VB.CommandButton Command1 
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   1320
               Width           =   1095
            End
            Begin VB.TextBox TxtChequeNumber 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   600
               Width           =   3015
            End
            Begin MSComCtl2.DTPicker DtpChequeDueDate 
               Height          =   315
               Left            =   30
               TabIndex        =   86
               Top             =   900
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   556
               _Version        =   393216
               Format          =   104595457
               CurrentDate     =   41640
            End
            Begin MSDataListLib.DataCombo DcboBankName 
               Height          =   315
               Left            =   30
               TabIndex        =   87
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboBox 
               Height          =   315
               Left            =   5880
               TabIndex        =   88
               Top             =   600
               Visible         =   0   'False
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   195
               Index           =   15
               Left            =   9540
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   240
               Visible         =   0   'False
               Width           =   1245
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   255
               Index           =   13
               Left            =   3480
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·Õ—ﬂ…"
               Height          =   285
               Index           =   19
               Left            =   3300
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   900
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·Õ—ﬂ…"
               Height          =   285
               Index           =   18
               Left            =   3240
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Õ«›Ÿ…"
               Height          =   285
               Index           =   17
               Left            =   3270
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   270
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·Œ“‰…"
               Height          =   285
               Index           =   16
               Left            =   5190
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   660
               Visible         =   0   'False
               Width           =   1215
            End
         End
         Begin VB.TextBox TxtBocketId 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5520
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Tag             =   "Õœœ «·„Õ›Ÿ…"
            Top             =   2400
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.TextBox TxtoprType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2520
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Text            =   "1"
            Top             =   0
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   360
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.TextBox TxtCompanyID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5520
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Tag             =   "Õœœ «·‘—ﬂÂ"
            Top             =   2760
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   4140
            Left            =   5040
            TabIndex        =   8
            Top             =   3360
            Width           =   8115
            _cx             =   14314
            _cy             =   7302
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
            BackColorBkg    =   16777215
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
            Rows            =   2
            Cols            =   19
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"ArrowsPurchase.frx":5348
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
            Height          =   3285
            Index           =   2
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   240
            Width           =   4800
            _cx             =   8467
            _cy             =   5794
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
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
            Begin VB.TextBox TxtSerial 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2520
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   480
               Width           =   705
            End
            Begin VB.OptionButton OpComm 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ”«» «·⁄„Ê·Â „‰ «ﬁ· „»·€ ⁄„Ê·Â"
               Height          =   255
               Index           =   2
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   2880
               Width           =   2655
            End
            Begin VB.OptionButton OpComm 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ”«» «·⁄„Ê·Â „‰ ‰”»… «·»‰ﬂ"
               Height          =   255
               Index           =   1
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   2640
               Width           =   2655
            End
            Begin VB.OptionButton OpComm 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ”«» «·⁄„Ê·Â „‰ ‰”»… «·‘»ﬂÂ"
               Height          =   255
               Index           =   0
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   2400
               Value           =   -1  'True
               Width           =   2655
            End
            Begin VB.TextBox txtCommvalue 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   1920
               Width           =   3015
            End
            Begin VB.TextBox txtCommpercentage 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   1560
               Width           =   3015
            End
            Begin VB.TextBox TxtPrice 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Tag             =   "Õœœ ”⁄— «·‘—«¡"
               Top             =   1200
               Width           =   3015
            End
            Begin VB.TextBox txtqty 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Tag             =   "Õœœ «·ﬂ„ÌÂ"
               Top             =   840
               Width           =   3015
            End
            Begin MSComCtl2.DTPicker DpOprdate 
               Height          =   270
               Left            =   240
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   480
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   476
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/M/d"
               Format          =   104595459
               CurrentDate     =   41640
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„”·”· «·⁄„·Ì…"
               Height          =   255
               Index           =   7
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·‘—«¡"
               Height          =   255
               Index           =   1
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» «·⁄„Ê·…"
               Height          =   255
               Index           =   6
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   2400
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„»·€ «·⁄„Ê·…"
               Height          =   255
               Index           =   5
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   2040
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "‰”»… «·⁄„Ê·…"
               Height          =   255
               Index           =   4
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "”⁄— «·‘—«¡"
               Height          =   255
               Index           =   3
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   18
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ﬂ„Ì…"
               Height          =   255
               Index           =   2
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «·‘—«¡"
               Height          =   255
               Index           =   0
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   480
               Width           =   855
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   375
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   840
               Width           =   7575
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   2445
            Index           =   3
            Left            =   120
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   5040
            Width           =   4800
            _cx             =   8467
            _cy             =   4313
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
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
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   1680
               Width           =   2775
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   2040
               Width           =   2775
            End
            Begin VB.TextBox TxtCurrentValue 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   480
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   1320
               Width           =   2745
            End
            Begin VB.TextBox txtName 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox TxtBocketCode 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   240
               Width           =   2775
            End
            Begin VB.TextBox txtSymbol 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox txtBalance 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   960
               Width           =   2775
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   2640
               Width           =   2775
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·ﬁÌ„… «·«›  «ÕÌ…"
               Height          =   255
               Index           =   15
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·ﬁÌ„… «·Õ«·Ì…"
               Height          =   255
               Index           =   14
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   2040
               Width           =   1095
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   375
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   840
               Width           =   7575
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·„Œ›Ÿ…"
               Height          =   255
               Index           =   12
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·‘—ﬂÂ"
               Height          =   255
               Index           =   11
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·ﬁÊÂ «·‘—«∆Ì…"
               Height          =   255
               Index           =   10
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·ﬁÌ„Â «·«”„ÌÂ"
               Height          =   255
               Index           =   9
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ï «·‘—«¡"
               Height          =   255
               Index           =   8
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   2640
               Width           =   1095
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid 
            Height          =   1725
            Left            =   5160
            TabIndex        =   42
            Top             =   360
            Width           =   7950
            _cx             =   14023
            _cy             =   3043
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
            BackColorAlternate=   -2147483643
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
            Rows            =   50
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"ArrowsPurchase.frx":5631
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
         Begin MSDataListLib.DataCombo DcboFinMarketId 
            Height          =   315
            Left            =   9120
            TabIndex        =   44
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   2520
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboCurrencyId 
            Height          =   315
            Left            =   7320
            TabIndex        =   82
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   2400
            Visible         =   0   'False
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbogroupId 
            Height          =   315
            Left            =   7320
            TabIndex        =   83
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   2760
            Visible         =   0   'False
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboCreditSide 
            Height          =   315
            Left            =   4920
            TabIndex        =   98
            Top             =   0
            Visible         =   0   'False
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õœœ «·»Ê—’Â"
            Height          =   255
            Index           =   4
            Left            =   11160
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«Œ — —ﬁ„ «·„Õ›Ÿ…"
            Height          =   255
            Left            =   10920
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«Œ — «·‘—ﬂÂ"
            Height          =   255
            Left            =   10440
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   3000
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   0
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   840
            Width           =   7575
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   7590
         Index           =   1
         Left            =   13935
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   420
         Width           =   13200
         _cx             =   23283
         _cy             =   13388
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
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
            Height          =   6060
            Left            =   0
            TabIndex        =   39
            Top             =   360
            Width           =   13155
            _cx             =   23204
            _cy             =   10689
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
            BackColorBkg    =   16777215
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
            Rows            =   2
            Cols            =   29
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"ArrowsPurchase.frx":5804
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
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "«Ã„«·Ì ﬁÌ„… «·«”Â„"
            Height          =   495
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   6600
            Width           =   1695
         End
         Begin VB.Label Lblnet 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   495
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   6600
            Width           =   1215
         End
         Begin VB.Label lbltotal 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   495
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   6600
            Width           =   1215
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "«Ã„«·Ì ⁄œœ «·«”Â„"
            Height          =   495
            Left            =   10200
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   6600
            Width           =   1695
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   1
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   840
            Width           =   7575
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   7590
         Index           =   4
         Left            =   14235
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   420
         Width           =   13200
         _cx             =   23283
         _cy             =   13388
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
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   840
            Width           =   7575
         End
      End
   End
End
Attribute VB_Name = "ArrowsPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim path As String
Dim NEW_interface As Boolean
Dim OpCommindex As Integer
 
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecID As String
Dim II As Long
Dim cSearch  As clsDCboSearch

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.name, True) = False Then
        Exit Sub
    End If

    If TxtVac_ID.text <> "" Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)

        If MSGType = vbYes Then
            RsSavRec.find "OprId=" & val(TxtVac_ID.text), , adSearchForward, 1
            Dim sql As String
            sql = "Delete   from notes where NoteID=" & val(TXTNoteID.text)
            Cn.Execute sql
        
            RsSavRec.delete
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            '------------------------------ Move Next ---------------------------.
            'FillGridWithData
            FillGridWithData2
            BtnNext_Click
        End If
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub BtnFirst_Click()
    'On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec.MoveFirst
    FiLLTXT

    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    'On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MoveLast
    FiLLTXT
    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnModify_Click()
    Dim Msg As String

    If DoPremis(Do_Edit, Me.name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID.text <> "" Then
        Ele(2).Enabled = True
        TxtModFlg = "E"
        'Frm2.Enabled = True
        Me.TxtQty.SetFocus
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & Chr(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
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

    If DoPremis(Do_New, Me.name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    'Frm2.Enabled = True
    '-----------------------------------
    clear_all Me
    Ele(2).Enabled = True
    FillGridWithData2
    FillGridWithData
    OpComm(0).value = True
  
    With Me.VSFlexGrid1
        .Rows = 2
        .Clear flexClearScrollable
    End With

    '-----------------------------------
    TxtModFlg.text = "N"

    My_SQL = "ArrowsTransactions"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.text = rs.RecordCount + 1
    Else
        TxtSerial.text = 1
    End If

    rs.Close
    'CmbType.ListIndex = 0
    TxtQty.SetFocus
CboPaymentType.ListIndex = 1

ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
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

    FiLLTXT
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MovePrevious

    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If

    FiLLTXT
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnSave_Click()
    'On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------

    For Each CtrlTxt In Me.Controls

        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next

    '------------------------------ check if Empcode exist ----------------------

    StrVacName = IsRecExist("ArrowsTransactions", "GovernmentName", Trim(TxtQty.text), "GovernmentName", "Vac_ID<>'" & Trim(TxtVac_ID.text) & "'")

    If Me.TxtModFlg.text <> "R" Then
        If Me.CboPaymentType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
            Else
                Msg = "Select Payment method ...!!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CboPaymentType.SetFocus
            Exit Sub
        End If

        If Me.CboPaymentType.ListIndex = 0 Then
            If Trim(Me.DcboBox.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
                Else
                    Msg = "Select Box..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                '            DcboBox.SetFocus
                '            SendKeys "{F4}"
                Exit Sub
            End If

        ElseIf Me.CboPaymentType.ListIndex = 1 Then

            If Me.DcboBankName.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                Else
                    Msg = "Select Bank...!!"
        
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                '            DcboBankName.SetFocus
                '            SendKeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.text) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                Else
                    Msg = "Enter Cheque No:...!!"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If

            '                        If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
            '                                            If SystemOptions.UserInterface = ArabicInterface Then
            '                                                Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
            '                                            Else
            '                                            Msg = "Cheque Due Date Not Valid...!!"
            '
            '                                            End If
            '                            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '                '            DtpChequeDueDate.SetFocus
            '                '            SendKeys "{F4}"
            '                            Exit Sub
            '                        End If
        End If

        If Me.TxtModFlg.text = "N" Then
            If Me.CboPaymentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    '          If CheckBoxAccount(Me.DcboBox.BoundText, Val(Me.XPTxtVal.Text), _
                    '              XPDtbTrans.value) = False Then
                    '              Exit Sub
                    '          End If
                End If
            End If

        ElseIf Me.TxtModFlg.text = "E" Then

            If Me.CboPaymentType.ListIndex = 0 Then
                '      If Val(Me.DcboBox.BoundText) <> 0 Then
                '          If CheckBoxAccount(Me.DcboBox.BoundText, Val(Me.XPTxtVal.Text), _
                '              XPDtbTrans.value, , , Val(Me.XPTxtID.Text)) = False Then
                '              Exit Sub
                '          End If
            End If
        End If
    End If

    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.text

            '------------------------------ new record ----------------------------
        Case "N"
    
            If DcboCurrencyId.text = "" Then
                MsgBox "«Œ — «·⁄„·Â"
                Exit Sub
            End If
    
            If DcboGroupID.text = "" Then
                MsgBox "«Œ — «·„Ã„Ê⁄Â"
                Exit Sub
            End If
    
            '------------------------- save record -----------------------------
            If TxtNoteSerial.text = "" Then
                If Notes_coding(branch_id, DpOprdate.value) = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                    Else
                        MsgBox " Cant't Create Journal Entry to this Process no You exceed the maximum number ": Exit Sub
                    End If

                ElseIf Notes_coding(val(branch_id), DpOprdate.value) = "" Then

                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                    Else
                        MsgBox "You must Define JE Coding ": Exit Sub
                    End If

                Else
                    TxtNoteSerial.text = Notes_coding(val(branch_id), DpOprdate.value)
                    TXTNoteID = CStr(new_id("Notes", "NoteID", "", True))
                End If
            End If

            CheckExistingCompany (val(TxtCompanyID))

            AddNewRec
            BtnLast_Click

        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
            Ele(2).Enabled = False
    End Select

    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title

End Sub

Function CheckExistingCompany(CompanyId As Integer)

    'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim RsSavRec As ADODB.Recordset
    Dim StrRecID As String
    GetArrowsCompanyData CompanyId

    If CompanyId = 0 Then

        My_SQL = "ArrowsCompanies"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect

        StrRecID = new_id("ArrowsCompanies", "CompanyId", "")
        TxtCompanyID = StrRecID
        RsSavRec.AddNew

        RsSavRec.Fields("CompanyId").value = IIf(StrRecID <> "", StrRecID, Null)
        RsSavRec.Fields("CompanyName").value = TxtName.text
        RsSavRec.Fields("CompanySymbol").value = txtSymbol.text
        RsSavRec.Fields("groupId").value = IIf(DcboGroupID.BoundText <> 0, val(DcboGroupID.BoundText), Null)
        RsSavRec.Fields("CurrencyId").value = IIf(DcboCurrencyId.BoundText <> 0, val(DcboCurrencyId.BoundText), Null)
        RsSavRec.Fields("FinMarketId").value = IIf(DcboFinMarketId.BoundText <> 0, val(DcboFinMarketId.BoundText), Null)

        RsSavRec.Fields("CurrentValue").value = val(TxtPrice.text)
        RsSavRec.Fields("StatusId").value = 1
        RsSavRec.update
        DcboFinMarketId_Change
    End If

End Function

Private Sub BtnUndo_Click()
    FindRec val(TxtVac_ID.text)
    Me.TxtModFlg.text = "R"
End Sub

Private Sub BtnUpdate_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    FristCount = RsSavRec.RecordCount
    RsSavRec.Requery
    LastCount = RsSavRec.RecordCount
    BtnUndo_Click

    If FristCount = LastCount Then
        Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
    Else
        Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
        End If
    End If

    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic

    Me.Caption = "ArrowsTransactions Data"
    Me.Label1(2).Caption = Me.Caption
    Label1(3).Caption = "Code"
    Label1(0).Caption = "Name"
    Label1(1).Caption = "Neighborhood"

    Label2(0).Caption = "Current Record"
    Label2(1).Caption = "NO. Recordes"

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "Ser"
        .TextMatrix(0, .ColIndex("OprId")) = "Id"
        .TextMatrix(0, .ColIndex("CityName")) = "Name"
        .TextMatrix(0, .ColIndex("GovernmentID")) = "Neighborhood"
    End With

End Sub

Private Sub CboPayMentType_Change()

    If Me.CboPaymentType.ListIndex = 0 Then
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
     
    ElseIf Me.CboPaymentType.ListIndex = 1 Then
        Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.text = ""
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
    Else
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
    End If

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub Command1_Click()
    ShowGL_cc Me.TxtNoteSerial.text, , 200
End Sub

Private Sub DcboBankName_Change()
    DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.DcboBankName.BoundText))
 
End Sub

Private Sub DcboBankName_Click(Area As Integer)
    DcboBankName_Change
End Sub

Private Sub DcboBox_Change()
 
    DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
 
End Sub

Private Sub DcboBox_Click(Area As Integer)
    DcboBox_Change
End Sub

Private Sub DcboFinMarketId_Change()
    FillCompanyDatagrid (val(Me.DcboFinMarketId.BoundText))

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
                btnSave_Click

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
    Set FrmVacancy = Nothing

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

    Set cSearch = Nothing
ErrTrap:
End Sub

Private Sub Form_Activate()
    Me.ZOrder 0
End Sub

Public Sub AddNewRec()
    'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("ArrowsTransactions", "OprId", "")

    RsSavRec.AddNew
    RsSavRec.Fields("OprId").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Public Sub FiLLRec()
    'On Error GoTo ErrTrap
    RsSavRec("NoteID").value = IIf(Trim(Me.TXTNoteID.text) = "", Null, TXTNoteID.text)
    RsSavRec("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, TxtNoteSerial.text)
 
    RsSavRec.Fields("oprTypeId").value = 1 'oprType
    RsSavRec.Fields("BocketId").value = IIf(IsNumeric(TxtBocketId.text), val(TxtBocketId.text), 0)
    RsSavRec.Fields("CompanyID").value = IIf(IsNumeric(TxtCompanyID.text), val(TxtCompanyID.text), 0)
    RsSavRec.Fields("Oprdate").value = Me.DpOprdate.value

    RsSavRec.Fields("noteserial").value = IIf((TxtNoteSerial.text) <> "", val(TxtNoteSerial.text), Null)
    RsSavRec.Fields("noteserial1").value = IIf((TxtNoteSerial1.text) <> "", val(TxtNoteSerial1.text), Null)

    RsSavRec.Fields("Commpercentage").value = IIf(IsNumeric(txtCommpercentage.text), val(txtCommpercentage.text), 0)
    RsSavRec.Fields("Commvalue").value = IIf(IsNumeric(txtCommvalue.text), val(txtCommvalue.text), 0)

    RsSavRec.Fields("CommType").value = OpCommindex

    RsSavRec.Fields("CurrentValue").value = IIf(IsNumeric(TxtCurrentValue.text), val(TxtCurrentValue.text), 0)

    RsSavRec.Fields("qty").value = IIf(IsNumeric(TxtQty.text), val(TxtQty.text), 0)
    RsSavRec.Fields("Price").value = IIf(IsNumeric(TxtPrice.text), val(TxtPrice.text), 0)
    RsSavRec.Fields("total").value = val(RsSavRec.Fields("qty").value) * val(RsSavRec.Fields("Price").value) + val(RsSavRec.Fields("Commvalue").value)

    If Me.CboPaymentType.ListIndex = 0 Then
        RsSavRec("BoxID").value = val(DcboBox.BoundText)
        RsSavRec("BankID").value = Null
        RsSavRec("ChqueNum").value = Null
        RsSavRec("DueDate").value = Null
       
    ElseIf Me.CboPaymentType.ListIndex = 1 Then
        RsSavRec("BoxID").value = Null
        RsSavRec("BankID").value = val(Me.DcboBankName.BoundText)
        RsSavRec("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        RsSavRec("DueDate").value = Me.DtpChequeDueDate.value
   
    End If
    
    RsSavRec.Fields("PaymentType").value = CboPaymentType.ListIndex
    
    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    FillGridWithData2
    TxtModFlg = "R"

    If CreateJL = False Then '«‰‘«¡ «·ﬁÌÊœ
        GoTo ErrTrap
    End If

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Function CreateJL() As Boolean
    CreateJL = False
    Dim LngDevID As Long
    Dim DepitAccount As String
    Dim CreditAccount1 As String
    Dim CreditAccount2 As String
    Dim Msg As String
    Dim Arrows_group As Integer
    Dim rsOut As New ADODB.Recordset
    Dim GroupID As Integer
    Dim total As Double
 
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!Arrows_group = True Then
            Arrows_group = 1
        Else
            Arrows_group = 0
        End If
    End If

    'CreditAccount2 = get_account_code_branch(46, my_branch)

    'If CreditAccount2 = "NO branch" Then
    '        If SystemOptions.UserInterface = ArabicInterface Then
    '        Msg = "·« ÌÊÃœ Õ”«»«  ·Â–« «·›—⁄"
    '        Else
    '        Msg = "No Accounts For This Branch"
    '        End If
    '        MsgBox Msg, vbCritical
    '        CreateJL = False
    'Exit Function

    'ElseIf CreditAccount2 = "NO account" Then
    '        If SystemOptions.UserInterface = ArabicInterface Then
    '        Msg = "Õ”«» Ê”Ìÿ «›  «ÕÌ ··«”Â„ €Ì— „Õœœ ›Ï «·›—⁄"
    '        Else
    '        Msg = "ArrowsOpening Balance Account Not Defined In this Branch"
    '        End If
    '        MsgBox Msg, vbCritical
    'CreateJL = False
    'Exit Function
    'End If

    total = val(TxtQty.text) * val(TxtPrice) + val(txtCommvalue.text)
    CreditAccount2 = DcboCreditSide.BoundText
    Dim sql As String
    sql = "Delete   from notes where NoteID=" & val(TXTNoteID.text)
    Cn.Execute sql
    '«‰‘«¡ «·ﬁÌÊœ
    Dim RsNotes As ADODB.Recordset
    Dim RsDev As ADODB.Recordset
    Dim NoteID As String
    Set RsNotes = New ADODB.Recordset
    RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  
    RsNotes.AddNew
    
    RsNotes("NoteID").value = CStr(TXTNoteID.text)
    RsNotes("Note_Value").value = total
    RsNotes("Remark").value = ""
    RsNotes("NoteType").value = 902
    RsNotes("NoteDate").value = Me.DpOprdate.value
    RsNotes("UserID").value = user_id
    RsNotes("NoteSerial").value = Trim$(Me.TxtNoteSerial.text) '„”·”· «·ﬁÌœ
    RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    RsNotes("sanad_year").value = year(DpOprdate.value)
    RsNotes("sanad_month").value = Month(DpOprdate.value)
    RsNotes("note_value_by_characters").value = WriteNo(Format(val(total), "0.00"), 0, True, ".")

    If Me.CboPaymentType.ListIndex = 0 Then
        RsNotes("BoxID").value = val(DcboBox.BoundText)
        RsNotes("BankID").value = Null
        RsNotes("ChqueNum").value = Null
        RsNotes("DueDate").value = Null
        RsNotes("NoteCashingType").value = 0
    ElseIf Me.CboPaymentType.ListIndex = 1 Then
        RsNotes("BoxID").value = Null
        RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
        RsNotes("NoteCashingType").value = 1
    End If
    
    RsNotes.update
    Dim des As String
    des = ""
 
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

    If Arrows_group = 0 Then
        '„œÌ‰
        DepitAccount = get_account_code_branch(43, my_branch)

        If ModAccounts.AddNewDev(LngDevID, 0, DepitAccount, val(total), 0, des, val(Me.TXTNoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.DpOprdate.value, user_id) = False Then
            GoTo ErrTrap
                    
        End If

    Else
        '„œÌ‰
        GetArrowsCompanyData val(TxtCompanyID.text), , , GroupID
        GetArrowsGroupAccount GroupID, , DepitAccount

        If ModAccounts.AddNewDev(LngDevID, 0, DepitAccount, val(total), 0, des, val(Me.TXTNoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.DpOprdate.value, user_id) = False Then
            GoTo ErrTrap
                    
        End If
            
    End If

    '            œ«∆‰ 1
    If ModAccounts.AddNewDev(LngDevID, 1, CreditAccount2, val(total), 1, des, val(Me.TXTNoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.DpOprdate.value, user_id) = False Then
        GoTo ErrTrap
                    
    End If

    CreateJL = True
    Exit Function
ErrTrap:
    CreateJL = False
End Function

Public Sub FiLLTXT()

    'On Error GoTo ErrTrap
    Dim i As Integer
    'Frm2.Enabled = False
    Dim currentvalue As Double
    Dim CompanySymbol As String
    Dim comanyname As String
    Dim BocketCode As String
    Dim Balance As Double
    Dim CompanyId As Integer
    Dim BocketId As Integer
    TxtCompanyID.text = IIf(IsNull(RsSavRec.Fields("CompanyID").value), "", RsSavRec.Fields("CompanyID").value)
    CompanyId = val(TxtCompanyID.text)

    GetArrowsCompanyData CompanyId, CompanySymbol, comanyname, , , currentvalue
    txtSymbol = CompanySymbol
    TxtName = comanyname
    TxtCurrentValue = currentvalue
    TxtCompanyID = CompanyId

    TxtBocketId.text = IIf(IsNull(RsSavRec.Fields("BocketId").value), "", RsSavRec.Fields("BocketId").value)
    BocketId = val(TxtBocketId.text)
    GetArrowsBocketData BocketId, BocketCode, Balance
    TxtBocketCode.text = BocketCode
    TxtBalance.text = Balance

    TxtVac_ID.text = IIf(IsNull(RsSavRec.Fields("OprId").value), "", RsSavRec.Fields("OprId").value)
    TxtoprType.text = IIf(IsNull(RsSavRec.Fields("oprTypeId").value), "", RsSavRec.Fields("oprTypeId").value)

    Me.DpOprdate.value = RsSavRec.Fields("Oprdate").value
    TxtNoteSerial.text = IIf(IsNull(RsSavRec.Fields("noteserial").value), "", RsSavRec.Fields("noteserial").value)
    TxtNoteSerial1.text = IIf(IsNull(RsSavRec.Fields("noteserial1").value), "", RsSavRec.Fields("noteserial1").value)
    TXTNoteID.text = IIf(IsNull(RsSavRec.Fields("Noteid").value), "", RsSavRec.Fields("Noteid").value)

    txtCommpercentage.text = IIf(IsNull(RsSavRec.Fields("Commpercentage").value), "", RsSavRec.Fields("Commpercentage").value)
    txtCommvalue.text = IIf(IsNull(RsSavRec.Fields("Commvalue").value), "", RsSavRec.Fields("Commvalue").value)
 
    OpCommindex = RsSavRec.Fields("CommType").value
    CalComm OpCommindex, Me.txtCommpercentage.text, Me.txtCommvalue.text

    TxtCurrentValue.text = IIf(IsNull(RsSavRec.Fields("CurrentValue").value), "", RsSavRec.Fields("CurrentValue").value)

    TxtQty.text = IIf(IsNull(RsSavRec.Fields("qty").value), "", RsSavRec.Fields("qty").value)
    TxtPrice.text = IIf(IsNull(RsSavRec.Fields("Price").value), "", RsSavRec.Fields("Price").value)
    TxtTotal.text = IIf(IsNull(RsSavRec.Fields("total").value), "", RsSavRec.Fields("total").value)
  
    CboPaymentType.ListIndex = IIf(IsNull(RsSavRec.Fields("PaymentType").value), -1, RsSavRec.Fields("PaymentType").value)
    DcboBox.BoundText = IIf(IsNull(RsSavRec.Fields("BoxID").value), "", RsSavRec.Fields("BoxID").value)
    DcboBankName.BoundText = IIf(IsNull(RsSavRec.Fields("BankID").value), "", RsSavRec.Fields("BankID").value)
    TxtChequeNumber.text = IIf(IsNull(RsSavRec.Fields("ChqueNum").value), "", RsSavRec.Fields("ChqueNum").value)

    Me.DtpChequeDueDate.value = IIf(IsNull(RsSavRec.Fields("DueDate").value), Date, RsSavRec.Fields("DueDate").value)
 
    CalComm OpCommindex, Me.txtCommpercentage.text, Me.txtCommvalue.text

    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With VSFlexGrid2

        For i = 1 To .Rows - 1

            If Trim(TxtVac_ID.text) = .TextMatrix(i, .ColIndex("OprId")) Then
                TxtSerial.text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Sub EditRec(StrTable As String, _
                   RecID As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("OprId")))
ErrTrap:
End Sub

Private Sub Timer1_Timer()

    If TxtPrice.backcolor = vbWhite Then
        TxtPrice.backcolor = vbBlue
    Else
        TxtPrice.backcolor = vbWhite

    End If

End Sub

Private Sub TxtPrice_Change()
    CalComm OpCommindex, Me.txtCommpercentage.text, Me.txtCommvalue.text
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecID As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "OprId=" & RecID, , adSearchForward, 1

    If Not (RsSavRec.EOF) Then
        FiLLTXT
    
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If

    'RsSavRec.Filter = adFilterNone
End Function

'Private Sub TxtVacCode_KeyPress(KeyAscii As Integer)
'KeyAscii = DataFormat(ChrOnly, KeyAscii)
'End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.text = "N" Then
     
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        'Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        '    btnNext.Enabled = False
        '    btnPrevious.Enabled = False
        '    btnFirst.Enabled = False
        '    btnLast.Enabled = False
    
    ElseIf TxtModFlg.text = "R" Then
     
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False

        If TxtVac_ID.text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If

        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
    
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
    
    ElseIf TxtModFlg.text = "E" Then
        'Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        '  Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    
    End If

End Sub

Public Sub FillGridWithData2()

    'On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Dim comanyname As String
    Dim CompanyId As Integer
    Set rs = New ADODB.Recordset
    My_SQL = "select * From ArrowsTransactions where oprTypeId=1  order by OprId"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
Dim total As Double
Dim Net As Double
    With Me.VSFlexGrid2
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst
        
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("OprId")) = IIf(IsNull(rs.Fields("OprId").value), "", rs.Fields("OprId").value)
            
                .TextMatrix(i, .ColIndex("Oprdate")) = IIf(IsNull(rs.Fields("Oprdate").value), "", rs.Fields("Oprdate").value)
            
                .TextMatrix(i, .ColIndex("qty")) = IIf(IsNull(rs.Fields("qty").value), "", rs.Fields("qty").value)
               total = total + val(.TextMatrix(i, .ColIndex("qty")))
            
               
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(rs.Fields("Price").value), "", rs.Fields("Price").value)
               
                '               .TextMatrix(i, .ColIndex("Commpercentage")) = IIf(IsNull(rs.Fields("Commpercentage").value), _
                                "", rs.Fields("Commpercentage").value)
         
                .TextMatrix(i, .ColIndex("Commvalue")) = IIf(IsNull(rs.Fields("Commvalue").value), "", rs.Fields("Commvalue").value)
         
                .TextMatrix(i, .ColIndex("Commvalue")) = IIf(IsNull(rs.Fields("Commvalue").value), "", rs.Fields("Commvalue").value)
             
                CompanyId = IIf(IsNull(rs.Fields("CompanyId").value), 0, rs.Fields("CompanyId").value)
            
                GetArrowsCompanyData CompanyId, , comanyname
          
                .TextMatrix(i, .ColIndex("Name")) = comanyname
           
                .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(rs.Fields("total").value), 0, rs.Fields("total").value) - val(.TextMatrix(i, .ColIndex("Commvalue")))
             
                .TextMatrix(i, .ColIndex("Net")) = IIf(IsNull(rs.Fields("total").value), "", rs.Fields("total").value)
             
                .TextMatrix(i, .ColIndex("CurrentValue")) = IIf(IsNull(rs.Fields("CurrentValue").value), "", rs.Fields("CurrentValue").value)
              Net = Net + val(.TextMatrix(i, .ColIndex("Net")))
                rs.MoveNext
            Next

            rs.Close
        End If
Me.LblTotal.Caption = total
Me.LblNet.Caption = Net

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

'-------------------------------------------------------------
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = Chr(13) + Chr(10)

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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
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

Private Function CheckDelCountry(Lngid As Long) As Boolean
    'Dim Rs As ADODB.Recordset
    'Dim StrSQL As String
    'StrSQL = "Select * From TblEmployee Where GovernmentID=" & Lngid & ""
    'Set Rs = New ADODB.Recordset
    'Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If Not (Rs.BOF Or Rs.EOF) Then
    '    CheckDelCountry = False
    'Else
    '    CheckDelCountry = True
    'End If
    'Rs.Close
    'Set Rs = Nothing
End Function

Public Sub FillGridWithData()

    'On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Dim BankName As String
    Dim BankID As String
    Set rs = New ADODB.Recordset
    My_SQL = "select * From Bockets order by BocketId"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("BocketId")) = IIf(IsNull(rs.Fields("BocketId").value), "", rs.Fields("BocketId").value)
               
                .TextMatrix(i, .ColIndex("BocketCode")) = IIf(IsNull(rs.Fields("BocketCode").value), "", rs.Fields("BocketCode").value)
                BankID = val(IIf(IsNull(rs.Fields("Bankid").value), "", rs.Fields("Bankid").value))
            
                Get_Name "BanksData", "BankID", False, BankID, "BankName", BankName
                .TextMatrix(i, .ColIndex("BankId")) = BankName

                If rs.Fields("AccountType").value = 0 Then
                    .TextMatrix(i, .ColIndex("AccountType")) = "Õ”«» »‰ﬂÌ"
                ElseIf rs.Fields("AccountType").value = 1 Then
                    .TextMatrix(i, .ColIndex("AccountType")) = "Õ”«» ‰ﬁœÌ - ÃÌ»"
                ElseIf rs.Fields("AccountType").value = 2 Then
                    .TextMatrix(i, .ColIndex("AccountType")) = "Õ”«» «” À„«—Ì „Õ›Ÿ…"
                Else
                    .TextMatrix(i, .ColIndex("AccountType")) = ""
                End If

                .TextMatrix(i, .ColIndex("Balance")) = IIf(IsNull(rs.Fields("Balance").value), 0, rs.Fields("Balance").value)
                      
                .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(rs.Fields("RecordDate").value), Date, rs.Fields("RecordDate").value)
 
                .TextMatrix(i, .ColIndex("BankComm")) = IIf(IsNull(rs.Fields("BankComm").value), 0, rs.Fields("BankComm").value)
            
                .TextMatrix(i, .ColIndex("NetComm")) = IIf(IsNull(rs.Fields("NetComm").value), 0, rs.Fields("NetComm").value)
            
                .TextMatrix(i, .ColIndex("MinComm")) = IIf(IsNull(rs.Fields("MinComm").value), 0, rs.Fields("MinComm").value)
            
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

Public Sub FillCompanyDatagrid(FinMarketId As Integer)

    If FinMarketId = 0 Then Exit Sub
    'On Error GoTo ErrTrap

    If FinMarketId = 1 Then

        With Me.VSFlexGrid1
            .Rows = 2
            .Clear flexClearScrollable
        End With

        Frame1.Visible = False
        NEW_interface = True
        path = "http://www.tadawul.com.sa/wps/portal/!ut/p/c1/04_SB8K8xLLM9MSSzPy8xBz9CP0os3g_A-ewIE8TIwMLj2AXA0_vQGNzY18g18cQKB-JJO8eEGZq4GniE2wUHOBlbOBpREB3cGKRvp9Hfm6qfkFuRDkAgpcLJw!!/dl2/d1/L2dJQSEvUUt3QS9ZQnB3LzZfTjBDVlJJNDIwMFM1MDBJNExWVENMRzMwMjY!/"
        WebBrowser1.Navigate2 path
        Exit Sub
    End If

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Dim BankName As String
    Dim BankID As String
    Dim GroupID As Integer
    Set rs = New ADODB.Recordset
    My_SQL = "select * From ArrowsCompanies where FinMarketId=" & FinMarketId & "  order by CompanySymbol"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    Dim HyperLink As String
    Dim GroupName As String
    Dim CurrencyId As Integer
    Dim CurrencyName As String

    With Me.VSFlexGrid1
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("CompanyId")) = IIf(IsNull(rs.Fields("CompanyId").value), "", rs.Fields("CompanyId").value)
             
                GroupName = ""
                CurrencyName = ""
                .TextMatrix(i, .ColIndex("groupId")) = IIf(IsNull(rs.Fields("groupId").value), "", rs.Fields("groupId").value)
            
                GroupID = val(.TextMatrix(i, .ColIndex("groupId")))
            
                GetArrowsGroupAccount GroupID, , , , , , GroupName
                .TextMatrix(i, .ColIndex("groupName")) = GroupName
            
                .TextMatrix(i, .ColIndex("CurrencyId")) = IIf(IsNull(rs.Fields("CurrencyId").value), "", rs.Fields("CurrencyId").value)
            
                CurrencyId = val(.TextMatrix(i, .ColIndex("CurrencyId")))
            
                GetCurrencyData CurrencyId, , CurrencyName
                .TextMatrix(i, .ColIndex("CurrencyName")) = CurrencyName
            
                .TextMatrix(i, .ColIndex("CompanySymbol")) = IIf(IsNull(rs.Fields("CompanySymbol").value), "", rs.Fields("CompanySymbol").value)
               
                .TextMatrix(i, .ColIndex("CompanyName")) = IIf(IsNull(rs.Fields("CompanyName").value), "", rs.Fields("CompanyName").value)
                       
                .TextMatrix(i, .ColIndex("CurrentValue")) = IIf(IsNull(rs.Fields("CurrentValue").value), 0, rs.Fields("CurrentValue").value)
            
                get_Financial_market_data val(DcboFinMarketId.BoundText), , , HyperLink
                     
                If FinMarketId = 3 Then
                    .TextMatrix(i, .ColIndex("Hyperlink")) = HyperLink
 
                Else
                    .TextMatrix(i, .ColIndex("Hyperlink")) = HyperLink & IIf(IsNull(rs.Fields("CompanySymbol").value), "", rs.Fields("CompanySymbol").value)
                       
                End If
                  
                .TextMatrix(i, .ColIndex("LastPrice")) = "«÷€ÿ Â‰« „— Ì‰"
                 
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

Private Sub Form_Load()
    'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos

    My_SQL = "ArrowsTransactions"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Me.TxtModFlg.text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL

    Set Dcombos = New ClsDataCombos
    Dcombos.getˆaArrowsGroup Me.DcboGroupID

    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.DcboBankName

    Dcombos.GetAccountingCodes Me.DcboCreditSide

    With Me.CboPaymentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "‘Ìﬂ"
    End With

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select id,name from currency  order by name  "
    Else
        My_SQL = "  select id,code from currency  order by code  "
    End If

    fill_combo DcboCurrencyId, My_SQL

    'ModFgLib.LinkFgColWithDataCombo VSFlexGrid1, VSFlexGrid1.ColIndex("groupName"), Me.DcbogroupId
    'ModFgLib.LinkFgColWithDataCombo VSFlexGrid1, VSFlexGrid1.ColIndex("CurrencyName"), Me.DcboCurrencyId

    'ArrowsGroup
   
    Set cSearch = New clsDCboSearch
    FillGridWithData2

    BtnFirst_Click
    ShowTip

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If

ErrTrap:
 
    Set Dcombos = New ClsDataCombos
    Dcombos.getFinMarkets DcboFinMarketId

    Resize_Form Me
    NEW_interface = False
    FillGridWithData
 
    'WebBrowser1.Navigate2 "http://www.tadawul.com.sa/Resources/Reports/DetailedDaily_ar.html"
End Sub

Private Sub Grid_Click()
    CalComm OpCommindex, Me.txtCommpercentage.text, Me.txtCommvalue.text

    With Grid

        If Not .TextMatrix(.Row, .ColIndex("BocketId")) = "" Then
            TxtBocketCode.text = .TextMatrix(.Row, .ColIndex("BocketCode"))
            TxtBalance.text = val(.TextMatrix(.Row, .ColIndex("Balance")))
            TxtBocketId.text = val(.TextMatrix(.Row, .ColIndex("BocketId")))
            
        End If

    End With

End Sub

Private Sub Label9_Click()
    Frame1.Visible = False
End Sub

Private Sub OpComm_Click(Index As Integer)
    CalComm Index, Me.txtCommpercentage.text, Me.txtCommvalue.text
    OpCommindex = Index
End Sub

Function CalComm(CommType As Integer, ByRef Commpercentage As String, ByRef commvalue As String)
    Dim totalprice As Double
    Dim Mincomm As Double
 
    With Grid

        If Not .TextMatrix(.Row, .ColIndex("BocketCode")) = "" Then
            totalprice = val(TxtQty.text) * val(Me.TxtPrice.text)
            Mincomm = .TextMatrix(.Row, .ColIndex("MinComm"))
            TxtTotal.text = totalprice

            Select Case CommType

                Case 0
                    Commpercentage = val(.TextMatrix(.Row, .ColIndex("NetComm")))
                    commvalue = Commpercentage * totalprice / 100
        
                    If commvalue < Mincomm Then
                        commvalue = Mincomm
                    End If
        
                Case 1
                    Commpercentage = val(.TextMatrix(.Row, .ColIndex("BankComm")))
                    commvalue = Commpercentage * totalprice / 100
        
                    If commvalue < Mincomm Then
                        commvalue = Mincomm
                    End If
        
                Case 2

                    If val(.TextMatrix(.Row, .ColIndex("NetComm"))) <= val(.TextMatrix(.Row, .ColIndex("BankComm"))) Then
                        Commpercentage = val(.TextMatrix(.Row, .ColIndex("NetComm")))
                    Else
                        Commpercentage = val(.TextMatrix(.Row, .ColIndex("BankComm")))
                    End If

                    commvalue = Commpercentage * totalprice / 100
        
                    If commvalue < Mincomm Then
                        commvalue = Mincomm
                    End If
        
            End Select

        End If

        txtCommpercentage.text = val(Commpercentage)
        txtCommvalue.text = val(commvalue)
    End With

End Function

Private Sub TxtPurchasePrice_Change()
    CalComm OpCommindex, Me.txtCommpercentage.text, Me.txtCommvalue.text
End Sub

Private Sub TxtQty_Change()

    CalComm OpCommindex, Me.txtCommpercentage.text, Me.txtCommvalue.text
End Sub

Public Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                 ByVal Col As Long)
    'check_cost_center
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
    Dim project_id As Integer

    With VSFlexGrid1

        Select Case .ColKey(Col)
 
            Case "groupName"
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("groupName"), False, True)
                .TextMatrix(Row, .ColIndex("groupid")) = StrAccountCode
                StrSQL = "Select  * from  ArrowsGroup where groupid=" & val(StrAccountCode)
               
                If StrAccountCode <> "" Then
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        .TextMatrix(Row, .ColIndex("groupName")) = IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
                          
                    End If
                End If
            
            Case "CurrencyName"
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CurrencyName"), False, True)
                .TextMatrix(Row, .ColIndex("CurrencyId")) = StrAccountCode
                StrSQL = "Select  * from  currency where id=" & val(StrAccountCode)
               
                If StrAccountCode <> "" Then
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        .TextMatrix(Row, .ColIndex("CurrencyName")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                          
                    End If
                End If

        End Select
 
        Me.DcboGroupID.BoundText = val(.TextMatrix(.Row, .ColIndex("groupId")))
        DcboCurrencyId.BoundText = val(.TextMatrix(.Row, .ColIndex("CurrencyId")))

    End With

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid1

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

        If val(.TextMatrix(.Row, .ColIndex("NetLastPrice"))) > 0 Then
            .ComboList = ""
            Exit Sub
        End If
  
        If .ColKey(Col) = "groupName" Then

        ElseIf .ColKey(Col) = "CurrencyName" Then

        Else
            .ComboList = ""
        End If
        
        '  Cancel = True
     
    End With

End Sub

Private Sub VSFlexGrid1_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 80

    End If

End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String
    Dim Msg As String
    Dim project_id As Integer
    Dim whrstring As String

    With VSFlexGrid1

        Select Case .ColKey(Col)

            Case "groupName"
   
                StrSQL = " select * from ArrowsGroup "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList1 = .BuildComboList(rs, "GroupName", "GroupID")
 
                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
            
            Case "CurrencyName"
      
                StrSQL = "  select id,name from currency "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList1 = .BuildComboList(rs, "id,name", "id")
 
                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
          
        End Select

    End With

End Sub

Private Sub VSFlexGrid1_Click()

    With VSFlexGrid1

        If Not .TextMatrix(.Row, .ColIndex("CompanyID")) = "" Then
            txtSymbol = .TextMatrix(.Row, .ColIndex("CompanySymbol"))
            TxtName = .TextMatrix(.Row, .ColIndex("CompanyName"))
            TxtCurrentValue = val(.TextMatrix(.Row, .ColIndex("CurrentValue")))
            TxtCompanyID = val(.TextMatrix(.Row, .ColIndex("CompanyID")))
            Me.DcboGroupID.BoundText = val(.TextMatrix(.Row, .ColIndex("groupId")))
            DcboCurrencyId.BoundText = val(.TextMatrix(.Row, .ColIndex("CurrencyId")))
   
            If val(.TextMatrix(.Row, .ColIndex("NetLastPrice"))) > 0 Then
                TxtPrice = val(.TextMatrix(.Row, .ColIndex("NetLastPrice")))
            Else
  
                TxtPrice = ""
            End If
 
        End If

    End With

End Sub

Private Sub VSFlexGrid1_DblClick()

    With VSFlexGrid1

        If Not .TextMatrix(.Row, .ColIndex("HyperLink")) = "" Then
            Frame1.Visible = True

            WebBrowser3.Navigate .TextMatrix(.Row, .ColIndex("Hyperlink"))
 
        End If

    End With

End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, _
                                         URL As Variant)
    On Error GoTo ErrTrap

    If NEW_interface = False Then Exit Sub
    Dim i As Integer

    Dim objTable As Object
 
    'The ninth table in the page is the Companies List
    Dim startLoad As Integer
    Dim Cols As Integer

    'On Error Resume Next
    DoEvents
    startLoad = 77 + 17
    Dim lastCompanyId As Integer
    lastCompanyId = CStr(new_id("ArrowsCompanies", "CompanyId", "", True))
    Set objTable = WebBrowser1.Document.getElementsByTagName("table").Item(13)

    With Me.VSFlexGrid1
 
        .Rows = objTable.getElementsByTagName("tr").Length - 1
 
        For i = startLoad To .Rows
            Cols = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Length
            Dim HyperLink  As String
            Dim SymbolNo As Integer

            If Cols >= 2 Then
                '      .TextMatrix((i - startLoad) + 1, .ColIndex("LineNo")) = (i - startLoad) + 1
                .TextMatrix((i - startLoad) + 1, .ColIndex("CompanyName")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(0).innerText
      
            End If
      
            Dim CompanyId As Integer
            Dim GroupID As Integer
            Dim CurrencyId As Integer
            Dim currentvalue As Double
            Dim CompanySymbol As String
            Dim GroupName As String
            Dim CurrencyName  As String
      
            If Cols = 14 Then
                HyperLink = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("a")
                SymbolNo = right(HyperLink, 4)
                .TextMatrix((i - startLoad) + 1, .ColIndex("CompanySymbol")) = SymbolNo
       
                .TextMatrix((i - startLoad) + 1, .ColIndex("LastPrice")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(1).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("NetLastPrice")) = .TextMatrix((i - startLoad) + 1, .ColIndex("LastPrice"))
     
                CompanyId = 0
                GroupID = 0
                CurrencyId = 0
                currentvalue = 0
                CompanySymbol = 0
                CompanySymbol = SymbolNo
                GetArrowsCompanyData CompanyId, CompanySymbol, , GroupID, , currentvalue, CurrencyId

                If CompanyId <> 0 Then
                    .TextMatrix((i - startLoad) + 1, .ColIndex("CompanyId")) = CompanyId
                Else
                    .TextMatrix((i - startLoad) + 1, .ColIndex("CompanyId")) = lastCompanyId
                    lastCompanyId = lastCompanyId + 1
                End If
 
                .TextMatrix((i - startLoad) + 1, .ColIndex("groupid")) = GroupID
                .TextMatrix((i - startLoad) + 1, .ColIndex("CurrentValue")) = currentvalue
                .TextMatrix((i - startLoad) + 1, .ColIndex("CurrencyId")) = CurrencyId
           
                GroupName = ""
                CurrencyName = ""
                GetArrowsGroupAccount GroupID, , , , , , GroupName
                .TextMatrix((i - startLoad) + 1, .ColIndex("groupName")) = GroupName
           
                GetCurrencyData CurrencyId, , CurrencyName
                .TextMatrix((i - startLoad) + 1, .ColIndex("CurrencyName")) = CurrencyName
 
            End If

        Next i

        .AutoSize 0, .Cols - 1, False
        Dim j As Integer
        Dim lastindex As Integer

        For j = .Rows - 1 To 2 Step -1

            If .TextMatrix(j, .ColIndex("CompanyName")) <> "" Then
                lastindex = j + 1
                GoTo LL
            End If

        Next j

LL:
        .Rows = lastindex + 1
    End With

    Set objTable = Nothing
    Exit Sub
ErrTrap:
    MsgBox "·«»œ „‰ «·« ’«· »«·«‰ —‰  «Ê·«"
End Sub
 
