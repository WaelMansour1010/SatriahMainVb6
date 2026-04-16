VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmEmpExpir2 
   Caption         =   "«”„«¡ «·„ÊŸ›Ì‰ «· Ì ” ‰ ÂÌ «ﬁ«„ Â„"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   20160
   Icon            =   "FrmEmpExpir2.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   20160
      _cx             =   35560
      _cy             =   15266
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   1095
         Left            =   8880
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   7440
         Width           =   10080
         _cx             =   17780
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
         Begin VB.CommandButton Command2 
            Caption         =   "ÿ»«⁄Â"
            Height          =   315
            Left            =   630
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   720
            Width           =   1260
         End
         Begin VB.CommandButton Command1 
            Caption         =   "«” ⁄·«„"
            Height          =   315
            Left            =   630
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   360
            Width           =   1260
         End
         Begin Dynamic_Byte.NourHijriCal Txt_to_H 
            Height          =   255
            Left            =   5475
            TabIndex        =   33
            Top             =   720
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   450
         End
         Begin Dynamic_Byte.NourHijriCal Txt_from_H 
            Height          =   300
            Left            =   5475
            TabIndex        =   34
            Top             =   360
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   529
         End
         Begin MSComCtl2.DTPicker d1 
            Height          =   315
            Left            =   2670
            TabIndex        =   35
            Top             =   360
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   556
            _Version        =   393216
            Format          =   126550017
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker d2 
            Height          =   315
            Left            =   2670
            TabIndex        =   36
            Top             =   720
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   556
            _Version        =   393216
            Format          =   126550017
            CurrentDate     =   38784
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "«·Ï"
            Height          =   300
            Left            =   9015
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   720
            Width           =   435
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   " «—ÌŒ «·«‰ Â«¡ „‰"
            Height          =   300
            Left            =   8655
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.Frame Frm2 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   675
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   10365
         Width           =   5115
         Begin VB.TextBox TxtVacName 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   75
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·√Ã«“…"
            Top             =   285
            Visible         =   0   'False
            Width           =   3750
         End
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
            Left            =   2400
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   285
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.ComboBox CmbType 
            BackColor       =   &H80000018&
            Height          =   315
            ItemData        =   "FrmEmpExpir2.frx":058A
            Left            =   2280
            List            =   "FrmEmpExpir2.frx":059A
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   870
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·ÊŸÌ›…"
            Height          =   285
            Index           =   0
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   0
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·ÊŸÌ›…"
            Height          =   195
            Index           =   3
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   360
            Visible         =   0   'False
            Width           =   990
         End
      End
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   0
         Width           =   20250
         Begin VB.Frame Frmo2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   450
            Visible         =   0   'False
            Width           =   3105
            Begin MSDataListLib.DataCombo DCUser 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   -255
               TabIndex        =   5
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
               TabIndex        =   6
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
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Text            =   "modflag"
            Top             =   120
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
            TabIndex        =   2
            Top             =   510
            Visible         =   0   'False
            Width           =   945
         End
         Begin MSComctlLib.ImageList GrdImageList 
            Left            =   480
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
                  Picture         =   "FrmEmpExpir2.frx":05B3
                  Key             =   "CompanyName"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmEmpExpir2.frx":094D
                  Key             =   "Ser"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmEmpExpir2.frx":0CE7
                  Key             =   "Vac_Name"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmEmpExpir2.frx":1081
                  Key             =   "ShareCount"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmEmpExpir2.frx":141B
                  Key             =   "Dis_Count"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmEmpExpir2.frx":17B5
                  Key             =   "Bouns"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmEmpExpir2.frx":1B4F
                  Key             =   "SharesValue"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmEmpExpir2.frx":20E9
                  Key             =   "BuyValue"
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«ﬁ«„«  «· Ì ” ‰ ÂÌ "
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
            Left            =   16680
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   120
            Width           =   3120
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   1020
         Left            =   0
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   9495
         Width           =   6840
         _cx             =   12065
         _cy             =   1799
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
            Left            =   4125
            TabIndex        =   15
            Top             =   555
            Visible         =   0   'False
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
            ButtonImage     =   "FrmEmpExpir2.frx":2483
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   2580
            TabIndex        =   16
            Top             =   555
            Visible         =   0   'False
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
            ButtonImage     =   "FrmEmpExpir2.frx":281D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   3345
            TabIndex        =   17
            Top             =   555
            Visible         =   0   'False
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
            ButtonImage     =   "FrmEmpExpir2.frx":2BB7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   1815
            TabIndex        =   18
            Top             =   555
            Visible         =   0   'False
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
            ButtonImage     =   "FrmEmpExpir2.frx":2F51
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   1050
            TabIndex        =   19
            Top             =   555
            Visible         =   0   'False
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
            ButtonImage     =   "FrmEmpExpir2.frx":32EB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   5880
            TabIndex        =   20
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
            ButtonImage     =   "FrmEmpExpir2.frx":3885
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   6045
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   105
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
            ButtonImage     =   "FrmEmpExpir2.frx":3C1F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   3765
            TabIndex        =   22
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
            ButtonImage     =   "FrmEmpExpir2.frx":3FB9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   2505
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   225
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   1
            Left            =   810
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   225
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   210
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   240
            Width           =   675
         End
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   210
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   225
            Width           =   540
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   6315
         Left            =   30
         TabIndex        =   27
         Top             =   990
         Width           =   20160
         _cx             =   35560
         _cy             =   11139
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
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmEmpExpir2.frx":4353
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
         ExplorerBar     =   1
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
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   330
         Left            =   450
         TabIndex        =   28
         Top             =   7170
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
         ButtonImage     =   "FrmEmpExpir2.frx":465A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid2 
         Height          =   6675
         Left            =   270
         TabIndex        =   39
         Top             =   540
         Width           =   19530
         _cx             =   34449
         _cy             =   11774
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
         Cols            =   36
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmEmpExpir2.frx":49F4
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
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "«÷€ÿ „— Ì‰ ⁄·Ì «Ì „ÊŸ› ·⁄—÷ »Ì«‰« …"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   5130
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   7170
         Width           =   3975
      End
   End
End
Attribute VB_Name = "FrmEmpExpir2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim My_SQL As String
Dim date_type As Integer
Dim xApp As New CRAXDRT.Application
Dim EmpReport As ClsEmployeeReport
Dim Askinterval As String
Dim Askcount As Integer
Public mIsEndType As Integer
Public mWhere As String
Public mNewDate As Date

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    If TxtVac_ID.Text <> "" Then
        If CheckDelJobType(val(Me.TxtVac_ID.Text)) = False Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)

        If MSGType = vbYes Then
            RsSavRec.Find "JobTypeID=" & val(TxtVac_ID.Text), , adSearchForward, 1
            RsSavRec.delete
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            '------------------------------ Move Next ---------------------------.
            FillGridWithData
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
    On Error GoTo ErrTrap

    Dim Msg As String
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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

    If TxtVac_ID.Text <> "" Then
        TxtModFlg = "E"
        Frm2.Enabled = True
        Me.TxtVacName.SetFocus
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Private Sub btnNew_Click()
 
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    TxtModFlg.Text = "N"

    My_SQL = "TblEmpJobsTypes"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.Text = rs.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If

    rs.Close
    CmbType.ListIndex = 0
    TxtVacName.SetFocus
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------

    For Each CtrlTxt In Me.Controls

        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next

    '------------------------------ check if Empcode exist ----------------------

    StrVacName = IsRecExist("TblEmpJobsTypes", "JobTypeName", Trim(TxtVacName.Text), "JobTypeName", "Vac_ID<>'" & Trim(TxtVac_ID.Text) & "'")

    If StrVacName <> "" Then
        Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
         
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtVacName.SetFocus
    
        Exit Sub

    End If

    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            AddNewRec
            BtnLast_Click

        Case "E"

            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select

    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title

End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtVac_ID.Text)
    Me.TxtModFlg.Text = "R"
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

Private Sub Command1_Click()
If mWhere <> "" Then Exit Sub
If mIsEndType <> 0 Then Exit Sub
    On Error GoTo ErrTrap
 If date_type = 0 Then date_type = 1
    If date_type = 1 Then
        d1.value = Format$(ToGregorianDate(Txt_from_H.value), "dd-mm-yyyy")
        d2.value = Format$(ToGregorianDate(Txt_to_H.value), "dd-mm-yyyy")
    End If

  '  My_SQL = "SELECT   * from dbo.TblEmployee WHERE     (DateEndekama >= CONVERT(DATETIME, '" & Format$(d1.value, "dd-mm-yyyy") & " 00:00:00', 102)) AND (DateEndekama <= CONVERT(DATETIME, '" & Format$(d2.value, "dd-mm-yyyy") & " 00:00:00', 102))"
    
    My_SQL = "SELECT   Dayss = DateDiff(d, " & SQLDate(d2.value, True) & ", DateEndekama),* from dbo.emp_all_details WHERE    dbo.emp_all_details.NationlID <>1 and    (NOT (dbo.emp_all_details.NumEkama IS NULL ))  "
    
   If mIsEndType = 1 Then
        My_SQL = My_SQL & " AND DateEndekama>'" & SQLDate(DateAdd(Askinterval, Askcount, SQLDate(d1.value, True))) & "'"
    ElseIf mIsEndType = 2 Then
        My_SQL = My_SQL & "  AND (   DateEndekama   >=" & SQLDate(d1.value, True) & "AND DateEndekama <=" & SQLDate(d2.value, True) & ")"
    Else
        My_SQL = My_SQL & " AND (   DateEndekama   >=" & SQLDate(d1.value, True) & "AND DateEndekama <=" & SQLDate(d2.value, True) & ")"
    End If
    My_SQL = My_SQL & mWhere
    
    
    My_SQL = My_SQL & " order by DateEndekama,fullcode"
    FillGridWithData

    'My_SQL = "select * From TblEmployee  where    (MONTH(DateEndekama) <= MONTH(GETDATE()))"
    'End If

    Exit Sub
ErrTrap:
    MsgBox "«œŒ·   «—ÌŒ ÂÃ—Ì Œ«ÿÌ¡", vbCritical
    
    
End Sub

Private Sub Command2_Click()
Command1_Click
    Dim rs As New ADODB.Recordset
    Dim mtxt As String
    Dim xReport As New CRAXDRT.Report

    '    Sql = "SELECT * from emp_all_details WHERE emp_code='" & FrmEmployee.TxtEmp_Code.text & "'"
    rs.Open My_SQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
    If mIsEndType = 2 Then
        Set xReport = xApp.OpenReport(App.path & "\reports\emp\REPORT12.rpt")
        mtxt = App.path & "\reports\emp\REPORT12.rpt"
    ElseIf mIsEndType = 5 Then
        Set xReport = xApp.OpenReport(App.path & "\reports\emp\REPORT12New.rpt")
        mtxt = App.path & "\reports\emp\REPORT12New.rpt"
     ElseIf mIsEndType = 3 Then
     Set xReport = xApp.OpenReport(App.path & "\reports\emp\REPORTVisa2.rpt")
     mtxt = App.path & "\reports\emp\REPORTVisa2.rpt"
      ElseIf mIsEndType = 4 Then
      Set xReport = xApp.OpenReport(App.path & "\reports\emp\REPORTVisa1.rpt")
      mtxt = App.path & "\reports\emp\REPORTVisa1.rpt"
    Else
         Set xReport = xApp.OpenReport(App.path & "\reports\emp\REPORT12.rpt")
        mtxt = App.path & "\reports\emp\REPORT12.rpt"
    End If
    xReport.Database.SetDataSource rs
    Dim FrmReport As New FrmReportViewer
    '   FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.txtPath = (mtxt)
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    xReport.ParameterFields(1).AddCurrentValue Txt_from_H.value
    xReport.ParameterFields(2).AddCurrentValue Txt_to_H.value
     
    Screen.MousePointer = vbDefault
    ' xReport.ReportTitle = X
    SendKeys "{RIGHT}"
      
End Sub

Private Sub d1_Change()
    date_type = 2
     Txt_from_H.value = ToHijriDate(d1.value)
     
End Sub

Private Sub d1_GotFocus()
    date_type = 2
End Sub

Private Sub d2_Change()
    date_type = 2
        Txt_from_H.value = ToHijriDate(d2.value)
End Sub

Private Sub d2_GotFocus()
    date_type = 2
End Sub


Private Sub Form_Load()
    d1.value = Date
    d2.value = Date
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
    ' On Error GoTo ErrTrap
    Dim RsDev As New ADODB.Recordset
    Dim i As Integer
    Dim StrSQL  As String
    My_SQL = "TblEmployee"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    If IsNull(mNewDate) Then
        mNewDate = Date
    End If
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_ExpireEkama", "D")
    Askcount = GetSetting(StrAppRegPath, "Setting", "count_ExpireEkama", 0)
'    My_SQL = "SELECT     * from dbo.TblEmployee Where DateEndekama<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
 
 Dim mCount As Long
If mIsEndType = 4 Then
    Grid.Visible = False
    GRID2.Visible = True
        
  GRID2.ColHidden(GRID2.ColIndex("PeriodStill")) = True
 '  Grid2.ColHidden(Grid2.ColIndex("Priod")) = True
   'Grid2.ColHidden(Grid2.ColIndex("ArriveDate")) = True
  GRID2.ColHidden(GRID2.ColIndex("Office")) = False
   'Grid2.ColHidden(Grid2.ColIndex("StarDate")) = True
  GRID2.ColHidden(GRID2.ColIndex("Ex")) = True
 
StrSQL = "  SELECT DISTINCT"
StrSQL = StrSQL & "         dbo.TbVisaDeti.VisaID,"
StrSQL = StrSQL & "         TbVisaDeti.remarks,"
StrSQL = StrSQL & "         TblOffice.Name   AS OfficeName,"
StrSQL = StrSQL & "         TblOffice.Namee  AS OfficeNamee,"
StrSQL = StrSQL & "         dbo.TbVisaDeti.JobID,"
StrSQL = StrSQL & "         dbo.TblEmpJobsTypes.JobTypeName,"
StrSQL = StrSQL & "         dbo.TblEmpJobsTypes.JobTypeNamee,"
StrSQL = StrSQL & "         dbo.TbVisaDeti.NotionalID,"
StrSQL = StrSQL & "         dbo.Nationality.name,"
StrSQL = StrSQL & "         dbo.Nationality.namee,"
StrSQL = StrSQL & "         dbo.TblCountriesGovernments.GovernmentName,"
StrSQL = StrSQL & "         dbo.TbVisaDeti.[count],"
StrSQL = StrSQL & "         dbo.TbVisaDeti.Place,"
StrSQL = StrSQL & "         TbVisa.orderNo,"
StrSQL = StrSQL & "         TbVisa.visano,"
StrSQL = StrSQL & "         TbVisa.Priod,"
StrSQL = StrSQL & "         TbVisa.StarDate,"
StrSQL = StrSQL & "         TbVisa.EndDate,TbVisa.ArriveDate,"
StrSQL = StrSQL & "         DateDiff(d," & SQLDate(mNewDate, True) & "    , TbVisa.ArriveDate) as Peri,"
StrSQL = StrSQL & "         TblEmployee.KafelName"
StrSQL = StrSQL & "  From dbo.TblEmpJobsTypes"
StrSQL = StrSQL & "         RIGHT OUTER JOIN dbo.TblCountriesGovernments"
StrSQL = StrSQL & "         RIGHT OUTER JOIN dbo.TbVisaDeti"
StrSQL = StrSQL & "              ON  dbo.TblCountriesGovernments.GovernmentID = dbo.TbVisaDeti.CityID"
StrSQL = StrSQL & "         LEFT OUTER JOIN dbo.Nationality"
StrSQL = StrSQL & "              ON  dbo.TbVisaDeti.NotionalID = dbo.Nationality.id"
StrSQL = StrSQL & "              ON  dbo.TblEmpJobsTypes.JobTypeID = dbo.TbVisaDeti.JobID"
StrSQL = StrSQL & "         LEFT OUTER JOIN dbo.TblEmployee"
StrSQL = StrSQL & "              ON  dbo.TbVisaDeti.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & "         LEFT OUTER JOIN dbo.TblOffice"
StrSQL = StrSQL & "              ON  dbo.TbVisaDeti.OfficeId = dbo.TblOffice.ID"
StrSQL = StrSQL & "         LEFT OUTER JOIN dbo.TbVisa"
StrSQL = StrSQL & "              ON  dbo.TbVisaDeti.VisaID = dbo.TbVisa.ID"


  
StrSQL = StrSQL & "  Where 1 = 1 And TbVisaDeti.[Type] = 0 " & mWhere

StrSQL = StrSQL & "  Group By"
StrSQL = StrSQL & "         dbo.TbVisaDeti.VisaID,"
StrSQL = StrSQL & "         TbVisaDeti.remarks,TbVisa.ArriveDate,"
StrSQL = StrSQL & "         TblOffice.Name,"
StrSQL = StrSQL & "         TblOffice.Namee,"
StrSQL = StrSQL & "         dbo.TbVisaDeti.JobID,"
StrSQL = StrSQL & "         dbo.TblEmpJobsTypes.JobTypeName,"
StrSQL = StrSQL & "         dbo.TblEmpJobsTypes.JobTypeNamee,"
StrSQL = StrSQL & "         dbo.TbVisaDeti.NotionalID,"
StrSQL = StrSQL & "         dbo.Nationality.name,"
StrSQL = StrSQL & "         dbo.Nationality.namee,"
StrSQL = StrSQL & "         dbo.TblCountriesGovernments.GovernmentName,"
StrSQL = StrSQL & "         dbo.TbVisaDeti.[count],"
StrSQL = StrSQL & "         dbo.TbVisaDeti.Place,"
StrSQL = StrSQL & "         TbVisa.orderNo,"
StrSQL = StrSQL & "         TbVisa.visano,"
StrSQL = StrSQL & "         TbVisa.Priod,"
StrSQL = StrSQL & "         TbVisa.StarDate,"
StrSQL = StrSQL & "         TbVisa.EndDate,"
StrSQL = StrSQL & "         TblEmployee.KafelName"
My_SQL = StrSQL

 Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.GRID2
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
              '  .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(RsDev("Emp_ID").value), "", RsDev("Emp_ID").value)
            
               ' .TextMatrix(i, .ColIndex("Emp_code")) = IIf(IsNull(RsDev("Fullcode").value), "", RsDev("Fullcode").value)
                 .TextMatrix(i, .ColIndex("Office")) = IIf(IsNull(RsDev("OfficeName").value), "", RsDev("OfficeName").value)
               
                If SystemOptions.UserInterface = ArabicInterface Then
                '.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
                .TextMatrix(i, .ColIndex("Job")) = IIf(IsNull(RsDev("JobTypeName").value), "", RsDev("JobTypeName").value)
                .TextMatrix(i, .ColIndex("Notional")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
                Else
               ' .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(RsDev("Emp_Namee").value), "", RsDev("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("Office")) = IIf(IsNull(RsDev("OfficeNamee").value), "", RsDev("OfficeNamee").value)
                .TextMatrix(i, .ColIndex("Job")) = IIf(IsNull(RsDev("JobTypeNamee").value), "", RsDev("JobTypeNamee").value)
                .TextMatrix(i, .ColIndex("Notional")) = IIf(IsNull(RsDev("namee").value), "", RsDev("namee").value)
                End If
              '  .TextMatrix(i, .ColIndex("HododNo")) = IIf(IsNull(RsDev("HododNo").value), "", RsDev("HododNo").value)
                .TextMatrix(i, .ColIndex("City")) = IIf(IsNull(RsDev("GovernmentName").value), "", RsDev("GovernmentName").value)
                .TextMatrix(i, .ColIndex("NotionalID")) = IIf(IsNull(RsDev("NotionalID").value), "", RsDev("NotionalID").value)
                .TextMatrix(i, .ColIndex("JobID")) = IIf(IsNull(RsDev("JobID").value), "", RsDev("JobID").value)
               ' .TextMatrix(i, .ColIndex("CityID")) = IIf(IsNull(RsDev("CityID").value), "", RsDev("CityID").value)
              '  .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), "", RsDev("Price").value)
                .TextMatrix(i, .ColIndex("remarks")) = IIf(IsNull(RsDev("remarks").value), "", RsDev("remarks").value)
            
                            .TextMatrix(i, .ColIndex("orderNo")) = IIf(IsNull(RsDev("orderNo").value), "", RsDev("orderNo").value)
                .TextMatrix(i, .ColIndex("visano")) = IIf(IsNull(RsDev("visano").value), "", RsDev("visano").value)
                .TextMatrix(i, .ColIndex("Priod")) = IIf(IsNull(RsDev("Priod").value), "", RsDev("Priod").value)
                .TextMatrix(i, .ColIndex("StarDate")) = IIf(IsNull(RsDev("StarDate").value), "", RsDev("StarDate").value)
                .TextMatrix(i, .ColIndex("EndDate")) = IIf(IsNull(RsDev("EndDate").value), "", RsDev("EndDate").value)
                .TextMatrix(i, .ColIndex("ArriveDate")) = IIf(IsNull(RsDev("ArriveDate").value), "", RsDev("ArriveDate").value)
                .TextMatrix(i, .ColIndex("PeriodStill")) = DateDiff("d", mNewDate, IIf(IsNull(RsDev("ArriveDate").value), Date, RsDev("ArriveDate").value))
                .TextMatrix(i, .ColIndex("Priod")) = DateDiff("d", mNewDate, IIf(IsNull(RsDev("ArriveDate").value), Date, RsDev("ArriveDate").value))
                .TextMatrix(i, .ColIndex("KafelName")) = IIf(IsNull(RsDev("KafelName").value), "", RsDev("KafelName").value)
           
            
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
    Txt_from_H.Visible = False
    Txt_to_H.Visible = False
    d1.Visible = False
    d2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Command1.Visible = False
    
ElseIf mIsEndType = 3 Then
    Grid.Visible = False
    GRID2.Visible = True

    Txt_from_H.Visible = False
    Txt_to_H.Visible = False
    d1.Visible = False
    d2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Command1.Visible = False
   'Grid2.ColHidden(Grid2.ColIndex("PeriodStill")) = True
   GRID2.ColHidden(GRID2.ColIndex("Priod")) = True
   'Grid2.ColHidden(Grid2.ColIndex("ArriveDate")) = True
   'Grid2.ColHidden(Grid2.ColIndex("Office")) = True
   'Grid2.ColHidden(Grid2.ColIndex("StarDate")) = True
   'Grid2.ColHidden(Grid2.ColIndex("Ex")) = True
   
StrSQL = "  SELECT DISTINCT"
StrSQL = StrSQL & "         dbo.TbVisaDeti.VisaID,"
StrSQL = StrSQL & "         TbVisaDeti.remarks,"
StrSQL = StrSQL & "         TblOffice.Name   AS OfficeName,"
StrSQL = StrSQL & "         TblOffice.Namee  AS OfficeNamee,"
StrSQL = StrSQL & "         dbo.TbVisaDeti.JobID,"
StrSQL = StrSQL & "         dbo.TblEmpJobsTypes.JobTypeName,"
StrSQL = StrSQL & "         dbo.TblEmpJobsTypes.JobTypeNamee,"
StrSQL = StrSQL & "         dbo.TbVisaDeti.NotionalID,"
StrSQL = StrSQL & "         dbo.Nationality.name,"
StrSQL = StrSQL & "         dbo.Nationality.namee,"
StrSQL = StrSQL & "         dbo.TblCountriesGovernments.GovernmentName,"
StrSQL = StrSQL & "         dbo.TbVisaDeti.[count],"
StrSQL = StrSQL & "         dbo.TbVisaDeti.Place,"
StrSQL = StrSQL & "         TbVisa.orderNo,"
StrSQL = StrSQL & "         TbVisa.visano,"
StrSQL = StrSQL & "         TbVisa.Priod,"
StrSQL = StrSQL & "         TbVisa.StarDate,"
StrSQL = StrSQL & "         TbVisa.EndDate,TbVisa.ArriveDate,"

StrSQL = StrSQL & "         DateDiff(d," & SQLDate(mNewDate, True) & "    , TbVisa.ArriveDate) as Peri,"
'" & SQLDate(DateDiff('d',  Askcoun, mNewDate)) & "'"
StrSQL = StrSQL & "         TblEmployee.KafelName"
StrSQL = StrSQL & "  From dbo.TblEmpJobsTypes"
StrSQL = StrSQL & "         RIGHT OUTER JOIN dbo.TblCountriesGovernments"
StrSQL = StrSQL & "         RIGHT OUTER JOIN dbo.TbVisaDeti"
StrSQL = StrSQL & "              ON  dbo.TblCountriesGovernments.GovernmentID = dbo.TbVisaDeti.CityID"
StrSQL = StrSQL & "         LEFT OUTER JOIN dbo.Nationality"
StrSQL = StrSQL & "              ON  dbo.TbVisaDeti.NotionalID = dbo.Nationality.id"
StrSQL = StrSQL & "              ON  dbo.TblEmpJobsTypes.JobTypeID = dbo.TbVisaDeti.JobID"
StrSQL = StrSQL & "         LEFT OUTER JOIN dbo.TblEmployee"
StrSQL = StrSQL & "              ON  dbo.TbVisaDeti.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & "         LEFT OUTER JOIN dbo.TblOffice"
StrSQL = StrSQL & "              ON  dbo.TbVisaDeti.OfficeId = dbo.TblOffice.ID"
StrSQL = StrSQL & "         LEFT OUTER JOIN dbo.TbVisa"
StrSQL = StrSQL & "              ON  dbo.TbVisaDeti.VisaID = dbo.TbVisa.ID"


  
StrSQL = StrSQL & "  Where 1 = 1 And TbVisaDeti.[Type] = 1 " & mWhere

StrSQL = StrSQL & "  Group By"
StrSQL = StrSQL & "         dbo.TbVisaDeti.VisaID,"
StrSQL = StrSQL & "         TbVisaDeti.remarks,TbVisa.ArriveDate,"
StrSQL = StrSQL & "         TblOffice.Name,"
StrSQL = StrSQL & "         TblOffice.Namee,"
StrSQL = StrSQL & "         dbo.TbVisaDeti.JobID,"
StrSQL = StrSQL & "         dbo.TblEmpJobsTypes.JobTypeName,"
StrSQL = StrSQL & "         dbo.TblEmpJobsTypes.JobTypeNamee,"
StrSQL = StrSQL & "         dbo.TbVisaDeti.NotionalID,"
StrSQL = StrSQL & "         dbo.Nationality.name,"
StrSQL = StrSQL & "         dbo.Nationality.namee,"
StrSQL = StrSQL & "         dbo.TblCountriesGovernments.GovernmentName,"
StrSQL = StrSQL & "         dbo.TbVisaDeti.[count],"
StrSQL = StrSQL & "         dbo.TbVisaDeti.Place,"
StrSQL = StrSQL & "         TbVisa.orderNo,"
StrSQL = StrSQL & "         TbVisa.visano,"
StrSQL = StrSQL & "         TbVisa.Priod,"
StrSQL = StrSQL & "         TbVisa.StarDate,"
StrSQL = StrSQL & "         TbVisa.EndDate,"
StrSQL = StrSQL & "         TblEmployee.KafelName"
My_SQL = StrSQL
 Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.GRID2
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
              '  .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(RsDev("Emp_ID").value), "", RsDev("Emp_ID").value)
            
               ' .TextMatrix(i, .ColIndex("Emp_code")) = IIf(IsNull(RsDev("Fullcode").value), "", RsDev("Fullcode").value)
                 .TextMatrix(i, .ColIndex("Office")) = IIf(IsNull(RsDev("OfficeName").value), "", RsDev("OfficeName").value)
               
                If SystemOptions.UserInterface = ArabicInterface Then
                '.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
                .TextMatrix(i, .ColIndex("Job")) = IIf(IsNull(RsDev("JobTypeName").value), "", RsDev("JobTypeName").value)
                .TextMatrix(i, .ColIndex("Notional")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
                Else
               ' .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(RsDev("Emp_Namee").value), "", RsDev("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("Office")) = IIf(IsNull(RsDev("OfficeNamee").value), "", RsDev("OfficeNamee").value)
                .TextMatrix(i, .ColIndex("Job")) = IIf(IsNull(RsDev("JobTypeNamee").value), "", RsDev("JobTypeNamee").value)
                .TextMatrix(i, .ColIndex("Notional")) = IIf(IsNull(RsDev("namee").value), "", RsDev("namee").value)
                End If
              '  .TextMatrix(i, .ColIndex("HododNo")) = IIf(IsNull(RsDev("HododNo").value), "", RsDev("HododNo").value)
                .TextMatrix(i, .ColIndex("City")) = IIf(IsNull(RsDev("GovernmentName").value), "", RsDev("GovernmentName").value)
                .TextMatrix(i, .ColIndex("NotionalID")) = IIf(IsNull(RsDev("NotionalID").value), "", RsDev("NotionalID").value)
                .TextMatrix(i, .ColIndex("JobID")) = IIf(IsNull(RsDev("JobID").value), "", RsDev("JobID").value)
               ' .TextMatrix(i, .ColIndex("CityID")) = IIf(IsNull(RsDev("CityID").value), "", RsDev("CityID").value)
              '  .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), "", RsDev("Price").value)
                .TextMatrix(i, .ColIndex("remarks")) = IIf(IsNull(RsDev("remarks").value), "", RsDev("remarks").value)
            
                            .TextMatrix(i, .ColIndex("orderNo")) = IIf(IsNull(RsDev("orderNo").value), "", RsDev("orderNo").value)
                .TextMatrix(i, .ColIndex("visano")) = IIf(IsNull(RsDev("visano").value), "", RsDev("visano").value)
                .TextMatrix(i, .ColIndex("Priod")) = IIf(IsNull(RsDev("Priod").value), "", RsDev("Priod").value)
                .TextMatrix(i, .ColIndex("StarDate")) = IIf(IsNull(RsDev("StarDate").value), "", RsDev("StarDate").value)
                .TextMatrix(i, .ColIndex("EndDate")) = IIf(IsNull(RsDev("EndDate").value), "", RsDev("EndDate").value)
                .TextMatrix(i, .ColIndex("ArriveDate")) = IIf(IsNull(RsDev("ArriveDate").value), "", RsDev("ArriveDate").value)
                .TextMatrix(i, .ColIndex("PeriodStill")) = DateDiff("d", mNewDate, IIf(IsNull(RsDev("ArriveDate").value), Date, RsDev("ArriveDate").value))
                .TextMatrix(i, .ColIndex("Priod")) = DateDiff("d", mNewDate, IIf(IsNull(RsDev("ArriveDate").value), Date, RsDev("ArriveDate").value))
                .TextMatrix(i, .ColIndex("KafelName")) = IIf(IsNull(RsDev("KafelName").value), "", RsDev("KafelName").value)
           
            
            
                RsDev.MoveNext
            Next i
        End With

    End If

'StrSQL = StrSQL & " Where (dbo.TbVisaDeti.VisaID = " & val(Me.xptxtid.text) & ")"

  
Else

If mIsEndType = 5 Then
    Grid.ColHidden(Grid.ColIndex("name")) = True
    Grid.ColHidden(Grid.ColIndex("LocationName")) = True
    Grid.ColHidden(Grid.ColIndex("placeEkama")) = True
    Grid.ColHidden(Grid.ColIndex("DateExpoekama")) = True
    Grid.ColHidden(Grid.ColIndex("DateEndekama")) = True
    Grid.ColHidden(Grid.ColIndex("ToMDateNew")) = True
    Grid.ColHidden(Grid.ColIndex("DateExpoekamaH")) = True
    Grid.ColHidden(Grid.ColIndex("days")) = True
    
End If
    Grid.Visible = True
    GRID2.Visible = False
    
     My_SQL = "SELECT   Dayss = DateDiff(d, " & SQLDate(mNewDate, True) & ", DateEndekama),  * from dbo.emp_all_details Where  dbo.emp_all_details.NationlID <>1 and  (NOT (dbo.emp_all_details.NumEkama IS NULL)) "
    If mIsEndType = 1 Then
        My_SQL = My_SQL & " AND DateEndekama>'" & SQLDate(DateAdd(Askinterval, Askcount, mNewDate)) & "'"
    ElseIf mIsEndType = 2 Then
        My_SQL = My_SQL & " AND DateEndekama<='" & SQLDate(DateAdd(Askinterval, Askcount, mNewDate)) & "'"
    Else
        My_SQL = My_SQL & " AND DateEndekama<='" & SQLDate(DateAdd(Askinterval, Askcount, mNewDate)) & "'"
    End If
    StrSQL = StrSQL & mWhere
        My_SQL = My_SQL & " order by DateEndekama,fullcode"
    With Me.Grid
        .Cell(flexcpPicture, 0, .ColIndex("Emp_Name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon

        '    .Cell(flexcpPicture, 0, .ColIndex("DateEndPasp")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
        For i = 0 To .Cols - 1
            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
        Next
   
        .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
        .RowHeight(-1) = 300
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    FillGridWithData
End If
   ' FillGridWithData
    
    'My_SQL = "SELECT     * from dbo.TblEmployee Where (Month(DateEndekama) <= Month(GetDate())) And (year(DateEndekama) <= year(GetDate()))"

  
    'BtnFirst_Click
    ShowTip

ErrTrap:
End Sub

Function ChangeLang()
    Me.Caption = "Expire Residence"
    Label1(2).Caption = Me.Caption
   ' Frame1.Caption = "Query"
    Label3.Caption = "From"
    Label4.Caption = "To"
    Command1.Caption = "Search"
    Command2.Caption = "Print"
    btnCancel.Caption = "Exit"
Label5.Caption = "Double Click To Show Employee Data"
    With Me.Grid
        .TextMatrix(0, .ColIndex("emp_code")) = "Emp Code"
        .TextMatrix(0, .ColIndex("emp_name")) = "Emp Name"
        .TextMatrix(0, .ColIndex("NumEkama")) = "Num."
        .TextMatrix(0, .ColIndex("placeEkama")) = "Issue Place"
        .TextMatrix(0, .ColIndex("DateExpoekamaH")) = "Issue Date H"
        .TextMatrix(0, .ColIndex("DateEndekamah")) = "Expire Date H"
        .TextMatrix(0, .ColIndex("DateExpoekama")) = "Issue Date G"
        .TextMatrix(0, .ColIndex("DateEndekama")) = "Expire Date G"
        .TextMatrix(0, .ColIndex("days")) = "Remain Days"
        .TextMatrix(0, .ColIndex("name")) = "Remain Days"
        .TextMatrix(0, .ColIndex("LocationName")) = "Work Status"
        FillGridWithData
    End With

End Function

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

ErrTrap:
End Sub

Private Sub Form_Activate()
    Me.ZOrder 0
End Sub

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblEmpJobsTypes", "JobTypeID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("JobTypeID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Public Sub FiLLRec()
    On Error GoTo ErrTrap

    RsSavRec.Fields("JobTypeName").value = IIf(TxtVacName.Text <> "", Trim(TxtVacName.Text), Null)

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    FillGridWithData
    TxtModFlg = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLTXT()

    On Error GoTo ErrTrap
    Dim i As Integer
    Frm2.Enabled = False
    TxtVac_ID.Text = IIf(IsNull(RsSavRec.Fields("JobTypeID").value), "", RsSavRec.Fields("JobTypeID").value)
    TxtVacName.Text = IIf(IsNull(RsSavRec.Fields("JobTypeName").value), "", RsSavRec.Fields("JobTypeName").value)
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

        For i = 1 To .Rows - 1

            If Trim(TxtVac_ID.Text) = .TextMatrix(i, .ColIndex("JobTypeID")) Then
                TxtSerial.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Sub EditRec(StrTable As String, _
                   RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub Grid_DblClick()
FrmEmployee.show
 FrmEmployee.Retrive val(Grid.TextMatrix(Grid.Row, 1))
End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("JobTypeID")))
ErrTrap:
End Sub

Private Sub Txt_from_H_GotFocus()
    date_type = 1
    
End Sub

Private Sub Txt_from_H_LostFocus()
  d1.value = ToGregorianDate(Txt_from_H.value)
      
End Sub

Private Sub Txt_to_H_GotFocus()
    date_type = 1
 
End Sub

Private Sub Txt_to_H_LostFocus()
   d2.value = ToGregorianDate(Txt_to_H.value)
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "JobTypeID=" & RecId, , adSearchForward, 1

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

    If TxtModFlg.Text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        '    btnNext.Enabled = False
        '    btnPrevious.Enabled = False
        '    btnFirst.Enabled = False
        '    btnLast.Enabled = False
    
    ElseIf TxtModFlg.Text = "R" Then
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False

        If TxtVac_ID.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If

        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
    
        '    btnNext.Enabled = True
        '    btnPrevious.Enabled = True
        '    btnFirst.Enabled = True
        '    btnLast.Enabled = True
    
    ElseIf TxtModFlg.Text = "E" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        Grid.Enabled = False
        '    btnNext.Enabled = False
        '    btnPrevious.Enabled = False
        '    btnFirst.Enabled = False
        '    btnLast.Enabled = False
    
    End If

End Sub

Public Sub FillGridWithData()

    'On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset

    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1

                .TextMatrix(i, .ColIndex("Ser")) = i
                '
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Fullcode").value), "", rs.Fields("Fullcode").value)
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(rs.Fields("Emp_id").value), "", rs.Fields("Emp_id").value)
If SystemOptions.UserInterface = ArabicInterface Then
          .TextMatrix(i, .ColIndex("LocationName")) = IIf(IsNull(rs.Fields("LocationName").value), "", rs.Fields("LocationName").value)
 Else
           .TextMatrix(i, .ColIndex("LocationName")) = IIf(IsNull(rs.Fields("LocationNameE").value), "", rs.Fields("LocationNameE").value)
 End If
 
  .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)

                
                
If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
 Else
 .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Namee").value), "", rs.Fields("Emp_Namee").value)
 
 End If
                'If FrmEmployee.OptExpirEkama = True Then
     
                .TextMatrix(i, .ColIndex("NumEkama")) = IIf(IsNull(rs.Fields("NumEkama").value), "", rs.Fields("NumEkama").value)
                
                .TextMatrix(i, .ColIndex("InsuranceRenewA")) = IIf(IsNull(rs.Fields("InsuranceRenewA").value), "", rs.Fields("InsuranceRenewA").value)
                .TextMatrix(i, .ColIndex("ToMA")) = IIf(IsNull(rs.Fields("ToMA").value), "", rs.Fields("ToMA").value)
                .TextMatrix(i, .ColIndex("InsuranceRenewDate")) = IIf(IsNull(rs.Fields("InsuranceRenewDate").value), "", rs.Fields("InsuranceRenewDate").value)
                .TextMatrix(i, .ColIndex("ToMDateNew")) = IIf(IsNull(rs.Fields("ToMDateNew").value), "", rs.Fields("ToMDateNew").value)
                
            
                .TextMatrix(i, .ColIndex("placeEkama")) = IIf(IsNull(rs.Fields("placeEkama").value), "", rs.Fields("placeEkama").value)
            
                .TextMatrix(i, .ColIndex("DateExpoekamaH")) = IIf(IsNull(rs.Fields("DateExpoekamaH").value), "", rs.Fields("DateExpoekamaH").value)
                .TextMatrix(i, .ColIndex("DateEndekamah")) = IIf(IsNull(rs.Fields("DateEndekamah").value), "", rs.Fields("DateEndekamah").value)
            
                .TextMatrix(i, .ColIndex("DateExpoekama")) = IIf(IsNull(rs.Fields("DateExpoekama").value), "", rs.Fields("DateExpoekama").value)
                .TextMatrix(i, .ColIndex("DateEndekama")) = IIf(IsNull(rs.Fields("DateEndekama").value), "", rs.Fields("DateEndekama").value)
            
                .TextMatrix(i, .ColIndex("days")) = IIf(IsNull(rs.Fields("DateEndekama").value), "", DateDiff("d", Date, rs.Fields("DateEndekama").value))
                '           If .TextMatrix(I, .ColIndex("Days")) = 0 Then
                '            .Cell(flexcpBackColor, I, 9, I, 9) = vbRed
                '            End If
 
                rs.MoveNext
            Next

            rs.Close
        End If
 .AutoSize 0, .Cols - 1, False
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
        '    .AddControl btnCancel, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        '    .AddControl btnFirst, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        '    .AddControl btnPrevious, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        '    .AddControl btnNext, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        '    .AddControl btnLast, Msg, True
    End With

ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
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
    'If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
    '    If btnFirst.Enabled = False Then Exit Sub
    '    BtnFirst_Click
    'End If
    'Move Previous---------------------------------------------------------
    'If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
    '    If btnPrevious.Enabled = False Then Exit Sub
    '    BtnPrevious_Click
    'End If

    'Move Next---------------------------------------------------------
    'If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
    '    If btnNext.Enabled = False Then Exit Sub
    '    BtnNext_Click
    'End If

    'Move Last---------------------------------------------------------
    'If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
    '    If btnLast.Enabled = False Then Exit Sub
    '    BtnLast_Click
    'End If

    'End If

    Exit Sub
ErrTrap:
End Sub

Private Function CheckDelJobType(LngJobTypeID As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select * From TblEmployee Where JobTypeID=" & LngJobTypeID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        CheckDelJobType = False
    Else
        CheckDelJobType = True
    End If

    rs.Close
    Set rs = Nothing
End Function

