VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmGoTime 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„Ê«⁄Ìœ «·«‰’—«› ..."
   ClientHeight    =   5280
   ClientLeft      =   2685
   ClientTop       =   2475
   ClientWidth     =   6360
   HelpContextID   =   560
   Icon            =   "FrmGoTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   6360
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   15
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   15
      Width           =   6045
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Text            =   "modflag"
         Top             =   15
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox TxtPresentTime_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   -15
         Visible         =   0   'False
         Width           =   945
      End
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   3405
         Top             =   15
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGoTime.frx":038A
               Key             =   "Emp_Name"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGoTime.frx":0724
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGoTime.frx":0ABE
               Key             =   "Emp_Code"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGoTime.frx":1058
               Key             =   "Emp_Salary"
            EndProperty
         EndProperty
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   60
         TabIndex        =   14
         Top             =   60
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
         ButtonImage     =   "FrmGoTime.frx":13F2
         ColorButton     =   14871017
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   525
         TabIndex        =   13
         Top             =   60
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
         ButtonImage     =   "FrmGoTime.frx":178C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1125
         TabIndex        =   12
         Top             =   60
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
         ButtonImage     =   "FrmGoTime.frx":1B26
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1590
         TabIndex        =   11
         Top             =   60
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
         ButtonImage     =   "FrmGoTime.frx":1EC0
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ê«⁄Ìœ «·«‰’—«› ..."
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
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   90
         Width           =   2205
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3180
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   585
      Width           =   6060
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   " ÊﬁÌ  «·√‰’—«› ›Ï «·‘—ﬂ…"
         Height          =   1425
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   60
         Width           =   6000
         Begin VB.ComboBox CmbTimeType 
            ForeColor       =   &H000000FF&
            Height          =   315
            ItemData        =   "FrmGoTime.frx":225A
            Left            =   1680
            List            =   "FrmGoTime.frx":225C
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   810
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox TxtHour 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   4110
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   810
            Width           =   480
         End
         Begin VB.TextBox TxtMinute 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   3585
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   810
            Width           =   480
         End
         Begin VB.ComboBox CmbTime 
            Height          =   315
            ItemData        =   "FrmGoTime.frx":225E
            Left            =   2865
            List            =   "FrmGoTime.frx":2260
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   810
            Width           =   690
         End
         Begin MSComCtl2.DTPicker DTDate 
            Height          =   330
            Left            =   2865
            TabIndex        =   0
            Top             =   195
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   582
            _Version        =   393216
            Format          =   100073475
            CurrentDate     =   38887
         End
         Begin VB.Label LabWork 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Label5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   630
            Width           =   1680
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "œ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   3780
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   555
            Width           =   90
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "”"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   4275
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   555
            Width           =   225
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "› —…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   3015
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   570
            Width           =   390
         End
         Begin VB.Label LabDayName 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   330
            Left            =   518
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   195
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Õ÷Ê—"
            Height          =   195
            Index           =   0
            Left            =   4890
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   270
            Width           =   915
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Caption         =   " ÊﬁÌ  « ·«‰’—«›"
            Height          =   195
            Left            =   4710
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   885
            Width           =   1095
         End
      End
      Begin VB.Frame FrmBrngTime 
         BackColor       =   &H00E2E9E9&
         Height          =   1470
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1440
         Width           =   6000
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1155
            Index           =   3
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   150
            Width           =   2370
            Begin MSComCtl2.DTPicker DtpDeparture 
               Height          =   330
               Left            =   60
               TabIndex        =   53
               Top             =   300
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               Format          =   100073475
               CurrentDate     =   38887
            End
            Begin MSComCtl2.DTPicker DtpDepHour 
               Height          =   405
               Left            =   60
               TabIndex        =   54
               Top             =   660
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   714
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   100073475
               UpDown          =   -1  'True
               CurrentDate     =   39240
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”«⁄…"
               Height          =   195
               Index           =   6
               Left            =   1740
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   780
               Width           =   465
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÌÊ„"
               Height          =   225
               Index           =   2
               Left            =   1890
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   360
               Width           =   315
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ì⁄«œ «·«‰’—«›"
            Height          =   1095
            Left            =   105
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   165
            Width           =   2310
            Begin VB.ComboBox CmbBringTime 
               Height          =   315
               ItemData        =   "FrmGoTime.frx":2262
               Left            =   0
               List            =   "FrmGoTime.frx":2264
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Tag             =   "⁄›Ê« Ì—Ã∆ ≈œŒ«· › —… «·Õ÷Ê—"
               Top             =   4770
               Width           =   2340
            End
            Begin VB.ComboBox CmbHour 
               Height          =   315
               ItemData        =   "FrmGoTime.frx":2266
               Left            =   1530
               List            =   "FrmGoTime.frx":2268
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Tag             =   "⁄›Ê« Ì—Ã∆ ≈œŒ«· ”«⁄… «·Õ÷Ê—"
               Top             =   615
               Width           =   585
            End
            Begin VB.ComboBox CmbMinute 
               Height          =   315
               ItemData        =   "FrmGoTime.frx":226A
               Left            =   840
               List            =   "FrmGoTime.frx":226C
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Tag             =   "⁄›Ê« Ì—Ã∆ ≈œŒ«· œﬁÌﬁ… «·Õ÷Ê—"
               Top             =   615
               Width           =   675
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁ"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   3
               Left            =   1185
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   345
               Width           =   135
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "”"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   4
               Left            =   1815
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   345
               Width           =   225
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "› —…"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   5
               Left            =   435
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   345
               Width           =   390
            End
         End
         Begin VB.TextBox TxtCode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3675
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   210
            Width           =   1230
         End
         Begin VB.TextBox TxtEmp_Code 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3675
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Tag             =   "⁄›Ê« Ì—Ã∆ ≈œŒ«· ﬂÊœ «·„ÊŸ›"
            Top             =   570
            Width           =   1230
         End
         Begin MSDataListLib.DataCombo DCEmp_Name 
            Height          =   315
            Left            =   2520
            TabIndex        =   7
            Top             =   930
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   "DCEmp_Name"
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   195
            Index           =   1
            Left            =   5385
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   195
            Index           =   0
            Left            =   5070
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   990
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„ÊŸ›"
            Height          =   195
            Index           =   3
            Left            =   5055
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   630
            Width           =   810
         End
      End
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   240
      TabIndex        =   36
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   3810
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
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1050
      Left            =   120
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4170
      Width           =   6210
      _cx             =   10954
      _cy             =   1852
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
         Left            =   5235
         TabIndex        =   39
         Top             =   555
         Width           =   600
         _ExtentX        =   1058
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
         ButtonImage     =   "FrmGoTime.frx":226E
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   330
         Left            =   3216
         TabIndex        =   40
         Top             =   555
         Width           =   660
         _ExtentX        =   1164
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
         ButtonImage     =   "FrmGoTime.frx":2608
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   330
         Left            =   4215
         TabIndex        =   41
         Top             =   555
         Width           =   705
         _ExtentX        =   1244
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
         ButtonImage     =   "FrmGoTime.frx":29A2
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   330
         Left            =   2169
         TabIndex        =   42
         Top             =   555
         Width           =   720
         _ExtentX        =   1270
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
         ButtonImage     =   "FrmGoTime.frx":2D3C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   330
         Left            =   1242
         TabIndex        =   43
         Top             =   555
         Width           =   600
         _ExtentX        =   1058
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
         ButtonImage     =   "FrmGoTime.frx":30D6
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   5010
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   1110
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "»ÕÀ"
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
         ButtonImage     =   "FrmGoTime.frx":3670
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   3960
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   1110
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
         ButtonImage     =   "FrmGoTime.frx":3A0A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnPrint 
         Height          =   285
         Left            =   2940
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1140
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
         ButtonImage     =   "FrmGoTime.frx":3DA4
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   330
         Left            =   165
         TabIndex        =   47
         Top             =   555
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
         ButtonImage     =   "FrmGoTime.frx":413E
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   165
         Width           =   540
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   165
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   210
         Index           =   1
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   165
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   210
         Index           =   2
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   165
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„” Œœ„"
      Height          =   285
      Index           =   13
      Left            =   2730
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   3840
      Width           =   765
   End
End
Attribute VB_Name = "FrmGoTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim cSearch As clsDCboSearch
Dim RecID As String
Dim II As Long

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    On Error GoTo ErrTrap

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String

    If TxtPresentTime_ID.text <> "" Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbYesNo + vbMsgBoxRight, App.Title)
        
        If MSGType = vbYes Then
            RsSavRec.find "Present_ID=" & val(TxtPresentTime_ID.text), , adSearchForward, 1
            RsSavRec.delete
                     
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.Title
            '------------------------------ Move Next ---------------------------.
            
            BtnNext_Click
        End If
    
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnModify_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If TxtPresentTime_ID.text <> "" Then
        '        If Not (RsSavRec.BOF Or RsSavRec.EOF) Then
        '            RsSavRec.MoveNext
        '            RsSavRec.MoveLast
        '        End If
        
        TxtModFlg = "E"
        Frm2.Enabled = True
        Me.DCEmp_Name.SetFocus
    End If

    Me.DCUser.BoundText = user_id
    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & Chr(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Private Sub btnNew_Click()
    On Error GoTo ErrTrap
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    TxtModFlg.text = "N"
    
    My_SQL = "select * From tblPresentTime where Present_Type=1"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        TXTCode.text = rs.RecordCount + 1
    Else
        TXTCode.text = 1
    End If

    rs.Close
    DTDate_Click

    If TxtEmp_Code.Enabled = True Then
        TxtEmp_Code.SetFocus
    End If

    Me.DCUser.BoundText = user_id
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    Dim Msg As String
    On Error GoTo ErrTrap
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    Dim Msg As String
    On Error GoTo ErrTrap

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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrTrap

    'Dim StrVacCode As String
    'Dim StrVacName As String
    Dim Msg As String
    Dim CtrlTxt As Control

    '---------------------- check if data Vaclete -----------------------
    If CmbTimeType.ListIndex = 0 Then

        For Each CtrlTxt In Me.Controls

            If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
                If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                    MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.Title
                    CtrlTxt.SetFocus
                    Exit Sub
                End If
            End If

        Next

        If DCEmp_Name.BoundText = "" Then
            Msg = "⁄›Ê« Ì—ÃÏ  ÕœÌœ «”„ «·„ÊŸ›"
            MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCEmp_Name.SetFocus
            Exit Sub
        End If
    
        '    Msg = "⁄›Ê« Â–« «·„ÊŸ› ·„ ÌÕ÷— «·ÌÊ„"
        '    If ChkEmpComeToday = False Then
        '        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        Exit Sub
        '    End If
        '
        '    Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· «·«‰’—«› ·Â–« «·„ÊŸ› „‰ ﬁ»·"
        '    If ChkEmpExist = True Then
        '        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        Exit Sub
        '    End If
        '
        '    Msg = "⁄›Ê« ·„ Ì „  ”ÃÌ· Õ÷Ê— ·Â–« «·„ÊŸ› «·ÌÊ„"
        '    If ChkEmpBringTime = False Then
        '        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        Exit Sub
        '    End If
    
        ' -------------------------------------- txtmodflg type -------------------
        Select Case Me.TxtModFlg.text

                '------------------------------ new record ----------------------------
            Case "N"
          
                '------------------------- save record -----------------------------
                'AddNewRec
                'BtnLast_Click
            Case "E"
    
                '----------------------------- save edit -------------------------------
                FiLLRec
        End Select

    Else
        Msg = "⁄›Ê« Â–« ÌÊ„ ⁄ÿ·…"
        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If

    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title

End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtPresentTime_ID.text)
    Me.TxtModFlg.text = "R"
End Sub

Private Sub BtnUpdate_Click()
    On Error GoTo ErrTrap

    Dim FristCount As Long
    Dim LastCount As Long
    Dim Msg As String
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

    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.Title
ErrTrap:
End Sub

Private Sub DCEmp_Name_Click(Area As Integer)
    TxtEmp_Code.text = GetEmpCode("Emp_Code", " Emp_ID=" & val(DCEmp_Name.BoundText))
End Sub

Private Sub DTDate_Change()
    DTDate_Click
End Sub

Private Sub DTDate_Click()
    LabDayName.Caption = Format(DtDate.value, "dddd")
    GetTimeDetails

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

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Resize_Form Me
    My_SQL = "select * From tblPresentTime where Present_Type=1"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
    Me.TxtModFlg.text = "R"

    'load tblEmployee -----------------------------------------------
    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DCEmp_Name, True
    Dcombos.GetUsers Me.DCUser
    DCUser.BoundText = user_id
    Set cSearch = New clsDCboSearch
    Set cSearch.Client = DCEmp_Name

    'DTTime.Value = Time
    DtDate.value = Date
    CmbTime.AddItem "’"
    CmbTime.ItemData(CmbTime.NewIndex) = 0
    CmbTime.AddItem "„"
    CmbTime.ItemData(CmbTime.NewIndex) = 1

    CmbBringTime.AddItem "’"
    CmbBringTime.ItemData(CmbBringTime.NewIndex) = 0
    CmbBringTime.AddItem "„"
    CmbBringTime.ItemData(CmbBringTime.NewIndex) = 1
    '----------------------------------------------------------------------------
    CmbTimeType.AddItem "⁄„·"
    CmbTimeType.ItemData(CmbTimeType.NewIndex) = 0
    CmbTimeType.AddItem "⁄ÿ·…"
    CmbTimeType.ItemData(CmbTimeType.NewIndex) = 1

    For i = 1 To 12
        CmbHour.AddItem i
    Next

    For i = 0 To 59
        CmbMinute.AddItem i
    Next

    SetDtpickerDate Me.DtDate
    DtDate.value = Date
    DTDate_Click
    BtnFirst_Click
    ShowTip
ErrTrap:
End Sub

Private Sub Form_Terminate()
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

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("tblPresentTime", "Present_ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("Present_ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Public Sub FiLLRec()
    On Error GoTo ErrTrap

    RsSavRec.update

    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.Title

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
    TxtPresentTime_ID.text = IIf(IsNull(RsSavRec.Fields("Present_ID").value), "", RsSavRec.Fields("Present_ID").value)
    DCEmp_Name.BoundText = IIf(IsNull(RsSavRec.Fields("Emp_ID").value), "", RsSavRec.Fields("Emp_ID").value)
    CmbBringTime.ListIndex = IIf(IsNull(RsSavRec.Fields("Present_Time").value), -1, RsSavRec.Fields("Present_Time").value)
    DtDate.value = IIf(IsNull(RsSavRec.Fields("PresentDate").value), "", RsSavRec.Fields("PresentDate").value)

    TXTCode.text = IIf(IsNull(RsSavRec.Fields("Present_Code").value), "", RsSavRec.Fields("Present_Code").value)
    CmbHour.ListIndex = IIf(IsNull(RsSavRec.Fields("Present_Hour").value), -1, RsSavRec.Fields("Present_Hour").value - 1)
    CmbMinute.ListIndex = IIf(IsNull(RsSavRec.Fields("Present_Minute").value), -1, RsSavRec.Fields("Present_Minute").value)
    TxtEmp_Code.text = GetEmpCode("Emp_Code", " Emp_ID=" & val(DCEmp_Name.BoundText))
    'DCUser.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").Value), "", RsSavRec.Fields("UserID").Value)

    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    DTDate_Click
ErrTrap:

End Sub

Public Sub EditRec(StrTable As String, _
                   RecID As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub TxtEmp_Code_KeyUp(KeyCode As Integer, _
                              Shift As Integer)
    DCEmp_Name.BoundText = GetEmpCode("Emp_ID", " Emp_Code='" & TxtEmp_Code.text & "'")

End Sub

Private Sub TxtPresentTime_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecID As Long)
    On Error GoTo ErrTrap

    RsSavRec.find "Present_ID=" & RecID, , adSearchForward, 1

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

Private Sub TxtModFlg_Change()

    If TxtModFlg.text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
    
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    
    ElseIf TxtModFlg.text = "R" Then
        Frm2.Enabled = False
  
        btnModify.Enabled = False
        btnDelete.Enabled = False

        If TxtPresentTime_ID.text <> "" Then
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
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    
    End If

End Sub

Private Sub GetTimeDetails()
    On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset

    'My_SQL = "select * From tblTimeSetting where day ='" & Trim(LabDayName.Caption) & "' "
    My_SQL = "select * From tblTimeSetting where Is_WorkDay =0"

    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    'If Not Rs.EOF Then
    'Is_WorkDay
    If rs.RecordCount > 0 Then
    
        CmbTimeType.ListIndex = IIf(IsNull(rs.Fields("Is_WorkDay").value), -1, rs.Fields("Is_WorkDay").value)
       
        TxtHour.text = IIf(IsNull(rs.Fields("Go_HourTime").value), "", rs.Fields("Go_HourTime").value)
                                    
        TxtMinute.text = IIf(IsNull(rs.Fields("Go_MinuteTime").value), "", rs.Fields("Go_MinuteTime").value)
    
        CmbTime.ListIndex = IIf(IsNull(rs.Fields("Go_Time").value), -1, rs.Fields("Go_Time").value)
    
    End If

    If CmbTimeType.ListIndex = 0 Then
        FrmBrngTime.Enabled = True
    Else
        FrmBrngTime.Enabled = False
        TxtEmp_Code.text = ""
        DCEmp_Name.text = ""
        '    TxtBringHour.Text = ""
        '    TxtBringMinute.Text = ""
        CmbBringTime.ListIndex = -1
    End If

    LabWork.Caption = CmbTimeType.text
    rs.Close
    Set rs = Nothing
ErrTrap:
End Sub

Private Function GetEmpCode(ByVal Fild As String, _
                            ByVal whr As String) As String
    On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    GetEmpCode = ""
    My_SQL = "select " & Fild & " From TblEmployee where " & whr
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        GetEmpCode = IIf(IsNull(rs.Fields(Fild).value), "", rs.Fields(Fild).value)
    End If

    rs.Close
    Set rs = Nothing
ErrTrap:
End Function

Private Function ChkEmpExist() As Boolean
    On Error GoTo ErrTrap
    Dim My_SQL As String

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    ChkEmpExist = False

    My_SQL = "select * From tblPresentTime where  Present_Type=1 and Emp_ID=" & DCEmp_Name.BoundText & " and  PresentDate=" & SQLDate(DtDate.value, True) & " and Present_ID <>'" & Trim(TxtPresentTime_ID.text) & "'"
            
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        ChkEmpExist = True
    End If

    rs.Close
    Set rs = Nothing
ErrTrap:
End Function

Private Function ChkEmpBringTime() As Boolean
    On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    ChkEmpBringTime = False
    My_SQL = "select * From tblPresentTime where Present_Type=0 and  Emp_ID=" & DCEmp_Name.BoundText & " and  PresentDate=" & SQLDate(DtDate.value, True) & ""
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        ChkEmpBringTime = True
    End If

    rs.Close
    Set rs = Nothing
ErrTrap:
End Function

Private Function ChkEmpComeToday() As Boolean
    On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    ChkEmpComeToday = True
    My_SQL = "select * From QryAbsentEmp where AbsDate='" & DtDate.value & "' and Emp_ID=" & DCEmp_Name.BoundText
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        ChkEmpComeToday = False
    End If

    rs.Close
    Set rs = Nothing
ErrTrap:
End Function

'-------------------------------------------------------------
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

