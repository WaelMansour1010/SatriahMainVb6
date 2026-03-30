VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form shipmentA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„ «»⁄Â «·‘Õ‰"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11940
   Icon            =   "shipmentA.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   11940
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  «·‘Õ‰Â"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   5640
      Width           =   11775
      Begin VSFlex8UCtl.VSFlexGrid Grid 
         Height          =   1515
         Left            =   1080
         TabIndex        =   50
         Top             =   360
         Width           =   9765
         _cx             =   17224
         _cy             =   2672
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
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"shipmentA.frx":000C
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„ «»⁄Â «·‘Õ‰Â"
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
      Height          =   4815
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   11775
      Begin VB.CheckBox ChkRefounded_gurantee 
         Caption         =   "Check1"
         Height          =   255
         Left            =   4680
         TabIndex        =   52
         Top             =   4320
         Width           =   255
      End
      Begin VB.TextBox txt_Shipment_no 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TXT_AWB_OR_BL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   2520
         Width           =   3015
      End
      Begin VB.TextBox txt_forward_AGENT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox TXT_CLR_AGENT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   2880
         Width           =   3015
      End
      Begin VB.TextBox TXT_INS_REQ_NO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox TXT_POLICY_NO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox TXT_GURANTY_NO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   3960
         Width           =   3015
      End
      Begin VB.TextBox txt_portal_of_sale 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txt_insurance_company 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtVessel_or_Flight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   2880
         Width           =   2775
      End
      Begin MSDataListLib.DataCombo dc_shipment_mode 
         Height          =   315
         Left            =   7080
         TabIndex        =   23
         Top             =   2520
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txt_portal_of_dest 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox TXTid 
         Height          =   285
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTP_Exp_date_ship 
         Height          =   330
         Left            =   8280
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   720
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   94175235
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_Exp_date_arriv 
         Height          =   330
         Left            =   8280
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   94175235
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_Act_date_clr 
         Height          =   330
         Left            =   8280
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   94175235
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_doc_ref_date 
         Height          =   330
         Left            =   8280
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   94175235
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_g_delivery_date 
         Height          =   330
         Left            =   8040
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   94175235
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_act_date_shipm 
         Height          =   330
         Left            =   3360
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   720
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   94175235
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_act_date_arrival 
         Height          =   330
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   94175235
         CurrentDate     =   37140
      End
      Begin MSDataListLib.DataCombo dc_guranty_type 
         Height          =   315
         Left            =   7080
         TabIndex        =   26
         Top             =   3600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DTP_Exp_date_site 
         Height          =   330
         Left            =   120
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   94175235
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_date_clr 
         Height          =   330
         Left            =   3360
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   94175235
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTPi_act_date_site 
         Height          =   330
         Left            =   3360
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   94175235
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_G_Refound_date 
         Height          =   330
         Left            =   2880
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   4320
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   94175235
         CurrentDate     =   37140
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  «⁄«œÂ  √”Ì”"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   41
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «·÷„«‰"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   40
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·»Ê·Ì’Â"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   39
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·⁄„Ì·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   38
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„Ì·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   37
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "AWB or B/L NO."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   36
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊﬂÌ·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   35
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„Ì‰«¡ «·‘Õ‰"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   33
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label DTP_1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  «·Ê’Ê·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   30
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  «·«” ·«„"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   29
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  «·„€«œ—Â"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   27
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  Ê’Ê· ›⁄·Ï"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  ‘Õ‰ ›⁄·Ï"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  Ê’Ê· «·÷„«‰"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9480
         TabIndex        =   12
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·÷„«‰"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   11
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘—ﬂÂ «· √„Ì‰"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   10
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”›Ì‰Â/ÿÌ—«‰"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   9
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—Ìﬁ «·‘Õ‰"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   8
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„Ì‰«¡ «·Ê’Ê·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   7
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  «· ”ÃÌ·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  «· Ê“Ì⁄"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  Ê’Ê· „ Êﬁ⁄"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  ‘Õ‰ „ Êﬁ⁄"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·‘Õ‰Â"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   9960
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   765
      Index           =   5
      Left            =   0
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   0
      Width           =   12150
      _cx             =   21431
      _cy             =   1349
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
      Picture         =   "shipmentA.frx":009D
      Caption         =   ""
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
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1695
         TabIndex        =   54
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "shipmentA.frx":0D77
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
         TabIndex        =   55
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "shipmentA.frx":1111
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
         TabIndex        =   56
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "shipmentA.frx":14AB
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
         TabIndex        =   57
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "shipmentA.frx":1845
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„ «»⁄… «·‘Õ‰"
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
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   120
         Width           =   3750
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic7 
      Height          =   8730
      Left            =   0
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   0
      Width           =   11925
      _cx             =   21034
      _cy             =   15399
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   780
         Left            =   120
         TabIndex        =   61
         Top             =   7800
         Width           =   11745
         Begin ImpulseButton.ISButton Cmd 
            Height          =   285
            Index           =   0
            Left            =   10665
            TabIndex        =   62
            Top             =   330
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
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
            ColorToggledText=   -2147483631
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   285
            Index           =   1
            Left            =   9360
            TabIndex        =   63
            Top             =   330
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
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
            Height          =   285
            Index           =   2
            Left            =   8550
            TabIndex        =   64
            Top             =   315
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
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
            Height          =   285
            Index           =   3
            Left            =   7125
            TabIndex        =   65
            Top             =   330
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
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
            Height          =   285
            Index           =   4
            Left            =   5835
            TabIndex        =   66
            Top             =   360
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   503
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
            Height          =   285
            Index           =   5
            Left            =   3075
            TabIndex        =   67
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
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
            Height          =   345
            Index           =   6
            Left            =   120
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   240
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
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
            Height          =   225
            Index           =   7
            Left            =   4365
            TabIndex        =   69
            Top             =   360
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   397
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â"
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
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   345
            Left            =   1575
            TabIndex        =   70
            Top             =   240
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "„”«⁄œÂ"
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
      End
   End
End
Attribute VB_Name = "shipmentA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim RsDev As ADODB.Recordset
Dim month_salary As Double
Dim day_salary As Double
 
Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0
     
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            Me.TXTid.Text = CStr(new_id("Shipments", "id", "", True))
 
            '   Txt_DateEndLincH.value = ToHijriDate(Date)
        Case 1
 
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            Me.Grid.Rows = Me.Grid.Rows + 1

        Case 2
    
            '  calc_total
            ' cal_interval
            SaveData
        
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_ProfData

        Case 5
 
            '  FrmEmployeeSearch.Show ' vbModal
        Case 6
            Unload Me

        Case 7
 
            printingReport (val(TXTid.Text))
    End Select

    Exit Sub
ErrTrap:
 
End Sub
 
Public Function printingReport(Optional ID As String, _
                               Optional date_from)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From Shipment_all_details where id=" & ID

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "Shipment Status Report_eng.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "Shipment Status Report_eng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "No data to show"
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = "Shipment Status Report" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = "Shipment Status Report"
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Sub Retrive(Optional Lngid As Long = 0)
    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        ' XPTxtCurrent.Caption = 0
        ' XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        ' If Lngid <> 0 Then
        '     Rs.find "Emp_ID=" & Lngid, , adSearchForward, adBookmarkFirst
        '     If Rs.EOF Or Rs.BOF Then
        '         Exit Sub
        '     End If
        ' End If
    End If

    Me.TXTid.Text = IIf(IsNull(rs("id").value), 0, val(rs("id").value))
    Me.txt_Shipment_no.Text = IIf(IsNull(rs("Shipment_no").value), "", rs("Shipment_no").value)

    Me.DTP_Exp_date_ship.value = IIf(Not IsDate(rs("Exp_date_ship").value), Date, rs("Exp_date_ship").value)
    Me.DTP_Exp_date_arriv.value = IIf(Not IsDate(rs("Exp_date_arriv").value), Date, rs("Exp_date_arriv").value)
    Me.DTP_Act_date_clr.value = IIf(Not IsDate(rs("Act_date_clr").value), Date, rs("Act_date_clr").value)
    Me.DTP_doc_ref_date.value = IIf(Not IsDate(rs("doc_ref_date").value), Date, rs("doc_ref_date").value)
    Me.DTP_g_delivery_date.value = IIf(Not IsDate(rs("g_delivery_date").value), Date, rs("g_delivery_date").value)
    Me.DTP_act_date_shipm.value = IIf(Not IsDate(rs("act_date_shipm").value), Date, rs("act_date_shipm").value)
    Me.DTP_act_date_arrival.value = IIf(Not IsDate(rs("act_date_arrival").value), Date, rs("act_date_arrival").value)
    Me.DTP_Exp_date_site.value = IIf(Not IsDate(rs("Exp_date_site").value), Date, rs("Exp_date_site").value)
    Me.DTP_date_clr.value = IIf(Not IsDate(rs("date_clr").value), Date, rs("date_clr").value)
    Me.DTPi_act_date_site.value = IIf(Not IsDate(rs("act_date_site").value), Date, rs("act_date_site").value)
    Me.DTP_G_Refound_date.value = IIf(Not IsDate(rs("G_Refound_date").value), Date, rs("G_Refound_date").value)
 
    Me.txt_portal_of_dest.Text = IIf(IsNull(rs("portal_of_dest").value), "", Trim(rs("portal_of_dest").value))
    Me.txtVessel_or_Flight.Text = IIf(IsNull(rs("Vessel_or_Flight").value), "", Trim(rs("Vessel_or_Flight").value))
    Me.txt_insurance_company.Text = IIf(IsNull(rs("insurance_company").value), "", Trim(rs("insurance_company").value))
    Me.txt_portal_of_sale.Text = IIf(IsNull(rs("portal_of_sale").value), "", Trim(rs("portal_of_sale").value))
    Me.txt_forward_AGENT.Text = IIf(IsNull(rs("forward_AGENT").value), "", Trim(rs("forward_AGENT").value))
    Me.TXT_AWB_OR_BL.Text = IIf(IsNull(rs("AWB_OR_BL").value), "", Trim(rs("AWB_OR_BL").value))
    Me.TXT_CLR_AGENT.Text = IIf(IsNull(rs("CLR_AGENT").value), "", Trim(rs("CLR_AGENT").value))
    Me.TXT_INS_REQ_NO.Text = IIf(IsNull(rs("INS_REQ_NO").value), "", Trim(rs("INS_REQ_NO").value))
    Me.TXT_POLICY_NO.Text = IIf(IsNull(rs("POLICY_NO").value), "", Trim(rs("POLICY_NO").value))
    Me.TXT_GURANTY_NO.Text = IIf(IsNull(rs("GURANTY_NO").value), "", Trim(rs("GURANTY_NO").value))

    dc_shipment_mode.BoundText = IIf(rs("shipment_mode").value = 0, "", Trim(rs("shipment_mode").value))
    dc_guranty_type.BoundText = IIf(rs("guranty_type").value = 0, "", Trim(rs("guranty_type").value))

    If rs("Refounded_gurantee").value = True Then
        ChkRefounded_gurantee.value = vbChecked
    Else
        ChkRefounded_gurantee.value = Unchecked
    End If
 
    FillGridWithData

    Exit Sub
ErrTrap:
End Sub

Function FillGridWithData()
    StrSQL = "Select * from  Shipments_details  where  Shipments_id=" & val(Me.TXTid.Text)
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
       
        If .Rows > 1 Then
            .Clear 1, 1
            .FixedRows = 1
            .Rows = .FixedRows + 1
        End If
 
    End With
    
    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
   
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1

                .TextMatrix(i, .ColIndex("cno")) = IIf(IsNull(RsDev("Container_no").value), "", RsDev("Container_no").value)
            
                .TextMatrix(i, .ColIndex("ono")) = IIf(IsNull(RsDev("order_no").value), "", RsDev("order_no").value)
            
                .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("des").value), "", RsDev("des").value)
 
                RsDev.MoveNext
            Next i
 
        End With

    End If

    RsDev.Close
End Function

Private Sub SaveData()
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then

        '    If Me.DcboEmp.text = "" Then
        '        Msg = "ÌÃ» «Œ Ì«— «”„ «·„ÊŸ› "
        '        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        DcboEmp.SetFocus
        '       SendKeys "{F4}"
        '       Screen.MousePointer = vbDefault
        '        Exit Sub
        '  End If

        ' If Me.dcby.BoundText = "" Then
        '        Msg = "ÌÃ» «Œ Ì«— «”„ «·ﬁ«∆„ »«·⁄„·Ì… "
        '        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        dcby.SetFocus
        '       SendKeys "{F4}"
        '       Screen.MousePointer = vbDefault
        '        Exit Sub
        'End If
 
        '  If Me.dctype.BoundText = "" Then
        '        Msg = "ÌÃ» «Œ Ì«— ‰Ê⁄ ⁄„·Ì… «·«‰Â«¡ "
        '        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        dctype.SetFocus
        ''       SendKeys "{F4}"
        '      Screen.MousePointer = vbDefault
        '       Exit Sub
        'End If
 
        If Me.txt_Shipment_no.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·«»œ „‰ ﬂ‰«»Â —ﬁ„ «·‘Ã‰Â"
            Else
                Msg = "shipment No. is a must"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.txt_Shipment_no.SetFocus
            SelectText Me.txt_Shipment_no
            Exit Sub

        End If
 
        Cn.BeginTrans
        BeginTrans = True
    
        If TxtModFlg.Text = "N" Then
            '  Dim RsTemp As New ADODB.Recordset
            '            StrSQL = "select * From End_of_service where emp_code=" & Val(Me.txtEmpCode.text)
            '            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            '            If RsTemp.RecordCount > 0 Then
            '                Msg = " „ Õ”«» «·„ﬂ«›√… ·Â–« «·„ÊŸ› „‰ ﬁ»·" & Chr(13)
            '                Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & Chr(13)
            ''                Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
            '               MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '               Exit Sub
            '           End If
            '
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
    
            StrSQL = "Delete   Shipments_details Where Shipments_id=" & val(Me.TXTid.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
        End If

        rs("id").value = val(Me.TXTid.Text)
        rs("Shipment_no").value = IIf(txt_Shipment_no.Text = "", Null, Trim(txt_Shipment_no.Text))
        rs("Exp_date_ship").value = Me.DTP_Exp_date_ship.value
        rs("Exp_date_arriv").value = Me.DTP_Exp_date_arriv.value
        rs("Act_date_clr").value = Me.DTP_Act_date_clr.value
        rs("doc_ref_date").value = Me.DTP_doc_ref_date.value
        rs("g_delivery_date").value = Me.DTP_g_delivery_date.value
        rs("act_date_shipm").value = Me.DTP_act_date_shipm.value
        rs("act_date_arrival").value = Me.DTP_act_date_arrival.value
        rs("Exp_date_site").value = Me.DTP_Exp_date_site.value
        rs("date_clr").value = Me.DTP_date_clr.value
        rs("act_date_site").value = Me.DTPi_act_date_site.value
        rs("G_Refound_date").value = Me.DTP_G_Refound_date.value

        If ChkRefounded_gurantee.value = vbChecked Then
            rs("Refounded_gurantee").value = 1
        Else
            rs("Refounded_gurantee").value = 0
        End If
 
        rs("portal_of_dest").value = IIf(Me.txt_portal_of_dest.Text = "", Null, Me.txt_portal_of_dest.Text)

        rs("Vessel_or_Flight").value = IIf(Me.txtVessel_or_Flight.Text = "", Null, Me.txtVessel_or_Flight.Text)
        rs("insurance_company").value = IIf(Me.txt_insurance_company.Text = "", Null, Me.txt_insurance_company.Text)
        rs("portal_of_sale").value = IIf(Me.txt_portal_of_sale.Text = "", Null, Me.txt_portal_of_sale.Text)
        rs("forward_AGENT").value = IIf(Me.txt_forward_AGENT.Text = "", Null, Me.txt_forward_AGENT.Text)
        rs("AWB_OR_BL").value = IIf(Me.TXT_AWB_OR_BL.Text = "", Null, Me.TXT_AWB_OR_BL.Text)
        rs("CLR_AGENT").value = IIf(Me.TXT_CLR_AGENT.Text = "", Null, Me.TXT_CLR_AGENT.Text)
        rs("INS_REQ_NO").value = IIf(Me.TXT_INS_REQ_NO.Text = "", Null, Me.TXT_INS_REQ_NO.Text)
        rs("POLICY_NO").value = IIf(Me.TXT_POLICY_NO.Text = "", Null, Me.TXT_POLICY_NO.Text)
        rs("GURANTY_NO").value = IIf(Me.TXT_GURANTY_NO.Text = "", Null, Me.TXT_GURANTY_NO.Text)

        rs("shipment_mode").value = IIf(Me.dc_shipment_mode.BoundText = "", 0, Me.dc_shipment_mode.BoundText)
        rs("guranty_type").value = IIf(Me.dc_guranty_type.BoundText = "", 0, Me.dc_guranty_type.BoundText)
      
        rs.update
    
        Cn.CommitTrans
        BeginTrans = False
        '    XPTxtCurrent.Caption = Rs.AbsolutePosition
        '    XPTxtCount.Caption = Rs.RecordCount
    
        Select Case Me.TxtModFlg.Text

            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ  ›«’Ì· «·‘Õ‰  " & CHR(13)
                    Msg = Msg + "Â·  —Ìœ «÷«›…  ›«’Ì· ‘Õ‰Â «Œ—Ï"
                Else
                    Msg = "Shipment data was saved ... do you want to add another shipment"
                End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Edits was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
        End Select

        TxtModFlg.Text = "R"
    End If

    Set RsDev = New ADODB.Recordset
    RsDev.Open "[Shipments_details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    With Me.Grid

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("cno")) <> "" And val(Me.TXTid.Text) <> 0 Then
                RsDev.AddNew
                RsDev("Shipments_id").value = val(Me.TXTid.Text)
                RsDev("Container_no").value = .TextMatrix(i, .ColIndex("cno"))
                RsDev("order_no").value = .TextMatrix(i, .ColIndex("ono"))
                RsDev("des").value = .TextMatrix(i, .ColIndex("des"))
                RsDev.update
            End If

        Next i

    End With

    RsDev.Close

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ «·Õ›Ÿ   " & CHR(13)
            Msg = Msg + " Œÿ√ ›Ì «·»Ì«‰«  «·„œŒ·Â " & CHR(13)
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If BeginTrans = True Then
        Cn.RollbackTrans
        BeginTrans = False
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÕœÀ Œÿ√ «À‰«¡ «·Õ›Ÿ " & CHR(13)
    Else
        Msg = "Error while saving"
        End If
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
            rs.find "id='" & val(Me.TXTid.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub CmdExit_Click()
    Frame2.Visible = False
End Sub
 
Private Sub Del_ProfData()
    Dim Msg As String
    On Error GoTo ErrTrap

    If Me.TXTid.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
            Msg = Msg + (Me.TXTid.Text) & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Are you sure you want to delete this record?"
        End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    '  XPTxtCurrent.Caption = 0
                    '  XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            Msg = "this operation is not available due to lack of records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–… «·⁄„·Ì… "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
   
    'ScreenNameArabic = " „ «»⁄Â «·‘Õ‰  "
    'ScreenNameEnglish = " Shipment Follow "
    'RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    DTP_Exp_date_ship.value = Date
    DTP_Exp_date_arriv.value = Date
    DTP_Act_date_clr.value = Date
    DTP_doc_ref_date.value = Date
    DTP_g_delivery_date.value = Date
    DTP_act_date_shipm.value = Date
    DTP_act_date_arrival.value = Date
    DTP_Exp_date_site.value = Date
    DTP_date_clr.value = Date
    DTPi_act_date_site.value = Date
    DTP_G_Refound_date.value = Date

    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    'Dcombos.GetEmployees Me.DcboEmp
    'Dcombos.GetEmployees Me.dcby

    My_SQL = "  select  id,name  from Shipment_mode  "

    fill_combo dc_shipment_mode, My_SQL

    My_SQL = "  select  id,name  from gurantee_type  "

    fill_combo dc_guranty_type, My_SQL

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    'Resize_Form Me
    'AddTip
    Set rs = New ADODB.Recordset
    rs.Open "[Shipments]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    ' If SystemOptions.UserInterface = EnglishInterface Then
    '    SetInterface Me
    '    ChangeLang
    '    End If
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    Else
    End If
    Exit Sub
ErrTrap:

End Sub
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    
    Me.Caption = "Transportation Tracking"
    Label1(2).Caption = Me.Caption
    Frame1.Caption = Me.Caption
    Label1(0).Caption = "Shipment No."
    Label2.Caption = "Expected shipping D"
    Label3.Caption = "ETA"
    Label4.Caption = "Distribution date"
    Label5.Caption = "Registration date"
    Label6.Caption = "Arriving Port"
    Label7.Caption = "Shipping route"
    Label8.Caption = "Ship/Plane"
    Label9.Caption = "Insurance Company"
    Label10.Caption = "Insurance Type"
    Label11.Caption = "Expected Insurance arriving D"
    Label12.Caption = "Actual shipping D"
    Label15.Caption = "Receiving D"
    DTP_1.Caption = "arriving D"
    Label17.Caption = "Shipping port"
    Label18.Caption = "Agent"
    Label20.Caption = "Client"
    Label21.Caption = "Client No."
    Label22.Caption = "Policy No."
    Label23.Caption = "Insurance"
    Label24.Caption = "Re-establishment D"
    Label13.Caption = "Actual Arriving D"
    Label14.Caption = "Departure D"
    Frame2.Caption = "Shipment data"
    
    With Grid
        .TextMatrix(0, .ColIndex("cno")) = "Shipment No."
        .TextMatrix(0, .ColIndex("ono")) = "Order No."
        .TextMatrix(0, .ColIndex("Des")) = "Description"
        .TextMatrix(0, .ColIndex("Ser")) = "No."
    End With
    
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(7).Caption = "Print"
    Cmd(5).Caption = "Search"
    CmdHelp.Caption = "Help"
    Cmd(6).Caption = "Exit"
    
    
End Sub
  
Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim StrMSG As String
    Dim IntResult As String
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

Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)

    With Me.Grid

        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

    End With

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
 
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            Me.TXTid.locked = True

            '  Me.txtEmpCode.locked = True
            '  Me.DcboEmp.locked = True
        
            Frame4.Enabled = False
            'Me.date2.Enabled = False
            'txtnum.locked = True
            'Frame1.Enabled = False
            
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            
            End If

        Case "N"
      
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
        
            Me.TXTid.locked = False

            'Me.txtEmpCode.locked = False
            'Me.DcboEmp.locked = False
        
            ' Frame4.Enabled = True
            'Me.date2.Enabled = True
            'txtnum.locked = False
            'Frame1.Enabled = True

        Case "E"

            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            Me.TXTid.locked = False

            'Me.txtEmpCode.locked = False
            'Me.DcboEmp.locked = False
        
            ' Frame4.Enabled = True
            'Me.date2.Enabled = True
            'txtnum.locked = False
            'Frame1.Enabled = True

    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

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

