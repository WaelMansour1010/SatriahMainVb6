VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form shipment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entry Of  Shipment"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10920
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   10920
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame2 
      Caption         =   "Containers Information"
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
      TabIndex        =   55
      Top             =   5760
      Width           =   10575
      Begin VSFlex8UCtl.VSFlexGrid Grid 
         Height          =   1515
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   10125
         _cx             =   17859
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
         FormatString    =   $"shipment.frx":0000
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
         RightToLeft     =   0   'False
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
      Caption         =   "Shipment Info"
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
      TabIndex        =   6
      Top             =   840
      Width           =   10695
      Begin VB.CheckBox ChkRefounded_gurantee 
         Caption         =   "Check1"
         Height          =   255
         Left            =   5160
         TabIndex        =   68
         Top             =   4320
         Width           =   255
      End
      Begin VB.TextBox txt_Shipment_no 
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TXT_AWB_OR_BL 
         Height          =   285
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   2520
         Width           =   3015
      End
      Begin VB.TextBox txt_forward_AGENT 
         Height          =   285
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox TXT_CLR_AGENT 
         Height          =   285
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   2880
         Width           =   3015
      End
      Begin VB.TextBox TXT_INS_REQ_NO 
         Height          =   285
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox TXT_POLICY_NO 
         Height          =   285
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox TXT_GURANTY_NO 
         Height          =   285
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   3960
         Width           =   3015
      End
      Begin VB.TextBox txt_portal_of_sale 
         Height          =   285
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txt_insurance_company 
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtVessel_or_Flight 
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   2880
         Width           =   2775
      End
      Begin MSDataListLib.DataCombo dc_shipment_mode 
         Height          =   315
         Left            =   1920
         TabIndex        =   29
         Top             =   2520
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txt_portal_of_dest 
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox TXTid 
         Height          =   285
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTP_Exp_date_ship 
         Height          =   330
         Left            =   1920
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   720
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   96468995
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_Exp_date_arriv 
         Height          =   330
         Left            =   1920
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   96468995
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_Act_date_clr 
         Height          =   330
         Left            =   1920
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   96468995
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_doc_ref_date 
         Height          =   330
         Left            =   1920
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   96468995
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_g_delivery_date 
         Height          =   330
         Left            =   1920
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   4320
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   96468995
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_act_date_shipm 
         Height          =   330
         Left            =   5520
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   720
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   96468995
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_act_date_arrival 
         Height          =   330
         Left            =   5520
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   96468995
         CurrentDate     =   37140
      End
      Begin MSDataListLib.DataCombo dc_guranty_type 
         Height          =   315
         Left            =   1920
         TabIndex        =   32
         Top             =   3600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DTP_Exp_date_site 
         Height          =   330
         Left            =   5520
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   96468995
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_date_clr 
         Height          =   330
         Left            =   9000
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   96468995
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTPi_act_date_site 
         Height          =   330
         Left            =   9000
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   96468995
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_G_Refound_date 
         Height          =   330
         Left            =   7200
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   4320
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   96468995
         CurrentDate     =   37140
      End
      Begin VB.Label Label24 
         Caption         =   "G. Refound Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   47
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "Guranty  No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   46
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "Policy No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   45
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Ins.Requisit No."
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
         Left            =   5520
         TabIndex        =   44
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "Clr Agent"
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
         Left            =   5520
         TabIndex        =   43
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label19 
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
         Height          =   375
         Left            =   5520
         TabIndex        =   42
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Forward Agent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   41
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Portal of Sale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   39
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label DTP_1 
         Caption         =   "Act Date Site."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         TabIndex        =   36
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Exp Date Clr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         TabIndex        =   35
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Exp Date Site :"
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
         Left            =   3720
         TabIndex        =   33
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Act  Date Arrival"
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
         Left            =   3720
         TabIndex        =   20
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   " Act Date .Shipment "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   19
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "G.Delivery Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Guranty Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Insurance Company"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Vessel/Flight"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Shipment Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Port of Destinat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Doc.Rec.Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Act  Date Clr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Exp Date Arrival"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Exp Date .Shipment "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Shipment No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   675
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10920
      _cx             =   19262
      _cy             =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   24
         Charset         =   178
         Weight          =   700
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
      Caption         =   "Entry Of  Shipment"
      Align           =   1
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
      PicturePos      =   1
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
         BackColor       =   &H000000FF&
         Height          =   285
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   90
         Visible         =   0   'False
         Width           =   555
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   9705
         TabIndex        =   2
         Top             =   120
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
         ButtonImage     =   "shipment.frx":0098
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
         Left            =   8640
         TabIndex        =   3
         Top             =   120
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
         ButtonImage     =   "shipment.frx":0432
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
         Left            =   10230
         TabIndex        =   4
         Top             =   120
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
         ButtonImage     =   "shipment.frx":07CC
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
         Left            =   9165
         TabIndex        =   5
         Top             =   120
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
         ButtonImage     =   "shipment.frx":0B66
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
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   900
      Index           =   1
      Left            =   600
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   7920
      Width           =   8655
      _cx             =   15266
      _cy             =   1588
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
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   0
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
         Height          =   645
         Index           =   0
         Left            =   -30
         TabIndex        =   58
         Top             =   135
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1138
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "New"
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
         Height          =   645
         Index           =   1
         Left            =   900
         TabIndex        =   59
         Top             =   135
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1138
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Modify"
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
         Height          =   645
         Index           =   2
         Left            =   1800
         TabIndex        =   60
         Top             =   135
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1138
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Save"
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
         Height          =   645
         Index           =   3
         Left            =   2700
         TabIndex        =   61
         Top             =   135
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1138
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Undo"
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
         Height          =   645
         Index           =   4
         Left            =   3795
         TabIndex        =   62
         Top             =   135
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1138
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Delete"
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
         Height          =   645
         Index           =   5
         Left            =   5595
         TabIndex        =   63
         Top             =   135
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   1138
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Search"
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
         Height          =   645
         Index           =   6
         Left            =   7710
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   135
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1138
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Exit"
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
         Height          =   645
         Index           =   7
         Left            =   4770
         TabIndex        =   65
         Top             =   135
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1138
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Print"
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
         Height          =   645
         Left            =   6810
         TabIndex        =   66
         Top             =   135
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1138
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Help"
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
Attribute VB_Name = "shipment"
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
 
            TxtModFlg.text = "N"
            clear_all Me
            Me.txtid.text = CStr(new_id("Shipments", "id", "", True))
 
            '   Txt_DateEndLincH.value = ToHijriDate(Date)
        Case 1
 
            TxtModFlg.text = "E"
            Me.Grid.Rows = Me.Grid.Rows + 1

        Case 2
    
            '  calc_total
            ' cal_interval
            SaveData
        
        Case 3
            Undo

        Case 4
 
            Del_ProfData

        Case 5
 
            '  FrmEmployeeSearch.Show ' vbModal
        Case 6
            Unload Me

        Case 7
 
            printingReport (val(txtid.text))
    End Select

    Exit Sub
ErrTrap:
 
End Sub
 
Public Function printingReport(Optional id As String, _
                               Optional date_from)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From Shipment_all_details where id=" & id

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

    Me.txtid.text = IIf(IsNull(rs("id").value), 0, val(rs("id").value))
    Me.txt_Shipment_no.text = IIf(IsNull(rs("Shipment_no").value), "", rs("Shipment_no").value)

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
 
    Me.txt_portal_of_dest.text = IIf(IsNull(rs("portal_of_dest").value), "", Trim(rs("portal_of_dest").value))
    Me.txtVessel_or_Flight.text = IIf(IsNull(rs("Vessel_or_Flight").value), "", Trim(rs("Vessel_or_Flight").value))
    Me.txt_insurance_company.text = IIf(IsNull(rs("insurance_company").value), "", Trim(rs("insurance_company").value))
    Me.txt_portal_of_sale.text = IIf(IsNull(rs("portal_of_sale").value), "", Trim(rs("portal_of_sale").value))
    Me.txt_forward_AGENT.text = IIf(IsNull(rs("forward_AGENT").value), "", Trim(rs("forward_AGENT").value))
    Me.TXT_AWB_OR_BL.text = IIf(IsNull(rs("AWB_OR_BL").value), "", Trim(rs("AWB_OR_BL").value))
    Me.TXT_CLR_AGENT.text = IIf(IsNull(rs("CLR_AGENT").value), "", Trim(rs("CLR_AGENT").value))
    Me.TXT_INS_REQ_NO.text = IIf(IsNull(rs("INS_REQ_NO").value), "", Trim(rs("INS_REQ_NO").value))
    Me.TXT_POLICY_NO.text = IIf(IsNull(rs("POLICY_NO").value), "", Trim(rs("POLICY_NO").value))
    Me.TXT_GURANTY_NO.text = IIf(IsNull(rs("GURANTY_NO").value), "", Trim(rs("GURANTY_NO").value))

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
    StrSQL = "Select * from  Shipments_details  where  Shipments_id=" & val(Me.txtid.text)
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

    If Me.TxtModFlg.text <> "R" Then

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
 
        '    If Me.txtnum.text <> "" Then
        '        If Not (IsNumeric(Me.txtnum.text)) Then
        '            Msg = "«·ﬁÌ„… ·«»œ «‰  ﬂÊ‰ —ﬁ„"
        '            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '            Me.txtnum.SetFocus
        '            SelectText Me.txtnum
        '            Exit Sub
        '        End If
        '    End If
 
        Cn.BeginTrans
        BeginTrans = True
    
        If TxtModFlg.text = "N" Then
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
        ElseIf Me.TxtModFlg.text = "E" Then
    
            StrSQL = "Delete   Shipments_details Where Shipments_id=" & val(Me.txtid.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
        End If

        rs("id").value = val(Me.txtid.text)
        rs("Shipment_no").value = IIf(txt_Shipment_no.text = "", Null, Trim(txt_Shipment_no.text))
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
 
        rs("portal_of_dest").value = IIf(Me.txt_portal_of_dest.text = "", Null, Me.txt_portal_of_dest.text)

        rs("Vessel_or_Flight").value = IIf(Me.txtVessel_or_Flight.text = "", Null, Me.txtVessel_or_Flight.text)
        rs("insurance_company").value = IIf(Me.txt_insurance_company.text = "", Null, Me.txt_insurance_company.text)
        rs("portal_of_sale").value = IIf(Me.txt_portal_of_sale.text = "", Null, Me.txt_portal_of_sale.text)
        rs("forward_AGENT").value = IIf(Me.txt_forward_AGENT.text = "", Null, Me.txt_forward_AGENT.text)
        rs("AWB_OR_BL").value = IIf(Me.TXT_AWB_OR_BL.text = "", Null, Me.TXT_AWB_OR_BL.text)
        rs("CLR_AGENT").value = IIf(Me.TXT_CLR_AGENT.text = "", Null, Me.TXT_CLR_AGENT.text)
        rs("INS_REQ_NO").value = IIf(Me.TXT_INS_REQ_NO.text = "", Null, Me.TXT_INS_REQ_NO.text)
        rs("POLICY_NO").value = IIf(Me.TXT_POLICY_NO.text = "", Null, Me.TXT_POLICY_NO.text)
        rs("GURANTY_NO").value = IIf(Me.TXT_GURANTY_NO.text = "", Null, Me.TXT_GURANTY_NO.text)

        rs("shipment_mode").value = IIf(Me.dc_shipment_mode.BoundText = "", 0, Me.dc_shipment_mode.BoundText)
        rs("guranty_type").value = IIf(Me.dc_guranty_type.BoundText = "", 0, Me.dc_guranty_type.BoundText)
      
        rs.update
    
        Cn.CommitTrans
        BeginTrans = False
        '    XPTxtCurrent.Caption = Rs.AbsolutePosition
        '    XPTxtCount.Caption = Rs.RecordCount
    
        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = " Shipment Information Saved " & Chr(13)
                Msg = Msg + "Do you want to Add  anew Shipment Information"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        End Select

        TxtModFlg.text = "R"
    End If

    Set RsDev = New ADODB.Recordset
    RsDev.Open "[Shipments_details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    With Me.Grid

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("cno")) <> "" And val(Me.txtid.text) <> 0 Then
                RsDev.AddNew
                RsDev("Shipments_id").value = val(Me.txtid.text)
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
        Msg = "Cant' Save  " & Chr(13)
        Msg = Msg + "Check entered Values " & Chr(13)
        'Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If BeginTrans = True Then
        Cn.RollbackTrans
        BeginTrans = False
    End If

    Msg = "Error During Saving " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "id='" & val(Me.txtid.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
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

    If Me.txtid.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + (Me.txtid.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

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
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–… «·⁄„·Ì… "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

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
    Me.TxtModFlg.text = "R"

    ' If SystemOptions.UserInterface = EnglishInterface Then
    '    SetInterface Me
    '    ChangeLang
    '    End If
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub ChangeLang()

End Sub
  
Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim StrMSG As String
    Dim IntResult As String
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
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
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

    Select Case Me.TxtModFlg.text

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
            Me.txtid.locked = True

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
        
            Me.txtid.locked = False

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
            Me.txtid.locked = False

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

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
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

