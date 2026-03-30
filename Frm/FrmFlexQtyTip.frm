VERSION 5.00
Begin VB.Form FrmFlexQtyTip 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "«·ﬂ„Ì… «·„ «Õ…"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   3075
      Index           =   0
      Left            =   -3720
      ScaleHeight     =   3045
      ScaleWidth      =   3825
      TabIndex        =   0
      Top             =   30
      Width           =   3855
      Begin VB.Image Img 
         Height          =   240
         Index           =   2
         Left            =   1560
         Picture         =   "FrmFlexQtyTip.frx":0000
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   1
         Left            =   2940
         Picture         =   "FrmFlexQtyTip.frx":038A
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label LblQty 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "«·ﬂ„Ì…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label LblUnit 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "«·ÊÕœ…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2010
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1380
         Width           =   945
      End
      Begin VB.Shape ShapContaner 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         Height          =   1695
         Index           =   0
         Left            =   450
         Top             =   1290
         Width           =   2985
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   930
         Width           =   2295
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   630
         Width           =   2295
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   330
         Width           =   2295
      End
      Begin VB.Label LblData 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   30
         Width           =   2295
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„Œ“‰:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2370
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   930
         Width           =   1425
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„Ã„Ê⁄…:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2370
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   630
         Width           =   1425
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·’‰›:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2370
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   330
         Width           =   1425
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·’‰›:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   60
         Width           =   1425
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   0
         Left            =   3540
         Picture         =   "FrmFlexQtyTip.frx":0714
         Top             =   30
         Width           =   240
      End
      Begin VB.Shape ShapContaner 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   1695
         Index           =   1
         Left            =   450
         Top             =   1320
         Width           =   2985
      End
   End
   Begin VB.PictureBox PicContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   4275
      Index           =   1
      Left            =   30
      ScaleHeight     =   4245
      ScaleWidth      =   5025
      TabIndex        =   11
      Top             =   30
      Width           =   5055
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«⁄·Ï ”⁄—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1950
         Width           =   1425
      End
      Begin VB.Label LblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   16
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   2010
         Width           =   3225
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«œ‰Ï ”⁄—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   3570
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1680
         Width           =   1425
      End
      Begin VB.Label LblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   15
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1710
         Width           =   3225
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·›« Ê—…:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   2580
         Width           =   1305
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   8040
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬂÊœ «·ﬁœÌ„ :-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   10380
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label LblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   13
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1410
         Width           =   3225
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”⁄— «·»Ì⁄ ··œÌ·—-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   3570
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1380
         Width           =   1425
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   3780
         Width           =   3225
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   3360
         Width           =   3225
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   2610
         Width           =   3225
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   3060
         Width           =   3225
      End
      Begin VB.Label LblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1110
         Width           =   3225
      End
      Begin VB.Label LblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   840
         Width           =   3225
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”⁄— «·»Ì⁄ ··⁄„Ì·:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   3570
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1110
         Width           =   1425
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”⁄— «·»Ì⁄ ··„” Â·ﬂ:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   840
         Width           =   1755
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   8
         Left            =   3330
         Picture         =   "FrmFlexQtyTip.frx":0A9E
         Top             =   5280
         Width           =   240
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·⁄„Ì·:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   3840
         Width           =   1305
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   7
         Left            =   4710
         Picture         =   "FrmFlexQtyTip.frx":0E28
         Top             =   3870
         Width           =   240
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·”⁄— -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   3420
         Width           =   1305
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   6
         Left            =   4710
         Picture         =   "FrmFlexQtyTip.frx":11B2
         Top             =   3450
         Width           =   240
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·›« Ê—…:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2970
         Width           =   1305
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   4
         Left            =   4710
         Picture         =   "FrmFlexQtyTip.frx":153C
         Top             =   2730
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   3
         Left            =   4680
         Picture         =   "FrmFlexQtyTip.frx":1AC6
         Top             =   2370
         Width           =   240
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "√Œ— ›« Ê—… »Ì⁄ ··’‰›:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   4
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   2310
         Width           =   1755
      End
      Begin VB.Shape ShapContaner 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         Height          =   1635
         Index           =   2
         Left            =   0
         Top             =   2580
         Width           =   5145
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   5
         Left            =   4740
         Picture         =   "FrmFlexQtyTip.frx":1E50
         Top             =   30
         Width           =   240
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·’‰›:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   30
         Width           =   1425
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·’‰›:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   3570
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„Ã„Ê⁄…:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   3570
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   540
         Width           =   1425
      End
      Begin VB.Label LblData 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   -1800
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   30
         Width           =   5025
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Index           =   6
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   120
         Width           =   3225
      End
      Begin VB.Label LblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   540
         Width           =   3225
      End
      Begin VB.Shape ShapContaner 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         Height          =   1305
         Index           =   3
         Left            =   210
         Top             =   2670
         Width           =   3345
      End
   End
End
Attribute VB_Name = "FrmFlexQtyTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        lbl(0).Caption = "Item Code:-"
        lbl(1).Caption = "Item Name:-"
        lbl(2).Caption = "Group Name:-"
        lbl(3).Caption = "Store Name:-"
        LblUnit(0).Caption = "Unit"
        lblqty(0).Caption = "Quantity"
    
        lbl(7).Caption = "Item Code:-"
        lbl(6).Caption = "Item Name:-"
        lbl(5).Caption = "Group Name:-"
        lbl(11).Caption = "User Price:-"
        lbl(12).Caption = "Customer Price:-"
        lbl(13).Caption = "Dealer Price:-"
        lbl(14).Caption = "Part No:-"
        lbl(4).Caption = "Last Item Invoice:-"
        lbl(8).Caption = "Invoice Date:-"
        lbl(9).Caption = "Item Price:-"
        lbl(10).Caption = "Customer Name:-"
    
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           Y As Single)
    Unload Me
End Sub

Private Sub Form_Resize()
    Me.PicContainer(0).Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth - 50, Me.ScaleHeight - 50
    Me.PicContainer(1).Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth - 50, Me.ScaleHeight - 50
End Sub

Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          x As Single, _
                          Y As Single)
    Unload Me
End Sub

Private Sub LblData_MouseMove(Index As Integer, _
                              Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              Y As Single)
    Unload Me
End Sub

Private Sub LblUnit_MouseMove(Index As Integer, _
                              Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              Y As Single)
    Unload Me
End Sub

Private Sub PicContainer_MouseMove(Index As Integer, _
                                   Button As Integer, _
                                   Shift As Integer, _
                                   x As Single, _
                                   Y As Single)
    Unload Me
End Sub
