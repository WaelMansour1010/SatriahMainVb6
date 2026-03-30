VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPrinterDialog 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÅÎÊíÇÑ ÇáØÇÈÚÉ æÎÕÇÆÕ ÇáØÈÇÚÉ"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "FrmPrinterDialog.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   2550
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   4020
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   4140
      Top             =   3900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÚÏÏ ÇáäÓÎ( ÚÏÏ ÇáäÓÎÉ ãä ßá æÑÞÉ)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1785
      Index           =   2
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1980
      Width           =   3405
      Begin VB.TextBox TxtCopies 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   300
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Text            =   "1"
         Top             =   420
         Width           =   1845
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   345
         Left            =   60
         TabIndex        =   17
         Top             =   420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox ChkCollate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÊÑÊíÈ ÇáäÓÎ"
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         ForeColor       =   &H000000C0&
         Height          =   165
         Index           =   10
         Left            =   3180
         TabIndex        =   24
         Top             =   1050
         Width           =   105
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         ForeColor       =   &H000000C0&
         Height          =   165
         Index           =   9
         Left            =   3060
         TabIndex        =   23
         Top             =   1170
         Width           =   105
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         ForeColor       =   &H000000C0&
         Height          =   165
         Index           =   8
         Left            =   2400
         TabIndex        =   22
         Top             =   990
         Width           =   105
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         ForeColor       =   &H000000C0&
         Height          =   165
         Index           =   7
         Left            =   2250
         TabIndex        =   21
         Top             =   1110
         Width           =   105
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H000000C0&
         Height          =   165
         Index           =   6
         Left            =   2940
         TabIndex        =   20
         Top             =   1320
         Width           =   105
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H000000C0&
         Height          =   165
         Index           =   5
         Left            =   2100
         TabIndex        =   19
         Top             =   1230
         Width           =   105
      End
      Begin VB.Shape Shp 
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   5
         Left            =   1890
         Top             =   960
         Width           =   345
      End
      Begin VB.Shape Shp 
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   4
         Left            =   2040
         Top             =   840
         Width           =   345
      End
      Begin VB.Shape Shp 
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   3
         Left            =   2190
         Top             =   720
         Width           =   345
      End
      Begin VB.Shape Shp 
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   2
         Left            =   2700
         Top             =   1050
         Width           =   345
      End
      Begin VB.Shape Shp 
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   1
         Left            =   2850
         Top             =   900
         Width           =   345
      End
      Begin VB.Shape Shp 
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   0
         Left            =   2970
         Top             =   780
         Width           =   345
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÚÏÏ ÇáäÓÎ:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   2190
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   450
         Width           =   1125
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "äØÇÞ ÇáÕÝÍÇÊ(ÇáÕÝÍÇÊ ÇáãÑÇÏ ØÈÇÚÊåÇ)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1785
      Index           =   1
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1980
      Width           =   3435
      Begin VB.TextBox TxtPages 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   900
         Width           =   2235
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÇáÕÝÍÇÊ:"
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
         Left            =   2310
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   900
         Width           =   975
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÇáÕÝÍÉ ÇáÍÇáíÉ"
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
         Left            =   1860
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   1425
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Çáßá"
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
         Left            =   1860
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   300
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   13
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   570
         Width           =   1695
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   12
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
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
         Height          =   435
         Index           =   4
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1260
         Width           =   3375
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÇáØÇÈÚÇÊ ÇáãæÌæÏÉ æÇáãÚÑÝÉ Ýì ÇáÌåÇÒ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1905
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   6945
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4260
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrinterDialog.frx":038A
               Key             =   "Printer"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrinterDialog.frx":0724
               Key             =   "Options"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageCombo IcboPrinters 
         Height          =   330
         Left            =   1440
         TabIndex        =   1
         Top             =   270
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "ImageList1"
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   405
         Index           =   0
         Left            =   30
         TabIndex        =   7
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÎÕÇÆÕ..."
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
         ButtonImage     =   "FrmPrinterDialog.frx":0ABE
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   405
         Index           =   1
         Left            =   30
         TabIndex        =   27
         Top             =   720
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÅÖÇÝÉ ØÇÈÚÉ"
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
         ButtonImage     =   "FrmPrinterDialog.frx":0E58
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÇáÃÓã:"
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
         Left            =   6090
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÇáÊÚáíÞ:"
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
         Index           =   14
         Left            =   6090
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÇáãæÞÚ:"
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
         Left            =   6090
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÇáäæÚ:"
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
         Left            =   6090
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÇáÍÇáÉ:"
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
         Left            =   6090
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   660
         Width           =   735
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   375
      Index           =   2
      Left            =   90
      TabIndex        =   28
      Top             =   3960
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÅáÛÇÁ"
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
      ButtonImage     =   "FrmPrinterDialog.frx":11F2
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   3
      Left            =   1110
      TabIndex        =   29
      Top             =   3960
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ØÈÇÚÉ"
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
      ButtonImage     =   "FrmPrinterDialog.frx":178C
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      RightToLeft     =   -1  'True
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6960
      X2              =   30
      Y1              =   3870
      Y2              =   3870
   End
End
Attribute VB_Name = "FrmPrinterDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_UserCancel As Boolean

Private Sub Cmd_Click(Index As Integer)
    Dim cPrinter As ClsPrinters
    Dim cTempObject As ClsPrinters
    Dim Msg As String

    Select Case Index

        Case 0

            If Not IcboPrinters.SelectedItem Is Nothing Then
                Set cPrinter = New ClsPrinters
                cPrinter.ShowPrinterProperties IcboPrinters.text, Me.hWnd
                Set cPrinter = Nothing
            End If

        Case 1
            Set cPrinter = New ClsPrinters
            cPrinter.AddPrinter
            Set cPrinter = Nothing

        Case 2
            Me.UserCancel = True
            Me.Hide

        Case 3

            If Me.IcboPrinters.SelectedItem Is Nothing Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "íÌÈ ÅÎÊíÇÑ ÇáØÇÈÚÉ ÇáÊì ÊÑíÏ ÅÓÊÎÏÇãåÇ Ýì ÚãáíÉ ÇáØÈÇÚÉ..!"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    Msg = "Please choose your printer...!!"
                    MsgBox Msg, vbExclamation, App.Title
                End If

                Exit Sub
            ElseIf Me.IcboPrinters.SelectedItem.Index = 1 Then

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "íÌÈ ÅÎÊíÇÑ ÇáØÇÈÚÉ ÇáÊì ÊÑíÏ ÅÓÊÎÏÇãåÇ Ýì ÚãáíÉ ÇáØÈÇÚÉ..!"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    Msg = "Please choose your printer...!!"
                    MsgBox Msg, vbExclamation, App.Title
                End If

                Exit Sub
            End If

            If Me.Opt(2).value = True Then
                If Trim(Me.TxtPages.text) = "" Then
                    Msg = "íÌÈ ßÊÇÈÉ ÃÑÞÇã ÇáÕÝÍÇÊ ÇáãÑÇÏ ØÈÇÚÊåÇ"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If

            Me.UserCancel = False
            Me.Hide
    End Select

End Sub

Private Sub cmdCommand1_Click()
    Me.Com.ShowPrinter
    Me.Com.PrinterDefault = True

End Sub

Private Sub Form_Load()
    Dim Msg As String
    Dim IntDefIndex As Integer
    Dim xComboItem As MSComctlLib.ComboItem
    CenterForm Me

    FormPostion Me, GetPostion

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÇÏÎá ÃÑÞÇã ÇáÕÝÍÇÊ ãÝÕæáÉ ÈÜ , ãËá 1,3,4  "
        Msg = Msg & Chr(13) & "æÝì ÍÇáÉ ØÈÇÚÉ ãÏì Çæ äØÇÞ ÅÓÊÎÏã - ãËá 1-5"
        lbl(4).Caption = Msg
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Msg = "ÇÏÎá ÃÑÞÇã ÇáÕÝÍÇÊ ãÝÕæáÉ ÈÜ , ãËá 1,3,4  "
    Msg = Msg & Chr(13) & "æÝì ÍÇáÉ ØÈÇÚÉ ãÏì Çæ äØÇÞ ÅÓÊÎÏã - ãËá 1-5"
    lbl(4).Caption = Msg

    With Me.IcboPrinters
        Set .ImageList = Me.ImageList1
        .ComboItems.Add 1, , "ÅÎÊÑ ØÇÈÚÉ", 2, 2

        If Printers.count > 0 Then

            For i = 0 To Printers.count - 1
                Set xComboItem = .ComboItems.Add(, , Printers(i).DeviceName, 1, 1)

                If Printer.DeviceName = Printers(i).DeviceName Then
                    IntDefIndex = i
                    Set IcboPrinters.SelectedItem = xComboItem
                End If

            Next i

        End If

    End With

    Opt(0).value = True
    Opt_Click 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    If UnloadMode <> VBRUN.QueryUnloadConstants.vbFormCode Then
        Me.UserCancel = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub Opt_Click(Index As Integer)
    Me.TxtPages.Enabled = Opt(2).value
End Sub

Private Sub TxtCopies_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtCopies.text, 1)
End Sub

Private Sub TxtPages_KeyPress(KeyAscii As Integer)
    Dim Temp As String

    If KeyAscii = Asc(",") Or KeyAscii = Asc("æ") Or KeyAscii = Asc("-") Then
        Temp = Trim(Me.TxtPages.text)

        If Temp = "" Then
            KeyAscii = 0
            Exit Sub
        End If

        If KeyAscii = Asc(right$(Temp, 1)) Then
            KeyAscii = 0
        ElseIf KeyAscii = Asc(",") And right$(Temp, 1) = "," Then
            KeyAscii = 0
        ElseIf KeyAscii = Asc("-") And right$(Temp, 1) = "-" Then
            KeyAscii = 0
        Else
            KeyAscii = KeyAscii
        End If

    Else
        KeyAscii = KeyAscii_Num(KeyAscii, TxtPages.text, 1)
    End If

End Sub

Private Sub UpDown1_DownClick()

    If val(Me.TxtCopies.text) <= 1 Then
        Me.TxtCopies.text = 1
    Else
        Me.TxtCopies.text = val(Me.TxtCopies.text) - 1
    End If

End Sub

Private Sub UpDown1_UpClick()

    If val(Me.TxtCopies.text) >= 200 Then
        Me.TxtCopies.text = 200
    Else
        Me.TxtCopies.text = val(Me.TxtCopies.text) + 1
    End If

End Sub

Private Sub ChangeLang()
    Me.Caption = "Print Properties"
    Fra(0).Caption = "Choose Printer"
    lbl(3).Caption = "Name:"
    lbl(0).Caption = "Status:"
    lbl(1).Caption = "Type:"
    lbl(2).Caption = "Postion:"
    lbl(14).Caption = "Comment:"
    Fra(1).Caption = "Pages"
    Opt(0).Caption = "All"
    Opt(1).Caption = "Current Page"
    Opt(2).Caption = "Pages"
    Fra(2).Caption = "Copy"
    lbl(11).Caption = "Copies Number"
    ChkCollate.Caption = "Collate"
    Cmd(0).Caption = "Properties..."
    Cmd(1).Caption = "Add Printer..."
    Cmd(2).Caption = "Cancel"
    Cmd(3).Caption = "Print"
End Sub

Public Property Get UserCancel() As Boolean
    UserCancel = m_UserCancel
End Property

Public Property Let UserCancel(ByVal vNewValue As Boolean)
    m_UserCancel = vNewValue
End Property
