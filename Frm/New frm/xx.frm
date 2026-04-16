VERSION 5.00
Object = "{0F0877EF-2A93-4AE6-8BA8-4129832C32C3}#230.0#0"; "SmartMenuXP.ocx"
Begin VB.Form xx 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   420
   ClientLeft      =   210
   ClientTop       =   210
   ClientWidth     =   2355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   420
   ScaleWidth      =   2355
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VBSmartXPMenu.SmartMenuXP SmartMenuXP1 
      Align           =   4  'Align Right
      Height          =   375
      Left            =   1500
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shadow          =   0   'False
   End
End
Attribute VB_Name = "xx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    pBuildMenus
    '    Me.Width = MDIFrmMain.Width - 1000
    '    Me.Height = MDIFrmMain.Height - 1000
    '  Me.left = (MDIFrmMain.Width - Me.Width) / 2
    '    Me.top = (MDIFrmMain.Height - Me.Height) / 2 - 500
End Sub

Private Function pGetPicture(sFileName As String) As StdPicture
    ' - This example uses LoadPicture() to load the menu images from disk
    ' - You can also use an ImageList object for this purpose...
    Set pGetPicture = LoadPicture(App.path + "\Images\" + sFileName + ".ico")
End Function

Private Sub pBuildMenus()
    
    With SmartMenuXP1.MenuItems
        
        ' Root > File...
        .Add 0, "keyFile", , " ÇáãÏÇÑÓ æÇáãÚÇåÏ ÇáÊÚáíãíÉ"
        .Add "keyFile", "keyNew", , "ÈíÇäÇÊ ÇÓÇÓíÉ"
        .Add "keyFile", "keyOpen", , "ÇáÊÓÌíá æ ÇáŞÈæá"
        .Add "keyFile", "keyexam", , "ÍÑßÉ ÇáÇãÊÍÇäÇÊ"
        .Add "keyFile", "keystudentsalarm", , "ãÊÇÈÚå ÇáØáÇÈ"
        .Add "keyFile", "keyparent", , "ãÊÇÈÚå ÇæáíÇÁ ÇáÇãæÑ"
        
        .Add "keyFile", "keytable", , "ÇáÌÏæá ÇáÏÑÇÓí"
        .Add "keyFile", "keybook", , "ÇáßÊÈ ÇáÏÑÇÓíÉ"
        .Add "keyFile", "keybox", , "ÍÑßÉ ÇáÎÒíäÉ"
        .Add "keyFile", "keycard", , "ÍÑßÉ ÇáßÇÑäíåÇÊ"
        .Add "keyFile", "keyreport", , "ÇáÊŞÇÑíÑ"
        .Add "keyFile", "System_manger2", , "ÊÑŞíã ÇáãÓÊäÏÇÊ"
            
        .Add "keyFile", , smiSeparator
        
        ' Root > File > New...
        .Add "keyNew", "keygrades", , "ÇáÓäæÇÊ ÇáÏÑÇÓíÉ"
        .Add "keyNew", "keySPEC", , "ÊÎÕÕÇÊ ÇáãÏÑÓíä"
        .Add "keyNew", "keymister", , "ÈíÇäÇÊ ÇáãÏÑÓíä"
        .Add "keyNew", "keykest", , "ÇäæÇÚ ÇáÇŞÓÇØ"
        .Add "keyNew", "keyfines", , "ÇäæÇÚ ÇáÛÑÇãÇÊ"
        .Add "keyNew", "keySubscription", , "ÇäæÇÚ ÇáÇÔÊÑÇßÇÊ"
        .Add "keyNew", "keyactivity", , "ÇäæÇÚ ÇáÇäÔØÉ"
        .Add "keyNew", "keyalarmtype", , "ÇäæÇÚ ÇäĞÇÑÇÊ ÇáİÕá"
        .Add "keyNew", "keyrevenue", , "ÇäæÇÚ ÇáÇíÑÇÏÇÊ"
        .Add "keyNew", "keyexpanses", , "ÇäæÇÚ ÇáãÕÑæİÇÊ"
        .Add "keyNew", "keyhay", , "ÈíÇäÇÊ ÇáÇÍíÇÁ"
        .Add "keyNew", "keystreet", , "ÈíÇäÇÊ ÇáÔæÇÑÚ"
        .Add "keyNew", "keybus", , "ÊÚÑíİ ÇáÍÇİáÉ"
        .Add "keyNew", "keythisyear", , "ÇáÓäÉ ÇáÏÑÇÓíÉ ÇáÍÇáíÉ"
        .Add "keyNew", "keymanager", , "ÇÓã ãÏíÑ ÇáãÏÑÓÉ ÇáÍÇáÉ"
        .Add "keyNew", "ked_types", , "ÇäæÇÚ ÇáŞíæÏ"
                      
        .Add "keyNew", "keystudent", , "ãáİ ÇáØÇáÈ"
    
        ' Root > File > Open...
        .Add "keyopen", "keyapp", , "ØáÈ ÊÓÌíá"
        .Add "keyopen", "keyrenew", , "ÊÌÏíÏ ÇáÇáÊÍÇŞ"
        .Add "keyopen", "keykestsave", , "ÊÓÌíá ÇáÇŞÓÇØ"
        .Add "keyopen", "keyrefinesave", , "ÊÓÌíá ÇáÛÑÇãÇÊ"
        .Add "keyopen", "keyactivitysave", , "ãÊÇÈÚÉ ÇáÇäÔØÉ"
        ' .Add "keyopen", "keylost", , "ÈÏá İÇÆÏ"
        
        .Add "keyactivitysave", "activitynew", , " ÇÖÇİÉ äÔÇØ ÌÏíÏ áØÇáÈ"
        .Add "keyactivitysave", "activityrenew", , "ÊÌÏíÏ äÔÇØ ØÇáÈ "
        .Add "keyactivitysave", "activitydelete", , " ÍĞİ äÔÇØ ØÇáÈ"
 
        .Add "keybox", "keyboxnewmember", , "ÓÏÇÏ ÑÓæã ÇáÇáÊÍÇŞ "
        .Add "keybox", "keyboxrenewmember", , "ÓÏÇÏ ÑÓæã ÊÌÏíÏ ÇáÇáÊÍÇŞ "
        '.Add "keybox", "keyboxactivitypay", , "ÓÏÇÏ ŞíãÉ ÇáÇäÔØÉ"
 
        '.Add "keybox", "keyboxexpanses", , "ÊÓÌíá ÇáãÕÑæİÇÊ "
        '.Add "keybox", "keyboxrevenue ", , "ÊÓÌíá ÇáÇíÑÇÏÇÊ "
        '.Add "keybox", "keyboxlost", , " ÏİÚ ÑÓæã ÈÏá İÇÆß ááßÇÑäíÉ"

        .Add "keyexam", "keyexam1", , " ÊÚÑíİ ÇáÇãÊÍÇäÇÊ"
        .Add "keyexam", "keyexam2", , " ÊÓÌíá äÊÇÆÌ ÇáÇãÊÍÇäÇÊ"

        .Add "keystudentsalarm", "keystudentsalarm1", , "ÊäÈíÉ ÇáØáÇÈ ÇáãÓÊÍŞ Úáíåã ÇŞÓÇØ æáã ÊÓÏÏ"
        .Add "keystudentsalarm", "keystudentsalarm2", , "ØáÇÈ ÇáŞÇÆãÉ ÇáÓæÏÇÁ"

        '.Add " ", " ", , " "
        .Add "keyparent", "keyparent1", , "Ïáíá ÇáÊáíİæäÇÊ"
        .Add "keyparent", "keyparent2", , "ÇÏÇÑÉ ÇáÑÓÇÆá"
        .Add "keyparent", "keyparent3", , "ãÊÇÈÚå ÇáÛíÇÈ "
        .Add "keyparent", "keyparent4", , " ÇäĞÇÑÇÊ ÇáİÕá"
        .Add "keyparent", "keyparent5", , "ÇÌÊãÇÚ ÇæáíÇÁ ÇáÇãæÑ "
        .Add "keyparent", "keyparent6", , "ÇáÊÍæíá ÇáØÈí ááØÇáÈ "

        .Add "keytable", "keytable1", , " ÇÚÏÇÏ ÇáÌÏæá ÇáÏÑÇÓí"
        .Add "keytable", "keytable2", , "ØÈÇÚå ÇáÌÏæá ÇáÏÑÇÓí "
        .Add "keytable", "keytable3", , "ØÈÇÚå ÌÏÇæá ÇáãÏÑÓíä "
 
        .Add "keybook", "keybook1", , " ÊÚÑíİ ÇáßÊÈ"
        .Add "keybook", "keybook2", , "ÊÓáíã ÇáßÊÈ ááØáÇÈ "

        .Add "keyreport", "keyreport1", , " ÊŞÑíÑ ÍÇáÉ ÇáØÇáÈ"
        .Add "keyreport", "keyreport1", , " ÊŞÑíÑ ãÊÇÈÚå ÇÏÇÁ ÇáãÏÑÓíä"

        .Add "keycard", "keycardready", , "ÇáßÇÑäíåÇÊ ÇáÌÇåÒÉ ááØÈÇÚå "
        .Add "keycard", "keycardprinted", , "ÇáßÇÑäíåÇÊ ÇáãØÈæÚå æãÚåÏå ááÊÓáíã "
        .Add "keycard", "keycardreprint", , "ÇÚÇÏå ØÈÇÚå ÇáßÇÑäíåÇÊ "
      
    End With
    
    SmartMenuXP1.Font.name = "Ms Sans Serif"
    SmartMenuXP1.Font.size = 9

End Sub

Public Sub SmartMenuXP1_Click(ByVal id As Long)

    With SmartMenuXP1.MenuItems
        '   Text1.text = "Menu Item (" + Format(id, "00") + ") = " + .text(id) + vbCrLf + Text1.text
        
        Select Case .key(id)

            Case "keyExit"
            
                ' - The "End" statement is not a recomended way for closing aplications
                ' - It gives lots of problems when subclassing or using Hooks
                ' - You should always try to use "Unload Me"
                ' - However, as you can see, SmartMenuXP supports this feature!
                End
                
            Case "keygrades"
                

            Case "keySPEC"
                FrmEmpSpecifications.show

                ' specefic.Show
            Case "keymister"
                OpenScreen EmployeesScreen

                ' MISTER.Show
            Case "keykest"
               
                         
            Case "keyfines"
                

            Case "keySubscription"
                

            Case "keyactivity"
              

            Case "keyalarmtype"
               
                        
            Case "keyrevenue"
                OpenScreen RevenuesTypes

            Case "keyexpanses"
                OpenScreen ExpensesTypes

            Case "keyhay"
                FrmGovernCitiesData.show

            Case "keystreet"
                streets.show
                        
            Case "keybus"
                
        
            Case "keystudent"
             
            Case "keythisyear"
                

            Case "keymanager"
             

            Case "keyexam1"
               

            Case "keyexam2"
               

            Case "keystudentsalarm1"

                '            alram_frm.Show
            Case "keystudentsalarm2"
        
            
            Case "keyapp"
             

            Case "keyrenew"
             

      
            Case "ked_types"
               
               
            Case "System_manger2"
                System_manger2.show
                         
        End Select
        
    End With
    
End Sub

