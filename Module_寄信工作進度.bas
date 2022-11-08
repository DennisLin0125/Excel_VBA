Attribute VB_Name = "Module寄信工作進度"
Option Explicit
Sub 工作進度()
      Dim strBody$
      strBody = "<B>Fortran</B><br><br>" & _
      "<B>這禮拜工作進度如下:</B>" & _
      "<H3><B>All the best</B></H3>" & _
      "<B>Dennis Lin</B><br>" & _
      "<B>**********************************</B><br>" & _
      "<B>Kaitek Coproration</B><br>" & _
      "<B>14F., No. 659, Bannan Rd, Zhonghe Dist.,</B><br>" & _
      "<B>New Taipei City 23557, Taiwan R.O.C.</B><br>" & _
      "<B>TEL: 02-3234-8222 #251</B><br>" & _
      "<B>**********************************</B>"
              
        Dim outApp As Object
        Set outApp = CreateObject("Outlook.Application")
        
        Dim outMail As Object
        Set outMail = outApp.CreateItem(0)
        
        With outMail
                .To = "fortran.wu@kaitek.com.tw"
                .CC = ""
                .BCC = ""
                .Subject = Year(Date - 4) & Month(Date - 4) & Day(Date - 4) & "~" & Year(Date) & Month(Date) & Day(Date) & "工作進度"
                .HtmlBody = strBody
                '.Body = "請幫忙領料 謝謝"
                .Display
        End With
        Set outApp = Nothing
        Set outMail = Nothing
End Sub
