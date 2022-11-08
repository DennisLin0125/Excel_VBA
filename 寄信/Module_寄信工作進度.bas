Attribute VB_Name = "Module寄信工作進度"
Option Explicit
Sub 工作進度()
      Dim strBody$
      strBody = ""     
        Dim outApp As Object
        Set outApp = CreateObject("Outlook.Application")
        
        Dim outMail As Object
        Set outMail = outApp.CreateItem(0)
        
        With outMail
                .To = ""
                .CC = ""
                .BCC = ""
                .Subject = Year(Date - 4) & Month(Date - 4) & Day(Date - 4) & "~" & Year(Date) & Month(Date) & Day(Date) & "工作進度"
                .HtmlBody = strBody
                .Display
        End With
        Set outApp = Nothing
        Set outMail = Nothing
End Sub
