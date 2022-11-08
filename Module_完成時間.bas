Attribute VB_Name = "Module完成時間"
Sub 完成時間()
        Myname = InputBox("請輸入RMA")
        
        If Myname = "" Then Exit Sub
        
        Dim rng As Range
        Dim oROW%
        Dim machine As Range
        
        oROW = Range("A" & Rows.Count).End(xlUp).Row
        
        Set machine = Range("A1:A" & oROW).Find(What:=Myname, LookAt:=xlWhole)
        
        If Not machine Is Nothing Then
                oROW = machine.Row
                Range("A" & oROW & ":O" & oROW).Cut
                Rows("2:2").Insert Shift:=xlDown
                
                
                Range("G2") = Date
                Range("G2").NumberFormatLocal = "mm/dd/yy"
                
                Range("J2") = Range("G2") - Range("C2")
                
                Range("J2").NumberFormatLocal = "0_)"
                
                With Range("J2").Font
                        .name = "Tahoma"
                        .Size = 10
                End With
                
                Range("K2:O2").ClearContents
                
                Range("A2").Interior.Pattern = xlNone
                
                Range("A1").Select
                
                strBody = ""
                
                Dim outApp As Object
                Set outApp = CreateObject("Outlook.Application")
        
                Dim outMail As Object
                Set outMail = outApp.CreateItem(0)
                
                With outMail
                          .To = ""
                          .CC = ""
                          .BCC = ""
                          .Subject = Myname & " 完成維修"
                          .HtmlBody = strBody
                          .Attachments.Add ""
                          .Display
                End With
                  
                Set outApp = Nothing
                Set outMail = Nothing
                Set machine = Nothing
        Else
                MsgBox "查無 " & Myname & " RMA"
        End If
End Sub

