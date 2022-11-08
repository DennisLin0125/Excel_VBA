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
                
                strBody = "<B>Hi Jack</B><br><br>" & _
                                "<B>" & Myname & "已完成 謝謝</B><br><br>" & _
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
                          .To = "jack.chen@kaitek.com.tw"
                          .CC = "fortran.wu@kaitek.com.tw"
                          .BCC = ""
                          .Subject = Myname & " 完成維修"
                          .HtmlBody = strBody
                          .Attachments.Add "P:\Service\RMA\WR\2022\" & Myname & ".xls"
                          .Display
                End With
                  
                Set outApp = Nothing
                Set outMail = Nothing
                Set machine = Nothing
        Else
                MsgBox "查無 " & Myname & " RMA"
        End If
End Sub

