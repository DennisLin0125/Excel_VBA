Attribute VB_Name = "Module找尋RMA欄位次數"
Sub 找尋RMA欄位()
        Application.ScreenUpdating = False
        Dim mainSh As Worksheet
        Set mainSh = ActiveSheet
        
        Dim rg As Range
        Set rg = Range("A2:A" & Range("A" & Rows.Count).End(3).Row)
        
        Dim sh As Worksheet
        Dim ans As Range
        Dim RMA As Range, c As Range, str$, k%
        For Each c In rg
                For Each sh In ActiveWorkbook.Worksheets
                        sh.Activate
                        If sh.name <> "下周排程" Then
                                Set RMA = Cells.Find(What:=c, LookAt:=xlPart)
                                If Not RMA Is Nothing Then
                                        temp = c.Row
                                        Dim firstRng As Range
                                        Set firstRng = RMA
                                        Do
                                                k = k + 1
                                                str = str & "頁面 : " & sh.name & Space(5) & "欄位: "" & RMA.Address & vbCrLf
                                                Set RMA = Cells.FindNext(RMA)
                                        Loop Until RMA.Address = firstRng.Address
                                End If
                        End If
                Next sh
                mainSh.Range("J" & temp) = str
                mainSh.Range("K" & temp) = k
                
                
                If k <= 2 Then
                        mainSh.Range("K" & temp).Interior.Color = vbRed
                Else
                        mainSh.Range("K" & temp).Interior.Color = xlNone
                End If
                str = ""
                k = 0
        Next c
        mainSh.Activate
                Application.ScreenUpdating = True
End Sub
