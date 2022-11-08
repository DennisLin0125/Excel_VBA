Attribute VB_Name = "Module搜尋RMA"
Sub 搜尋RMA()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim arr
Dim WRAns As Boolean, CompAns As Boolean
Dim wb As Workbook
Dim Keyword$, RMA$
Dim tempWR$, tempComplete$
Dim sh As Worksheet
Dim sh2 As Worksheet
Dim a As Range, b As Range

If [A2] = "" Then Exit Sub

Set sh = Workbooks("待修分析.xlsm").Worksheets("資料")

Range("I3:I100") = ""

LF = 3

myTime = Time

Keyword = Cells(5, "D")

Set fs = CreateObject("Scripting.FileSystemObject")

arr = Cells(1, 1).CurrentRegion

For j = 2 To UBound(arr)
        RMA = arr(j, 1)
        
        For i = 2021 To 2006 Step -1
        
                tempWR = ""
                tempComplete = ""
                
                WRAns = fs.FileExists(tempWR)
                CompAns = fs.FileExists(tempComplete)
        
                If WRAns Then
                
                        Set wb = Workbooks.Open(tempWR, UpdateLinks:=0)
                        
                        For Each sh2 In ActiveWorkbook.Worksheets
                                sh2.Select
                                If sh2.name = "User parts" Then
                                        Worksheets("User parts").name = "Use parts"
                                End If
                        Next
                        
                        Set a = Sheets("RMA").Range("B46:B100").Find(What:=Keyword, LookAt:=xlWhole)
                        Set b = Sheets("Use parts").Range("A4:A100").Find(What:=Keyword, LookAt:=xlWhole)
                        
                        If Not a Is Nothing Then
                                sh.Cells(LF, "I") = RMA
                                LF = LF + 1
                        ElseIf Not b Is Nothing Then
                                sh.Cells(LF, "I") = RMA
                                LF = LF + 1
                        ElseIf InStr(Sheets("RMA").Range("J19"), Keyword) Then
                                sh.Cells(LF, "I") = RMA
                                LF = LF + 1
                        End If
                        
                        Set a = Nothing
                        Set b = Nothing
                        
                        wb.Close False
                        Exit For
                ElseIf CompAns Then
                
                        Set wb = Workbooks.Open(tempComplete, UpdateLinks:=0)
                        
                        For Each sh2 In ActiveWorkbook.Worksheets
                                sh2.Select
                                If sh2.name = "User parts" Then
                                        Worksheets("User parts").name = "Use parts"
                                End If
                        Next
                        
                        Set a = Sheets("RMA").Range("B46:B100").Find(What:=Keyword, LookAt:=xlWhole)
                        Set b = Sheets("Use parts").Range("A4:A100").Find(What:=Keyword, LookAt:=xlWhole)
                        
                        If Not a Is Nothing Then
                                sh.Cells(LF, "I") = RMA
                                LF = LF + 1
                        ElseIf Not b Is Nothing Then
                                sh.Cells(LF, "I") = RMA
                                LF = LF + 1
                        ElseIf InStr(Sheets("RMA").Range("J19"), Keyword) Then
                                sh.Cells(LF, "I") = RMA
                                LF = LF + 1
                        End If
                        
                        Set a = Nothing
                        Set b = Nothing
                        
                        wb.Close False
                        Exit For
                End If
        Next i
        mum = mum + 1
Next j

Set wb = Nothing
Set fs = Nothing

Application.DisplayAlerts = True
Application.ScreenUpdating = True

myTime = Time - myTime
myMin = Minute(myTime)
mySec = Second(myTime)

MsgBox "搜尋完畢" & Chr(10) & Chr(10) & "搜尋時間" & myMin & "分" & mySec & "秒" & Chr(10) & Chr(10) & "共搜尋" & mum & "次"
End Sub
