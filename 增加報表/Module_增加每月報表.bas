Attribute VB_Name = "Module增加每月報表"
Option Explicit

Sub 增加每月報表()
        Dim fs As Object, path$, myMonth$, nameStr, nameStrs
        Set fs = CreateObject("Scripting.FileSystemObject")

        
        path = ""
        myMonth = Month(Date) + 1
        
        nameStrs = Array("Dennis", "Cena", "Eric", "Tracy", "週進度報告")
        With fs
                .CreateFolder (path & Year(Date) & "." & myMonth)
                For Each nameStr In nameStrs
                        .CreateFolder (path & Year(Date) & "." & myMonth & "\" & nameStr)
                Next
        End With
        Set fs = Nothing
        
        Dim savePath$, Save_name$
        For Each nameStr In nameStrs
                If nameStr <> "週進度報告" Then
                        savePath = path & Year(Date) & "." & myMonth & "\" & nameStr & "\"
                        Save_name = "R & D Personal From" & Year(Date) & "." & myMonth & "(" & nameStr & ").xlsx"
                        Dim NewBook As Workbook
                        Set NewBook = Workbooks.Add
                        With NewBook
                                .SaveAs Filename:=savePath & Save_name
                                .Activate
                        End With
                        GetDateINFO
                        ActiveWorkbook.Close True
                End If
        Next
End Sub

Sub GetDateINFO()
        Dim i%, myDate As Date, Firstday As Date, Endday As Date, a%, j%
        Columns("A:B").ColumnWidth = 1
        Columns("C").ColumnWidth = 13.38
        Columns("D").ColumnWidth = 132.75
        [C2] = "Date"
        [D2] = "事項"
        
        Range("C2:D33").Select
        With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                With .Font
                        .name = "新細明體"
                        .Size = 16
                End With
        End With
      
        Rows("2:40").RowHeight = 27.75
        
        With [C2:D33].Borders
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
        End With
        myDate = DateAdd("m", 1, Date) '取得下個月
        Firstday = DateSerial(Year(myDate), Month(myDate), 1) '當月第一天
        Endday = DateSerial(Year(myDate), Month(myDate) + 1, 1 - 1) '當月最後一天
        a = Endday - Firstday
        j = 0
        If a = 29 Then
                For i = 0 To 29     '小月
                        Range("C" & i + 3) = Format(Firstday + j, "yyyy.mm.dd")
                        Range("C" & i + 3).Font.Bold = False
                        If Weekday(Firstday + j) = 1 Or Weekday(Firstday + j) = 7 Then
                                Range("D" & i + 3) = "假日"
                                Range("D" & i + 3).Font.Bold = False
                                Range("D" & i + 3).HorizontalAlignment = xlLeft
                                Range("D" & i + 3).VerticalAlignment = xlCenter
                                Range("D" & i + 3).Font.Color = vbRed
                        End If
                        j = j + 1
                Next
        Else
                For i = 0 To 30    '大月
                        Range("C" & i + 3) = Format(Firstday + j, "yyyy.mm.dd")
                        Range("C" & i + 3).Font.Bold = False
                        If Weekday(Firstday + j) = 1 Or Weekday(Firstday + j) = 7 Then
                                Range("D" & i + 3) = "假日"
                                Range("D" & i + 3).Font.Bold = False
                                Range("D" & i + 3).HorizontalAlignment = xlLeft
                                Range("D" & i + 3).VerticalAlignment = xlCenter
                                Range("D" & i + 3).Font.Color = vbRed
                        End If
                        j = j + 1
                Next
        End If
End Sub



