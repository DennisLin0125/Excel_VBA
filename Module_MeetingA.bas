Attribute VB_Name = "ModulemeetingA"
Option Explicit
Function 計算維修W3M() As Integer
        Dim W3MRP%, strLen$, i%
        W3MRP = 0
        If [A2] <> "" Then
                For i = 2 To [A1].End(xlDown).Row
                        strLen = Range("I" & i)
                        If InStr(strLen, "*") > 0 Then
                            W3MRP = W3MRP + 1
                        End If
                Next i
        Else
                W3MRP = 0
        End If
            
        計算維修W3M = W3MRP
    
End Function
Function 計算W3M() As Integer

        Dim W3MTime%, LF%, i%, strLen$
        If Range("H2").Value = "" Then
                W3MTime = 0
        Else
                Dim arr(20), oROW
                oROW = 0
                LF = Range("H1").End(xlDown).Row
                W3MTime = 0
                For i = 2 To LF
                strLen = Range("H" & i)
                        If InStr(strLen, "Dennis") > 0 Then
                                W3MTime = W3MTime + 1
                                arr(oROW) = Range("A" & i)
                                oROW = oROW + 1
                        End If
                Next i
        End If
         
        If W3MTime > 0 Then
                MsgBox "你有 " & W3MTime & " 台保固"
                Dim sPath$
                sPath = "D:\Users\Dlin\Desktop\W3M.xlsx"
                Dim wb2 As Workbook
                Set wb2 = Workbooks.Open(sPath, UpdateLinks:=0)
                
                wb2.Sheets("Sheet1").Activate
                If Range("A1") = "" Then
                        Range("A1").Resize(UBound(arr) + 1) = Application.WorksheetFunction.Transpose(arr)
                Else
                        Dim myRow%
                        myRow = Range("A" & Rows.Count).End(xlUp).Row
                        Range("A" & myRow + 1).Resize(UBound(arr) + 1) = Application.WorksheetFunction.Transpose(arr)
                End If
                wb2.Close True
        Else
                MsgBox "恭喜沒有保固,請好好保持 ^^"
        End If
            
        計算W3M = W3MTime
        
End Function
Sub meetingA()
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
    
        Dim meetAns As Boolean, RPAns As Boolean, RPNumAns As Boolean, CpleRateAns As Boolean
        Dim MeetAddress$, RPNumAddress$, CpleRateAddress$, RPAddress$
        Dim machine%, myTime2$, myYear$, myMonth$, myDay$
'*****************************************************************************************************************
        Dim RP%, LF%
        If [A2] = "" Then
                RP = 0
        Else
                LF = [A1].End(xlDown).Row
                RP = LF - 1
        End If
'*****************************************************************************************************************
        Dim oWR%, oWFP%, oWFC%, oKT%, i%
        For i = 1 To Range("G" & Rows.Count).End(xlUp).Row
                If Range("G" & i) = "WR" Then
                        oWR = oWR + 1
                ElseIf Range("G" & i) = "WFC" Then
                        oWFC = oWFC + 1
                ElseIf Range("G" & i) = "WFP" Then
                        oWFP = oWFP + 1
                End If
                If Range("B" & i) = "KAITEK" Then
                        oKT = oKT + 1
                End If
        Next i
'*****************************************************************************************************************
        Dim fs As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
    
        myTime2 = Format(Date + 1, "yyyymmdd")
        
        myYear = Mid(myTime2, 1, 4) '取年分
        
        myMonth = Mid(myTime2, 5, 2) '取月份
        
        myDay = Mid(myTime2, 7, 2) '取日期
    
        MeetAddress = ""  '開啟Weekly
        meetAns = fs.FileExists(MeetAddress)
        
        RPAddress = ""     '開啟工時
        RPAns = fs.FileExists(RPAddress)
        
        RPNumAddress = ""     '開啟維修台數
        RPNumAns = fs.FileExists(RPAddress)
        
        CpleRateAddress = ""     '開啟達成率
        CpleRateAns = fs.FileExists(RPAddress)
'**************************************************************************************************************************************************************
        '偵測檔案存在
        If Not meetAns Then
                MsgBox "Weekly檔案不存在"
                Exit Sub
        ElseIf Not RPAns Then
                MsgBox "工時檔案不存在"
                Exit Sub
        ElseIf Not RPNumAns Then
                MsgBox "維修台數檔案不存在"
                Exit Sub
        ElseIf Not CpleRateAns Then
                MsgBox "達成率檔案不存在"
                Exit Sub
        End If
        
'*************************************************************************************************************************************************
        '偵測唯讀
        Dim wbMeet As Workbook, UpdateLinks%
        Set wbMeet = Workbooks.Open(MeetAddress, UpdateLinks:=0)
        If wbMeet.ReadOnly Then
                wbMeet.Close False
                MsgBox "請注意, Weekly目前為唯讀"
                Exit Sub
        End If
                
        Dim wbRP As Workbook
        Set wbRP = Workbooks.Open(RPAddress, UpdateLinks:=0)
        If wbRP.ReadOnly Then
                wbMeet.Close False
                wbRP.Close False
                MsgBox "請注意, 工時目前為唯讀"
                Exit Sub
        End If

        Dim wbRPNum As Workbook
        Set wbRPNum = Workbooks.Open(RPNumAddress, UpdateLinks:=0)
        If wbRPNum.ReadOnly Then
                wbMeet.Close False
                wbRP.Close False
                wbRPNum.Close False
                MsgBox "請注意, 維修台數目前為唯讀"
                Exit Sub
        End If

        Dim wbCpleRate As Workbook
        Set wbCpleRate = Workbooks.Open(CpleRateAddress, UpdateLinks:=0)
        If wbCpleRate.ReadOnly Then
                wbMeet.Close False
                wbRP.Close False
                wbRPNum.Close False
                wbCpleRate.Close False
                MsgBox "請注意, 達成率目前為唯讀"
                Exit Sub
        End If
'*****************************************************************************************************************
        '處理達成率
        Application.ScreenUpdating = True
        
        wbCpleRate.Worksheets("達成率").Select
        Dim temp%, LF3, strLen$, k%
        Dim workTime, myTemp
        temp = 0
        LF3 = Range("A1").End(xlDown).Row
        For i = 1 To LF3
                strLen = Range("A" & i)
                If InStr(strLen, "Dennis") > 0 Then
                        k = (Range("A" & i).End(xlToRight).Column) + 1
                        Cells(i, k).Select
                        workTime = InputBox("請輸入Dennis資料", "輸入")
                        Cells(i, k) = workTime
                         If temp = 0 Then
                                myTemp = workTime
                                temp = temp + 1
                         End If
                End If
        Next
'*****************************************************************************************************************
        '處理Weekly
        wbMeet.Activate
        wbMeet.Worksheets("Dennis").Delete
        Workbooks("RMA by Dennis.xls").Worksheets("Meeting").Copy After:=wbMeet.Worksheets("Jay")
        Worksheets("Meeting").name = "Dennis"
        
        Dim W3MRP%
        W3MRP = 計算維修W3M
        
        Workbooks("RMA by Dennis.xls").Activate
        
        Application.ScreenUpdating = False

        machine = InputBox("請輸入下周預排台數", "開啟")
        
        wbMeet.Worksheets("This Week").Activate
        
        Dim nameRng As Range
        LF = Range("A5").End(xlDown).Row
        Set nameRng = Range("A5:A" & LF).Find(What:="Dennis Lin", LookAt:=xlWhole)
        
        Range("D" & nameRng.Row) = Range("C" & nameRng.Row) '複製下周台數
        Range("C" & nameRng.Row) = machine      '安排台數
        Range("G" & nameRng.Row) = RP '本周維修
        Range("H" & nameRng.Row) = oWR - oKT '待修
      
        Range("I" & nameRng.Row) = oWFC 'WFC
        Range("J" & nameRng.Row) = oWFP 'WFP
        Range("N" & nameRng.Row) = Range("N" & nameRng.Row) + Range("G" & nameRng.Row) '總維修台數
        Range("Q" & nameRng.Row) = W3MRP '修了多少W3M
        Range("K" & nameRng.Row) = oKT '備品
          
        Worksheets("W3M").Activate
        Dim W3M%
        W3M = 計算W3M
        Worksheets("This Week").Activate
        Range("O" & nameRng.Row) = Range("O" & nameRng.Row) + W3M

 '********************************************************************************************************************************************************
        '處理工時
        wbRP.Worksheets("Analysis").Activate
        
        Dim oDennis As Range, oROW%, oColumn%
        Set oDennis = Cells.Find(What:="Dennis", LookAt:=xlWhole)
        
        oROW = oDennis.Row
        oColumn = oDennis.Column
        
        Cells(oROW + 12, oColumn) = RP
        Cells(oROW + 11, oColumn) = myTemp
        
'*******************************************************************************************************************************************
        '處理維修台數
        wbRPNum.Worksheets("repair list").Activate
        
        Dim LF1%
        LF1 = 1314

        For i = 4 To LF1
                strLen = Cells(1, i)
                If InStr(strLen, "Dennis") > 0 Then
                        Cells(1, i).Select
                        Selection.End(xlDown).Select
                        
                        Cells(ActiveCell.Row + 1, i) = RP
                        Cells(ActiveCell.Row + 1, i + 1) = 0
                        Cells(ActiveCell.Row + 1, i + 2) = W3M
                        Exit For
                End If
        Next

        Worksheets("Test list").Select
        Dim LF2%
        LF2 = Range("C1").End(xlToRight).Column

        For i = 2 To LF2
                strLen = Cells(1, i)
                If InStr(strLen, "Dennis") > 0 Then
                        Cells(1, i).Select
                        Selection.End(xlDown).Select
                        
                        Cells(ActiveCell.Row + 1, i) = 0
                        Cells(ActiveCell.Row + 2, i) = 0
                        Cells(ActiveCell.Row + 3, i) = 0
                        Cells(ActiveCell.Row + 4, i) = 0
                End If
        Next i
        
        wbMeet.Close True
        wbRP.Close True
        wbRPNum.Close True
        wbCpleRate.Close True
        
        Set wbCpleRate = Nothing
        Set wbRPNum = Nothing
        Set wbRP = Nothing
        Set wbMeet = Nothing
        
        MsgBox "Meeting 已全部完成", vbInformation
        
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
End Sub


