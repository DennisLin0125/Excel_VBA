Attribute VB_Name = "ModuleT8畫波形"
Sub T8畫圖()
        Application.ScreenUpdating = False
        Dim arr, oROW%
'*******************************普通畫圖**************************************************************
'        oRow = 7800
'        arr = Range("C2:C" & oRow)
'
'        Dim oArr
'        ReDim oArr(UBound(arr))
'
'        Dim i%
'        For i = 0 To UBound(arr) - 1
'                oArr(i) = Abs(arr(i + 1, 1) * 1000)
'        Next i
'*************************************T8畫圖**********************************************************
       
        oROW = Range("D" & Rows.Count).End(xlUp).Row
        arr = Range("D2:D" & oROW)
        
        Dim oArr
        ReDim oArr(UBound(arr))
        
        Dim i%, temp As Double, AvgNum As Double, Min As Double, Max As Double
        For i = 0 To UBound(arr) - 1
                oArr(i) = Abs(arr(i + 1, 1) * 0.4)
                temp = temp + oArr(i)
        Next i

        '平均
        AvgNum = temp / (oROW - 1)

        '修飾數值
        Dim upLim As Boolean, downLim As Boolean
        For i = 0 To UBound(oArr)
                Min = Math.Round(AvgNum * 0.9, 2)
                Max = Math.Round(AvgNum * 1.09, 2)

                upLim = Min > oArr(i)
                downLim = oArr(i) > Max

                If upLim Or downLim Then
                        oArr(i) = Math.Round(AvgNum * 1.001, 3)
                End If

                If i = oROW - 2 Then Exit For
        Next i

        [G2].Resize(oROW) = Application.WorksheetFunction.Transpose(oArr)
        
 '=============畫圖===============================================
        Dim myChart As ChartObject, sngLeft As Single, sngTop As Single
        
        Dim sh As Worksheet
        Set sh = ActiveSheet
        
        sngLeft = sh.Range("J9").Left
        sngTop = sh.Range("J9").Top
        
        Set myChart = ActiveSheet.ChartObjects.Add(sngLeft, sngTop, 600, 200)
        
        With myChart.Chart
                .SetSourceData Source:=Range("G2:G" & oROW), PlotBy:=xlColumns
                .ChartType = xlLine
                .HasTitle = True
                .HasLegend = False
                .ChartTitle.Text = "idle voltage out"
                '.ChartTitle.Text = "Forward Power"
                
               With .Axes(xlValue, xlPrimary)
                        .MaximumScale = 5
                        .MinimumScale = -5
                        .MajorUnit = 1
'                        .MaximumScale = 5000
'                        .MinimumScale = 0
'                        .MajorUnit = 1000
                        .HasTitle = True
                        .AxisTitle.Text = "idle voltage out  (mV)"
                        '.AxisTitle.Text = "Forward Power  (W)"
               End With
               With .Axes(xlCategory, xlPrimary)
                        .HasTitle = True
                        .AxisTitle.Text = "time (s)"
               End With
        End With

        Application.ScreenUpdating = True
End Sub
