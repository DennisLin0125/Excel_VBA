Attribute VB_Name = "Module阻抗"
Sub 阻抗()
        
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

        
        Dim snDennis As Worksheet
        Set snDennis = Workbooks("待修分析.xlsm").Worksheets("阻抗")
     
        If [B5] <> "" Then
                Dim rg As Range
                Set rg = snDennis.Range("B5:Q" & Range("B" & Rows.Count).End(xlUp).Row)
                rg.ClearContents
                Set rg = Nothing
        End If
        
        
        Dim rng As Range, DennisRow%
        DennisRow = Range("A" & Rows.Count).End(xlUp).Row
        Set rng = snDennis.Range("A5:A" & DennisRow)
        

'=======================================================================
        Dim RMAnum As String
        Dim Year As String, fs As Object
        Dim WRAns, CompAns As Boolean

        Set fs = CreateObject("Scripting.FileSystemObject")


        For Each c In rng

                RMAnum = c.Value
                DRow = c.Row

                Year = "2021"

                For i = Year To 2006 Step -1

                        WRAns = fs.FileExists("")
                        CompAns = fs.FileExists("")

                        If WRAns = True Then
                                
                                Set wb = Workbooks.Open("")
                                If wb.Sheets(1).[B12] = "新竹科學工業園區力行六路8號" Then
                                        snDennis.Range("B" & DRow) = "TSMC12-P1"
                                ElseIf wb.Sheets(1).[B12] = "741 台南科學園區南科北路1之1號" Then
                                        snDennis.Range("B" & DRow) = "TSMC14"
                                ElseIf wb.Sheets(1).[B12] = "台中中部科學園區科雅六路1號" Then
                                        snDennis.Range("B" & DRow) = "TSMC15"
                                ElseIf wb.Sheets(1).[B12] = "新竹市科學園區研新一路9號 (3廠)" Then
                                        snDennis.Range("B" & DRow) = "TSMC3"
                                ElseIf wb.Sheets(1).[B12] = "新竹科學園區園區三路121號" Then
                                        snDennis.Range("B" & DRow) = "TSMC5"
                                ElseIf wb.Sheets(1).[B12] = "741 台南科學園區南科北路1號 (6廠)" Then
                                        snDennis.Range("B" & DRow) = "TSMC6"
                                ElseIf wb.Sheets(1).[B12] = "新竹市科學園區力行路25號 (8廠)" Then
                                        snDennis.Range("B" & DRow) = "TSMC8"
                                        
                                ElseIf Left(wb.Sheets(1).[B11], 4) = "聯華電子" Then
                                        snDennis.Range("B" & DRow) = Right(wb.Sheets(1).[B11], 8)
                                        
                                ElseIf wb.Sheets(1).[B12] = "新竹科學工業園區園區三路123號" Then
                                        snDennis.Range("B" & DRow) = "VISC1"
                                ElseIf wb.Sheets(1).[B12] = "新竹科學工業園區力行路9號" Then
                                        snDennis.Range("B" & DRow) = "VISC2"
                                ElseIf wb.Sheets(1).[B12] = "桃園縣蘆竹鄉南崁路一段336號 " Then
                                        snDennis.Range("B" & DRow) = "VISC3"
                                        
                                ElseIf wb.Sheets(1).[B12] = "新竹市香山區牛埔南路17巷20號 " Then
                                        snDennis.Range("B" & DRow) = "碩輝科技"
                                ElseIf wb.Sheets(1).[B12] = "新竹市香山區牛埔南路17巷20號 " Then
                                        snDennis.Range("B" & DRow) = "碩輝科技"
                                ElseIf wb.Sheets(1).[B12] = "新竹市高翠路327巷4弄32號 " Then
                                        snDennis.Range("B" & DRow) = "寶虹"
                                End If
                                
                                
                                snDennis.Range("C" & DRow) = wb.Sheets(1).[F8]
                                snDennis.Range("D" & DRow) = wb.Sheets(1).[F9]
                                snDennis.Range("E" & DRow) = wb.Sheets(1).[F11]
                                
                                If wb.Sheets(1).[F10] = 3 Then
                                        snDennis.Range("F" & DRow) = "*"
                                Else
                                        snDennis.Range("F" & DRow) = ""
                                End If
                                
                                snDennis.Range("G" & DRow) = wb.Sheets(1).[E13]
                                
                                snDennis.Range("H" & DRow) = wb.Sheets(2).[K17]
                                snDennis.Range("I" & DRow) = wb.Sheets(2).[J17]
                                snDennis.Range("J" & DRow) = wb.Sheets(2).[L17]
                                
                                snDennis.Range("K" & DRow) = wb.Sheets(2).[N17]
                                snDennis.Range("L" & DRow) = wb.Sheets(2).[M17]
                                snDennis.Range("M" & DRow) = wb.Sheets(2).[O17]
                                
                                snDennis.Range("N" & DRow) = wb.Sheets(2).[M32]
                                snDennis.Range("O" & DRow) = wb.Sheets(2).[N32]
                                
                                snDennis.Range("P" & DRow) = wb.Sheets(2).[M36]
                                snDennis.Range("Q" & DRow) = wb.Sheets(2).[N36]
                                wb.Close
                                Exit For
                        ElseIf CompAns = True Then
                                
                                Set wb = Workbooks.Open("")
                                If wb.Sheets(1).[B12] = "新竹科學工業園區力行六路8號" Then
                                        snDennis.Range("B" & DRow) = "TSMC12-P1"
                                ElseIf wb.Sheets(1).[B12] = "741 台南科學園區南科北路1之1號" Then
                                        snDennis.Range("B" & DRow) = "TSMC14"
                                ElseIf wb.Sheets(1).[B12] = "台中中部科學園區科雅六路1號" Then
                                        snDennis.Range("B" & DRow) = "TSMC15"
                                ElseIf wb.Sheets(1).[B12] = "新竹市科學園區研新一路9號 (3廠)" Then
                                        snDennis.Range("B" & DRow) = "TSMC3"
                                ElseIf wb.Sheets(1).[B12] = "新竹科學園區園區三路121號" Then
                                        snDennis.Range("B" & DRow) = "TSMC5"
                                ElseIf wb.Sheets(1).[B12] = "741 台南科學園區南科北路1號 (6廠)" Then
                                        snDennis.Range("B" & DRow) = "TSMC6"
                                ElseIf wb.Sheets(1).[B12] = "新竹市科學園區力行路25號 (8廠)" Then
                                        snDennis.Range("B" & DRow) = "TSMC8"
                                        
                                ElseIf Left(wb.Sheets(1).[B11], 4) = "聯華電子" Then
                                        snDennis.Range("B" & DRow) = Right(wb.Sheets(1).[B11], 8)
                                        
                                ElseIf wb.Sheets(1).[B12] = "新竹科學工業園區園區三路123號" Then
                                        snDennis.Range("B" & DRow) = "VISC1"
                                ElseIf wb.Sheets(1).[B12] = "新竹科學工業園區力行路9號" Then
                                        snDennis.Range("B" & DRow) = "VISC2"
                                ElseIf wb.Sheets(1).[B12] = "桃園縣蘆竹鄉南崁路一段336號 " Then
                                        snDennis.Range("B" & DRow) = "VISC3"
                                        
                                ElseIf wb.Sheets(1).[B12] = "新竹市香山區牛埔南路17巷20號 " Then
                                        snDennis.Range("B" & DRow) = "碩輝科技"
                                ElseIf wb.Sheets(1).[B12] = "新竹市香山區牛埔南路17巷20號 " Then
                                        snDennis.Range("B" & DRow) = "碩輝科技"
                                ElseIf wb.Sheets(1).[B12] = "新竹市高翠路327巷4弄32號 " Then
                                        snDennis.Range("B" & DRow) = "寶虹"
                                End If
                                snDennis.Range("C" & DRow) = wb.Sheets(1).[F8]
                                snDennis.Range("D" & DRow) = wb.Sheets(1).[F9]
                                snDennis.Range("E" & DRow) = wb.Sheets(1).[F11]
                                
                                If wb.Sheets(1).[F10] = 3 Then
                                        snDennis.Range("F" & DRow) = "*"
                                Else
                                        snDennis.Range("F" & DRow) = ""
                                End If
                                
                                snDennis.Range("G" & DRow) = wb.Sheets(1).[E13]
                                
                                snDennis.Range("H" & DRow) = wb.Sheets(2).[K17]
                                snDennis.Range("I" & DRow) = wb.Sheets(2).[J17]
                                snDennis.Range("J" & DRow) = wb.Sheets(2).[L17]
                                
                                snDennis.Range("K" & DRow) = wb.Sheets(2).[N17]
                                snDennis.Range("L" & DRow) = wb.Sheets(2).[M17]
                                snDennis.Range("M" & DRow) = wb.Sheets(2).[O17]
                                
                                snDennis.Range("N" & DRow) = wb.Sheets(2).[M32]
                                snDennis.Range("O" & DRow) = wb.Sheets(2).[N32]
                                
                                snDennis.Range("P" & DRow) = wb.Sheets(2).[M36]
                                snDennis.Range("Q" & DRow) = wb.Sheets(2).[N36]
                                wb.Close
                                Exit For
                        End If
                Next i
        Next c
        MsgBox "查詢完成"
        Application.DisplayAlerts = True
End Sub


