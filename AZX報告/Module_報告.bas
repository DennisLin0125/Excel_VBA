Attribute VB_Name = "ModuleA報告"
Option Explicit
Sub 波形圖(ByVal oROW As Integer, ByVal colum As Integer)
        Dim fd As FileDialog, myWidth%, myHeight%, sPath, iTop
        Set fd = Application.FileDialog(msoFileDialogFilePicker)

        With fd
                .AllowMultiSelect = True
                .Title = "請選擇照片"
                .ButtonName = "就是你了!!!!"
        
                myWidth = 395
                myHeight = 295
                
                Dim rng As Range
                Dim sShape As Shape
                
                If .Show Then
                        For Each sPath In .SelectedItems
                                Set rng = Cells(oROW, colum)
                                Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, myWidth, myHeight)
                                sShape.Cut
                                rng.Select
                                ActiveSheet.Paste
                                colum = colum + 4
                                Set rng = Nothing
                                Set sShape = Nothing
                        Next
                End If
        End With
        Set fd = Nothing
End Sub

Sub AZX報告()
        
        Dim myTime As Date
        myTime = Time

        If ActiveSheet.name <> "RMA" Then
                MsgBox "請到RMA頁面執行", vbCritical
                Exit Sub
        End If
        
        If [F10] = 2 Then
                AzxNormal.Show
                [H12] = "Yes"
                [F42] = "6"
                [B41] = "0.5"
                [D41] = Date
        Else
                AZXW3M.Show
        End If
        
        AZXRMA.Show
        
        Dim cus$
        cus = [B12]
        
        If [H9] = "" Then
                [H9] = "=H8"
                [H10] = "=H8"
        Else
                [H10] = "=H9"
        End If
        
        Dim sh As Worksheet
        
        For Each sh In ActiveWorkbook.Worksheets
                sh.Select
                Select Case sh.name
                Case Is = "Test Table Tuner (-020,-023)", "Test Table Tuner (-036,-039)", "Test Table Tuner (-014)", "Test Table Tuner-014"
                
                        If cus = "新竹市科學園區力行路25號 (8廠)" Then
                                AZX表單.Show
                                Dim IdleV$, IdleI$, AfterSelect2$
                                IdleV = [K36]
                                IdleI = [L36]
                        
                                AfterSelect2 = [P36]
                                
                                MsgBox "T8選擇2張史密斯圖"
                                Call 波形圖(37, 1)
                                
                                Worksheets("進出廠照片").Copy Before:=Sheets("Failure Photo")
                                Worksheets("進出廠照片 (2)").name = "Failure Photo(客戶)"
                                Worksheets("Failure Photo(客戶)").Copy Before:=Sheets("Failure Photo")
                                Worksheets("Failure Photo(客戶) (2)").name = "Failure Photo(客戶-2)"
                                
                                Dim MystrT8(4) As String
                                MystrT8(0) = "Customer request"
                                MystrT8(1) = "1. The input impedance of phase mag board: 0.1 ohms"
                                MystrT8(2) = "2. Idle V/I = " & IdleV & "mV/" & IdleI & "mV"
                                MystrT8(3) = "3. Chuck On V/I = 2.45V/" & AfterSelect2 & "V "
                                MystrT8(4) = "4. Chuck On V/I(Max) = 2.45V/" & AfterSelect2 & "V "
                                
                                Worksheets("RMA").[E33] = Join(MystrT8, vbCrLf)
        
                                With Worksheets("RMA").[E33]
                                        .HorizontalAlignment = xlGeneral
                                        .VerticalAlignment = xlTop
                                End With
                                
                                Worksheets("Failure Photo(客戶)").Activate
                                [A17:E17] = ""
                                MsgBox "選擇給客戶圖片(各一張就好)"
                                貼上損壞照片
                                
                                
                                Worksheets("Failure Photo(客戶-2)").Activate
                                [A17:E17] = ""
                                With Range("A36:H36").Borders
                                        .LineStyle = xlContinuous
                                End With
                                
                                With Range("A58:D58").Borders
                                        .LineStyle = xlContinuous
                                End With
                                
                                With Range("A36:H36")
                                        .Merge
                                        .HorizontalAlignment = xlCenter
                                        .VerticalAlignment = xlCenter
                                End With
                                
                                With Range("A58:D58")
                                        .Merge
                                        .HorizontalAlignment = xlCenter
                                        .VerticalAlignment = xlCenter
                                End With
                                
                                [A36] = "Monitor ESC voltage out"
                                [A58] = "MN"
                                
                                With [A36].Font
                                        .name = "Tahoma"
                                        .Size = 12
                                End With
                                
                                With [A58].Font
                                        .name = "Tahoma"
                                        .Size = 12
                                End With
                        Else
                                AZX表單.Show
                                MsgBox "選擇1張史密斯圖"
                                Call 波形圖(36, 2)
                        End If
                        
                Case Is = "Test Table Tuner-020-023"
                
                        If cus = "741 台南科學園區南科北路1號 (6廠)" Then
                                AZX表單.Show
                                MsgBox "T6選擇1張史密斯圖"
                                Call 波形圖(41, 5)
                        End If
                
                Case Is = "Test Table Tuner (-037)", "Test Table Tuner (-043)", "Test Table Tuner", "Test Table Tuner (-039)"
                        AZX表單.Show
                        MsgBox "選擇1張史密斯圖"
                        Call 波形圖(36, 2)
                
                Case Is = "Failure Photo"
                        MsgBox "請選 " & ActiveSheet.name & " (可複選)"
                        貼上損壞照片
                        AZXError.Show
                        
                Case Is = "進出廠照片"
                        MsgBox "請選 " & ActiveSheet.name & " (可複選)"
                        貼上進出廠圖片
         
                End Select
        Next sh
        
        Dim myMin%, mySec%
        myTime = Time - myTime
        myMin = Minute(myTime)
        mySec = Second(myTime)
        
        Worksheets("RMA").Select
        
        MsgBox "處理完成" & Chr(10) & Chr(10) & "執行時間" & myMin & "分" & mySec & "秒。", vbInformation
        
End Sub
