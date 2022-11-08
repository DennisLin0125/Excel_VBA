Attribute VB_Name = "ModuleAERPS"
Private Sub 一次貼圖片(ByVal Row As Integer)

        Application.ScreenUpdating = True
        
        ActiveWindow.SmallScroll Down:=-200
        
        Application.ScreenUpdating = False
        
        MsgBox "請一次選 完OFF LINE 和 熱機 共7張波形"
        
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        fd.AllowMultiSelect = True
        
        fd.Filters.Add "*.*", "*.*"
        
        fd.Title = "請選擇照片"
        
        clum = 1
        
        Dim rng As Range
        Dim sShape As Shape
        
        If fd.Show = -1 Then
        
                For Each sPath In fd.SelectedItems
                
                        Set rng = Cells(Row, clum)
                        Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, 384, 283)
                        
                        sShape.Cut
                        rng.Select
                        ActiveSheet.Paste
                        
                        If Row = 18 Then
                                Row = Row + 13
                        ElseIf Row = 35 Then
                                Row = Row + 20
                        Else
                                Row = Row + 20
                                If Row = 91 Or Row = 115 Then
                                        clum = 5
                                        If Row = 115 Then
                                                Row = 35
                                        ElseIf Row = 91 Then
                                                Row = 18
                                        End If
                                End If
                        End If
                        Set rng = Nothing
                        Set sShape = Nothing
                        Set fd = Nothing
                Next
        End If
        
        Application.ScreenUpdating = True
        Application.ScreenUpdating = False

End Sub

Sub 貼上AE資料()
Attribute 貼上AE資料.VB_ProcData.VB_Invoke_Func = "e\n14"
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
      
        myTime = Time
        
        If ActiveSheet.name <> "RMA" Then
                MsgBox "請到RMA頁面執行", vbCritical
                Exit Sub
        End If
      
        Dim oFunction As clsFunction
        Set oFunction = New clsFunction
        
        Dim RMAname$, Engineername$, SN$, MN$, MKS3L$
        
        If Range("F7").Value = "" Then Exit Sub
        
        RMAname = Range("F7")
        MN = Range("F8")
        SN = Range("F9")
        MKS3L = Trim(Range("B8"))
        
        Dim wb As Workbook
        Set wb = Workbooks(RMAname & ".xls")
        
        Dim sh1 As Worksheet
        
        If Range("F10").Value = 2 Then
                AENormal.Show
        Else
                AEW3M.Show
        End If
            
        Engineername = Range("F11")
        
        wb.Worksheets("RMA").Activate
        
        For Each sh1 In wb.Sheets
        
                sh1.Select
                
                Select Case sh1.name
                
                Case Is = "RMA"
                        
                        AERPSRMA.Show
                        
                        oFunction.AEtxt
                        
                        Range("D41").Value = Date
               
                        If Range("H9").Value = "" Then
                           Range("H9").Value = "=H8"
                           Range("H10").Value = "=H8"
                        Else
                           Range("H10").Value = "=H9"
                        End If
                        
                Case Is = "報價"
                        Worksheets("報價").Move After:=Worksheets("進出廠照片")
            
                Case Is = "報價 (2)"
                        Worksheets("報價 (2)").Move After:=Worksheets("進出廠照片")
                
                Case Is = "Source報價"
                        Worksheets("Source報價").Move After:=Worksheets("進出廠照片")
                 
                Case Is = "Failure Photo", "Failure Photo (1)", "Failure Photo (2)", "Failure Photo (3)"
                        Application.ScreenUpdating = True
                        Call oFunction.Photo(18, 21)
                        AERPSError.Show
                        ActiveWindow.Zoom = 75
                
                Case Is = "進出廠照片"
                        Application.ScreenUpdating = True
                        Call oFunction.Photo(18, 20)
                        ActiveWindow.Zoom = 75
                        
                Case Is = "Test Table RPS"
                
                        Application.ScreenUpdating = True
                        
                        If [C22] = "" Then
                                AEPower.Show
                        End If
                        
                        一次貼圖片 35
                        
                Case Is = "Test Table RPS 1"
                
                        Application.ScreenUpdating = True
                        
                        If Worksheets("Test Table RPS 1").Range("C51").Value = "" Then
                                AEPower.Show
                        End If
                        
                Case Is = "Test Table RPS 2"
                
                        Application.ScreenUpdating = True
                        
                        一次貼圖片 18
                
                End Select
                
        Next
             
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Worksheets("RMA").Activate
        
        myTime = Time - myTime
        myMin = Minute(myTime)
        mySec = Second(myTime)
    
        MsgBox "處理完成" & Chr(10) & Chr(10) & "執行時間" & myMin & "分" & mySec & "秒。", vbInformation
           
End Sub


