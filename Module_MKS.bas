Attribute VB_Name = "ModuleMKS"
Sub �K�WLOG���()
Attribute �K�WLOG���.VB_ProcData.VB_Invoke_Func = "l\n14"
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
      
        myTime = Time
        
        If ActiveSheet.name <> "RMA" Then
                MsgBox "�Ш�RMA��������", vbCritical
                Exit Sub
        End If
      
        Dim oFunction As clsFunction
        Set oFunction = New clsFunction
        
        Dim RMAname$, Engineername$, SN$, MN$, MKS3L$
        
        RMAname = [F7]
        MN = Range("F8")
        SN = Range("F9")
        MKS3L = Trim(Range("B8"))
    
        Workbooks(RMAname & ".xls").Activate
        
        Dim sh1 As Worksheet
        
        If Range("F10").Value = 2 Then
                Normal.Show
        Else
                W3M.Show
        End If
            
        Engineername = Range("F11")
        
        Workbooks(RMAname & ".xls").Worksheets("RMA").Activate
        
        For Each sh1 In Workbooks(RMAname & ".xls").Sheets
        
                sh1.Select
                
                Select Case sh1.name
                
                Case Is = "RMA"
                
                        If Range("H9").Value = "" Then
                           Range("H9").Value = "=H8"
                           Range("H10").Value = "=H8"
                        Else
                           Range("H10").Value = "=H9"
                        End If
                        
                        Range("D41").Value = Date
                
                Case Is = "Test Table MKS (3L)"
                        Application.ScreenUpdating = True
                        Call oFunction.���J�I���q��(40, 4)
                        ActiveWindow.Zoom = 75
                        
                Case Is = "Test Table MKS (2L)"
                        Application.ScreenUpdating = True
                        Photo2L
                        ActiveWindow.Zoom = 75
                        
                Case Is = "Test Table MKS (8L)", "Test Table MKS (15L)", "Test Table MKS (6L)", "Test Table MKS (22L)"
                        Application.ScreenUpdating = True
                        Call oFunction.���J�I���q��(43, 4)
                        ActiveWindow.Zoom = 75
               
                Case Is = "����"
                        Worksheets("����").Move After:=Worksheets("�i�X�t�Ӥ�")
            
                Case Is = "���� (2)"
                        Worksheets("���� (2)").Move After:=Worksheets("�i�X�t�Ӥ�")
                
                Case Is = "Source����"
                        Worksheets("Source����").Move After:=Worksheets("�i�X�t�Ӥ�")
                 
                Case Is = "Failure Photo", "Failure Photo (2)", "Failure Photo (3)"
                        Application.ScreenUpdating = True
                        Call oFunction.Photo(18, 21)
                        Error.Show
                        ActiveWindow.Zoom = 75
                
                Case Is = "�i�X�t�Ӥ�"
                        Application.ScreenUpdating = True
                        Call oFunction.Photo(18, 20)
                        ActiveWindow.Zoom = 75
                
                Case Is = "Nozzle"
                        Application.ScreenUpdating = True
                        Call oFunction.���JNozzle�Ϥ�
                        
                Case Is = "Test Table MKS"
                        Application.ScreenUpdating = True
                        
                        Workbooks(RMAname & ".xls").Worksheets("RMA").Activate
                        Range("A1").Select
                        
                        If Worksheets("Test Table MKS").Range("C21").Value = "" Then
                            Power.Show
                        End If
                        
                        Call oFunction.LOGdata(RMAname, MN, SN, Engineername)
                
                Case Is = "Log"
                
                        Call oFunction.�}��LogData(RMAname)
                
                        For i = 2 To 10
                                If Range("A" & i) <> "" Then
                                        Dim StrTemp$, RunTime$
                                        StrTemp = Mid(Range("A" & i), 1, InStr(Range("A" & i), ":") - 1)
                                        RunTime = Mid(StrTemp, InStrRev(StrTemp, """") + 1)
                                        Range("A1").Select
                                        Exit For
                                End If
                        Next i

                        Workbooks(RMAname & ".xls").Worksheets("RMA").[E33] = "1. PA date code: " & Chr(10) & _
                                                                                                                          "2. Run hour: " & RunTime & " hours" & Chr(10) & _
                                                                                                                          "3. AC Input Current: " & "      " & "A"
                End Select
                
        Next
             
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Worksheets("RMA").Activate
        
        s = MsgBox("�n�ƻs�Ӥ���P�Ѷ�?", vbYesNo)
        
        If s = vbYes Then
                Call oFunction.CopyPhoto(RMAname, SN, Engineername)
        End If
        
        myTime = Time - myTime
        myMin = Minute(myTime)
        mySec = Second(myTime)
    
        MsgBox "�B�z����" & Chr(10) & Chr(10) & "����ɶ�" & myMin & "��" & mySec & "��C", vbInformation
           
End Sub
Sub Photo2L()
        ActiveWindow.Zoom = 75

        Application.ScreenUpdating = True
        
        Application.ScreenUpdating = False
    
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        Row = 40
        index = 4
        With fd
                .AllowMultiSelect = True
                .Title = "�п�ܷӤ�"
                .ButtonName = "�N�O�A�F!!!!"
                MsgBox "�п���I���q�� (�i�ƿ�)"
        
                colum = 1
                
                a = Range("A1:D1").Width
                c = Range("E1:H1").Width
                b = Range("A40:A61").Height
                
                b = b * 0.99
                
                temp = Row
                
                Dim rng As Range
                Dim sShape As Shape
                k = 0
                If .Show Then
                        For Each sPath In .SelectedItems
                                Set rng = Cells(temp, colum)
                                
                                If k = 0 Then
                                        Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, a, b)
                                        k = k + 1
                                Else
                                        Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, c, b)
                                End If
                                
                                sShape.Cut
                                rng.Select
                                ActiveSheet.Paste
                                colum = colum + index
                                Set rng = Nothing
                                Set sShape = Nothing
                        Next
                End If
        End With
    
        Set fd = Nothing
        
        Application.ScreenUpdating = True
        
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd
                .AllowMultiSelect = True
                .ButtonName = "�N�O�A�F!!!!"
                .Title = "�п�ܷӤ�"
                MsgBox "�п���|�Ӥ� (�i�ƿ�)"
                
                colum = 1
                Row = Row + 21
                
                k = 0
                If .Show Then
                        For Each sPath In .SelectedItems
                                Set rng = Cells(Row, colum)
                                
                                If k = 0 Then
                                        Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, a, b)
                                        k = k + 1
                                Else
                                        Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, c, b)
                                End If
                                
                                sShape.Cut
                                rng.Select
                                ActiveSheet.Paste
                                colum = colum + index
                                Set rng = Nothing
                                Set sShape = Nothing
                        Next
                End If
        End With
        Set fd = Nothing
End Sub


