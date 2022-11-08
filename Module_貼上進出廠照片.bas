Attribute VB_Name = "Module貼上進出廠照片"
Sub 貼上進出廠圖片()
        ActiveWindow.Zoom = 75
        
        With Application
                .ScreenUpdating = True
               ' ActiveWindow.SmallScroll Down:=-200
                .ScreenUpdating = False
        End With
        
        a = Range("A1:D1").Width
        
        c = Range("E1:H1").Width
        
        b = Range("A18:A37").Height
        
        b = b * 0.986
        
        MsgBox "請一次選完維修照片"
        
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd
                .AllowMultiSelect = True
                .InitialFileName = ""
            
                .Filters.Add "*.*", "*.*"
                .Title = "請選擇照片"
                
                Dim rng As Range
                Dim sShape As Shape, Row%
                Row = 18
                If .Show = -1 Then
                        For Each sPath In .SelectedItems
                                Set rng = Range("A" & Row)
                                Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, a, b)
                                sShape.Cut
                                rng.Select
                                ActiveSheet.Paste
                                Row = Row + 20
                                Set rng = Nothing
                                Set sShape = Nothing
                        Next
                End If
        End With
        
        Set fd = Nothing
        
        With Application
                .ScreenUpdating = True
                .ScreenUpdating = False
        End With
        
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
                .AllowMultiSelect = True
                .InitialFileName = ""
                .Title = "請選擇照片"
        
                Row = 18
                
                If .Show = -1 Then
                        For Each sPath In .SelectedItems
                        Set rng = Range("E" & Row)
                        Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, c, b)
                        sShape.Cut
                        rng.Select
                        ActiveSheet.Paste
                        Row = Row + 20
                        Set rng = Nothing
                        Set sShape = Nothing
                        Next
                End If
        End With
        Set fd = Nothing
        Application.ScreenUpdating = True
End Sub
