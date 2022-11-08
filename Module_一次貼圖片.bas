Attribute VB_Name = "Module一次貼圖片"
Private Sub 一次貼圖片()

        Application.ScreenUpdating = True
        
        ActiveWindow.SmallScroll Down:=-200
        
        Application.ScreenUpdating = False
        
        MsgBox "請一次選 完OFF LINE 和 熱機 共7張波形"
        
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        fd.AllowMultiSelect = True
        
        fd.Filters.Add "*.*", "*.*"
        
        fd.Title = "請選擇照片"
        
        Row = 18
        
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
                        Else
                                Row = Row + 20
                                If Row = 91 Then
                                        clum = 5
                                        Row = 18
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
