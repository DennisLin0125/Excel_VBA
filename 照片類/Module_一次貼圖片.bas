Attribute VB_Name = "Module�@���K�Ϥ�"
Private Sub �@���K�Ϥ�()

        Application.ScreenUpdating = True
        
        ActiveWindow.SmallScroll Down:=-200
        
        Application.ScreenUpdating = False
        
        MsgBox "�Ф@���� ��OFF LINE �M ���� �@7�i�i��"
        
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        fd.AllowMultiSelect = True
        
        fd.Filters.Add "*.*", "*.*"
        
        fd.Title = "�п�ܷӤ�"
        
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
