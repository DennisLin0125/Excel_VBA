Attribute VB_Name = "Module�K�W�l�a�Ӥ�"
Sub �K�W�l�a�Ӥ�()
        ActiveWindow.Zoom = 75
        
        With Application
                .ScreenUpdating = True
                .ScreenUpdating = False
        End With
    
        'MsgBox "�Ф@���粒���׷Ӥ�"
        
        a = Range("A1:D1").Width
        
        c = Range("E1:H1").Width
        
        b = Range("A18:A37").Height
        b = b * 0.986
    
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd
                .AllowMultiSelect = True
                .InitialFileName = "D:\Users\Dlin\Desktop\�Ӥ�\" & [F9]
                .Title = "�п�ܷӤ�"
                
                Row = 18
                
                Dim rng As Range
                Dim sShape As Shape
                
                If .Show Then
                        For Each sPath In .SelectedItems
                                Set rng = Range("A" & Row)
                                Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, a, b)
                                sShape.Cut
                                rng.Select
                                ActiveSheet.Paste
                                Row = Row + 21
                        Next
                End If
        End With
        
        Set fd = Nothing
        
        With Application
                .ScreenUpdating = True
                .ScreenUpdating = False
        End With
        
        a = 0
        
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd
                .AllowMultiSelect = True
                .InitialFileName = "D:\Users\Dlin\Desktop\�Ӥ�\" & [F9]
                .Title = "�п�ܷӤ�"
                'MsgBox "�Ф@���粒  ���׫�Ӥ�"
                Row = 18
                
                If .Show Then
                        For Each sPath In .SelectedItems
                            Set rng = Range("E" & Row)
                            Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, c, b)
                            sShape.Cut
                            rng.Select
                            ActiveSheet.Paste
                            Row = Row + 21
                        Next
                End If
        End With
       Application.ScreenUpdating = True
End Sub
