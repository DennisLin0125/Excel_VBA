Attribute VB_Name = "Module�K�W�i�X�t�Ӥ�"
Sub �K�W�i�X�t�Ϥ�()
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
        
        MsgBox "�Ф@���粒���׷Ӥ�"
        
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd
                .AllowMultiSelect = True
                .InitialFileName = "D:\Users\Dlin\Desktop\�Ӥ�\" & [F9]
                '.InitialFileName = "P:\Service\Repair Picture\RF Product\MKS RPS\10441"
                .Filters.Add "*.*", "*.*"
                .Title = "�п�ܷӤ�"
                
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
                .InitialFileName = "D:\Users\Dlin\Desktop\�Ӥ�\" & [F9]
                .Title = "�п�ܷӤ�"
                '.InitialFileName = "P:\Service\Repair Picture\RF Product\MKS RPS\10441"
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
