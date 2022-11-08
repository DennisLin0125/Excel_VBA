Attribute VB_Name = "ModuleBOM"
Sub �إ߾���BOM()
        
        Application.ScreenUpdating = False
        
        machine = InputBox("�п�J����")
        
        Dim sh As Worksheet
        Set sh = ActiveSheet
        
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        [A1] = "�s�W�Ƹ�"
        [B1] = "�s��y�z"
        [c1] = "�s�W���"
        [D1] = "�ϥξ���"
        [E1] = "�Ϥ�"
        [F1] = "�w�s��"
        
        Range("A1:F1").Interior.Color = vbYellow
        
        With fd
                .AllowMultiSelect = True
                .InitialFileName = "D:\Users\Dlin\Desktop\�Ӥ�"
                .Filters.Add "*.*", "*.*"
                .Title = "�п�ܷӤ�"
                
                Dim rng As Range
                Dim sShape As Shape, Row%
                MsgBox "�п��PCB�Ӥ�"
                Row = 2
                If .Show = -1 Then
                        For Each sPath In .SelectedItems
                                
                                Set rng = Range("E" & Row)
                                Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, 384, 283)
                                
                                With rng
                                       '.ClearComments
                                        .AddComment
                                        With .Comment
                                                .Shape.Fill.UserPicture sPath
                                                .Visible = False
                                                .Shape.Width = 250
                                                .Shape.Height = 208
                                        End With
                                End With
                                
                                sShape.Select
                                With Selection
                                        .ShapeRange.Width = 75
                                        .ShapeRange.Height = 50
                                End With

                                rng.ColumnWidth = 11.9
                                rng.RowHeight = 50
                                
                                sShape.Cut
                                rng.Select
                                ActiveSheet.Paste
                                Range("C" & Row) = Date
                                Range("F" & Row) = "�w��"
                                Range("D" & Row) = machine
                                Row = Row + 1
                        Next
                End If
                
        End With
        
        Set fd = Nothing
        
        oROW = Range("F" & Rows.Count).End(xlUp).Row
        
        With Range("A1:F" & oROW)
                .HorizontalAlignment = xlGeneral
                .HorizontalAlignment = xlCenter
                With .Borders
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                End With
                
                With .Font
                        .name = "Tahoma"
                        .Size = 10
                End With
                
        End With
        
        Application.ScreenUpdating = True
End Sub
