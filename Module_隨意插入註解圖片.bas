Attribute VB_Name = "Module�H�N���J���ѹϤ�"
Sub �H�N���J����()
        
        Application.ScreenUpdating = False
        
        Dim sh As Worksheet
        Set sh = ActiveSheet
        
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)

        With fd
                .AllowMultiSelect = True
                .InitialFileName = "D:\Users\Dlin\Desktop\�Ӥ�"
                .Filters.Add "*.*", "*.*"
                .Title = "�п�ܷӤ�"
                
                Dim rng As Range
                Dim sShape As Shape, Row%
                MsgBox "�п�ܹϧ�"
                Row = 2
                If .Show = -1 Then
                        For Each sPath In .SelectedItems
                                Set rng = Range(ActiveCell.Address)
                                With rng
                                        .ClearComments
                                        .AddComment
                                        With .Comment
                                                .Shape.Fill.UserPicture sPath
                                                .Visible = False
                                                .Shape.Width = 300
                                                .Shape.Height = 258
                                        End With
                                End With
                        Next
                End If
                
        End With
        
        Set fd = Nothing
        
End Sub

