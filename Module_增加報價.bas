Attribute VB_Name = "Module�W�[����"
Sub �W�[����()
    Application.ScreenUpdating = False
    
    Dim sPath, sFile As String
    Dim Width, Length, y_Position, x_Position As Integer
    
    Range("F11") = "Dennis"
    Range("H12") = "Yes"
    Range("F42") = "8"
    Range("B46").Interior.Color = 65535
    Range("A33") = "1. Machine cleaning." & Chr(10) & "2. According the test proccedure tested --- pass."
    
    Worksheets("�i�X�t�Ӥ�").Copy Before:=Sheets("Failure Photo")
    Worksheets("�i�X�t�Ӥ� (2)").name = "Failure Photo (����)"
    'Range("C5") = "�l �a �� �i"
    
    [A17:E17] = ""
    
    Application.ScreenUpdating = True
    
    ActiveWindow.SmallScroll Down:=-200

    Application.ScreenUpdating = False
    
   
    MsgBox "�Ф@���粒���׷Ӥ�"
    
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    fd.AllowMultiSelect = True
    
    fd.Title = "�п�ܷӤ�"
    
    Row = 18
    
    Dim rng As Range
    Dim sShape As Shape
    
    If fd.Show = -1 Then
    
        For Each sPath In fd.SelectedItems
        
            Set rng = Range("A" & Row)
            Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, 384, 283)
            
            sShape.Cut
            rng.Select
            ActiveSheet.Paste
            
            Range("A" & Row + 20 & ":D" & Row + 20).Select
    
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
            End With
            
            With Selection
                .HorizontalAlignment = xlCenter
            End With
            Selection.Merge
            
            Range("A" & Row + 20).Select
                With Selection.Font
                    .name = "Tahoma"
                    .Size = 12
                End With
           ' Range("A" & Row + 20).Value = "Bus�q�e�l�a�A�ݧ󴫷s�~�A���`��390uF (�зǬ�510��10% uF)"
            
            Row = Row + 21
            
        Next
        
    End If
    
    ActiveWindow.Zoom = 40
    
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    fd.AllowMultiSelect = True
    
    Row = 18
    
    If fd.Show = -1 Then
    
        For Each sPath In fd.SelectedItems
        
            Set rng = Range("E" & Row)
            Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, 386, 302)
            
            sShape.Cut
            rng.Select
            ActiveSheet.Paste
            
            Range("E" & Row + 20 & ":H" & Row + 20).Select
    
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
            End With
            
            With Selection
                .HorizontalAlignment = xlCenter
            End With
            Selection.Merge
            
            Range("E" & Row + 20).Select
                With Selection.Font
                    .name = "Tahoma"
                    .Size = 12
                End With
            'Range("E" & Row + 20) = "Bus�q�e�l�a�A�ݧ󴫷s�~�A���`��390uF (�зǬ�510��10% uF)"
            
            Row = Row + 21
        Next
        
    End If
    
    ActiveWindow.Zoom = 75
    
    Application.ScreenUpdating = True
    
    MsgBox "�B�z����", vbInformation
    
End Sub

