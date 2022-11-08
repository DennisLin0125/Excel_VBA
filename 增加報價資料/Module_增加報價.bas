Attribute VB_Name = "Module增加報價"
Sub 增加報價()
    Application.ScreenUpdating = False
    
    Dim sPath, sFile As String
    Dim Width, Length, y_Position, x_Position As Integer
    
    Range("F11") = "Dennis"
    Range("H12") = "Yes"
    Range("F42") = "8"
    Range("B46").Interior.Color = 65535
    Range("A33") = "1. Machine cleaning." & Chr(10) & "2. According the test proccedure tested --- pass."
    
    Worksheets("進出廠照片").Copy Before:=Sheets("Failure Photo")
    Worksheets("進出廠照片 (2)").name = "Failure Photo (報價)"
    
    [A17:E17] = ""
    
    Application.ScreenUpdating = True
    
    ActiveWindow.SmallScroll Down:=-200

    Application.ScreenUpdating = False
    
   
    MsgBox "請一次選完維修照片"
    
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    fd.AllowMultiSelect = True
    
    fd.Title = "請選擇照片"
    
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
                End Wit
            
            Row = Row + 21
        Next
        
    End If
    
    ActiveWindow.Zoom = 75
    
    Application.ScreenUpdating = True
    
    MsgBox "處理完成", vbInformation
    
End Sub

