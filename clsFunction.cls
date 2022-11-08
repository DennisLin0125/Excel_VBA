VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub AEtxt()
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd
                .Title = "請選擇LOG筆記本檔案"
                .ButtonName = "就是你了!!!"
                .Filters.Add "筆記本檔案", "*.*"
                MsgBox "請選擇 出廠 LOG筆記本檔案"
                
                If .Show Then
                        Dim temp(14) As String
                        Open .SelectedItems(1) For Input As #1
                        Do Until EOF(1)
                                  Line Input #1, LineFromFile
                                  temp(i) = LineFromFile
                                  i = i + 1
                        Loop
                        Close #1
                        [E19] = Join(temp, Chr(10))
                End If
        End With
        Set fd = Nothing
End Sub
Function MKSAllTestList(ByVal RMAname As String) As Boolean
      
        MKSAllTestList = False
          
        Dim PathTest$
        PathTest = "P:\Service\Product\MKS RPS\all test list.xls"
        
        Dim wbtest As Workbook
        Set wbtest = Workbooks.Open(PathTest, UpdateLinks:=0)
      
        Dim sh As Worksheet
        For Each sh In wbtest.Worksheets
                sh.Activate
                Dim RMA As Range
                Set RMA = Cells.Find(What:=RMAname, LookAt:=xlWhole)
          
                If Not RMA Is Nothing Then
                        MKSAllTestList = True
                        wbtest.Close False
                        Exit For
                End If
        Next

        Set wbtest = Nothing
        Set RMA = Nothing
        
End Function

Sub CopyPhoto(ByVal RMA As String, ByVal SN As String, ByVal Engineername As String)
        
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        If fs.FolderExists("P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN) Then
                fs.DeleteFolder "P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN
        End If
        
        With fs
                .CreateFolder ("P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN)
                .CreateFolder ("P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN & "\" & RMA)
                .CreateFolder ("P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\LOG\")
                .CreateFolder ("P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\進出廠照片")
                .CreateFolder ("P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\故障圖片")
                .CreateFolder ("P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\故障圖片\維修前")
                .CreateFolder ("P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\故障圖片\維修後")
                .CreateFolder ("P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\進出廠照片\維修前")
                .CreateFolder ("P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\進出廠照片\維修後")
        End With
        
        Dim RpBeforPath$, RpAfterPath$, MachAfterPath$, MachBeforPath$, SourcePath$, LogPath$
        
        RpBeforPath = "P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\故障圖片\維修前\"
        RpAfterPath = "P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\故障圖片\維修後\"
        
        MachAfterPath = "P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\進出廠照片\維修後\"
        MachBeforPath = "P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\進出廠照片\維修前\"
        
        SourcePath = "P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN & "\" & RMA
        LogPath = "P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\LOG\"
         
        '維修前外觀照片
        Dim fdMachBefor As FileDialog
        Set fdMachBefor = Application.FileDialog(msoFileDialogFilePicker)
        
        With fdMachBefor
                .AllowMultiSelect = True
                .Title = "請選擇 進出廠照片的 (維修前) 照片"
                .ButtonName = "就是你了!!!!"
                MsgBox "接下來，將把照片和資料做分類" & Chr(10) & Chr(10) & "請選擇 進出廠照片的 (維修前) 照片"
                If .Show Then
                        For Each sPath In .SelectedItems
                                fs.CopyFile sPath, MachBeforPath
                        Next
                End If
        End With
        
        '維修後外觀照片
        Dim fdMachAfter As FileDialog
        Set fdMachAfter = Application.FileDialog(msoFileDialogFilePicker)
        
        With fdMachAfter
                .AllowMultiSelect = True
                .Title = "請選擇 進出廠照片的 (維修後) 照片"
                .ButtonName = "就是你了!!!!"
                MsgBox "請選擇 進出廠照片的(維修後)照片 (可複選)"
                If .Show Then
                        For Each sPath In .SelectedItems
                                fs.CopyFile sPath, MachAfterPath
                        Next
                End If
        End With
        
        '維修前故障照片
        Dim fdRpBefor As FileDialog
        Set fdRpBefor = Application.FileDialog(msoFileDialogFilePicker)
        
        With fdRpBefor
                .AllowMultiSelect = True
                .Title = "請選擇 故障照片的 (維修前) 照片"
                .ButtonName = "就是你了!!!!"
                MsgBox "請選擇 故障照片的(維修前)照片 (可複選)"
                If .Show Then
                        For Each sPath In .SelectedItems
                                fs.CopyFile sPath, RpBeforPath
                        Next
                End If
        End With
        
        '維修前故障照片
        Dim fdRpAfter As FileDialog
        Set fdRpAfter = Application.FileDialog(msoFileDialogFilePicker)
        
        With fdRpAfter
                .AllowMultiSelect = True
                .ButtonName = "就是你了!!!!"
                .Title = "請選擇 故障照片的 (維修後) 照片"
                MsgBox "請選擇 故障照片的(維修後)照片 (可複選)"
                If .Show Then
                        For Each sPath In .SelectedItems
                                fs.CopyFile sPath, RpAfterPath
                        Next
                End If
        End With
        
        'LOG 點火電壓
        Dim fdLog As FileDialog
        Set fdLog = Application.FileDialog(msoFileDialogFilePicker)
        
        With fdLog
                .AllowMultiSelect = True
                .ButtonName = "就是你了!!!!"
                .Title = "請選擇組裝人員照片以RMA命名的資料夾"
                 MsgBox "請選擇 點火電壓 excel檔 BurnIN Ignition ROR LOG 資料"
                If .Show Then
                        For Each sPath In .SelectedItems
                                fs.CopyFile sPath, LogPath
                        Next
                End If
        End With
        
        '腔體組裝照片
        Dim fdSource As FileDialog
        Set fdSource = Application.FileDialog(msoFileDialogFolderPicker)
        
        With fdSource
                .InitialFileName = "P:\Service\Repair Picture\RF Product\拆裝人員暫放Source"
                .Title = "請選擇組裝人員照片以RMA命名的資料夾"
                .ButtonName = "就是你了!!!!"
                MsgBox "請選擇 組裝人員照片以 RMA 命名的資料夾" & Chr(10) & Chr(10) & "如果腔體未拆請按 取消"
                
                If .Show Then
                        fs.CopyFolder .SelectedItems(1), SourcePath
                End If
        End With
        
        Set fs = Nothing
        Set fdLog = Nothing
        Set fdSource = Nothing
        Set fdRpAfter = Nothing
        Set fdRpBefor = Nothing
        Set fdMachAfter = Nothing
        Set fdMachBefor = Nothing
        
        '移動
        Dim oTimes%
        oTimes = 0
        
        a = MsgBox("是否要複製照片至 P槽?", vbOKCancel)
        
        If a = vbOK Then
                Source = "P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN
                path = "P:\Service\Repair Picture\RF Product\MKS RPS\" & SN
                Set fs = CreateObject("scripting.FileSystemObject")
                fs.CopyFolder Source, path
                oTimes = 1
                Set fs = Nothing
        End If
        
        If oTimes Then
                Source = "P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN
                Set fs = CreateObject("scripting.FileSystemObject")
                fs.DeleteFolder Source
                Set myTex = fs.CreateTextFile("P:\Service\技術討論專區\Engineer\Dennis\MKS\" & RMA & "_" & Engineername, True)
                myTex.Write ("RMA :" & RMA)
                myTex.WriteLine
                myTex.Write ("SN :" & SN)
                myTex.Close
                Set myTex = Nothing
                Set fs = Nothing
        Else
                Source = "P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN
                Set fs = CreateObject("scripting.FileSystemObject")
                fs.DeleteFolder Source
                Set myTex = fs.CreateTextFile("P:\Service\技術討論專區\Engineer\Dennis\MKS\" & RMA & "_" & Engineername & "_照片沒傳.txt", True)
                myTex.Write ("RMA :" & RMA)
                myTex.WriteLine
                myTex.Write ("SN :" & SN)
                myTex.Close
                Set myTex = Nothing
                Set fs = Nothing
        End If
End Sub

Function MKS3LChecklist(ByVal Path1 As String, ByVal RMAname As String) As String
        Dim wb As Workbook
        Set wb = Workbooks.Open(Path1)
        
        Dim source1 As Range
        Set source1 = wb.Worksheets("清單").[C27:F30]
        Dim spath1 As Range
        Set spath1 = Workbooks(RMAname & ".xls").Worksheets("RMA").[K12]
        
        Dim source2 As Range
        Set source2 = wb.Worksheets("清單").[C33:F36]
        Dim spath2 As Range
        Set spath2 = Workbooks(RMAname & ".xls").Worksheets("RMA").[Q12]
        
        Dim DateCord$
        DateCord = Right(wb.Worksheets("清單").[G22], 4)
        
        If wb.Worksheets("清單").[C33] = "" Then
            
                spath1.Resize(source1.Rows.Count, source1.Columns.Count) = source1.Value
                
                 With [K12:N15]
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                End With
            
                wb.Close False
            
                Set fs = Nothing
                Set wb = Nothing
                Set spath1 = Nothing
                Set source1 = Nothing
                Set source2 = Nothing
                Set spath2 = Nothing
        Else
            
                spath1.Resize(source1.Rows.Count, source1.Columns.Count) = source1.Value
                spath2.Resize(source2.Rows.Count, source2.Columns.Count) = source2.Value
            
                wb.Close False
                
                With [K12:T15]
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                End With
            
                Set fs = Nothing
                Set wb = Nothing
                Set spath1 = Nothing
                Set source1 = Nothing
                Set source2 = Nothing
                Set spath2 = Nothing
            
        End If
        MKS3LChecklist = DateCord
End Function

Function MKS3LPhotoCheck(ByVal MKS3L As String, ByVal RMAname As String) As Boolean
      
        MKS3LPhotoCheck = False
      
        If MKS3L = "MKS ASTRONi 3L REACTIVE GAS GENERATOR" Then
          
                Dim Path3L$
                Path3L = "P:\Service\技術討論專區\Engineer\4. Jacky\MKS 3L.xlsx"
                
                Dim wb3L As Workbook
                Set wb3L = Workbooks.Open(Path3L, UpdateLinks:=0)
              
                For Each sh In wb3L.Worksheets
                        sh.Select
                        Dim RMA3L As Range
                        Set RMA3L = Cells.Find(What:=RMAname, LookAt:=xlWhole)
                  
                        If Not RMA3L Is Nothing Then
                                MKS3LPhotoCheck = True
                                wb3L.Close False
                                Set wb3L = Nothing
                                Set RMA3L = Nothing
                                Exit For
                        End If
                Next
        Else
                MKS3LPhotoCheck = True
        End If
End Function
    
Sub 插入點火電壓(ByVal Row As Integer, ByVal index As Integer)
        
        ActiveWindow.Zoom = 75

        Application.ScreenUpdating = True
        
        Application.ScreenUpdating = False
    
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd
                .AllowMultiSelect = True
                .Title = "請選擇照片"
                .ButtonName = "就是你了!!!!"
                MsgBox "請選擇點火電壓 (可複選)"
        
                colum = 1
                
                a = Range("A1:D1").Width
                c = Range("E1:H1").Width
                b = Range("A" & Row & ":A" & Row + 20).Height
                
                b = b * 0.988
                
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
                .ButtonName = "就是你了!!!!"
                .Title = "請選擇照片"
                MsgBox "請選測漏照片 (可複選)"
                
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
Sub 插入Nozzle圖片()

        Application.ScreenUpdating = True
        
        Application.ScreenUpdating = False
        
        a = Range("A1:D1").Width
        b = Range("A18:A37").Height
        b = b * 0.986
        
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd
                .Title = "請選擇照片"
                .ButtonName = "就是你了!!!"
                MsgBox "請選擇 Nozzle的照片"
                myRow = 18
                
                Dim rng As Range
                Dim sShape As Shape
                
                If .Show Then
                        Set rng = Range("A" & myRow)
                        Set sShape = ActiveSheet.Shapes.AddPicture(.SelectedItems(1), msoFalse, msoCTrue, rng.Left, iTop, a, b)
                        sShape.Cut
                        rng.Select
                        ActiveSheet.Paste
                        Set rng = Nothing
                        Set sShape = Nothing
                End If
        End With
    
        Set fd = Nothing

        With [A38:D38]
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
        End With
       
        With [A38:D38].Borders
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
        End With
        
        [A38:D38].Merge
        [A38] = "出廠點火頭"
        With [A38].Font
                .name = "Tahoma"
                .Size = 12
        End With

End Sub
Sub Photo(ByVal Row As Integer, ByVal index As Integer)
        ActiveWindow.Zoom = 75
        Application.ScreenUpdating = False
        
        tempName = Mid(ActiveSheet.name, 1, 14)
        
    
        If ActiveSheet.name = "Failure Photo" Then
                MsgBox "請選 (Failure Photo) 頁面 維修前的照片 (可複選)"
            
        ElseIf ActiveSheet.name = "Failure Photo (2)" Then
                MsgBox "請選 (Failure Photo (2) 頁面 維修前的照片 (可複選)"
             
        ElseIf ActiveSheet.name = "Failure Photo (3)" Then
                MsgBox "請選 (Failure Photo (3) 頁面 維修前的照片 (可複選)"
            
        ElseIf ActiveSheet.name = "進出廠照片" Then
                MsgBox "請選 (進出廠照片) 頁面 進廠照片 (可複選)"
            
        ElseIf tempName = "Test Table MKS" Then
                MsgBox "請選 點火電壓 照片 (可複選)"
        End If
        
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd
                .AllowMultiSelect = True
                .Title = "請選擇照片"
                .ButtonName = "就是你了!!!!"
                Dim rng As Range
                Dim sShape As Shape
                
                a = Range("A1:D1").Width
        
                c = Range("E1:H1").Width
        
                b = Range("A18:A37").Height
                b = b * 0.986
                temp = Row
                
                If .Show Then
                        For Each sPath In .SelectedItems
                                Set rng = Range("A" & temp)
                                Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, a, b)
                                sShape.Cut
                                rng.Select
                                ActiveSheet.Paste
                                temp = temp + index
                                Set rng = Nothing
                                Set sShape = Nothing
                        Next
                End If
        End With
    
        Set fd = Nothing
        
        Application.ScreenUpdating = True
   
    
        If ActiveSheet.name = "Failure Photo" Then
                MsgBox "請選擇 (Failure Photo) 頁面 維修後照片 (可複選)"
            
        ElseIf ActiveSheet.name = "Failure Photo (2)" Then
                MsgBox "請選擇 (Failure Photo (2) 頁面 維修後照片 (可複選)"
            
        ElseIf ActiveSheet.name = "Failure Photo (3)" Then
                MsgBox "請選擇 (Failure Photo (3) 頁面 維修後照片 (可複選)"
            
        ElseIf ActiveSheet.name = "進出廠照片" Then
                MsgBox "請選擇 (進出廠照片) 頁面 出廠照片 (可複選)"
        End If
        
    
        Application.ScreenUpdating = False
        
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd
                .AllowMultiSelect = True
                .ButtonName = "就是你了!!!!"
                
                If .Show Then
                        For Each sPath In .SelectedItems
                                Set rng = Range("E" & Row)
                                Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, c, b)
                                sShape.Cut
                                rng.Select
                                ActiveSheet.Paste
                                Row = Row + index
                                Set rng = Nothing
                                Set sShape = Nothing
                        Next
                End If
        End With
    
        Set fd = Nothing
        
        Application.ScreenUpdating = True
        
End Sub
Sub 開啟LogData(ByRef y)
    
        Application.ScreenUpdating = False
        
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd
                .Title = "請選擇LOG筆記本檔案"
                .ButtonName = "就是你了!!!!"
                .Filters.Add "筆記本檔案", "*.txt"
                MsgBox "請選擇LOG筆記本檔案"
                
                If .Show Then
                        Workbooks.OpenText (.SelectedItems(1))
                        Dim wb As Workbook
                        Set wb = ActiveWorkbook
                        
                        Dim rng As Range
                        Set rng = wb.ActiveSheet.Columns("A:K")
                        
                        Dim source1 As Range
                        Set source1 = Workbooks(y & ".xls").Sheets("Log").Columns("A:K")
                        
                        rng.Copy source1
                        
                        wb.Close False
                        Set wb = Nothing
                        Set rng = Nothing
                        Set source1 = Nothing
                End If
        End With
        
        Set fd = Nothing
    
End Sub
Sub LOGdata(ByVal RMA As String, ByVal MN As String, ByVal SN As String, ByVal Engineername As String)
      
        Application.ScreenUpdating = False
        
        tempRMA = RMA & ".xls"
        
        Dim BurnIn As Range
        Set BurnIn = Workbooks(tempRMA).Worksheets("Auto burn in").Columns("A:N")
        
        Dim ignition As Range
        Set ignition = Workbooks(tempRMA).Worksheets("Auto ignition").Columns("A:N")
        
        Dim ROR As Range
        Set ROR = Workbooks(tempRMA).Worksheets("Ror").Columns("A:N")

        'Auto burn IN
        Dim fd1 As FileDialog
        Set fd1 = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd1
                .Title = "請選Auto burn IN資料"
                .ButtonName = "就是你了!!!"
                MsgBox "請選Auto burn IN資料"
        End With
        
        If fd1.Show Then
                Dim wb1 As Workbook
                Set wb1 = Workbooks.Open(fd1.SelectedItems(1))
                With wb1.ActiveSheet
                        .[D5] = "RMA NO:" & RMA
                        .[A7] = "M/N:" & MN
                        .[B7] = "S/N:" & SN
                        .[C7] = "Engineer :" & Engineername
                        .Columns("A:N").Copy BurnIn
                End With
                Workbooks(tempRMA).Worksheets("Auto burn in").Activate
                [A1].Select
                wb1.Close False
                Set wb1 = Nothing
                Set fd1 = Nothing
                Set BurnIn = Nothing
        End If
'=====================================================
        'Auto ignition
        Dim fd2 As FileDialog
        Set fd2 = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd2
                .Title = "請選Auto ignition資料"
                .ButtonName = "就是你了!!!"
                MsgBox "請選Auto ignition資料"
        End With

         If fd2.Show Then
                Dim wb2 As Workbook
                Set wb2 = Workbooks.Open(fd2.SelectedItems(1))
                With wb2.ActiveSheet
                        .[D5] = "RMA NO:" & RMA
                        .[A7] = "M/N:" & MN
                        .[B7] = "S/N:" & SN
                        .[C7] = "Engineer :" & Engineername
                        .[D11] = "時間間隔"
                        LF = .Range("A12").End(xlDown).Row
                        For i = 12 To LF
                                .Range("D" & i).Value = "=TIME(VALUE(MID(A" & i & ",4,2)),VALUE(MID(A" & i & ",7,2)),VALUE(MID(A" & i & ",10,2)))"
                        Next i
            
                        For j = 13 To LF
                                .Range("E" & j).Value = .Range("D" & j).Value - .Range("D" & j - 1).Value
                        Next j
                        
                        .Range("E13:E" & LF).NumberFormatLocal = "h:mm:ss"
                        .Range("E13:E" & LF).Copy
                        .[D13].PasteSpecial xlPasteValuesAndNumberFormats
                
                        Application.CutCopyMode = False
            
                        With .Range("D13:D" & LF)
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                        End With
                        
                        .Range("E13:E" & LF).ClearContents
                        .[D12].ClearContents
                        .Columns("A:N").Copy ignition
                End With
        
                Workbooks(tempRMA).Worksheets("Auto ignition").Activate
                Range("D13:D" & LF).Select
                ActiveWindow.SmallScroll Down:=-87
                ActiveSheet.Shapes.AddChart.Select
                ActiveChart.SetSourceData Source:=Range("'Auto ignition'!$D$13:$D$" & LF)
                ActiveChart.ChartType = xlLine
                
                Range("A1").Select
            
                wb2.Close False
                Set wb2 = Nothing
                Set fd2 = Nothing
                Set ignition = Nothing
        End If
'===================================================================
        'ROR
        
        Dim oSec As Variant
        oSec = Split("Time:,0S,30S,60S,90S,120S,150S,180S,210S,240S,270S,300S", ",")
        
        Dim oTorr As Variant
        oTorr = Split("(m)Torr:,-1,2,4,6,8,10,13,16,19,21,23", ",")
        
        Dim fd3 As FileDialog
        Set fd3 = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd3
                .Title = "請選ROR資料"
                .ButtonName = "就是你了!!!"
                MsgBox "請選ROR資料"
        End With
        
        If fd3.Show = -1 Then
                Dim wb3 As Workbook
                Set wb3 = Workbooks.Open(fd3.SelectedItems(1))
                
                With wb3.ActiveSheet
                        .[D5] = "RMA NO:" & RMA
                        .[A7] = "M/N:" & MN
                        .[B7] = "S/N:" & SN
                        .[C7] = "Engineer :" & Engineername
                        .[D11].Resize(UBound(oSec)) = Application.WorksheetFunction.Transpose(oSec)
                        .[E11].Resize(UBound(oTorr)) = Application.WorksheetFunction.Transpose(oTorr)
        
                        With [D11:E22]
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                        End With
                        .Columns("A:N").Copy ROR
                End With
                
                Workbooks(tempRMA).Worksheets("Ror").Activate
                Range("A1").Select
                
                wb3.Close False
                Set wb3 = Nothing
                Set fd3 = Nothing
                Set ROR = Nothing
    End If
    
End Sub
