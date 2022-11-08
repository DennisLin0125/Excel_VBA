Attribute VB_Name = "Module1複製照片到P槽"
Sub 建立照片至P槽()

        SN = InputBox("請輸入SN")
        If SN = "" Then Exit Sub
        
        Dim myPath$
        myPath = "D:\Users\Dlin\Desktop\照片\" & SN
        If Dir(myPath & "\") = "" Then
                MsgBox "照片資料夾裡，沒有SN : " & SN & " 的照片！", vbCritical
                Exit Sub
        End If
        
        RMA = InputBox("請輸入RMA")
        If RMA = "" Then Exit Sub
        
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
                .InitialFileName = myPath
                .ButtonName = "就是你了!!"
                .Title = "請選擇 進出廠照片的(維修前)照片"
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
                .ButtonName = "就是你了!!"
                MsgBox "請選擇 進出廠照片的(維修後)照片"
               
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
                .ButtonName = "就是你了!!"
                MsgBox "請選擇 故障照片的(維修前)照片"
               
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
                .Title = "請選擇 故障照片的 (維修後) 照片"
                .ButtonName = "就是你了!!"
                MsgBox "請選擇 故障照片的(維修後)照片"
                
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
                .Title = "請選擇Log資料"
                .ButtonName = "就是你了!!"
                MsgBox "請選擇 LOG 資料"
                
                If .Show Then
                        For Each sPath In .SelectedItems
                                fs.CopyFile sPath, LogPath
                        Next
                End If
        End With
        
        '腔體組裝照片
'        Dim fdSource As FileDialog
'        Set fdSource = Application.FileDialog(msoFileDialogFolderPicker)
'
'        With fdSource
'                .InitialFileName = "P:\Service\Repair Picture\RF Product\拆裝人員暫放Source"
'                .Title = "請選擇組裝人員照片以RMA命名的資料夾"
'                .ButtonName = "就是你了!!"
'                MsgBox "請選擇 組裝人員照片以 RMA 命名的資料夾" & Chr(10) & Chr(10) & "如果腔體未拆請按 取消"
'
'                If .Show Then
'                        fs.CopyFolder .SelectedItems(1), SourcePath
'                End If
'        End With
        
        Set fs = Nothing
        Set fdLog = Nothing
        Set fdSource = Nothing
        Set fdRpAfter = Nothing
        Set fdRpBefor = Nothing
        Set fdMachAfter = Nothing
        Set fdMachBefor = Nothing
        
        
        '移動
        Source = "P:\Service\技術討論專區\Engineer\Dennis\MKS\" & SN
        
        Dim fd2 As FileDialog, a%
        Set fd2 = Application.FileDialog(msoFileDialogFolderPicker)
        
        With fd2
                .Title = "請選擇要移動到哪個資料夾"
                .InitialFileName = "P:\Service\Product\Repair Picture"
                .ButtonName = "就是你了!!"
                MsgBox "請選擇要移動到哪個資料夾"
                    
                If .Show Then
                        tpath = .SelectedItems(1) & "\" & SN
                        Set fs = CreateObject("scripting.FileSystemObject")
                        fs.CopyFolder Source, tpath
                        fs.DeleteFolder Source
                        a = MsgBox("是否要刪除 " & SN & " 資料夾 ?", vbYesNo)
                        If a = vbYes Then
                                fs.DeleteFolder myPath
                        End If
                        MsgBox "處理完成"
                End If
        End With
        
        Set fd2 = Nothing
        Set fs = Nothing
End Sub

