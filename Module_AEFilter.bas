Attribute VB_Name = "ModuleAEFilter"
Sub AdvancedFilter()
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        Dim rg As Range
        Set rg = Workbooks("待修分析.xlsm").Worksheets("搜尋機種").[A12].CurrentRegion
        rg.Offset(1).ClearContents
        
        
        Dim wb As Workbook, sPath$
        sPath = "P:\Service\RMA\Main\Kaitek RMA " & [B7] & " main.xls"
        
        Set wb = Workbooks.Open(sPath, UpdateLinks = 0)
        
        
        If wb.Worksheets("Master").FilterMode Then
                wb.Worksheets("Master").ShowAllData
        End If
        
        Dim rgData As Range, rgCriteria As Range, rgOutput As Range
        
        Set rgData = wb.Worksheets("Master").[A1].CurrentRegion
        
        With wb.Worksheets("Master")
                .AutoFilter.Sort.SortFields.Clear
                .AutoFilter.Sort.SortFields.Add Key:=Range("A1:A" & rgData.Rows.Count), Order:=xlDescending
                .AutoFilter.Sort.Apply
       End With
       
        Set rgCriteria = Workbooks("待修分析.xlsm").Worksheets("搜尋機種").[A1].CurrentRegion
        Set rgOutput = Workbooks("待修分析.xlsm").Worksheets("搜尋機種").[A12].CurrentRegion
        
        rgData.AdvancedFilter xlFilterCopy, rgCriteria, rgOutput
        
        
        wb.Close False
        Set wb = Nothing
        
        [A1].Select
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True

End Sub

