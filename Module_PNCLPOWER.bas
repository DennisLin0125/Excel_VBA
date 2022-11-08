Attribute VB_Name = "ModulePNCLPOWER"
Sub PNCL20K()
        [C29] = 20000
        [C40] = 800
        [C51] = 50
        
        [D54] = 806
        [D55] = 46
        
        [E19:E29] = 30
        [E30:E40] = 117.5
        [E41:E51] = 8.3
        
        Range("D41:D51") = Range("C41:C51").Value
        Range("D30:D40") = Range("C30:C40").Value
'******************************************************************************************************
        Dim arr
        arr = Range("C19:C29")
        
        Dim temp(11)
        i = 0
        For Each c In arr
                temp(i) = Int(c * 睹计A)
                i = i + 1
        Next
        
        Range("D19").Resize(UBound(temp)) = Application.WorksheetFunction.Transpose(temp)
'*****************************************************************************************************
        arr = Range("C30:C40")
        i = 0
        For Each c In arr
                temp(i) = Int(c * 睹计A)
                i = i + 1
        Next
        Range("D30").Resize(UBound(temp)) = Application.WorksheetFunction.Transpose(temp)
End Sub

Function 睹计A() As Double
        Dim MyValu%
        Dim Min%
        Dim Max%
        
        Dim arr
        arr = Array("0.9921", "0.9935", "0.9956", "0.9962", "0.9978", "0.9989", "0.9998", "1.00912", "1.00924")
        
        Randomize
        
        Min = LBound(arr)
        Max = UBound(arr)
        MyValue = Int((Max - Min + 1) * Rnd() + Min)
               
        睹计A = arr(MyValue)
        
End Function
