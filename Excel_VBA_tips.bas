Attribute VB_Name = "Module2"
Function reshape(arr, n, m)
    'fucntion for reshaping array in sheet
    ' arr - input array
    ' n - result strings
    ' m - result columns
    
    Dim result() As Variant
    
    ReDim result(n, m)
    i = 0
    j = 0
    For Each obj In arr
    
        result(i, j) = obj
        
        If j = m - 1 Then
            j = 0
            i = i + 1
        Else
            j = j + 1
        End If
    Next obj
    
    reshape = result
    
End Function


Function forecast_error(for_x, x, stand_error)
    ' function calculation forecast error
    ' for_x - forecasting x
    ' x - range of x used for building regression
    ' stand_error - standart error of approximation
    
    n = x.Count
    av_val = Application.Average(x)
    disp = Application.Var_P(x)
    
    forecast_error = stand_error * ((1 + (1 / n) + ((for_x - av_val) ^ 2) / (n * disp)) ^ (1 / 2))

End Function
