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

Function forecast_error_multiple(for_x, x, stand_error)
    ' function calculation forecast error for multiple regression
    ' for_x - forecasting x vector
    ' x - range of x used for building regression
    ' (if model including loose member if shold have ones first column)
    ' stand_error - standart error of approximation
    
    forecast_error_multiple = Application.Transpose(x)
    
End Function

 Function basic_autocorrelation(range, t)
    ' autorrelation function using formula:
    ' r(t) = cov(y(1,2,...,n-t),y(t,...,n))/sigma(y(1,2,...,n-t))*sigma(y(t,...,n))
    
    Dim arr1() As Variant
    Dim arr2() As Variant
    
    n = range.Count
    
    ReDim arr1(1 To n - t)
    ReDim arr2(t + 1 To n)
    
    For i = 1 To n
        If i <= n - t Then
            arr1(i) = range(i).Value
        End If
        If i > t Then
            arr2(i) = range(i).Value
        End If
        
    Next i
    
    basic_autocorrelation = Application.Correl(arr1, arr2)

 End Function


Function QstatLB(r, n)
    ' funtion for computing Lewis Box statistics
    ' r - range of computed autocorrlation coefficients for current and previous autocorrelation coefficients
    ' n - size of of the sample under study
    
    form_sum = 0
    t = r.Count

    For i = 1 To t
        form_sum = form_sum + ((r(i) ^ 2) / (n - i))
    Next i

    QstatLB = n * (n + 2) * form_sum

End Function
