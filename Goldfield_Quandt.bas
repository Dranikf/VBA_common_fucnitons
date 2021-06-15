Attribute VB_Name = "Module1"
Sub GoldfildQand()
    ' ���� ��������� �������
    ' ����������� ��� ����������� ���������� ��� ����������
    ' �������������� ��� � ������ ������� - ����������� ����������
    ' ��������� ���������
    
    '���������� �������� ��������������!
    
    
    table_height = Selection.Rows.Count
    table_width = Selection.Columns.Count
    
    '������� ������� ������� �����
    If table_height Mod 3 <> 0 Then
        g_size = Int(table_height / 3) + 1
    Else
        g_size = table_height / 3
    End If
    
    '�������� ������������ ��������� ��� ������� ������
    Selection.Cells(1, table_width + 1) = "������������ ������ ������"
    For i = 1 To table_width - 1
        Selection.Cells(2, table_width + i) = "b" & Str(i)
    Next i
    
    Selection.Cells(2, 2 * table_width) = "a"
    Range(Selection.Cells(3, table_width + 1).Address & ":" & Selection.Cells(3, 2 * table_width).Address).FormulaArray = "=LINEST(" & Selection.Cells(1, 1).Address _
    & ":" & Selection.Cells(g_size, 1).Address & "," & Selection.Cells(1, 2).Address & ":" & Selection.Cells(g_size, table_width).Address & ")"
    
    ' �������� ������������ �������� ��� ������ ������
    Selection.Cells(4, table_width + 1) = "������������ ������� ������"
    For i = 1 To table_width - 1
        Selection.Cells(5, table_width + i) = "b" & Str(i)
    Next i
    Selection.Cells(5, 2 * table_width) = "a"
    
    Range(Selection.Cells(6, table_width + 1).Address & ":" & Selection.Cells(6, 2 * table_width).Address).FormulaArray = "=LINEST(" & Selection.Cells(table_height - g_size + 1, 1).Address _
    & ":" & Selection.Cells(table_height, 1).Address & "," & Selection.Cells(table_height - g_size + 1, 2).Address & ":" & Selection.Cells(table_height, table_width).Address & ")"
    

    
    '����� ��������� �� ������ ���� � ����
    padding = table_width * 2
    
    Selection.Cells(0, padding + 1) = "y^"
    ' ������� y � �������������� �
    Range(Selection.Cells(1, padding + 1).Address & ":" & Selection.Cells(g_size, padding + 1).Address).FormulaArray = _
    "=TRANSPOSE(MMULT(" & Selection.Cells(3, table_width + 1).Address & ":" & Selection.Cells(3, padding - 1).Address & ",TRANSPOSE(" & _
    Selection.Cells(1, 2).Address & ":" & Selection.Cells(g_size, table_width).Address & "))+" & Selection.Cells(3, padding).Address & ")"
    
    'Debug.Print (Selection.Cells(table_height - g_size + 1, padding + 1).Address & ":" & Selection.Cells(table_height, padding + 1).Address)
    Range(Selection.Cells(table_height - g_size + 1, padding + 1).Address & ":" & Selection.Cells(table_height, padding + 1).Address).FormulaArray = _
    "=TRANSPOSE(MMULT(" & Selection.Cells(6, table_width + 1).Address & ":" & Selection.Cells(6, padding - 1).Address & ",TRANSPOSE(" & _
    Selection.Cells(table_height - g_size + 1, 2).Address & ":" & Selection.Cells(table_height, table_width).Address & "))+" & Selection.Cells(6, padding).Address & ")"
    
    
    ' ������ ����������� ������
    Selection.Cells(0, padding + 2) = "e"
    Range(Selection.Cells(1, padding + 2).Address & ":" & Selection.Cells(g_size, padding + 2).Address).FormulaArray = _
    "=" & Selection.Cells(1, 1).Address & ":" & Selection.Cells(g_size, 1).Address & "-" & Selection.Cells(1, padding + 1).Address & ":" & Selection.Cells(g_size, padding + 1).Address
    
    Range(Selection.Cells(table_height - g_size + 1, padding + 2).Address & ":" & Selection.Cells(table_height, padding + 2).Address).FormulaArray = _
    "=" & Selection.Cells(table_height - g_size + 1, 1).Address & ":" & Selection.Cells(table_height, 1).Address & _
    "-" & Selection.Cells(table_height - g_size + 1, padding + 1).Address & ":" & Selection.Cells(table_height, padding + 1).Address
    
    ' �������� ������
    Selection.Cells(0, padding + 3) = "e^2"
    Range(Selection.Cells(1, padding + 3).Address & ":" & Selection.Cells(g_size, padding + 3).Address).FormulaArray = _
    "=" & Selection.Cells(1, padding + 2).Address & ":" & Selection.Cells(g_size, padding + 2).Address & "^2"
    
    Range(Selection.Cells(table_height - g_size + 1, padding + 3).Address & ":" & Selection.Cells(table_height, padding + 3).Address).FormulaArray = _
    "=" & Selection.Cells(table_height - g_size + 1, padding + 2).Address & ":" & Selection.Cells(table_height, padding + 2).Address & "^2"
    
    ' ����� ��������� ������
    Selection.Cells(0, padding + 4) = "S"
    Selection.Cells(1, padding + 4).Formula = "=Sum(" & Selection.Cells(1, padding + 3).Address & ":" & Selection.Cells(g_size, padding + 3).Address & ")"
    Selection.Cells(table_height - g_size + 1, padding + 4).Formula = "=Sum(" & Selection.Cells(table_height - g_size + 1, padding + 3).Address & ":" & Selection.Cells(table_height, padding + 3).Address & ")"
    
    ' ��������� F ����������
    Selection.Cells(0, padding + 5) = "F"
    Selection.Cells(1, padding + 5) = "=If(" & Selection.Cells(1, padding + 4).Address & "<" & Selection.Cells(table_height - g_size + 1, padding + 4).Address & _
     "," & Selection.Cells(table_height - g_size + 1, padding + 4).Address & "/" & Selection.Cells(1, padding + 4).Address & "," _
     & Selection.Cells(1, padding + 4).Address & "/" & Selection.Cells(table_height - g_size + 1, padding + 4).Address & ")"
    
    ' ���������� ������������ F
    Selection.Cells(0, padding + 6) = "Fkr"
    Selection.Cells(1, padding + 6).Formula = "=F.DIST.RT(" & Selection.Cells(1, padding + 5).Address & ", " & Str(g_size - table_width) & ", " & Str(g_size - table_width) & ")"
    
End Sub
