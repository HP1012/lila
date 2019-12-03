Sub Merge_Multi_Column()
  Dim rng As Range

  If TypeName(Selection) = "Range" Then
    Set rng = Selection
    Dim temp_addr As String: temp_addr = Replace(rng.Address, "$", "")
    start_cell_addr = Split(temp_addr, ":")(0)
    last_cell_addr = Split(temp_addr, ":")(1)
    
    start_cell_convert_row = Range(start_cell_addr).Row
    start_cell_convert_col = Range(start_cell_addr).Column
    last_cell_addr_row = Range(last_cell_addr).Row
    last_cell_addr_col = Range(last_cell_addr).Column

    For j = start_cell_convert_col To last_cell_addr_col
        Dim temp_rng_addr_start As String: temp_rng_addr_start = Cells(start_cell_convert_row, j).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        Dim temp_rng_addr_end As String: temp_rng_addr_end = Cells(last_cell_addr_row, j).Address(RowAbsolute:=False, ColumnAbsolute:=False)

        Dim temp_rng As String: temp_rng = temp_rng_addr_start & ":" & temp_rng_addr_end
        Dim temp_merge_rng As Range: Set temp_merge_rng = Application.Range(temp_rng)
        temp_merge_rng.Merge
    Next j
  Else
    MsgBox "Please select a range of cells before running this macro!"
  End If
End Sub

Sub Insert_N_Row()
    Dim input_n_row As String: input_n_row = InputBox("Please input the number of row")

    If IsNumeric(input_n_row) Then
        input_n_row = input_n_row - 1
        Dim lRow As Long
        If TypeName(Selection) = "Range" And IsNumeric(input_n_row) Then
            lRow = Selection.Row()
            Debug.Print (TypeName(Selection))
            For i = 0 To input_n_row - 1
                Rows(lRow + 1).Select
                Selection.Insert Shift:=xlDown
            Next i
        Else
            Debug.Print "Please select a range of cells before running this macro!"
        End If
    Else
        Debug.Print ("Please input the number not character")
    End If
End Sub