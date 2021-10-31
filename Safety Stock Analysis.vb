' CODE USED TO GENERATE CONSUMPTION RATE AS GRAPH IN SEPARATE EXCEL SHEET
' THIS ENABLES TO EASILY SEE VISUALLY HOW TO ADJUST SAFETY STOCK FOR EACH PART
' HELPING TO REDUCE STOCK BALANCE, AND SECURE PRODUCTION FLOW


Private Sub CommandButton1_Click()


    MSG1 = MsgBox("Chart generation takes approximately 10 minutes, you cannot use computer during that time. Excel may become ""Non Responding"", don't worry about it and wait. Are you sure you want to generate?", vbYesNo, "Generate Charts button was clicked.")

    If MSG1 = vbNo Then
      Exit Sub
    End If


On Error Resume Next
    Set main_worksheet = Worksheets(1)
    On Error GoTo 0
    'make sure we have at least one visible sheet
    If Not main_worksheet Is Nothing Then
        Application.DisplayAlerts = False
        For Each main_worksheet In ThisWorkbook.Worksheets
            If Not main_worksheet.Index = 1 Then main_worksheet.Delete
        Next main_worksheet
        Application.DisplayAlerts = True
    End If

    Worksheets(1).Columns("F:F").Delete Shift:=xlToLeft

'======================================================================================================================================================'
    'Count number of lines'
    Dim nr_of_lines As Long
    nr_of_lines = 6

    For Each Line In Worksheets(1).Range("A7:A1000000")
        If Line.Value <> "" Then
            nr_of_lines = nr_of_lines + 1
        End If
        If Line.Value = "" Then
            Exit For
        End If
    Next Line

'======================================================================================================================================================'
    Dim end_of_rim As String
    end_of_rim = "A" + CStr(nr_of_lines)
    Dim next_rim_count As Long
    next_rim_count = 8
    Dim next_rim As String
    next_rim = "A" + CStr(next_rim_count)
    Dim current_rim_count As Long
    current_rim_count = 7
    Dim current_rim As String
    current_rim = "A" + CStr(current_rim_count)
    Dim ranges_of_rims(1 To 1000000, 1 To 2) As String
    Dim single_range As Long
    single_range = 1

'======================================================================================================================================================'
    'Get endings of orders'
    For Each rim In Worksheets(1).Range("A7:" & end_of_rim)
        If rim = Worksheets(1).Range(next_rim).Value Then
            next_rim_count = next_rim_count + 1
            current_rim_count = current_rim_count + 1
            next_rim = "A" + CStr(next_rim_count)
            current_rim = "A" + CStr(current_rim_count)
        Else
            next_rim_count = next_rim_count + 1
            
            ranges_of_rims(single_range, 1) = "A" + CStr(current_rim_count)
            
            single_range = single_range + 1
            current_rim_count = current_rim_count + 1

            next_rim = "A" + CStr(next_rim_count)
            current_rim = "A" + CStr(current_rim_count)
        End If
    Next rim

    single_range = 1
    current_rim_count = 7
    next_rim_count = 6
    current_rim = "A" + CStr(current_rim_count)
    next_rim = "A" + CStr(next_rim_count)

    'Get beginnings of orders'
    For Each Order In Worksheets(1).Range("A7:" & end_of_rim)
        If Order = Worksheets(1).Range(next_rim).Value Then
            next_rim_count = next_rim_count + 1
            current_rim_count = current_rim_count + 1
            next_rim = "A" + CStr(next_rim_count)
            current_rim = "A" + CStr(current_rim_count)
        Else
            next_rim_count = next_rim_count + 1
            
            ranges_of_rims(single_range, 2) = "A" + CStr(current_rim_count)
            
            single_range = single_range + 1
            current_rim_count = current_rim_count + 1

            next_rim = "A" + CStr(next_rim_count)
            current_rim = "A" + CStr(current_rim_count)
        End If
    Next Order

'======================================================================================================================================================'
    Dim combined_count As Long
    combined_count = 1
    Dim combined_ranges_of_rims(1 To 100000) As String
    Dim every_second_row As Long
    every_second_row = 0
    Dim row_counter As Long
    Dim column_counter As Long

'======================================================================================================================================================'
    'Combine two level array into one'
    For row_counter = 1 To 100000
        For column_counter = 1 To 2
            combined_ranges_of_rims(combined_count) = ranges_of_rims(row_counter, column_counter) + combined_ranges_of_rims(combined_count)
        Next column_counter
        combined_count = combined_count + 1
    Next row_counter

'======================================================================================================================================================'
    Dim clear_ranges_of_rims As Variant
    Dim clear_counter As Long
    clear_counter = 1

'======================================================================================================================================================'
    'Clear combined_ranges_of_rims array from empty fields'
    ReDim clear_ranges_of_rims(LBound(combined_ranges_of_rims) To UBound(combined_ranges_of_rims))
    For i = LBound(combined_ranges_of_rims) To UBound(combined_ranges_of_rims)
        If combined_ranges_of_rims(i) <> "" Then
            j = j + 1
            clear_ranges_of_rims(j) = combined_ranges_of_rims(i)
        End If
    Next i
    ReDim Preserve clear_ranges_of_rims(LBound(combined_ranges_of_rims) To j)

'======================================================================================================================================================'
    Dim left_piece As String
    Dim right_piece As String
    Dim array_counter As Long
    Dim cut_strings() As String
    Dim preserve_counter As Long
    preserve_counter = 1
    array_counter = 1
    ReDim rim_numbers(1 To 1) As String
    ReDim combined_rim_range(1 To 1) As String
    ReDim combined_safetystock_range(1 To 1) As String
    ReDim combined_stock_range(1 To 1) As String
    ReDim combined_date_range(1 To 1) As String
    

'======================================================================================================================================================'
    'Weld together ranges for orders'
    For Each i In clear_ranges_of_rims
        cut_strings() = Split(i, "A")
        left_piece = cut_strings(1)
        right_piece = cut_strings(2)
        ReDim Preserve rim_numbers(1 To preserve_counter)
        ReDim Preserve combined_rim_range(1 To preserve_counter)
        ReDim Preserve combined_safetystock_range(1 To preserve_counter)
        ReDim Preserve combined_stock_range(1 To preserve_counter)
        ReDim Preserve combined_date_range(1 To preserve_counter)

        clear_ranges_of_rims(array_counter) = "A" + left_piece + ":A" + right_piece

        combined_rim_range(array_counter) = "A" + left_piece + ":" + "A" + right_piece
        combined_safetystock_range(array_counter) = "B" + left_piece + ":" + "B" + right_piece
        combined_stock_range(array_counter) = "C" + left_piece + ":" + "C" + right_piece
        combined_date_range(array_counter) = "D" + left_piece + ":" + "D" + right_piece

        rim_numbers(preserve_counter) = "A" + left_piece
        
        preserve_counter = preserve_counter + 1
        array_counter = array_counter + 1
    Next i

'======================================================================================================================================================'


'======================================================================================================================================================'
    Dim rim_range_counter As Long
    rim_range_counter = 0
    Dim rim_nr As Long
    rim_nr = 1
'======================================================================================================================================================'
    'Count number of orders
    For Each rim In combined_rim_range
        rim_range_counter = rim_range_counter + 1
    Next rim
'======================================================================================================================================================'
    ' Declarations for variables used in XML copy loop
    
    Dim paste_counter As Long
    Dim rim_range As String
    Dim safetystock_range As String
    Dim stock_range As String
    Dim date_range As String
    Dim worksheet_counter As Long
    worksheet_counter = 2
    Dim ws As Worksheet

    Dim new_sheet_name As String
    new_sheet_name = ""


    Dim paste_length_counter As Long
    paste_length_counter = 0

    Dim if_safetystock_saved As Long
    Dim safetystock_saved_counter As Long

    Dim succesCounter As Long
    succesCounter = 0


    Worksheets(1).Columns("A:A").Replace What:="/", Replacement:=" ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

'======================================================================================================================================================'
    ' Copy data from single order, save as XML, clear and move to next order
    For rim_nr = 1 To rim_range_counter
        
        paste_counter = 1
        rim_range = "A" + CStr(paste_counter)
        safetystock_range = "B" + CStr(paste_counter)
        stock_range = "C" + CStr(paste_counter)
        date_range = "D" + CStr(paste_counter)

       
        
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = Worksheets(1).Range(CStr(rim_numbers(rim_nr))).Value

        new_sheet_name = Worksheets(1).Range(CStr(rim_numbers(rim_nr))).Value



        ' Copy all necessary data from Arrow Excel to Rimaster Order Confirmation

        For Each i In Worksheets(1).Range(combined_rim_range(rim_nr))
            Worksheets(worksheet_counter).Range(rim_range).Value = CStr(i.Value)
            paste_counter = paste_counter + 1
            rim_range = "A" + CStr(paste_counter)
        Next i
        paste_counter = 1
        ' Copy RIM numbers
        For Each i In Worksheets(1).Range(combined_safetystock_range(rim_nr))
            Worksheets(worksheet_counter).Range(safetystock_range).Value = CStr(i.Value)
            paste_counter = paste_counter + 1
            safetystock_range = "B" + CStr(paste_counter)
        Next i
        paste_counter = 1
        ' Copy confirmed date
        For Each i In Worksheets(1).Range(combined_stock_range(rim_nr))
            Worksheets(worksheet_counter).Range(stock_range).Value = CLng(i.Value)
            paste_counter = paste_counter + 1
            stock_range = "C" + CStr(paste_counter)
        Next i
        paste_counter = 1
        ' Copy confirmed quantity
        For Each i In Worksheets(1).Range(combined_date_range(rim_nr))
            Worksheets(worksheet_counter).Range(date_range).Value = CStr(i.Value)
            paste_counter = paste_counter + 1
            date_range = "D" + CStr(paste_counter)
        Next i
        paste_counter = 1

        Worksheets(worksheet_counter).Range("G3").Select

        paste_length_counter = 0
        For Each i In Worksheets(worksheet_counter).Range("A1:A10000")
            If i.Value <> "" Then
                paste_length_counter = paste_length_counter + 1
            End If
            If i.Value = "" Then
                Exit For
            End If
        Next i



        Worksheets(worksheet_counter).Shapes.AddChart2(227, xlLine).Select
        Worksheets(worksheet_counter).ChartObjects(1).Activate
        ActiveChart.ChartTitle.Text = Worksheets(1).Range(CStr(rim_numbers(rim_nr))).Value
        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.FullSeriesCollection(1).Name = "='" & Worksheets(1).Range(CStr(rim_numbers(rim_nr))).Value & "'!$A$1"
        ActiveChart.FullSeriesCollection(1).Values = "='" & Worksheets(1).Range(CStr(rim_numbers(rim_nr))).Value & "'!$B$1:$B$" & CStr(paste_length_counter)
        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.FullSeriesCollection(2).Name = "='" & Worksheets(1).Range(CStr(rim_numbers(rim_nr))).Value & "'!$A$1"
        ActiveChart.FullSeriesCollection(2).Values = "='" & Worksheets(1).Range(CStr(rim_numbers(rim_nr))).Value & "'!$C$1:$C$" & CStr(paste_length_counter)
        ActiveChart.FullSeriesCollection(2).XValues = "='" & Worksheets(1).Range(CStr(rim_numbers(rim_nr))).Value & "'!$D$1:$D$" & CStr(paste_length_counter)
        Worksheets(worksheet_counter).Shapes(1).IncrementLeft -250
        Worksheets(worksheet_counter).Shapes(1).IncrementTop -150
        Worksheets(worksheet_counter).Shapes(1).ScaleWidth 2.5, msoFalse, msoScaleFromTopLeft
        Worksheets(worksheet_counter).Shapes(1).ScaleHeight 2.24, msoFalse, msoScaleFromTopLeft


               
        
        safetystock_saved_counter = 1

        For Each i In Worksheets(worksheet_counter).Range("C1:C" & CStr(paste_length_counter))
            If i.Value < Worksheets(worksheet_counter).Range("B" & safetystock_saved_counter).Value Then
                Worksheets(worksheet_counter).Range("A1").Style = "Zły"
                safetystock_saved_counter = safetystock_saved_counter + 1
            Else
                safetystock_saved_counter = safetystock_saved_counter + 1
            End If
        Next i




        worksheet_counter = worksheet_counter + 1
        succesCounter = succesCounter + 1

    Next rim_nr


    safetystock_saved_counter = 7

    Dim test_var As String

    For i = 2 To (worksheet_counter - 1)
        test_var = Worksheets(i).Index
        Worksheets(i).Range("A1").Copy
        Worksheets(1).Paste Destination:=Worksheets(1).Range("F" & safetystock_saved_counter)
        safetystock_saved_counter = safetystock_saved_counter + 1
    Next i


    If succesCounter > 0 Then
        MsgBox "Chart Generation Complete"
    End If

End Sub




