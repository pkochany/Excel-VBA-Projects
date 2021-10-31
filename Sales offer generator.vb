
Sub mark_filtered()
  Dim tableLength As Long
  tableLength = 1
  For Each i In Worksheets(1).Range("A2:A100000")
    If i <> "" Then
      tableLength = tableLength + 1
    Else
      Exit For
    End If
  Next i
  Worksheets(1).Range("AC2:AC" & tableLength).Value = "x"
'
End Sub
Sub clear_calc_list()

    For i = 1 To 29
        Worksheets(1).ListObjects("Tabela_quote_list").Range.AutoFilter _
            Field:=i
    Next i
    
    Worksheets(1).Columns("AC").ClearContents
    
    Worksheets(1).Range("AC1").Value = "Mark:"
    
    Worksheets(1).ListObjects("Tabela_quote_list").Range.AutoFilter Field:=29, _
        Criteria1:=RGB(198, 224, 180), Operator:=xlFilterCellColor
    
End Sub

Sub calc_show_history()

    Worksheets(1).ListObjects("Tabela_quote_list").Range.AutoFilter _
            Field:=29

End Sub

Sub calc_hide_history()

    Worksheets(1).ListObjects("Tabela_quote_list").Range.AutoFilter Field:=29, _
        Criteria1:=RGB(198, 224, 180), Operator:=xlFilterCellColor

End Sub

Sub hide_details()

  Worksheets(5).ListObjects("Tabela_quote_activities").Range.AutoFilter Field:=8 _
    , Criteria1:="=1", Operator:=xlOr, Criteria2:="=11"

End Sub

Sub show_details()

  Worksheets(5).ListObjects("Tabela_quote_activities").Range.AutoFilter _
    Field:=8

End Sub

Sub refresh_offer()
    
    ' Set proper titles for extender version '
    Worksheets(2).Range("E26:E26").Value = "Material"
    Worksheets(2).Range("F26:F26").Value = "Work cost"
    Worksheets(2).Range("G26:G26").Value = "Overheads"
    ' And borders for titles '
    Worksheets(2).Range("E26").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("F26").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("E116").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("F116").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("E206").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("F206").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("E296").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("F296").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("E386").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("F386").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("E476").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("F476").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("E566").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("F566").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("E656").Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets(2).Range("F656").Borders(xlEdgeRight).LineStyle = xlContinuous
    
    ' Set proper titles for simplified version '
    If Worksheets(2).Range("M13").Value = "Yes" Then
      Worksheets(2).Range("E26:E26").Value = "Reference"
      Worksheets(2).Range("F26:F26").Value = "Valid From"
      Worksheets(2).Range("G26:G26").Value = "Valid To"
    End If
    
    Dim currentPosition As Long
    currentPosition = 2

    Dim markedRowsArray() As Long

    Dim arrayCounter As Integer
    arrayCounter = 0

    For Each i In Worksheets(1).Range("AC2:AC100000")
      If i <> "" Then
        ReDim Preserve markedRowsArray(arrayCounter)
        markedRowsArray(arrayCounter) = currentPosition
        arrayCounter = arrayCounter + 1
      End If
      currentPosition = currentPosition + 1
    Next i

    If arrayCounter = 0 Then
      MsgBox "Please mark harnesses you want to offer."
      Exit Sub
    End If

    ' wyczyść stronę 1 oferty '
    Dim rowsPage1 As Long
    rowsPage1 = 26
    For Each i In Worksheets(2).Range("A27:A10000")
      If i <> "" Then
        rowsPage1 = rowsPage1 + 1
      Else
        Exit For
      End If
    Next i
    If rowsPage1 > 26 Then
      Set rowsClear = Worksheets(2).Range("A27:J" & rowsPage1)
      rowsClear.NumberFormat = "General"
      With rowsClear.Interior
          .Pattern = xlNone
          .TintAndShade = 0
          .PatternTintAndShade = 0
      End With
      rowsClear.Borders(xlDiagonalDown).LineStyle = xlNone
      rowsClear.Borders(xlDiagonalUp).LineStyle = xlNone
      rowsClear.Borders(xlEdgeLeft).LineStyle = xlNone
      rowsClear.Borders(xlEdgeTop).LineStyle = xlNone
      rowsClear.Borders(xlEdgeBottom).LineStyle = xlNone
      rowsClear.Borders(xlEdgeRight).LineStyle = xlNone
      rowsClear.Borders(xlInsideVertical).LineStyle = xlNone
      rowsClear.Borders(xlInsideHorizontal).LineStyle = xlNone
      With Worksheets(2).Range("A27:J27").Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      rowsClear.ClearContents
    End If
    ' wyczyść stronę 2 oferty '
    Dim rowsPage2 As Long
    rowsPage2 = 116
    For Each i In Worksheets(2).Range("A117:A10000")
      If i <> "" Then
        rowsPage2 = rowsPage2 + 1
      Else
        Exit For
      End If
    Next i
    If rowsPage2 > 116 Then
      Set rowsClear = Worksheets(2).Range("A117:J" & rowsPage2)
      rowsClear.NumberFormat = "General"
      With rowsClear.Interior
          .Pattern = xlNone
          .TintAndShade = 0
          .PatternTintAndShade = 0
      End With
      rowsClear.Borders(xlDiagonalDown).LineStyle = xlNone
      rowsClear.Borders(xlDiagonalUp).LineStyle = xlNone
      rowsClear.Borders(xlEdgeLeft).LineStyle = xlNone
      rowsClear.Borders(xlEdgeTop).LineStyle = xlNone
      rowsClear.Borders(xlEdgeBottom).LineStyle = xlNone
      rowsClear.Borders(xlEdgeRight).LineStyle = xlNone
      rowsClear.Borders(xlInsideVertical).LineStyle = xlNone
      rowsClear.Borders(xlInsideHorizontal).LineStyle = xlNone
      With Worksheets(2).Range("A117:J117").Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      rowsClear.ClearContents
    End If
    ' wyczyść stronę 3 oferty '
    Dim rowsPage3 As Long
    rowsPage3 = 206
    For Each i In Worksheets(2).Range("A207:A10000")
      If i <> "" Then
        rowsPage3 = rowsPage3 + 1
      Else
        Exit For
      End If
    Next i
    If rowsPage3 > 206 Then
      Set rowsClear = Worksheets(2).Range("A207:J" & rowsPage3)
      rowsClear.NumberFormat = "General"
      With rowsClear.Interior
          .Pattern = xlNone
          .TintAndShade = 0
          .PatternTintAndShade = 0
      End With
      rowsClear.Borders(xlDiagonalDown).LineStyle = xlNone
      rowsClear.Borders(xlDiagonalUp).LineStyle = xlNone
      rowsClear.Borders(xlEdgeLeft).LineStyle = xlNone
      rowsClear.Borders(xlEdgeTop).LineStyle = xlNone
      rowsClear.Borders(xlEdgeBottom).LineStyle = xlNone
      rowsClear.Borders(xlEdgeRight).LineStyle = xlNone
      rowsClear.Borders(xlInsideVertical).LineStyle = xlNone
      rowsClear.Borders(xlInsideHorizontal).LineStyle = xlNone
      With Worksheets(2).Range("A207:J207").Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      rowsClear.ClearContents
    End If
    ' wyczyść stronę 4 oferty '
    Dim rowsPage4 As Long
    rowsPage4 = 296
    For Each i In Worksheets(2).Range("A297:A10000")
      If i <> "" Then
        rowsPage4 = rowsPage4 + 1
      Else
        Exit For
      End If
    Next i
    If rowsPage4 > 296 Then
      Set rowsClear = Worksheets(2).Range("A297:J" & rowsPage4)
      rowsClear.NumberFormat = "General"
      With rowsClear.Interior
          .Pattern = xlNone
          .TintAndShade = 0
          .PatternTintAndShade = 0
      End With
      rowsClear.Borders(xlDiagonalDown).LineStyle = xlNone
      rowsClear.Borders(xlDiagonalUp).LineStyle = xlNone
      rowsClear.Borders(xlEdgeLeft).LineStyle = xlNone
      rowsClear.Borders(xlEdgeTop).LineStyle = xlNone
      rowsClear.Borders(xlEdgeBottom).LineStyle = xlNone
      rowsClear.Borders(xlEdgeRight).LineStyle = xlNone
      rowsClear.Borders(xlInsideVertical).LineStyle = xlNone
      rowsClear.Borders(xlInsideHorizontal).LineStyle = xlNone
      With Worksheets(2).Range("A297:J297").Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      rowsClear.ClearContents
    End If
    ' wyczyść stronę 5 oferty '
    Dim rowsPage5 As Long
    rowsPage5 = 386
    For Each i In Worksheets(2).Range("A387:A10000")
      If i <> "" Then
        rowsPage5 = rowsPage5 + 1
      Else
        Exit For
      End If
    Next i
    If rowsPage5 > 386 Then
      Set rowsClear = Worksheets(2).Range("A387:J" & rowsPage5)
      rowsClear.NumberFormat = "General"
      With rowsClear.Interior
          .Pattern = xlNone
          .TintAndShade = 0
          .PatternTintAndShade = 0
      End With
      rowsClear.Borders(xlDiagonalDown).LineStyle = xlNone
      rowsClear.Borders(xlDiagonalUp).LineStyle = xlNone
      rowsClear.Borders(xlEdgeLeft).LineStyle = xlNone
      rowsClear.Borders(xlEdgeTop).LineStyle = xlNone
      rowsClear.Borders(xlEdgeBottom).LineStyle = xlNone
      rowsClear.Borders(xlEdgeRight).LineStyle = xlNone
      rowsClear.Borders(xlInsideVertical).LineStyle = xlNone
      rowsClear.Borders(xlInsideHorizontal).LineStyle = xlNone
      With Worksheets(2).Range("A387:J387").Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      rowsClear.ClearContents
    End If
    ' wyczyść stronę 6 oferty '
    Dim rowsPage6 As Long
    rowsPage6 = 476
    For Each i In Worksheets(2).Range("A477:A10000")
      If i <> "" Then
        rowsPage6 = rowsPage6 + 1
      Else
        Exit For
      End If
    Next i
    If rowsPage6 > 476 Then
      Set rowsClear = Worksheets(2).Range("A477:J" & rowsPage6)
      rowsClear.NumberFormat = "General"
      With rowsClear.Interior
          .Pattern = xlNone
          .TintAndShade = 0
          .PatternTintAndShade = 0
      End With
      rowsClear.Borders(xlDiagonalDown).LineStyle = xlNone
      rowsClear.Borders(xlDiagonalUp).LineStyle = xlNone
      rowsClear.Borders(xlEdgeLeft).LineStyle = xlNone
      rowsClear.Borders(xlEdgeTop).LineStyle = xlNone
      rowsClear.Borders(xlEdgeBottom).LineStyle = xlNone
      rowsClear.Borders(xlEdgeRight).LineStyle = xlNone
      rowsClear.Borders(xlInsideVertical).LineStyle = xlNone
      rowsClear.Borders(xlInsideHorizontal).LineStyle = xlNone
      With Worksheets(2).Range("A477:J477").Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      rowsClear.ClearContents
    End If
    ' wyczyść stronę 7 oferty '
    Dim rowsPage7 As Long
    rowsPage7 = 566
    For Each i In Worksheets(2).Range("A567:A10000")
      If i <> "" Then
        rowsPage7 = rowsPage7 + 1
      Else
        Exit For
      End If
    Next i
    If rowsPage7 > 566 Then
      Set rowsClear = Worksheets(2).Range("A567:J" & rowsPage7)
      rowsClear.NumberFormat = "General"
      With rowsClear.Interior
          .Pattern = xlNone
          .TintAndShade = 0
          .PatternTintAndShade = 0
      End With
      rowsClear.Borders(xlDiagonalDown).LineStyle = xlNone
      rowsClear.Borders(xlDiagonalUp).LineStyle = xlNone
      rowsClear.Borders(xlEdgeLeft).LineStyle = xlNone
      rowsClear.Borders(xlEdgeTop).LineStyle = xlNone
      rowsClear.Borders(xlEdgeBottom).LineStyle = xlNone
      rowsClear.Borders(xlEdgeRight).LineStyle = xlNone
      rowsClear.Borders(xlInsideVertical).LineStyle = xlNone
      rowsClear.Borders(xlInsideHorizontal).LineStyle = xlNone
      With Worksheets(2).Range("A567:J567").Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      rowsClear.ClearContents
    End If
    ' wyczyść stronę 8 oferty '
    Dim rowsPage8 As Long
    rowsPage8 = 656
    For Each i In Worksheets(2).Range("A657:A10000")
      If i <> "" Then
        rowsPage8 = rowsPage8 + 1
      Else
        Exit For
      End If
    Next i
    If rowsPage8 > 656 Then
      Set rowsClear = Worksheets(2).Range("A657:J" & rowsPage8)
      rowsClear.NumberFormat = "General"
      With rowsClear.Interior
          .Pattern = xlNone
          .TintAndShade = 0
          .PatternTintAndShade = 0
      End With
      rowsClear.Borders(xlDiagonalDown).LineStyle = xlNone
      rowsClear.Borders(xlDiagonalUp).LineStyle = xlNone
      rowsClear.Borders(xlEdgeLeft).LineStyle = xlNone
      rowsClear.Borders(xlEdgeTop).LineStyle = xlNone
      rowsClear.Borders(xlEdgeBottom).LineStyle = xlNone
      rowsClear.Borders(xlEdgeRight).LineStyle = xlNone
      rowsClear.Borders(xlInsideVertical).LineStyle = xlNone
      rowsClear.Borders(xlInsideHorizontal).LineStyle = xlNone
      With Worksheets(2).Range("A657:J657").Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      rowsClear.ClearContents
    End If
    ' KONIEC czyszczenia oferty '

    Dim rowInsertCount As Long
    rowInsertCount = 27

    ' dodaj zaznaczone pozycje do oferty '
    For Each i In markedRowsArray
      
        Dim saveStartPosition As Long
        saveStartPosition = rowInsertCount

      ' From Qutoe list '
      Dim materialCost As String
      materialCost = Worksheets(1).Range("M" & CStr(i) & ":M" & CStr(i)).Value
      materialCost = Replace(materialCost, ",", ".")
      Dim workCost As String
      workCost = Worksheets(1).Range("Z" & CStr(i) & ":Z" & CStr(i)).Value
      workCost = Replace(workCost, ",", ".")
      Dim affoCost As String
      affoCost = Worksheets(1).Range("K" & CStr(i) & ":K" & CStr(i)).Value
      affoCost = Replace(affoCost, ",", ".")
      Dim subcontractorCost As String
      subcontractorCost = Worksheets(1).Range("L" & CStr(i) & ":L" & CStr(i)).Value
      subcontractorCost = Replace(subcontractorCost, ",", ".")

      ' From Settings '
      Dim currecyConvert As String
      currecyConvert = "M9"
      Dim transportCost As String
      transportCost = "M8"
      Dim materialMultipier As String
      materialMultipier = "M11"
      Dim workMultiplier As String
      workMultiplier = "M10"
      Dim profitCost As String
      profitCost = "M12"

      Dim stepOne As String
      stepOne = workCost & "+" & materialCost
      Dim stepTwo As String
      stepTwo = "(" & stepOne & ")*1.01"
      Dim stepThree As String
      stepThree = "(" & stepTwo & ")*1.05"

      ' skopiuj wartości '
      Worksheets(2).Range("A" & rowInsertCount).Value = Worksheets(1).Range("A" & CStr(i) & ":A" & CStr(i)).Value ' part number
      Worksheets(2).Range("B" & rowInsertCount).Value = Worksheets(1).Range("B" & CStr(i) & ":B" & CStr(i)).Value ' part name
      Worksheets(2).Range("C" & rowInsertCount).Value = Worksheets(1).Range("H" & CStr(i) & ":H" & CStr(i)).Value ' annual volume
      Worksheets(2).Range("D" & rowInsertCount).Value = Worksheets(1).Range("I" & CStr(i) & ":I" & CStr(i)).Value ' moq
      Worksheets(2).Range("E" & rowInsertCount).Value = "=(" & materialCost & "+" & subcontractorCost & ")/" & currecyConvert ' material cost + subcontractor cost
      


      Worksheets(2).Range("F" & rowInsertCount).Value = "=(" & workCost & "+((" & stepOne & ")*1%)+((" & stepTwo & ")*4%)+((" & stepThree & ")*5%))/" & currecyConvert  ' work cost + affo + random 5%



      'Worksheets(2).Range("G" & rowInsertCount).Value = "=(E" & rowInsertCount & "+F" & rowInsertCount & ")*" & affoCost ' affo
      Worksheets(2).Range("G" & rowInsertCount).Value = "=(E" & rowInsertCount & "+F" & rowInsertCount & ")*" & profitCost ' profit
      Worksheets(2).Range("H" & rowInsertCount).Value = "=(E" & rowInsertCount & "+F" & rowInsertCount & ")*" & transportCost ' transport
      Worksheets(2).Range("I" & rowInsertCount).Value = "=(E" & rowInsertCount & "+F" & rowInsertCount & "+G" & rowInsertCount & "+H" & rowInsertCount & ")" ' final price
      Worksheets(2).Range("J" & rowInsertCount).Value = "=((E" & rowInsertCount & "*" & materialMultipier & ")+(F" & rowInsertCount & "*" & workMultiplier & ")+G" & rowInsertCount & "+H" & rowInsertCount & ")" ' initial sample
      
      ' formatowanie nowej pozycji w ofercie '
      Set row27 = Worksheets(2).Range("A" & rowInsertCount & ":J" & rowInsertCount)

      Worksheets(2).Range("A" & rowInsertCount).Font.Bold = True
      Worksheets(2).Range("B" & rowInsertCount & ":C" & rowInsertCount).Font.Bold = False
      Worksheets(2).Range("D" & rowInsertCount).Font.Bold = True
      Worksheets(2).Range("E" & rowInsertCount & ":H" & rowInsertCount).Font.Bold = False
      Worksheets(2).Range("I" & rowInsertCount).Font.Bold = True
      Worksheets(2).Range("J" & rowInsertCount).Font.Bold = False
      
      Set currencyRange = Worksheets(2).Range("E" & rowInsertCount & ":J" & rowInsertCount)
      Set pickedCurrency = Worksheets(2).Range("N9")
      
      If pickedCurrency.Value = "EUR" Then
        currencyRange.NumberFormat = "_-[$€-x-euro2] * #,##0.00_-;-[$€-x-euro2] * #,##0.00_-;_-[$€-x-euro2] * ""-""??_-;_-@_-"
      ElseIf pickedCurrency.Value = "GBP" Then
        currencyRange.NumberFormat = "_-* #,##0.00 [$GBP]_-;-* #,##0.00 [$GBP]_-;_-* ""-""?? [$GBP]_-;_-@_-"
      ElseIf pickedCurrency.Value = "CHF" Then
        currencyRange.NumberFormat = "_-* #,##0.00 [$CHF]_-;-* #,##0.00 [$CHF]_-;_-* ""-""?? [$CHF]_-;_-@_-"
      ElseIf pickedCurrency.Value = "USD" Then
        currencyRange.NumberFormat = "_-* #,##0.00 [$USD]_-;-* #,##0.00 [$USD]_-;_-* ""-""?? [$USD]_-;_-@_-"
      ElseIf pickedCurrency.Value = "PLN" Then
        currencyRange.NumberFormat = "_-* #,##0.00 [$PLN]_-;-* #,##0.00 [$PLN]_-;_-* ""-""?? [$PLN]_-;_-@_-"
      ElseIf pickedCurrency.Value = "CNY" Then
        currencyRange.NumberFormat = "_-* #,##0.00 [$CNY]_-;-* #,##0.00 [$CNY]_-;_-* ""-""?? [$CNY]_-;_-@_-"
      ElseIf pickedCurrency.Value = "SEK" Then
        currencyRange.NumberFormat = "_-* #,##0.00 [$SEK]_-;-* #,##0.00 [$SEK]_-;_-* ""-""?? [$SEK]_-;_-@_-"
      Else
        MsgBox "Currency not set in file. Please contact file administrator."
        Exit Sub
      End If

      With row27.Font
          .Name = "Calibri"
          .Size = 14
          .Strikethrough = False
          .Superscript = False
          .Subscript = False
          .OutlineFont = False
          .Shadow = False
          .Underline = xlUnderlineStyleNone
          .ThemeColor = xlThemeColorLight1
          .TintAndShade = 0
          .ThemeFont = xlThemeFontMinor
      End With
      With row27
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .IndentLevel = 0
          .ShrinkToFit = False
          .ReadingOrder = xlContext
          .MergeCells = False
      End With
      With row27.Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With row27.Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With row27.Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With row27.Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With row27.Borders(xlInsideVertical)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With row27.Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      ' koniec formatowania '
      ' Set only total price for simplified version '
      If Worksheets(2).Range("M13").Value = "Yes" Then
        Worksheets(2).Range("H" & rowInsertCount).Copy
        Worksheets(2).Range("H" & rowInsertCount).PasteSpecial xlPasteValues
        Worksheets(2).Range("I" & rowInsertCount).Copy
        Worksheets(2).Range("I" & rowInsertCount).PasteSpecial xlPasteValues
        Worksheets(2).Range("J" & rowInsertCount).Copy
        Worksheets(2).Range("J" & rowInsertCount).PasteSpecial xlPasteValues
        Worksheets(2).Range("E" & rowInsertCount).NumberFormat = "@"
        Worksheets(2).Range("E" & rowInsertCount).Value = Worksheets(1).Range("AA" & CStr(i) & ":AA" & CStr(i)).Value
        Worksheets(2).Range("E" & rowInsertCount).Font.Bold = True
        Worksheets(2).Range("F" & rowInsertCount).NumberFormat = "General"
        Worksheets(2).Range("F" & rowInsertCount).FormulaR1C1 = "=CONCATENATE(YEAR(TODAY()),""-"",MONTH(TODAY()),""-"",DAY(TODAY()))"
        Worksheets(2).Range("F" & rowInsertCount).Copy
        Worksheets(2).Range("F" & rowInsertCount).PasteSpecial xlPasteValues
        Worksheets(2).Range("G" & rowInsertCount).NumberFormat = "General"
        Worksheets(2).Range("G" & rowInsertCount).FormulaR1C1 = "=CONCATENATE(YEAR(TODAY()),""-"",MONTH(TODAY())+1,""-"",DAY(TODAY()))"
        Worksheets(2).Range("G" & rowInsertCount).Copy
        Worksheets(2).Range("G" & rowInsertCount).PasteSpecial xlPasteValues
      End If

      If rowInsertCount = 71 Then
        rowInsertCount = 117
      ElseIf rowInsertCount = 161 Then
        rowInsertCount = 207
      ElseIf rowInsertCount = 251 Then
        rowInsertCount = 297
      ElseIf rowInsertCount = 341 Then
        rowInsertCount = 387
      ElseIf rowInsertCount = 431 Then
        rowInsertCount = 477
      ElseIf rowInsertCount = 521 Then
        rowInsertCount = 567
      ElseIf rowInsertCount = 611 Then
        rowInsertCount = 657
      ElseIf rowInsertCount > 700 Then
        MsgBox "Page limit rached. Mark less parts and try again or contact file administrator."
        Exit Sub
      End If

      Dim extraInfoCounter As Long
      extraInfoCounter = 0
      Dim makeLinePlease As Integer
      makeLinePlease = 1
      ' Insert additional costs and engineering info '
      Dim escapeLoop As Integer
      escapeLoop = 0
      For Each j In Worksheets("Data - extra info").Range("J1:J100000")
        extraInfoCounter = extraInfoCounter + 1


        If j = Worksheets(1).Range("AA" & CStr(i) & ":AA" & CStr(i)).Value Then

            rowInsertCount = rowInsertCount + 1
            If rowInsertCount = 71 Then
              rowInsertCount = 117
            ElseIf rowInsertCount = 161 Then
              rowInsertCount = 207
            ElseIf rowInsertCount = 251 Then
              rowInsertCount = 297
            ElseIf rowInsertCount = 341 Then
              rowInsertCount = 387
            ElseIf rowInsertCount = 431 Then
              rowInsertCount = 477
            ElseIf rowInsertCount = 521 Then
              rowInsertCount = 567
            ElseIf rowInsertCount = 611 Then
              rowInsertCount = 657
            ElseIf rowInsertCount > 700 Then
              MsgBox "Page limit rached. Mark less parts and try again or contact file administrator."
              Exit Sub
            End If
            ' Copy data '

            Worksheets(2).Range("A" & rowInsertCount).Value = Worksheets("Data - extra info").Range("F" & extraInfoCounter & ":F" & extraInfoCounter).Value ' narzedzie '
            
            If Worksheets("Data - extra info").Range("E" & extraInfoCounter & ":E" & extraInfoCounter).Value = 4 Then
                Worksheets(2).Range("A" & rowInsertCount).Value = "x"
                Worksheets(2).Range("B" & rowInsertCount).Value = Worksheets("Data - extra info").Range("G" & extraInfoCounter & ":G" & extraInfoCounter).Value ' text
            End If
            If Worksheets("Data - extra info").Range("E" & extraInfoCounter & ":E" & extraInfoCounter).Value = 2 Then
                Worksheets(2).Range("B" & rowInsertCount).Value = Worksheets("Data - extra info").Range("G" & extraInfoCounter & ":G" & extraInfoCounter).Value ' text
                Worksheets(2).Range("D" & rowInsertCount).Value = CStr(Worksheets("Data - extra info").Range("H" & extraInfoCounter & ":H" & extraInfoCounter).Value) ' ilosc
                Worksheets(2).Range("C" & rowInsertCount).Value = Worksheets("Data - extra info").Range("I" & extraInfoCounter & ":I" & extraInfoCounter).Value / Worksheets(2).Range("M9:M9").Value ' kwota
                Worksheets(2).Range("E" & rowInsertCount).Value = Worksheets(2).Range("C" & rowInsertCount).Value * Worksheets(2).Range("D" & rowInsertCount).Value
            End If

            

            ' Formatting '
            Worksheets(2).Range("A" & rowInsertCount).Font.Bold = False
            Worksheets(2).Range("B" & rowInsertCount & ":C" & rowInsertCount).Font.Bold = False
            Worksheets(2).Range("D" & rowInsertCount).Font.Bold = True
            Worksheets(2).Range("E" & rowInsertCount & ":H" & rowInsertCount).Font.Bold = False
            Worksheets(2).Range("I" & rowInsertCount).Font.Bold = False
            Worksheets(2).Range("J" & rowInsertCount).Font.Bold = False
            Set row27 = Worksheets(2).Range("A" & rowInsertCount & ":E" & rowInsertCount)
            With row27.Font
                .Name = "Calibri"
                .Size = 14
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With row27
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            With row27.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With row27.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With row27.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With row27.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With row27.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With row27.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With



            Worksheets(2).Range("B" & rowInsertCount & ":B" & rowInsertCount).HorizontalAlignment = xlCenter

            If Worksheets("Data - extra info").Range("E" & extraInfoCounter & ":E" & extraInfoCounter).Value = 4 Then
                Worksheets(2).Range("A" & rowInsertCount & ":B" & rowInsertCount).HorizontalAlignment = xlLeft
                With Worksheets(2).Range("A" & rowInsertCount & ":A" & rowInsertCount).Font
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
                Worksheets(2).Range("A" & rowInsertCount).Borders(xlEdgeRight).LineStyle = xlNone
                Worksheets(2).Range("B" & rowInsertCount).Borders(xlEdgeRight).LineStyle = xlNone
                Worksheets(2).Range("C" & rowInsertCount).Borders(xlEdgeRight).LineStyle = xlNone
                Worksheets(2).Range("D" & rowInsertCount).Borders(xlEdgeRight).LineStyle = xlNone
                
            End If



            Set toolCostCurrency = Worksheets(2).Range("C" & rowInsertCount & ":C" & rowInsertCount)
            Set toolCostCurrency2 = Worksheets(2).Range("E" & rowInsertCount & ":E" & rowInsertCount)

            
            If pickedCurrency.Value = "EUR" Then
              toolCostCurrency.NumberFormat = "_-[$€-x-euro2] * #,##0.00_-;-[$€-x-euro2] * #,##0.00_-;_-[$€-x-euro2] * ""-""??_-;_-@_-"
            ElseIf pickedCurrency.Value = "GBP" Then
              toolCostCurrency.NumberFormat = "_-* #,##0.00 [$GBP]_-;-* #,##0.00 [$GBP]_-;_-* ""-""?? [$GBP]_-;_-@_-"
            ElseIf pickedCurrency.Value = "CHF" Then
              toolCostCurrency.NumberFormat = "_-* #,##0.00 [$CHF]_-;-* #,##0.00 [$CHF]_-;_-* ""-""?? [$CHF]_-;_-@_-"
            ElseIf pickedCurrency.Value = "USD" Then
              toolCostCurrency.NumberFormat = "_-* #,##0.00 [$USD]_-;-* #,##0.00 [$USD]_-;_-* ""-""?? [$USD]_-;_-@_-"
            ElseIf pickedCurrency.Value = "PLN" Then
              toolCostCurrency.NumberFormat = "_-* #,##0.00 [$PLN]_-;-* #,##0.00 [$PLN]_-;_-* ""-""?? [$PLN]_-;_-@_-"
            ElseIf pickedCurrency.Value = "CNY" Then
              toolCostCurrency.NumberFormat = "_-* #,##0.00 [$CNY]_-;-* #,##0.00 [$CNY]_-;_-* ""-""?? [$CNY]_-;_-@_-"
            ElseIf pickedCurrency.Value = "SEK" Then
              toolCostCurrency.NumberFormat = "_-* #,##0.00 [$SEK]_-;-* #,##0.00 [$SEK]_-;_-* ""-""?? [$SEK]_-;_-@_-"
            Else
              MsgBox "Currency not set in file. Please contact file administrator."
              Exit Sub
            End If


            If pickedCurrency.Value = "EUR" Then
              toolCostCurrency2.NumberFormat = "_-[$€-x-euro2] * #,##0.00_-;-[$€-x-euro2] * #,##0.00_-;_-[$€-x-euro2] * ""-""??_-;_-@_-"
            ElseIf pickedCurrency.Value = "GBP" Then
              toolCostCurrency2.NumberFormat = "_-* #,##0.00 [$GBP]_-;-* #,##0.00 [$GBP]_-;_-* ""-""?? [$GBP]_-;_-@_-"
            ElseIf pickedCurrency.Value = "CHF" Then
              toolCostCurrency2.NumberFormat = "_-* #,##0.00 [$CHF]_-;-* #,##0.00 [$CHF]_-;_-* ""-""?? [$CHF]_-;_-@_-"
            ElseIf pickedCurrency.Value = "USD" Then
              toolCostCurrency2.NumberFormat = "_-* #,##0.00 [$USD]_-;-* #,##0.00 [$USD]_-;_-* ""-""?? [$USD]_-;_-@_-"
            ElseIf pickedCurrency.Value = "PLN" Then
              toolCostCurrency2.NumberFormat = "_-* #,##0.00 [$PLN]_-;-* #,##0.00 [$PLN]_-;_-* ""-""?? [$PLN]_-;_-@_-"
            ElseIf pickedCurrency.Value = "CNY" Then
              toolCostCurrency2.NumberFormat = "_-* #,##0.00 [$CNY]_-;-* #,##0.00 [$CNY]_-;_-* ""-""?? [$CNY]_-;_-@_-"
            ElseIf pickedCurrency.Value = "SEK" Then
              toolCostCurrency2.NumberFormat = "_-* #,##0.00 [$SEK]_-;-* #,##0.00 [$SEK]_-;_-* ""-""?? [$SEK]_-;_-@_-"
            Else
              MsgBox "Currency not set in file. Please contact file administrator."
              Exit Sub
            End If

            ' Make it more than 0 so it will stop after all info is copied '
            escapeLoop = escapeLoop + 1
        ElseIf escapeLoop > 0 Then
          Exit For
        End If
      Next j


        Dim howManyToAdd As Long
        howManyToAdd = saveStartPosition
    
    If Worksheets(2).Range("D" & saveStartPosition + 1) <> "" Then
      For Each k In Worksheets(2).Range("D" & saveStartPosition & ":D100000")
        If k = "" Then
          Exit For
        End If
        howManyToAdd = howManyToAdd + 1
      Next k
      Worksheets(2).Range("J" & saveStartPosition + 1).Value = "=SUM(E" & saveStartPosition + 1 & ":E" & howManyToAdd - 1 & ")"
      
      Worksheets(2).Range("I" & saveStartPosition + 1).Value = "Investment cost:"

      With Worksheets(2).Range("I" & saveStartPosition + 1 & ":" & "J" & saveStartPosition + 1).Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Worksheets(2).Range("I" & saveStartPosition + 1 & ":" & "J" & saveStartPosition + 1).Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Worksheets(2).Range("I" & saveStartPosition + 1 & ":" & "J" & saveStartPosition + 1).Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Worksheets(2).Range("I" & saveStartPosition + 1 & ":" & "J" & saveStartPosition + 1).Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Worksheets(2).Range("I" & saveStartPosition + 1 & ":" & "J" & saveStartPosition + 1).Borders(xlInsideVertical)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Worksheets(2).Range("I" & saveStartPosition + 1 & ":" & "J" & saveStartPosition + 1).Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      
      Set investCurrency = Worksheets(2).Range("J" & saveStartPosition + 1)
      
      If pickedCurrency.Value = "EUR" Then
              investCurrency.NumberFormat = "_-[$€-x-euro2] * #,##0.00_-;-[$€-x-euro2] * #,##0.00_-;_-[$€-x-euro2] * ""-""??_-;_-@_-"
            ElseIf pickedCurrency.Value = "GBP" Then
              investCurrency.NumberFormat = "_-* #,##0.00 [$GBP]_-;-* #,##0.00 [$GBP]_-;_-* ""-""?? [$GBP]_-;_-@_-"
            ElseIf pickedCurrency.Value = "CHF" Then
              investCurrency.NumberFormat = "_-* #,##0.00 [$CHF]_-;-* #,##0.00 [$CHF]_-;_-* ""-""?? [$CHF]_-;_-@_-"
            ElseIf pickedCurrency.Value = "USD" Then
              investCurrency.NumberFormat = "_-* #,##0.00 [$USD]_-;-* #,##0.00 [$USD]_-;_-* ""-""?? [$USD]_-;_-@_-"
            ElseIf pickedCurrency.Value = "PLN" Then
              investCurrency.NumberFormat = "_-* #,##0.00 [$PLN]_-;-* #,##0.00 [$PLN]_-;_-* ""-""?? [$PLN]_-;_-@_-"
            ElseIf pickedCurrency.Value = "CNY" Then
              investCurrency.NumberFormat = "_-* #,##0.00 [$CNY]_-;-* #,##0.00 [$CNY]_-;_-* ""-""?? [$CNY]_-;_-@_-"
            ElseIf pickedCurrency.Value = "SEK" Then
              investCurrency.NumberFormat = "_-* #,##0.00 [$SEK]_-;-* #,##0.00 [$SEK]_-;_-* ""-""?? [$SEK]_-;_-@_-"
            Else
              MsgBox "Currency not set in file. Please contact file administrator."
              Exit Sub
            End If
      
    End If
        rowInsertCount = rowInsertCount + 1
    Next i

    ' get QUOTATION no. '
    Dim quoteRefLocation As Long
    quoteRefLocation = 1
    For Each i In Worksheets(1).Range("AC2:AC100000")
      If i = "" Then
        quoteRefLocation = quoteRefLocation + 1
      Else
        quoteRefLocation = quoteRefLocation + 1
        Worksheets(2).Range("I2").Value = Worksheets(1).Range("F" & quoteRefLocation).Value
        Exit For
      End If
    Next i

    MsgBox "Offer complete."

End Sub

Sub saveQuotePDF()

    ' Figure out how many pages to print
    Dim endOfOffer As String

    ' Start saving
    Dim clientName As String
    clientName = CStr(Worksheets(2).Range("A9").Value)
    Dim quotatnionNo As String
    quotatnionNo = CStr(Worksheets(2).Range("I2").Value)
    Dim offerDate As String
    offerDate = CStr(Worksheets(2).Range("I4").Value)
    Dim clientReference As String
    clientReference = CStr(Worksheets(2).Range("I7").Value)
    Dim projectReference As String
    projectReference = CStr(Worksheets(2).Range("C18").Value)
    
    Dim fileName As String
    fileName = clientName + " - " + clientReference + " - " + offerDate + " - " + quotatnionNo + " - " + projectReference

    Dim localPathAndName As String
    localPathAndName = Environ("USERPROFILE") & "\Desktop\" & fileName
    
    Dim serverPathAndName As String
    serverPathAndName = "C:\Users\XxX\Google Drive\_GENERATOR\XxX\Sales offers\" & fileName

    Dim pageCount As Integer
    pageCount = Worksheets(2).Range("M15").Value

    Dim QuoteWB As Workbook
    Set QuoteWB = ActiveWorkbook
    
    QuoteWB.Worksheets(2).Copy
    Set NewBook = ActiveWorkbook

    NewBook.Worksheets(1).Range("A:J").Copy
    NewBook.Worksheets(1).Range("A:J").PasteSpecial xlPasteValues
    NewBook.Worksheets(1).Range("K:P").Delete Shift:=xlRight

    If pageCount = 1 Then
      endOfOffer = "J100" 'wiekszy koniec o 10 zeby pokryc general terms'
      NewBook.Worksheets(1).Range("91:720").Delete Shift:=xlUp
    ElseIf pageCount = 2 Then
      endOfOffer = "J190"
      NewBook.Worksheets(1).Range("181:720").Delete Shift:=xlUp
    ElseIf pageCount = 3 Then
      endOfOffer = "J280"
      NewBook.Worksheets(1).Range("271:720").Delete Shift:=xlUp
    ElseIf pageCount = 4 Then
      endOfOffer = "J370"
      NewBook.Worksheets(1).Range("361:720").Delete Shift:=xlUp
    ElseIf pageCount = 5 Then
      endOfOffer = "J460"
      NewBook.Worksheets(1).Range("451:720").Delete Shift:=xlUp
    ElseIf pageCount = 6 Then
      endOfOffer = "J550"
      NewBook.Worksheets(1).Range("541:720").Delete Shift:=xlUp
    ElseIf pageCount = 7 Then
      endOfOffer = "J640"
      NewBook.Worksheets(1).Range("631:720").Delete Shift:=xlUp
    ElseIf pageCount = 8 Then
      endOfOffer = "J730"
    Else
        MsgBox "Incorrect page count."
        Exit Sub
    End If

    Dim shape As Excel.shape
    Dim shapeCount As Integer
    shapeCount = 1

    For Each shape In ActiveSheet.Shapes
        If shapeCount > 1 Then
        shape.Delete
        End If
        shapeCount = shapeCount + 1
    Next
    
    ' CHANGE LAYOUT START '
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$A$1:$" + CStr(endOfOffer)
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    ' CHANGE LAYOUT FINISH '

    Dim offerEndRow As String
    offerEndRow = CStr(CInt(Replace(endOfOffer, "J", "")) - 9)

    ActiveWindow.View = xlPageBreakPreview

    NewBook.Worksheets(1).PageSetup.PrintArea = "$A$1:$" + CStr(endOfOffer)

    If pageCount = 1 Then
      Set ActiveSheet.HPageBreaks(1).Location = Range("A" & offerEndRow)
    ElseIf pageCount = 2 Then
      Set ActiveSheet.HPageBreaks(2).Location = Range("A" & offerEndRow)
    ElseIf pageCount = 3 Then
      Set ActiveSheet.HPageBreaks(3).Location = Range("A" & offerEndRow)
    ElseIf pageCount = 4 Then
      Set ActiveSheet.HPageBreaks(4).Location = Range("A" & offerEndRow)
    ElseIf pageCount = 5 Then
      Set ActiveSheet.HPageBreaks(5).Location = Range("A" & offerEndRow)
    ElseIf pageCount = 6 Then
      Set ActiveSheet.HPageBreaks(6).Location = Range("A" & offerEndRow)
    ElseIf pageCount = 7 Then
      Set ActiveSheet.HPageBreaks(7).Location = Range("A" & offerEndRow)
    ElseIf pageCount = 8 Then
      Set ActiveSheet.HPageBreaks(8).Location = Range("A" & offerEndRow)
        MsgBox "Incorrect page count."
        Exit Sub
    End If

    ActiveWindow.View = xlNormalView
    
    

On Error GoTo meh
      ' Check if file exists, if so delete and save new
    If Dir(localPathAndName & ".pdf", vbDirectory) <> "" Then
        ' Remove readonly
        SetAttr localPathAndName & ".pdf", vbNormal
        ' Delete file
        Kill localPathAndName & ".pdf"
        ' Save new
        Worksheets(1).ExportAsFixedFormat Type:=xlTypePDF, fileName:=localPathAndName & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Else
        Worksheets(1).ExportAsFixedFormat Type:=xlTypePDF, fileName:=localPathAndName & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    End If
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    NewBook.Close
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' SAVE COPY TO GOOGLE DRIVE '
    Call save_quote_excel_on_server
    MsgBox "Saved."
    
    Exit Sub
    
meh:
    MsgBox "Cannot save, file opened by another program."
    
End Sub

Sub save_quote_excel()

On Error GoTo meh

    Dim clientName As String
    clientName = CStr(Worksheets(2).Range("A9").Value)
    Dim quotatnionNo As String
    quotatnionNo = CStr(Worksheets(2).Range("I2").Value)
    Dim offerDate As String
    offerDate = CStr(Worksheets(2).Range("I4").Value)
    Dim clientReference As String
    clientReference = CStr(Worksheets(2).Range("I7").Value)
    Dim projectReference As String
    projectReference = CStr(Worksheets(2).Range("C18").Value)

    Dim fileName As String
    fileName = clientName + " - " + clientReference + " - " + offerDate + " - " + quotatnionNo + " - " + projectReference

    Dim localPathAndName As String
    localPathAndName = Environ("USERPROFILE") & "\Desktop\" & fileName
    
    Dim serverPathAndName As String
    serverPathAndName = "C:\Users\XxX\Google Drive\_GENERATOR\XxX\Sales offers\" & fileName

    ' Figure out how many pages to print
    Dim endOfOffer As String
    Dim pageCount As Integer
    pageCount = Worksheets(2).Range("M15").Value
    
    Dim QuoteWB As Workbook
    Set QuoteWB = ActiveWorkbook
    
    QuoteWB.Worksheets(2).Copy
    Set NewBook = ActiveWorkbook
    
    NewBook.Worksheets(1).Range("A:J").Copy
    NewBook.Worksheets(1).Range("A:J").PasteSpecial xlPasteValues
    NewBook.Worksheets(1).Range("K:P").Delete Shift:=xlRight
     
    If pageCount = 1 Then
        endOfOffer = "J100" 'wiekszy koniec o 100 zeby pokryc general terms'
        NewBook.Worksheets(1).Range("91:720").Delete Shift:=xlUp
    ElseIf pageCount = 2 Then
        endOfOffer = "J190"
        NewBook.Worksheets(1).Range("181:720").Delete Shift:=xlUp
        NewBook.Worksheets(1).Range("72:116").Delete Shift:=xlUp '1
    ElseIf pageCount = 3 Then
        endOfOffer = "J280"
        NewBook.Worksheets(1).Range("271:720").Delete Shift:=xlUp
        NewBook.Worksheets(1).Range("162:206").Delete Shift:=xlUp '2
        NewBook.Worksheets(1).Range("72:116").Delete Shift:=xlUp '1
    ElseIf pageCount = 4 Then
        endOfOffer = "J370"
        NewBook.Worksheets(1).Range("361:720").Delete Shift:=xlUp
        NewBook.Worksheets(1).Range("252:296").Delete Shift:=xlUp '3
        NewBook.Worksheets(1).Range("162:206").Delete Shift:=xlUp '2
        NewBook.Worksheets(1).Range("72:116").Delete Shift:=xlUp '1
    ElseIf pageCount = 5 Then
        endOfOffer = "J460"
        NewBook.Worksheets(1).Range("451:720").Delete Shift:=xlUp
        NewBook.Worksheets(1).Range("342:386").Delete Shift:=xlUp '4
        NewBook.Worksheets(1).Range("252:296").Delete Shift:=xlUp '3
        NewBook.Worksheets(1).Range("162:206").Delete Shift:=xlUp '2
        NewBook.Worksheets(1).Range("72:116").Delete Shift:=xlUp '1
    ElseIf pageCount = 6 Then
        endOfOffer = "J550"
        NewBook.Worksheets(1).Range("541:720").Delete Shift:=xlUp
        NewBook.Worksheets(1).Range("432:476").Delete Shift:=xlUp '5
        NewBook.Worksheets(1).Range("342:386").Delete Shift:=xlUp '4
        NewBook.Worksheets(1).Range("252:296").Delete Shift:=xlUp '3
        NewBook.Worksheets(1).Range("162:206").Delete Shift:=xlUp '2
        NewBook.Worksheets(1).Range("72:116").Delete Shift:=xlUp '1
    ElseIf pageCount = 7 Then
        endOfOffer = "J640"
        NewBook.Worksheets(1).Range("631:720").Delete Shift:=xlUp
        NewBook.Worksheets(1).Range("522:566").Delete Shift:=xlUp '6
        NewBook.Worksheets(1).Range("432:476").Delete Shift:=xlUp '5
        NewBook.Worksheets(1).Range("342:386").Delete Shift:=xlUp '4
        NewBook.Worksheets(1).Range("252:296").Delete Shift:=xlUp '3
        NewBook.Worksheets(1).Range("162:206").Delete Shift:=xlUp '2
        NewBook.Worksheets(1).Range("72:116").Delete Shift:=xlUp '1
    ElseIf pageCount = 8 Then
        endOfOffer = "J730"
        NewBook.Worksheets(1).Range("612:656").Delete Shift:=xlUp '7
        NewBook.Worksheets(1).Range("522:566").Delete Shift:=xlUp '6
        NewBook.Worksheets(1).Range("432:476").Delete Shift:=xlUp '5
        NewBook.Worksheets(1).Range("342:386").Delete Shift:=xlUp '4
        NewBook.Worksheets(1).Range("252:296").Delete Shift:=xlUp '3
        NewBook.Worksheets(1).Range("162:206").Delete Shift:=xlUp '2
        NewBook.Worksheets(1).Range("72:116").Delete Shift:=xlUp '1
    Else
        MsgBox "Incorrect page count."
        Exit Sub
    End If
    
    Dim offerRange As Integer
    offerRange = 26
    Dim validityPeriod As Integer
    validityPeriod = 26

    NewBook.Worksheets(1).Range(offerRange & ":" & validityPeriod - 3).Delete Shift:=xlUp '1

    Dim offerEndInt As Long
    offerEndInt = 0

    For Each i In NewBook.Worksheets(1).Range("A1:A10000")
      offerEndInt = offerEndInt + 1
      If i = "end" Then
        Exit For
      End If
    Next i

    endOfOffer = "J" & CStr(offerEndInt + 14)

   
    
    

    For Each i In NewBook.Worksheets(1).Range("A27:A10000")
      If i <> "" Then
        offerRange = offerRange + 1
      Else
        offerRange = offerRange + 1
        Exit For
      End If
    Next i

    For Each i In NewBook.Worksheets(1).Range("A27:A10000")
      validityPeriod = validityPeriod + 1
      If i = "Period of validity" Then
        validityPeriod = validityPeriod + 1
        Exit For
      End If
    Next i

    



    ' Start saving
    
    Dim shape As Excel.shape
    Dim shapeCount As Integer
    shapeCount = 1

    For Each shape In ActiveSheet.Shapes
        If shapeCount > 1 Then
        shape.Delete
        End If
        shapeCount = shapeCount + 1
    Next


    
    ' CHANGE LAYOUT START '
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$A$1:$" + CStr(endOfOffer)
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    ' CHANGE LAYOUT FINISH '

    Dim offerEndRow As String
    offerEndRow = CStr(CInt(Replace(endOfOffer, "J", "")) - 10)

    ActiveWindow.View = xlPageBreakPreview

    Set ActiveSheet.HPageBreaks(1).Location = Range("A" & offerEndRow)

    NewBook.Worksheets(1).PageSetup.PrintArea = "$A$1:$" + CStr(endOfOffer)
    
    ActiveWindow.View = xlNormalView
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    ' Save to Desktop
    NewBook.SaveAs fileName:=localPathAndName & ".xlsx", FileFormat:=51
    ' And to google drive
    NewBook.Close
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    ' SAVE COPY TO SERVER '
    Call save_quote_excel_on_server
  MsgBox "Saved."
Exit Sub

meh:
    MsgBox "Cannot save, file opened by another program."
    NewBook.Close
    
End Sub



Sub save_BOM_as_PDF()
  ' SAVING TO DESKTOP '
      ' Figure out how many pages to print
    Dim endOfBOM As Long
    endOfBOM = 9

    For Each i In Worksheets(3).Range("L10:L100000")
      If i <> "" Then
        endOfBOM = endOfBOM + 1
      Else
        Exit For
      End If
    Next i

    ' compensate for footer '
    endOfBOM = endOfBOM + 7

    ' Start saving
    Dim bomDate As String
    bomDate = CStr(Worksheets(3).Range("E2").Value)
    Dim harnessNumber As String
    harnessNumber = CStr(Worksheets(3).Range("E3").Value)
    Dim harnessMOQ As String
    harnessMOQ = CStr(Worksheets(3).Range("E4").Value)
    
    Dim fileName As String
    fileName = "BOM - " + harnessNumber + " - " + bomDate + " - " + harnessMOQ

    Dim localPathAndName As String
    localPathAndName = Environ("USERPROFILE") & "\Desktop\" & fileName
    
    Dim serverPathAndName As String
    serverPathAndName = "C:\Users\XxX\Google Drive\_GENERATOR\XxX\BOMs\" & fileName
    
    ' CHANGE LAYOUT START '
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$A$1:$L" + CStr(endOfBOM)
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    ' CHANGE LAYOUT FINISH '

    ActiveSheet.PageSetup.PrintArea = "$A$1:$L" + CStr(endOfBOM)

On Error GoTo meh
    ' Check if file exists, if so delete and save new - DESKTOP
    If Dir(localPathAndName & ".pdf", vbDirectory) <> "" Then
        ' Remove readonly
        SetAttr localPathAndName & ".pdf", vbNormal
        ' Delete file
        Kill localPathAndName & ".pdf"
        ' Save new on desktop
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=localPathAndName & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Else
        ' Save on desktop
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=localPathAndName & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    End If
    
    ' SAVE COPY ON SERVER '
    If Dir(serverPathAndName & ".pdf", vbDirectory) <> "" Then
        ' Remove readonly
        SetAttr serverPathAndName & ".pdf", vbNormal
        ' Delete file
        Kill serverPathAndName & ".pdf"
        ' Save new on desktop
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=serverPathAndName & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Else
        ' Save on desktop
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=serverPathAndName & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    End If
    
    MsgBox "Saved."
    
    Exit Sub
    
meh:
    MsgBox "Cannot save, file opened by another program."

End Sub
Sub calculate_BOM()

    Sheets("Customer BOM").Range("A9").ListObject.QueryTable.Refresh BackgroundQuery:=False

    Dim bom_length As Long
    bom_length = 8
    For Each i In Worksheets("Customer BOM").Range("A9:A10000")
        If i.Value <> "" Then
            bom_length = bom_length + 1
        End If
        If i.Value = "" Then
            Exit For
        End If
    Next i

    Dim line_to_delete As Long
    line_to_delete = 8
    Dim begin_delete As Long
    For Each i In Worksheets("Customer BOM").Range("L9:L" & bom_length)
        line_to_delete = line_to_delete + 1
        If i.Value = "DELETED" Then
            begin_delete = line_to_delete
            Exit For
        End If
    Next i

    If begin_delete > 1 Then
        Worksheets("Customer BOM").Range(begin_delete & ":" & bom_length).Delete Shift:=xlUp
    End If
    
  Dim rowCounter As Long
  rowCounter = 9
  For Each i In Worksheets(3).Range("A10:A100000")
    If i <> "" Then
      rowCounter = rowCounter + 1
    Else
      Exit For
    End If
  Next i

  Set pickedCurrency = Worksheets(3).Range("U4")

  For Each i In Worksheets(3).Range("L10:L" & rowCounter)

    i.Value = "=(((((RC[-3]*R4C5)+RC[-4])*(RC[-2]+1))*RC[-7])/R4C5)/R4C20"

    If pickedCurrency.Value = "EUR" Then
      i.NumberFormat = "_-[$€-x-euro2] * #,##0.00_-;-[$€-x-euro2] * #,##0.00_-;_-[$€-x-euro2] * ""-""??_-;_-@_-"
    ElseIf pickedCurrency.Value = "GBP" Then
      i.NumberFormat = "_-* #,##0.00 [$GBP]_-;-* #,##0.00 [$GBP]_-;_-* ""-""?? [$GBP]_-;_-@_-"
    ElseIf pickedCurrency.Value = "CHF" Then
      i.NumberFormat = "_-* #,##0.00 [$CHF]_-;-* #,##0.00 [$CHF]_-;_-* ""-""?? [$CHF]_-;_-@_-"
    ElseIf pickedCurrency.Value = "USD" Then
      i.NumberFormat = "_-* #,##0.00 [$USD]_-;-* #,##0.00 [$USD]_-;_-* ""-""?? [$USD]_-;_-@_-"
    ElseIf pickedCurrency.Value = "PLN" Then
      i.NumberFormat = "_-* #,##0.00 [$PLN]_-;-* #,##0.00 [$PLN]_-;_-* ""-""?? [$PLN]_-;_-@_-"
    ElseIf pickedCurrency.Value = "CNY" Then
      i.NumberFormat = "_-* #,##0.00 [$CNY]_-;-* #,##0.00 [$CNY]_-;_-* ""-""?? [$CNY]_-;_-@_-"
    ElseIf pickedCurrency.Value = "SEK" Then
      i.NumberFormat = "_-* #,##0.00 [$SEK]_-;-* #,##0.00 [$SEK]_-;_-* ""-""?? [$SEK]_-;_-@_-"
    Else
      MsgBox "Currency not set in file. Please contact file administrator."
      Exit Sub
    End If

  Next i
  
  Set i = Worksheets(3).Range("L" & rowCounter + 1)
  
  i.Formula = "=SUM($L$10:L" & rowCounter & ")*(1+$S$5)"
  
  If pickedCurrency.Value = "EUR" Then
    i.NumberFormat = "_-[$€-x-euro2] * #,##0.00_-;-[$€-x-euro2] * #,##0.00_-;_-[$€-x-euro2] * ""-""??_-;_-@_-"
  ElseIf pickedCurrency.Value = "GBP" Then
    i.NumberFormat = "_-* #,##0.00 [$GBP]_-;-* #,##0.00 [$GBP]_-;_-* ""-""?? [$GBP]_-;_-@_-"
  ElseIf pickedCurrency.Value = "CHF" Then
    i.NumberFormat = "_-* #,##0.00 [$CHF]_-;-* #,##0.00 [$CHF]_-;_-* ""-""?? [$CHF]_-;_-@_-"
  ElseIf pickedCurrency.Value = "USD" Then
    i.NumberFormat = "_-* #,##0.00 [$USD]_-;-* #,##0.00 [$USD]_-;_-* ""-""?? [$USD]_-;_-@_-"
  ElseIf pickedCurrency.Value = "PLN" Then
    i.NumberFormat = "_-* #,##0.00 [$PLN]_-;-* #,##0.00 [$PLN]_-;_-* ""-""?? [$PLN]_-;_-@_-"
  ElseIf pickedCurrency.Value = "CNY" Then
    i.NumberFormat = "_-* #,##0.00 [$CNY]_-;-* #,##0.00 [$CNY]_-;_-* ""-""?? [$CNY]_-;_-@_-"
  ElseIf pickedCurrency.Value = "SEK" Then
    i.NumberFormat = "_-* #,##0.00 [$SEK]_-;-* #,##0.00 [$SEK]_-;_-* ""-""?? [$SEK]_-;_-@_-"
  Else
    MsgBox "Currency not set in file. Please contact file administrator."
  Exit Sub
  End If
  
' ==================================================================================================== '

    bom_length = 8
    
    For Each i In Worksheets("Customer BOM").Range("A9:A10000")
        If i.Value <> "" Then
            bom_length = bom_length + 1
        End If
        If i.Value = "" Then
            Exit For
        End If
    Next i
    
    Dim bom_input_line As Long
    bom_input_line = bom_length + 1
    
    ' ==================================================================================================== '
    
    If Worksheets("Customer BOM").Range("E6").Value <> "" Then
        Dim prev_bom_length As Long
        prev_bom_length = 0
    
        For Each i In Worksheets("BOM_PREV_REV_COMPARE").Range("L1:L10000")
            If i.Text <> "" Then
                prev_bom_length = prev_bom_length + 1
            End If
            If i.Text = "" Then
                Exit For
            End If
    
            If i.Text = "#N/D!" Then
                Worksheets("Customer BOM").Range(bom_input_line & ":" & bom_input_line).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
                Worksheets("Customer BOM").Range("A" & bom_input_line).Value = Worksheets("BOM_PREV_REV_COMPARE").Range("A" & prev_bom_length)
                Worksheets("Customer BOM").Range("B" & bom_input_line).Value = Worksheets("BOM_PREV_REV_COMPARE").Range("B" & prev_bom_length)
                Worksheets("Customer BOM").Range("C" & bom_input_line).Value = Worksheets("BOM_PREV_REV_COMPARE").Range("C" & prev_bom_length)
                Worksheets("Customer BOM").Range("D" & bom_input_line).Value = Worksheets("BOM_PREV_REV_COMPARE").Range("D" & prev_bom_length)
                Worksheets("Customer BOM").Range("E" & bom_input_line).Value = Worksheets("BOM_PREV_REV_COMPARE").Range("E" & prev_bom_length)
                Worksheets("Customer BOM").Range("F" & bom_input_line).Value = Worksheets("BOM_PREV_REV_COMPARE").Range("F" & prev_bom_length)
                Worksheets("Customer BOM").Range("G" & bom_input_line).Value = Worksheets("BOM_PREV_REV_COMPARE").Range("G" & prev_bom_length)
                Worksheets("Customer BOM").Range("H" & bom_input_line).Value = Worksheets("BOM_PREV_REV_COMPARE").Range("H" & prev_bom_length)
                Worksheets("Customer BOM").Range("I" & bom_input_line).Value = Worksheets("BOM_PREV_REV_COMPARE").Range("I" & prev_bom_length)
                Worksheets("Customer BOM").Range("K" & bom_input_line).Value = Worksheets("BOM_PREV_REV_COMPARE").Range("K" & prev_bom_length)
                Worksheets("Customer BOM").Range("L" & bom_input_line).Value = "DELETED"
    
            End If
        Next i
    End If
    
    MsgBox "BOM calculated."

End Sub

Sub save_quote_excel_on_server()
On Error GoTo meh
    Dim clientName As String
    clientName = CStr(Worksheets(2).Range("A9").Value)
    Dim quotatnionNo As String
    quotatnionNo = CStr(Worksheets(2).Range("I2").Value)
    Dim offerDate As String
    offerDate = CStr(Worksheets(2).Range("I4").Value)
    Dim clientReference As String
    clientReference = CStr(Worksheets(2).Range("I7").Value)
    Dim projectReference As String
    projectReference = CStr(Worksheets(2).Range("C18").Value)
    Dim fileName As String
    fileName = clientName + " - " + clientReference + " - " + offerDate + " - " + quotatnionNo + " - " + projectReference
    Dim serverPathAndName As String
    serverPathAndName = "C:\Users\XxX\Google Drive\_GENERATOR\XxX\Sales offers\" & fileName
    ' Figure out how many pages to print
    Dim endOfOffer As String
    Dim pageCount As Integer
    pageCount = Worksheets(2).Range("M15").Value
    Dim QuoteWB As Workbook
    Set QuoteWB = ActiveWorkbook
    QuoteWB.Worksheets(2).Copy
    Set NewBook = ActiveWorkbook
    If pageCount = 1 Then
        endOfOffer = "J90"
        NewBook.Worksheets(1).Range("91:720").Delete Shift:=xlUp
    ElseIf pageCount = 2 Then
        endOfOffer = "J180"
        NewBook.Worksheets(1).Range("181:720").Delete Shift:=xlUp
        NewBook.Worksheets(1).Range("72:116").Delete Shift:=xlUp '1
    ElseIf pageCount = 3 Then
        endOfOffer = "J270"
        NewBook.Worksheets(1).Range("271:720").Delete Shift:=xlUp
        NewBook.Worksheets(1).Range("162:206").Delete Shift:=xlUp '2
        NewBook.Worksheets(1).Range("72:116").Delete Shift:=xlUp '1
    ElseIf pageCount = 4 Then
        endOfOffer = "J360"
        NewBook.Worksheets(1).Range("361:720").Delete Shift:=xlUp
        NewBook.Worksheets(1).Range("252:296").Delete Shift:=xlUp '3
        NewBook.Worksheets(1).Range("162:206").Delete Shift:=xlUp '2
        NewBook.Worksheets(1).Range("72:116").Delete Shift:=xlUp '1
    ElseIf pageCount = 5 Then
        endOfOffer = "J450"
        NewBook.Worksheets(1).Range("451:720").Delete Shift:=xlUp
        NewBook.Worksheets(1).Range("342:386").Delete Shift:=xlUp '4
        NewBook.Worksheets(1).Range("252:296").Delete Shift:=xlUp '3
        NewBook.Worksheets(1).Range("162:206").Delete Shift:=xlUp '2
        NewBook.Worksheets(1).Range("72:116").Delete Shift:=xlUp '1
    ElseIf pageCount = 6 Then
        endOfOffer = "J540"
        NewBook.Worksheets(1).Range("541:720").Delete Shift:=xlUp
        NewBook.Worksheets(1).Range("432:476").Delete Shift:=xlUp '5
        NewBook.Worksheets(1).Range("342:386").Delete Shift:=xlUp '4
        NewBook.Worksheets(1).Range("252:296").Delete Shift:=xlUp '3
        NewBook.Worksheets(1).Range("162:206").Delete Shift:=xlUp '2
        NewBook.Worksheets(1).Range("72:116").Delete Shift:=xlUp '1
    ElseIf pageCount = 7 Then
        endOfOffer = "J630"
        NewBook.Worksheets(1).Range("631:720").Delete Shift:=xlUp
        NewBook.Worksheets(1).Range("522:566").Delete Shift:=xlUp '6
        NewBook.Worksheets(1).Range("432:476").Delete Shift:=xlUp '5
        NewBook.Worksheets(1).Range("342:386").Delete Shift:=xlUp '4
        NewBook.Worksheets(1).Range("252:296").Delete Shift:=xlUp '3
        NewBook.Worksheets(1).Range("162:206").Delete Shift:=xlUp '2
        NewBook.Worksheets(1).Range("72:116").Delete Shift:=xlUp '1
    ElseIf pageCount = 8 Then
        endOfOffer = "J720"
        NewBook.Worksheets(1).Range("612:656").Delete Shift:=xlUp '7
        NewBook.Worksheets(1).Range("522:566").Delete Shift:=xlUp '6
        NewBook.Worksheets(1).Range("432:476").Delete Shift:=xlUp '5
        NewBook.Worksheets(1).Range("342:386").Delete Shift:=xlUp '4
        NewBook.Worksheets(1).Range("252:296").Delete Shift:=xlUp '3
        NewBook.Worksheets(1).Range("162:206").Delete Shift:=xlUp '2
        NewBook.Worksheets(1).Range("72:116").Delete Shift:=xlUp '1
    Else
        MsgBox "Incorrect page count."
        Exit Sub
    End If

    Dim offerRange As Integer
    offerRange = 26
    Dim validityPeriod As Integer
    validityPeriod = 26
    

    For Each i In NewBook.Worksheets(1).Range("A27:A10000")
      If i <> "" Then
        offerRange = offerRange + 1
      Else
        offerRange = offerRange + 1
        Exit For
      End If
    Next i
    For Each i In NewBook.Worksheets(1).Range("A27:A10000")
      validityPeriod = validityPeriod + 1
      If i = "Period of validity" Then
        validityPeriod = validityPeriod + 1
        Exit For
      End If
    Next i
    NewBook.Worksheets(1).Range(offerRange & ":" & validityPeriod - 3).Delete Shift:=xlUp '1
    ' Start saving
    
    Dim shape As Excel.shape
    Dim shapeCount As Integer
    shapeCount = 1
    For Each shape In ActiveSheet.Shapes
        If shapeCount > 1 Then
        shape.Delete
        End If
        shapeCount = shapeCount + 1
    Next
    
    ' CHANGE LAYOUT START '
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$A$1:$" + CStr(endOfOffer)
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    ' CHANGE LAYOUT FINISH '
    NewBook.Worksheets(1).PageSetup.PrintArea = "$A$1:$" + CStr(endOfOffer)
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    ' Save to google drive
    NewBook.SaveAs fileName:=serverPathAndName & ".xlsx", FileFormat:=51
    NewBook.Close
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
    MsgBox "Saved."
    Exit Sub
meh:
    MsgBox "Cannot save, file opened by another program."
    NewBook.Close
End Sub












