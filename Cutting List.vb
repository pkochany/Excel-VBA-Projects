' PROGRAM DESIGNED TO DISTINGUISH CUTTING LIST OF WIRES DESIGNED FOR PRODUCTION
' TO DATA CONTAINED IN ERP DATABASE
' THIS IS CLEARING HUMAN MISTAKES AFTER THE LIST IS CREATED
' SO THE ERP SYSTEM WILL HAVE THE SAME QUANTITIES AS PHYSICALLY CUT WIRE

Sub f1(control As IRibbonControl)
'
    ' declare counter for table of values
    Dim rowCount As Integer
    rowCount = 5

    ' figure out list length by number of wires (all cells in column D have to be filled)
    For Each itemNumber In Worksheets("Wires").Range("D6:D10000")
        If itemNumber.Value <> "" Then
            rowCount = rowCount + 1
        End If
        If itemNumber.Value = "" Then
            Exit For
        End If
    Next itemNumber
    
    ' figure out bill of material length
    Dim rowCountBiOfMa As Integer
    rowCountBiOfMa = 0
    For Each formLicz In Worksheets("bill").Range("F1:F5000")
        If formLicz.Value <> "" Then
            rowCountBiOfMa = rowCountBiOfMa + 1
        End If
        If formLicz.Value = "" Then
            Exit For
        End If
    Next formLicz

    ' check if bill of material is not empty, proceed if not empty '
    If rowCountBiOfMa > 1 And rowCount > 5 Then

        ' variable declaration '
        Dim cable ItemColumnEnd As String
        cable ItemColumnEnd = "D" + CStr(rowCount)

        Dim seal1End As String
        seal1End = "H" + CStr(rowCount)

        Dim contact1End As String
        contact1End = "I" + CStr(rowCount)

        Dim seal2End As String
        seal2End = "N" + CStr(rowCount)

        Dim contact2End As String
        contact2End = "O" + CStr(rowCount)

        Dim checkList As String
        checkList = "O" + CStr(rowCount)

'====================================================================================================================================================================================
'====================================================================================================================================================================================
        ' FUNCTION - ADD POSITIONS INTO bill of material SHEET WHICH ARE ON THE LIST BUT NOT IN PREPARATION
        
        ' clear data in DISTINGUISH sheet
        Sheets("DISTINGUISH").Columns("A:D").Delete Shift:=xlToLeft
        
        ' clear filters at cutting list'
        Sheets("Wires").Range("G5", checkList).AutoFilter

        ' copy 4 columns in cutting list into one at COMPARE, one onder the other '
        Dim compareEndCount1 As Integer
        compareEndCount1 = rowCount + 1

        Dim compareEnd1 As String
        compareEnd1 = "A" + CStr(compareEndCount1)
        ' === '
        Dim compareEndCount2 As Integer
        compareEndCount2 = rowCount + rowCount + 1

        Dim compareEnd2 As String
        compareEnd2 = "A" + CStr(compareEndCount2)
        ' === '
        Dim compareEndCount3 As Integer
        compareEndCount3 = rowCount + rowCount + rowCount + 1

        Dim compareEnd3 As String
        compareEnd3 = "A" + CStr(compareEndCount3)
        ' === '
        Dim compareEndCount4 As Integer
        compareEndCount4 = rowCount + rowCount + rowCount + rowCount + 1

        Dim compareEnd4 As String
        compareEnd4 = "A" + CStr(compareEndCount4)
        ' === '
        Dim compareEndAfterPasteCount As Integer
        compareEndAfterPasteCount = rowCount + rowCount + rowCount + rowCount + rowCount + 1

        Dim compareEndColumnA As String
        compareEndColumnA = "A" + CStr(compareEndAfterPasteCount)

        Dim compareEndColumnB As String
        compareEndColumnB = "B" + CStr(compareEndAfterPasteCount)

        Dim compareEndColumnC As String
        compareEndColumnC = "C" + CStr(compareEndAfterPasteCount)

        Sheets("Wires").Range("H6", seal1End).Copy
        Sheets("DISTINGUISH").Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

        Sheets("Wires").Range("I6", contact1End).Copy
        Sheets("DISTINGUISH").Range(compareEnd1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

        Sheets("Wires").Range("N6", seal2End).Copy
        Sheets("DISTINGUISH").Range(compareEnd2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

        Sheets("Wires").Range("O6", contact2End).Copy
        Sheets("DISTINGUISH").Range(compareEnd3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

        Sheets("Wires").Range("D6", cable ItemColumnEnd).Copy
        Sheets("DISTINGUISH").Range(compareEnd4).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

        ' count occurrences '
        Sheets("DISTINGUISH").Range("B2", compareEndColumnB).FormulaR1C1 = "=COUNTIF(C[-1],RC[-1])"

        ' convert occurrence to string
        Sheets("DISTINGUISH").Columns("B:B").Copy
        Sheets("DISTINGUISH").Columns("B:B").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        ' clear duplicates
        Sheets("DISTINGUISH").Range("A1", compareEndColumnB).RemoveDuplicates Columns:=1, Header:=xlNo

        ' figure out end of list
        Dim endCountNumber As Integer
        endCountNumber = 1

        For Each endAfterClearDuplicates In Worksheets("DISTINGUISH").Range("A2:A10000")
            If endAfterClearDuplicates.Value <> "" Then
                endCountNumber = endCountNumber + 1
            End If
            If endAfterClearDuplicates.Value = "" Then
                Exit For
            End If
        Next endAfterClearDuplicates

        Dim endCountColumnC As String
        endCountColumnC = "C" + CStr(endCountNumber)

        ' check if exist in bill of material '
        Sheets("DISTINGUISH").Range("C2", endCountColumnC).FormulaR1C1 = _
            "=VLOOKUP(RC[-2],Tabela_Kwerenda_z_001_xxx[[art_artnr]:[Kolumna1]],3,0)"

        ' convert data to string '
        Sheets("DISTINGUISH").Columns("C:C").Copy
        Sheets("DISTINGUISH").Columns("C:C").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        ' check if error = not in bill of material and need to be added '
        Dim artEndBiOfMa As String
        artEndBiOfMa = "F" + CStr(rowCountBiOfMa + 1)

        Dim kvantEndBiOfMa As String
        kvantEndBiOfMa = "G" + CStr(rowCountBiOfMa + 1)

        Dim counterToCopy As Integer
        counterToCopy = 2

        Dim itemToCopy As String
        itemToCopy = "A" + CStr(counterToCopy)

        Dim qtyToCopy As String
        qtyToCopy = "B" + CStr(counterToCopy)

        Dim checkVariable As String

        For Each addToBiOfMa In Worksheets("DISTINGUISH").Range("C2", endCountColumnC)

            checkVariable = CStr(addToBiOfMa.Value)

            If CStr(addToBiOfMa.Value) = "Error 2042" Then
            
                itemToCopy = "A" + CStr(counterToCopy)
                qtyToCopy = "B" + CStr(counterToCopy)

                Sheets("bill").Range(artEndBiOfMa).Value = Worksheets("DISTINGUISH").Range(itemToCopy).Value
                Sheets("bill").Range(kvantEndBiOfMa).Value = "0"
                rowCountBiOfMa = rowCountBiOfMa + 1
                
                artEndBiOfMa = "F" + CStr(rowCountBiOfMa + 1)
                kvantEndBiOfMa = "G" + CStr(rowCountBiOfMa + 1)

            End If

            If CStr(addToBiOfMa.Value) <> "" Then

                counterToCopy = counterToCopy + 1

            End If

        Next addToBiOfMa
        
'====================================================================================================================================================================================
'====================================================================================================================================================================================
        
        ' Variable declaration '
        Dim ifFormulaEnd As String
        ifFormulaEnd = "H" + CStr(rowCountBiOfMa)
        
        Dim sumFormulaEnd As String
        sumFormulaEnd = "I" + CStr(rowCountBiOfMa)
        
        Dim articeFormulaEnd As String
        articeFormulaEnd = "F" + CStr(rowCountBiOfMa)
        
        Dim jFormEnd As String
        jFormEnd = "J" + CStr(rowCountBiOfMa)
        
        Dim kFormEnd As String
        kFormEnd = "K" + CStr(rowCountBiOfMa)

        Dim iFormEnd As String
        iFormEnd = "I" + CStr(rowCountBiOfMa)

        ' Clear duplicates from monitor BiOfMa'
        Dim lFormEnd As String
        lFormEnd = "L" + CStr(rowCountBiOfMa)

        Dim gFormEnd As String
        gFormEnd = "G" + CStr(rowCountBiOfMa)

        For Each formFillIn In Worksheets("bill").Range("L2", lFormEnd)
            formFillIn.Value = "=SUMIF([art_artnr],[@[art_artnr]],[st_kvant])"
        Next formFillIn

        Worksheets("bill").Range("L2", lFormEnd).Copy
        Worksheets("bill").Range("G2", gFormEnd).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        Worksheets("bill").Range("xxx").RemoveDuplicates Columns:=4, Header:=xlYes

        
'====================================================================================================================================================================================
        ' MESSAGE BOX - CHECK IF LIST IS NOT EMPTY
    
        Dim msgPromptCuntFillIn As Integer
        msgPromptCuntFillIn = 0

        For Each msgPrompt In Worksheets("Wires").Range("C6", checkList)
            If msgPrompt.Value <> "" Then
                msgPromptCuntFillIn = msgPromptCuntFillIn + 1
            End If
        Next msgPrompt

        If msgPromptCuntFillIn = 0 Then
            MsgBox "message box" ' & vbNewLine & "Line 2"
        End If

'====================================================================================================================================================================================

        ' CHECK CABLE LENGTHS
        Sheets("bill").Range("J2", jFormEnd).FormulaR1C1 = "=SUMIF('Cutting list'!R6C4:R" & rowCount & "C4,[@[art_artnr]],'Cutting list'!R6C3:R" & rowCount & "C3)"
        
        Sheets("bill").Range("K2", kFormEnd).FormulaR1C1 = "=IF(RC[-1]=0,0,IF(COUNTIF(R1C17:R5000C17,RC[-5])>0,0,-(RC[-4]-(RC[-1]/1000))))"

        ' CHECK CONTACTS, SEALS AND SHRINKING TUBES
        Sheets("bill").Range("H2", ifFormulaEnd).FormulaR1C1 = "=COUNTIF('Cutting list'!R6C8:R" & rowCount & "C" & 15 & ",BiOfMa!RC[-2])"
        
        Sheets("bill").Range("I2", sumFormulaEnd).FormulaR1C1 = "=IF(RC[-1]=0,0,IF(COUNTIF(R1C17:R500C17,RC[-3])>0,0,-(RC[-2]-RC[-1])))"
        
        ' prepare space for copy
        Sheets("Wires").Range("W6", "Z1000").ClearContents

        ' COPY  Item NUMBERS
        Dim rowCountCopy As Integer
        rowCountCopy = 2
        Dim copy ItemF As String
        copy ItemF = "F2"
        Dim copyQtyK As String
        copyQtyK = "K2"

        Dim rowCountPasteAdd As Integer
        rowCountPasteAdd = 6
        Dim rowCountPasteSubstract As Integer
        rowCountPasteSubstract = 6
        Dim pasteQtyW As String
        pasteQtyW = "W6"
        Dim paste ItemX As String
        paste ItemX = "X6"
        Dim pasteQtyY As String
        pasteQtyY = "Y6"
        Dim paste ItemZ As String
        paste ItemZ = "Z6"
        
        For Each estimatedValue In Worksheets("bill").Range("K2", kFormEnd)
            ' ADD CABLES TO TABLE - ADD TO SYSTEM '
            If estimatedValue.Value <> "" Then
                If estimatedValue.Value > 0 Then

                    Sheets("Wires").Range(pasteQtyW).Value = Sheets("bill").Range(copyQtyK).Value
                    Sheets("Wires").Range(paste ItemX).Value = Sheets("bill").Range(copy ItemF).Value

                    rowCountPasteAdd = rowCountPasteAdd + 1
                    pasteQtyW = "W" + CStr(rowCountPasteAdd)
                    paste ItemX = "X" + CStr(rowCountPasteAdd)

                End If
            End If
            ' ADD CABLES TO TABLE - SUBSTRACT FROM SYSTEM '
            If estimatedValue.Value <> "" Then
                If estimatedValue.Value < 0 Then

                    Sheets("Wires").Range(pasteQtyY).Value = Sheets("bill").Range(copyQtyK).Value
                    Sheets("Wires").Range(paste ItemZ).Value = Sheets("bill").Range(copy ItemF).Value

                    rowCountPasteSubstract = rowCountPasteSubstract + 1
                    pasteQtyY = "Y" + CStr(rowCountPasteSubstract)
                    paste ItemZ = "Z" + CStr(rowCountPasteSubstract)

                End If
            End If

            If estimatedValue.Value = "" Then
                Exit For
            End If

            rowCountCopy = rowCountCopy + 1
            copyQtyK = "K" + CStr(rowCountCopy)
            copy ItemF = "F" + CStr(rowCountCopy)

        Next estimatedValue
        
        ' prepare space for copy
        Sheets("Wires").Range("S6", "V1000").ClearContents

        ' COPY  Item NUMBERS
        rowCountCopy = 2
        copy ItemF = "F2"
        Dim copyQtyI As String
        copyQtyI = "I2"

        rowCountPasteAdd = 6
        rowCountPasteSubstract = 6
        Dim pasteQtyS As String
        pasteQtyS = "S6"
        Dim paste ItemT As String
        paste ItemT = "T6"
        Dim pasteQtyU As String
        pasteQtyU = "U6"
        Dim paste ItemV As String
        paste ItemV = "V6"
        
        For Each estimatedValue In Worksheets("bill").Range("I2", iFormEnd)
            ' ADD CONTACTS TO TABLE - ADD TO SYSTEM '
            If estimatedValue.Value <> "" Then
                If estimatedValue.Value > 0 Then
                    Sheets("Wires").Range(pasteQtyS).Value = Sheets("bill").Range(copyQtyI).Value
                    Sheets("Wires").Range(paste ItemT).Value = Sheets("bill").Range(copy ItemF).Value

                    rowCountPasteAdd = rowCountPasteAdd + 1
                    pasteQtyS = "S" + CStr(rowCountPasteAdd)
                    paste ItemT = "T" + CStr(rowCountPasteAdd)
                End If
            End If

            ' ADD CONTACTS TO TABLE - SUBSTRACT FROM SYSTEM '
            If estimatedValue.Value <> "" Then
                If estimatedValue.Value < 0 Then
                    Sheets("Wires").Range(pasteQtyU).Value = Sheets("bill").Range(copyQtyI).Value
                    Sheets("Wires").Range(paste ItemV).Value = Sheets("bill").Range(copy ItemF).Value

                    rowCountPasteSubstract = rowCountPasteSubstract + 1
                    pasteQtyU = "U" + CStr(rowCountPasteSubstract)
                    paste ItemV = "V" + CStr(rowCountPasteSubstract)
                End If
            End If


            If estimatedValue.Value = "" Then
                Exit For
            End If

            rowCountCopy = rowCountCopy + 1
            copyQtyI = "I" + CStr(rowCountCopy)
            copy ItemF = "F" + CStr(rowCountCopy)

        Next estimatedValue

        
'====================================================================================================================================================================================
        
        ' FINISH MESSAGE BOX
        Dim msgPromptCunt As Integer
        msgPromptCunt = 0
        
        For Each msgPrompt In Worksheets("Wires").Range("S6", "Z6")
            If msgPrompt.Value <> "" Then
                msgPromptCunt = msgPromptCunt + 1
            End If
            If msgPromptCuntFillIn = 0 Then
                msgPromptCunt = msgPromptCunt + 1
            End If
        Next msgPrompt
        
        If msgPromptCunt = 0 Then
            MsgBox "Congrats!" ' & vbNewLine & "Line 2"
        End If
    End If

    If rowCountBiOfMa < 2 Then
        MsgBox "Some error"
    End If

    If rowCount = 5 Then
        MsgBox "Sime other error"
    End If
    
    
End Sub







