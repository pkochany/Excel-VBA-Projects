' PROGRAM DESIGNED TO COMPARE CUTTING LIST OF WIRES DESIGNED FOR PRODUCTION
' TO DATA CONTAINED IN ERP DATABASE
' THIS IS CLEARING HUMAN MISTAKES AFTER THE LIST IS CREATED
' SO THE ERP SYSTEM WILL HAVE THE SAME QUANTITIES AS PHYSICALLY CUT WIRE

Sub EvaluateCuttingList(control As IRibbonControl)
'
    ' CUTTING LIST DATA EVALUATION MACRO BY PIOTR KOCHANY - RIMASTER COMPANY - 27-02-2017 '
    ' LAST MODIF 10-04-2017 '
    
    ' declare counter for table of values
    Dim rowCount As Integer
    rowCount = 5

    ' figure out list length by number of wires (all cells in column D have to be filled)
    For Each rimNumber In Worksheets("Lista cięć").Range("D6:D10000")
        If rimNumber.Value <> "" Then
            rowCount = rowCount + 1
        End If
        If rimNumber.Value = "" Then
            Exit For
        End If
    Next rimNumber
    
    ' figure out BOM length
    Dim rowCountBOM As Integer
    rowCountBOM = 0
    For Each formLicz In Worksheets("BOM").Range("F1:F5000")
        If formLicz.Value <> "" Then
            rowCountBOM = rowCountBOM + 1
        End If
        If formLicz.Value = "" Then
            Exit For
        End If
    Next formLicz

    ' check if BOM is not empty, proceed if not empty '
    If rowCountBOM > 1 And rowCount > 5 Then
        
        ' variable declaration '
        Dim cableRIMColumnEnd As String
        cableRIMColumnEnd = "D" + CStr(rowCount)

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
        ' FUNCTION - ADD POSITIONS INTO BOM SHEET WHICH ARE ON THE LIST BUT NOT IN PREPARATION
        
        ' clear data in COMPARE sheet
        Sheets("COMPARE").Columns("A:D").Delete Shift:=xlToLeft
        
        ' clear filters at cutting list'
        Sheets("Lista cięć").Range("G5", checkList).AutoFilter

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

        Sheets("Lista cięć").Range("H6", seal1End).Copy
        Sheets("COMPARE").Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

        Sheets("Lista cięć").Range("I6", contact1End).Copy
        Sheets("COMPARE").Range(compareEnd1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

        Sheets("Lista cięć").Range("N6", seal2End).Copy
        Sheets("COMPARE").Range(compareEnd2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

        Sheets("Lista cięć").Range("O6", contact2End).Copy
        Sheets("COMPARE").Range(compareEnd3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

        Sheets("Lista cięć").Range("D6", cableRIMColumnEnd).Copy
        Sheets("COMPARE").Range(compareEnd4).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

        ' count occurrences '
        Sheets("COMPARE").Range("B2", compareEndColumnB).FormulaR1C1 = "=COUNTIF(C[-1],RC[-1])"

        ' convert occurrence to string
        Sheets("COMPARE").Columns("B:B").Copy
        Sheets("COMPARE").Columns("B:B").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        ' clear duplicates
        Sheets("COMPARE").Range("A1", compareEndColumnB).RemoveDuplicates Columns:=1, Header:=xlNo

        ' figure out end of list
        Dim endCountNumber As Integer
        endCountNumber = 1

        For Each endAfterClearDuplicates In Worksheets("COMPARE").Range("A2:A10000")
            If endAfterClearDuplicates.Value <> "" Then
                endCountNumber = endCountNumber + 1
            End If
            If endAfterClearDuplicates.Value = "" Then
                Exit For
            End If
        Next endAfterClearDuplicates

        Dim endCountColumnC As String
        endCountColumnC = "C" + CStr(endCountNumber)

        ' check if exist in BOM '
        Sheets("COMPARE").Range("C2", endCountColumnC).FormulaR1C1 = _
            "=VLOOKUP(RC[-2],Tabela_Kwerenda_z_001_Rimaster_Poland_Sp._z_o.o[[art_artnr]:[Kolumna1]],3,0)"

        ' convert data to string '
        Sheets("COMPARE").Columns("C:C").Copy
        Sheets("COMPARE").Columns("C:C").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        ' check if error = not in BOM and need to be added '
        Dim artEndBOM As String
        artEndBOM = "F" + CStr(rowCountBOM + 1)

        Dim kvantEndBOM As String
        kvantEndBOM = "G" + CStr(rowCountBOM + 1)

        Dim counterToCopy As Integer
        counterToCopy = 2

        Dim rimToCopy As String
        rimToCopy = "A" + CStr(counterToCopy)

        Dim qtyToCopy As String
        qtyToCopy = "B" + CStr(counterToCopy)

        Dim checkVariable As String

        For Each addToBOM In Worksheets("COMPARE").Range("C2", endCountColumnC)

            checkVariable = CStr(addToBOM.Value)

            If CStr(addToBOM.Value) = "Error 2042" Then
            
                rimToCopy = "A" + CStr(counterToCopy)
                qtyToCopy = "B" + CStr(counterToCopy)

                Sheets("BOM").Range(artEndBOM).Value = Worksheets("COMPARE").Range(rimToCopy).Value
                Sheets("BOM").Range(kvantEndBOM).Value = "0"
                rowCountBOM = rowCountBOM + 1
                
                artEndBOM = "F" + CStr(rowCountBOM + 1)
                kvantEndBOM = "G" + CStr(rowCountBOM + 1)

            End If

            If CStr(addToBOM.Value) <> "" Then

                counterToCopy = counterToCopy + 1

            End If

        Next addToBOM
        
'====================================================================================================================================================================================
'====================================================================================================================================================================================
        
        ' Variable declaration '
        Dim ifFormulaEnd As String
        ifFormulaEnd = "H" + CStr(rowCountBOM)
        
        Dim sumFormulaEnd As String
        sumFormulaEnd = "I" + CStr(rowCountBOM)
        
        Dim articeFormulaEnd As String
        articeFormulaEnd = "F" + CStr(rowCountBOM)
        
        Dim jFormEnd As String
        jFormEnd = "J" + CStr(rowCountBOM)
        
        Dim kFormEnd As String
        kFormEnd = "K" + CStr(rowCountBOM)

        Dim iFormEnd As String
        iFormEnd = "I" + CStr(rowCountBOM)

        ' Clear duplicates from monitor BOM'
        Dim lFormEnd As String
        lFormEnd = "L" + CStr(rowCountBOM)

        Dim gFormEnd As String
        gFormEnd = "G" + CStr(rowCountBOM)

        For Each formFillIn In Worksheets("BOM").Range("L2", lFormEnd)
            formFillIn.Value = "=SUMIF([art_artnr],[@[art_artnr]],[st_kvant])"
        Next formFillIn

        Worksheets("BOM").Range("L2", lFormEnd).Copy
        Worksheets("BOM").Range("G2", gFormEnd).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        Worksheets("BOM").Range("Tabela_Kwerenda_z_001_Rimaster_Poland_Sp._z_o.o").RemoveDuplicates Columns:=4, Header:=xlYes

        
'====================================================================================================================================================================================
        ' MESSAGE BOX - CHECK IF LIST IS NOT EMPTY
    
        Dim msgPromptCuntFillIn As Integer
        msgPromptCuntFillIn = 0

        For Each msgPrompt In Worksheets("Lista cięć").Range("C6", checkList)
            If msgPrompt.Value <> "" Then
                msgPromptCuntFillIn = msgPromptCuntFillIn + 1
            End If
        Next msgPrompt

        If msgPromptCuntFillIn = 0 Then
            MsgBox "Uzupełnij listę połączeń" ' & vbNewLine & "Line 2"
        End If

'====================================================================================================================================================================================

        ' CHECK CABLE LENGTHS
        Sheets("BOM").Range("J2", jFormEnd).FormulaR1C1 = "=SUMIF('Lista cięć'!R6C4:R" & rowCount & "C4,[@[art_artnr]],'Lista cięć'!R6C3:R" & rowCount & "C3)"
        
        Sheets("BOM").Range("K2", kFormEnd).FormulaR1C1 = "=IF(RC[-1]=0,0,IF(COUNTIF(R1C17:R5000C17,RC[-5])>0,0,-(RC[-4]-(RC[-1]/1000))))"

        ' CHECK CONTACTS, SEALS AND SHRINKING TUBES
        Sheets("BOM").Range("H2", ifFormulaEnd).FormulaR1C1 = "=COUNTIF('Lista cięć'!R6C8:R" & rowCount & "C" & 15 & ",BOM!RC[-2])"
        
        Sheets("BOM").Range("I2", sumFormulaEnd).FormulaR1C1 = "=IF(RC[-1]=0,0,IF(COUNTIF(R1C17:R500C17,RC[-3])>0,0,-(RC[-2]-RC[-1])))"
        
        ' prepare space for copy
        Sheets("Lista cięć").Range("W6", "Z1000").ClearContents

        ' COPY RIM NUMBERS
        Dim rowCountCopy As Integer
        rowCountCopy = 2
        Dim copyRimF As String
        copyRimF = "F2"
        Dim copyQtyK As String
        copyQtyK = "K2"

        Dim rowCountPasteAdd As Integer
        rowCountPasteAdd = 6
        Dim rowCountPasteSubstract As Integer
        rowCountPasteSubstract = 6
        Dim pasteQtyW As String
        pasteQtyW = "W6"
        Dim pasteRimX As String
        pasteRimX = "X6"
        Dim pasteQtyY As String
        pasteQtyY = "Y6"
        Dim pasteRimZ As String
        pasteRimZ = "Z6"
        
        For Each estimatedValue In Worksheets("BOM").Range("K2", kFormEnd)
            ' ADD CABLES TO TABLE - ADD TO SYSTEM '
            If estimatedValue.Value <> "" Then
                If estimatedValue.Value > 0 Then

                    Sheets("Lista cięć").Range(pasteQtyW).Value = Sheets("BOM").Range(copyQtyK).Value
                    Sheets("Lista cięć").Range(pasteRimX).Value = Sheets("BOM").Range(copyRimF).Value

                    rowCountPasteAdd = rowCountPasteAdd + 1
                    pasteQtyW = "W" + CStr(rowCountPasteAdd)
                    pasteRimX = "X" + CStr(rowCountPasteAdd)

                End If
            End If
            ' ADD CABLES TO TABLE - SUBSTRACT FROM SYSTEM '
            If estimatedValue.Value <> "" Then
                If estimatedValue.Value < 0 Then

                    Sheets("Lista cięć").Range(pasteQtyY).Value = Sheets("BOM").Range(copyQtyK).Value
                    Sheets("Lista cięć").Range(pasteRimZ).Value = Sheets("BOM").Range(copyRimF).Value

                    rowCountPasteSubstract = rowCountPasteSubstract + 1
                    pasteQtyY = "Y" + CStr(rowCountPasteSubstract)
                    pasteRimZ = "Z" + CStr(rowCountPasteSubstract)

                End If
            End If

            If estimatedValue.Value = "" Then
                Exit For
            End If

            rowCountCopy = rowCountCopy + 1
            copyQtyK = "K" + CStr(rowCountCopy)
            copyRimF = "F" + CStr(rowCountCopy)

        Next estimatedValue
        
        ' prepare space for copy
        Sheets("Lista cięć").Range("S6", "V1000").ClearContents

        ' COPY RIM NUMBERS
        rowCountCopy = 2
        copyRimF = "F2"
        Dim copyQtyI As String
        copyQtyI = "I2"

        rowCountPasteAdd = 6
        rowCountPasteSubstract = 6
        Dim pasteQtyS As String
        pasteQtyS = "S6"
        Dim pasteRimT As String
        pasteRimT = "T6"
        Dim pasteQtyU As String
        pasteQtyU = "U6"
        Dim pasteRimV As String
        pasteRimV = "V6"
        
        For Each estimatedValue In Worksheets("BOM").Range("I2", iFormEnd)
            ' ADD CONTACTS TO TABLE - ADD TO SYSTEM '
            If estimatedValue.Value <> "" Then
                If estimatedValue.Value > 0 Then
                    Sheets("Lista cięć").Range(pasteQtyS).Value = Sheets("BOM").Range(copyQtyI).Value
                    Sheets("Lista cięć").Range(pasteRimT).Value = Sheets("BOM").Range(copyRimF).Value

                    rowCountPasteAdd = rowCountPasteAdd + 1
                    pasteQtyS = "S" + CStr(rowCountPasteAdd)
                    pasteRimT = "T" + CStr(rowCountPasteAdd)
                End If
            End If

            ' ADD CONTACTS TO TABLE - SUBSTRACT FROM SYSTEM '
            If estimatedValue.Value <> "" Then
                If estimatedValue.Value < 0 Then
                    Sheets("Lista cięć").Range(pasteQtyU).Value = Sheets("BOM").Range(copyQtyI).Value
                    Sheets("Lista cięć").Range(pasteRimV).Value = Sheets("BOM").Range(copyRimF).Value

                    rowCountPasteSubstract = rowCountPasteSubstract + 1
                    pasteQtyU = "U" + CStr(rowCountPasteSubstract)
                    pasteRimV = "V" + CStr(rowCountPasteSubstract)
                End If
            End If


            If estimatedValue.Value = "" Then
                Exit For
            End If

            rowCountCopy = rowCountCopy + 1
            copyQtyI = "I" + CStr(rowCountCopy)
            copyRimF = "F" + CStr(rowCountCopy)

        Next estimatedValue

        
'====================================================================================================================================================================================
        
        ' FINISH MESSAGE BOX
        Dim msgPromptCunt As Integer
        msgPromptCunt = 0
        
        For Each msgPrompt In Worksheets("Lista cięć").Range("S6", "Z6")
            If msgPrompt.Value <> "" Then
                msgPromptCunt = msgPromptCunt + 1
            End If
            If msgPromptCuntFillIn = 0 Then
                msgPromptCunt = msgPromptCunt + 1
            End If
        Next msgPrompt
        
        If msgPromptCunt = 0 Then
            MsgBox "Gratulacje, Lista Połączeń zgadza się z Preparation!" ' & vbNewLine & "Line 2"
        End If
    End If

    If rowCountBOM < 2 Then
        MsgBox "Brak materiału w liście BOM. Uzupełnij numer wiązki i kliknij Dane -> Odśwież wszystko."
    End If

    If rowCount = 5 Then
        MsgBox "Brak połączeń w kolumnie D. Uzupełnij listę połączeń."
    End If
    
    
End Sub







