' CODE USED TO GENERTE TOTAL LOSS PER MONTH ON OVERPAID PURCHASE OREDRS
' IT IS ALSO SHOWING POSSIBLE LOSS ON OPEN PURCHASE ORDERS

Sub CalculatePurchaseAtSubSupplier(control As IRibbonControl)
    Worksheets("Analysis").ListObjects("Tabela_Analiza_drozszych_zakuow___2").Range. _
        AutoFilter Field:=11

    Worksheets("Analysis").ListObjects("Tabela_Analiza_drozszych_zakuow___2").Range. _
        AutoFilter Field:=5
    
    Worksheets("ModLog").Columns("H:H").Delete Shift:=xlToLeft
    
    Dim modLogEnd As Integer
    modLogEnd = 1

    For Each calculateLastModRow In Worksheets("ModLog").Range("B2", "B100000")

        If calculateLastModRow.Value <> "" Then
            modLogEnd = modLogEnd + 1
        End If

        If calculateLastModRow = "" Then
            Exit For
        End If

    Next calculateLastModRow

    Dim modLogEndCell As String
    modLogEndCell = "B" + CStr(modLogEnd)

    ' copy changed supplier to next column
    For Each rowToCopy In Worksheets("ModLog").Range("B2", modLogEndCell)

        If rowToCopy.Value <> "" Then
            For Each rQ In Worksheets("ModLog").Range("H1", "H100000")
                If rQ.Value = "" Then
                    rQ.Value = rowToCopy.Value
                    Exit For
                End If
            Next rQ
        End If

        If rowToCopy.Value = "" Then
            Exit For
        End If
    Next rowToCopy
'========================================================================'

    Worksheets("Analysis").Columns("J:M").Delete Shift:=xlToLeft
    
    Dim analysisEnd As Integer
    analysisEnd = 1

    For Each calculateLastAnalysisRow In Worksheets("Analysis").Range("A2", "A1000000")

        If calculateLastAnalysisRow.Value <> "" Then
            analysisEnd = analysisEnd + 1
        End If

        If calculateLastAnalysisRow = "" Then
            Exit For
        End If

    Next calculateLastAnalysisRow

    Dim dateSearchFormEnd As String
    dateSearchFormEnd = "J" + CStr(analysisEnd)

    Dim timeCompareFormEnd As String
    timeCompareFormEnd = "K" + CStr(analysisEnd)

    Dim overpayPerOrderEnd As String
    overpayPerOrderEnd = "L" + CStr(analysisEnd)

    Dim overpayTimesOrerQty As String
    overpayTimesOrerQty = "M" + CStr(analysisEnd)

    Dim totalOverpayLocation As String
    totalOverpayLocation = "M" + CStr(analysisEnd + 1)

    Worksheets("Analysis").Range("J2", dateSearchFormEnd).NumberFormat = "m/d/yyyy"

    Worksheets("Analysis").Range("J2", dateSearchFormEnd).FormulaR1C1 = "=VLOOKUP(C[-9],ModLog!C[-8]:C[-3],6,0)"

    Worksheets("Analysis").Range("K2", timeCompareFormEnd).FormulaR1C1 = "=IF(C[-1]<C[-6],0,1)"

    Worksheets("Analysis").Range("L2", overpayPerOrderEnd).FormulaR1C1 = "=C[-8]-C[-9]"

    Worksheets("Analysis").Range("M2", overpayTimesOrerQty).FormulaR1C1 = "=C[-1]*C[-7]"

    Worksheets("Analysis").ListObjects("Tabela_Analiza_drozszych_zakuow___2").Resize Range( _
        "$A$1", overpayTimesOrerQty)

    Worksheets("Analysis").ListObjects("Tabela_Analiza_drozszych_zakuow___2").Range. _
        AutoFilter Field:=11, Criteria1:="=0", Operator:=xlOr, Criteria2:= _
        "=#N/D!"


    Dim dateInput As String
    dateInput = ""
    dateInput = Worksheets("Data").Range("C1:C1").Value

    Worksheets("Analysis").ListObjects("Tabela_Analiza_drozszych_zakuow___2").Range. _
        AutoFilter Field:=5, Operator:=xlFilterValues, Criteria2:=Array(1, _
        dateInput)


    Worksheets("Analysis").Range(totalOverpayLocation, totalOverpayLocation).Formula = "=SUBTOTAL(109,[Kolumna4])"

    'stare "=SUM($M$2:" & overpayTimesOrerQty & ")"'

    '=SUM(R[-463]C:R[-7]C)

    


    Worksheets("Data").Range("C3:C3").Value = Worksheets("Analysis").Range(totalOverpayLocation, totalOverpayLocation).Value

End Sub






