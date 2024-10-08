Attribute VB_Name = "game"
Public totalRevenue, clickRevenue, profitTotals(14) As Double
Public generatorBaserate(12), generatorRate(12), generatorRevenue(12), generatorBasecost(12), generatorCost(12), generatorCC(12), generatorRC(12) As Double
Public itemRate(13), itemBasecost(13), itemCost(13), itemCC(13), itemRC(13) As Double
Public buildSpend, itemSpend As Double
Public generatorQty(12), itemQty(13), x, y As Integer

Sub MainLoop()
    With Workbooks("cotclicker.xlsm").Worksheets("Sheet1")
        totalRevenue = Cells(2, 11).Value
        clickRevenue = Cells(3, 11).Value
        itemSpend = Cells(2, 8).Value
        buildSpend = Cells(2, 3).Value
        
        
        'define base costs
        generatorBasecost(0) = 15
        generatorBasecost(1) = 100
        generatorBasecost(2) = 600
        generatorBasecost(3) = 3000
        generatorBasecost(4) = 10000
        generatorBasecost(5) = 25000
        generatorBasecost(6) = 150000
        generatorBasecost(7) = 3500000
        generatorBasecost(8) = 150000000
        generatorBasecost(9) = 5000000000#
        generatorBasecost(10) = 30000000000#
        generatorBasecost(11) = 900000000000#
        generatorBasecost(12) = 25000000000000#
        
        itemBasecost(0) = 150
        itemBasecost(1) = 1000
        itemBasecost(2) = 5000
        itemBasecost(3) = 5000
        itemBasecost(4) = 5000
        itemBasecost(5) = 20000
        itemBasecost(6) = 100000
        itemBasecost(7) = 500000
        itemBasecost(8) = 1500000
        itemBasecost(9) = 10000000
        itemBasecost(10) = 100000000
        itemBasecost(11) = 7500000000#
        itemBasecost(12) = 10000000000#
        itemBasecost(13) = 250000000000#
        
        'define base rate
        generatorBaserate(0) = 0.2
        generatorBaserate(1) = 1
        generatorBaserate(2) = 5
        generatorBaserate(3) = 20
        generatorBaserate(4) = 50
        generatorBaserate(5) = 100
        generatorBaserate(6) = 500
        generatorBaserate(7) = 10000
        generatorBaserate(8) = 350000
        generatorBaserate(9) = 10000000
        generatorBaserate(10) = 50000000
        generatorBaserate(11) = 1000000000#
        generatorBaserate(12) = 20000000000#
        
        'define cost coefficients > 0
        generatorCC(0) = 1.03
        generatorCC(1) = 1.03
        generatorCC(2) = 1.025
        generatorCC(3) = 1.02
        generatorCC(4) = 1.02
        generatorCC(5) = 1.02
        generatorCC(6) = 1.018
        generatorCC(7) = 1.017
        generatorCC(8) = 1.013
        generatorCC(9) = 1.009
        generatorCC(10) = 1.002
        generatorCC(11) = 1.0005
        generatorCC(12) = 1.00001
        
        itemCC(0) = 1.5
        itemCC(1) = 1.01
        itemCC(2) = 1.07
        itemCC(3) = 1.1
        itemCC(4) = 1.05
        itemCC(5) = 1.1
        itemCC(6) = 1.05
        itemCC(7) = 1.05
        itemCC(8) = 1.03
        itemCC(9) = 1.02
        itemCC(10) = 1.01
        itemCC(11) = 1.03
        itemCC(12) = 1.005
        itemCC(13) = 1.001
    
        'define rate coefficients (gen > 1 | item > 0)
        generatorRC(0) = 1.3
        generatorRC(1) = 1.05
        generatorRC(2) = 1.02
        generatorRC(3) = 1.01
        generatorRC(4) = 1.001
        generatorRC(5) = 1.0005
        generatorRC(6) = 1.0005
        generatorRC(7) = 1.0005
        generatorRC(8) = 1.0005
        generatorRC(9) = 1.0001
        generatorRC(10) = 1.0001
        generatorRC(11) = 1.00001
        generatorRC(12) = 1.000005
        
        itemRC(0) = 2
        itemRC(1) = 0.1
        itemRC(2) = 0.15
        itemRC(3) = 0.3
        itemRC(4) = 0.15
        itemRC(5) = 0.2
        itemRC(6) = 0.1
        itemRC(7) = 0.15
        itemRC(8) = 0.1
        itemRC(9) = 0.07
        itemRC(10) = 0.06
        itemRC(11) = 0.04
        itemRC(12) = 0.05
        itemRC(13) = 0.03
    
    
'init
        x = 0
        For x = 0 To UBound(generatorRate)
            generatorCost(x) = generatorBasecost(x)
            generatorRevenue(x) = Cells(x + 4, 11).Value
            generatorQty(x) = Cells(x + 4, 5).Value
            If generatorQty(x) > 0 Then
                generatorRate(x) = generatorBaserate(x) * (generatorQty(x) ^ generatorRC(x))
            Else
                generatorRate(x) = 0
            End If
            Cells(x + 4, 4).Value = generatorRate(x) * itemRate(x + 1)
            
            For y = 1 To generatorQty(x)
                generatorCost(x) = generatorCost(x) ^ generatorCC(x)
            Next y
            Cells(x + 4, 3).Value = generatorCost(x)
        Next x
        
        x = 0
        For x = 0 To UBound(itemRate)
            itemCost(x) = itemBasecost(x)
            itemQty(x) = Cells(x + 3, 10).Value
            
            For y = 1 To itemQty(x)
                itemCost(x) = itemCost(x) ^ itemCC(x)
            Next y
            Cells(x + 3, 8).Value = itemCost(x)
            
            If itemQty(x) > 0 Then
                itemRate(x) = (itemQty(x) + 1) ^ itemRC(x)
                Cells(x + 3, 6).Value = itemRate(x)
            Else
                itemRate(x) = 1
            End If
            
            Cells(x + 3, 6).Value = itemRate(x)
        
        Next x
        
        x = 0
        For x = 0 To UBound(profitTotals)
            profitTotals(x) = Cells(x + 2, 12).Value
        Next x
            
        ActiveSheet.Rows("6:16").EntireRow.Hidden = True
        x = 6
        For x = 5 To 16
            If Cells(x, 5).Value >= 10 Then
                ActiveSheet.Rows(x + 1).EntireRow.Hidden = False
            End If
        Next x
            
        alertTime = Now + TimeValue("00:00:01")
        Application.OnTime alertTime, "Loop1000"
        
        
'main loop
        Do While totalRevenue < 1.5 * (10 ^ 50)
            DoEvents
            Cells(2, 11) = totalRevenue
            Cells(3, 11) = clickRevenue
            x = 0
            For x = 0 To UBound(generatorRevenue)
                Cells(x + 4, 11) = generatorRevenue(x)
            Next x
            
            profitTotals(0) = Application.Sum(Range(Cells(3, 12), Cells(16, 12)))
            
            x = 0
            For x = 0 To UBound(profitTotals)
                Cells(x + 2, 12).Value = profitTotals(x)
            Next x
        Loop
        
    End With
    
End Sub
Sub Loop1000()
    'generator loop
    y = 0
    For y = 0 To UBound(generatorRevenue)
        If generatorQty(y) > 0 Then
            generatorRevenue(y) = generatorRevenue(y) + (generatorRate(y) * itemRate(y + 1))
            Cells(y + 4, 4).Value = generatorRate(y) * itemRate(y + 1)
            totalRevenue = totalRevenue + (generatorRate(y) * itemRate(y + 1))
            Cells(y + 4, 11).Value = generatorRevenue(y)
            profitTotals(y + 2) = profitTotals(y + 2) + (generatorRate(y) * itemRate(y + 1))
        End If
    Next y
    
    'totals
    Cells(2, 3).Value = buildSpend
    Cells(2, 4).Value = Application.Sum(Range(Cells(4, 4), Cells(16, 4)))
    Cells(2, 5).Value = Application.Sum(Range(Cells(4, 5), Cells(16, 5)))
    Cells(2, 8).Value = itemSpend
    Cells(2, 10).Value = Application.Sum(Range(Cells(3, 10), Cells(16, 10)))
    
    alertTime = Now + TimeValue("00:00:01")
    Application.OnTime alertTime, "Loop1000"
    
End Sub
Sub ItemButton(index As Integer)
    If itemCost(index) <= 0 Then Exit Sub
    If profitTotals(0) >= itemCost(index) Then
        itemSpend = itemSpend + itemCost(index)
        profitTotals(index + 1) = profitTotals(index + 1) - itemCost(index)
        
        itemQty(index) = itemQty(index) + 1
        
        itemRate(index) = (itemQty(index) + 1) ^ itemRC(index)
        
        Cells(index + 3, 10).Value = itemQty(index)
        Cells(index + 3, 6).Value = itemRate(index)
        
        itemCost(index) = itemCost(index) ^ itemCC(index)
        Cells(index + 3, 8).Value = itemCost(index)
    End If
End Sub
Sub GeneratorButton(index)
    If generatorCost(index) <= 0 Then Exit Sub
    If profitTotals(0) >= generatorCost(index) Then
        buildSpend = buildSpend + generatorCost(index)
        profitTotals(index + 2) = profitTotals(index + 2) - generatorCost(index)
        
        generatorQty(index) = generatorQty(index) + 1
        If generatorQty(index) >= 10 Then
            ActiveSheet.Rows(index + 5).EntireRow.Hidden = False
        End If
        
        generatorRate(index) = generatorBaserate(index) * (generatorQty(index) ^ generatorRC(index))
        
        Cells(index + 4, 5).Value = generatorQty(index)
        Cells(index + 4, 4).Value = generatorRate(index) * itemRate(index + 1)
        
        generatorCost(index) = generatorCost(index) ^ generatorCC(index)
        Cells(index + 4, 3).Value = generatorCost(index)
    End If
End Sub
