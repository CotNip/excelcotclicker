Attribute VB_Name = "buttons"

Sub Quit()
    ActiveWorkbook.Save
    Application.Quit
End Sub
Sub StopCode()
        Stop
    End
End Sub

Sub Reset()
    If MsgBox("Are you sure?" & vbNewLine & "(All progress will be reset)", vbYesNo + vbExclamation) = vbYes Then
        Range("J2:L16").ClearContents
        Range("E2:E16").ClearContents
        Cells(2, 3).ClearContents
        Cells(2, 8).ClearContents
    End If
End Sub
Sub Click()
    totalRevenue = totalRevenue + itemRate(0)
    clickRevenue = clickRevenue + itemRate(0)
    profitTotals(1) = profitTotals(1) + itemRate(0)
End Sub

'-------------------------------------------------- buildings --------------------------------------------------
Sub BuildIntern() 'index 0
    Call GeneratorButton(0)
End Sub
Sub BuildJradmin() 'index 1
    Call GeneratorButton(1)
End Sub
Sub BuildSnradmin() 'index 2
    Call GeneratorButton(2)
End Sub
Sub BuildOperations() 'index 3
    Call GeneratorButton(3)
End Sub
Sub BuildConsultant() 'index 4
    Call GeneratorButton(4)
End Sub
Sub BuildAsstmgr() 'index 5
    Call GeneratorButton(5)
End Sub
Sub BuildSnrmgr() 'index 6
    Call GeneratorButton(6)
End Sub
Sub BuildAcct() 'index 7
    Call GeneratorButton(7)
End Sub
Sub BuildSnracct() 'index 8
    Call GeneratorButton(8)
End Sub
Sub BuildHR() 'index 9
    Call GeneratorButton(9)
End Sub
Sub BuildChiefFin() 'index 10
    Call GeneratorButton(10)
End Sub
Sub BuildChiefOps() 'index 11
    Call GeneratorButton(11)
End Sub
Sub BuildChiefExec() 'index 12
    Call GeneratorButton(12)
End Sub

'-------------------------------------------------- Items --------------------------------------------------
Sub ItemClicks() 'index 0
    Call ItemButton(0)
End Sub
Sub ItemTea() 'index 1
    Call ItemButton(1)
End Sub
Sub ItemHours() 'index 2
    Call ItemButton(2)
End Sub
Sub ItemOvertime() 'index 3
    Call ItemButton(3)
End Sub
Sub ItemPantry() 'index 4
    Call ItemButton(4)
End Sub
Sub ItemOversight() 'index 5
    Call ItemButton(5)
End Sub
Sub ItemBenefits() 'index 6
    Call ItemButton(6)
End Sub
Sub ItemRaises() 'index 7
    Call ItemButton(7)
End Sub
Sub ItemOptions() 'index 8
    Call ItemButton(8)
End Sub
Sub ItemTrust() 'index 9
    Call ItemButton(9)
End Sub
Sub ItemStaffinfo() 'index 10
    Call ItemButton(10)
End Sub
Sub ItemShares() 'index 11
    Call ItemButton(11)
End Sub
Sub ItemCars() 'index 12
    Call ItemButton(12)
End Sub
Sub ItemTax() 'index 13
    Call ItemButton(13)
End Sub

