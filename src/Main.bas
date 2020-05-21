Attribute VB_Name = "Main"

' begin program
Sub Main()
HollierForm.Show
End Sub

' begin solving for Hollier
Sub HollierMethod()
Dim Solver As StepNode
Dim tableValues() As Variant

Set Solver = New StepNode
Set Solver2 = New StepNode
'tableValues = HollierForm.InputRange

' Testing
Dim aRange As Range
Set aRange = Range("B17")

tableValues = Range("Sheet1!$B$2:$F$6")
'tableValues = Range("Sheet1!$C$3:$F$6")
Solver.InitializeVariables tableValues, True
Solver.SolveHollier
Solver.OutputSolution aRange, True

' surround by if when wrapping up
Solver2.InitializeVariables tableValues, True, hollier2:=True
Solver2.SolveHollier
Solver2.OutputSolution aRange, True

End Sub
