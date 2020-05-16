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
'tableValues = HollierForm.InputRange
'StepNode.InitializeVariables tableValues, HollierForm.CheckBox1.Value

' Testing
tableValues = Range("Sheet1!$B$2:$F$6")
'tableValues = Range("Sheet1!$C$3:$F$6")
Solver.InitializeVariables tableValues, True
Solver.SolveHollier

End Sub
