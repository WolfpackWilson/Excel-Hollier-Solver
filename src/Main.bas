Attribute VB_Name = "Main"

' begin program
Sub Main()
HollierForm.Show
End Sub

' begin solving for Hollier
Sub HollierMethod()
Dim Solver As StepNode
Dim tableValues() As Variant
Dim outputRange As Range

Set Solver = New StepNode
Set Solver2 = New StepNode
tableValues = HollierForm.InputRange

' choose output range based on the form selection
If (HollierForm.OptionButton1.value = True) Then
    Set outputRange = HollierForm.OutputCell
Else
    Set outputRange = Range("A1")
End If

' solve using method 1
Solver.InitializeVariables tableValues, HollierForm.CheckBox1.value
Solver.SolveHollier
Solver.OutputSolution outputRange, HollierForm.CheckBox2.value

' solve using method 2, if desired
If (HollierForm.CheckBox3.value = True) Then
    Solver2.InitializeVariables tableValues, HollierForm.CheckBox1.value, hollier2:=True
    Solver2.SolveHollier
    Solver2.OutputSolution outputRange, HollierForm.CheckBox2.value
End If

End Sub
