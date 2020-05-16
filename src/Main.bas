Attribute VB_Name = "Main"

' begin program
Sub Main()
HollierForm.Show
End Sub

' begin solving for Hollier
Sub HollierMethod()
Dim Solver As StepNode
Dim tableRange As Range

Set Solver = New StepNode
'Set tableRange = HollierForm.InputRange
'StepNode.InitializeVariables tableRange

' Testing
Set tableRange = Range("Sheet1!$B$2:$F$6")
'Set tableRange = Range("Sheet1!$C$3:$F$6")
HollierForm.CheckBox1.Value = True
Solver.InitializeVariables tableRange
End Sub
