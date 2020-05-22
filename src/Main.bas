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
        Dim shtName As String
        Dim exists As Boolean
        Dim i As Integer
        
        shtName = "Hollier Output"
        exists = True
        i = 1
        
        If (WorksheetExists(shtName)) Then
            While (exists)
                shtName = "Hollier Output " & i
                exists = WorksheetExists(shtName)
                i = i + 1
            Wend
        End If
        
        Sheets.Add.Name = shtName
        Set outputRange = Range("B2")
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

' check if a worksheet exists
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function
