VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HollierSolver"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private pFromToRange As Range
Private pFromToValues() As Variant
Private pFromTotals() As Double
Private pToTotals() As Double
Private pMachineNbs() As Variant
Private pSolution() As Variant
Private pTableSize As Integer

Private Sub Class_Initialize()
' initialize in a different sub
End Sub

' extract table information to initalize variables
Public Sub InitializeVariables(ByRef tableRange As Range, Optional ByRef MachineNbs As Variant)
If (HollierForm.CheckBox1.Value = True) Then
    Set pFromToRange = tableRange.Offset(1, 1).Resize(tableRange.Rows.count - 1, _
         tableRange.Columns.count - 1)
    pFromToValues() = tableRange
    pTableSize = UBound(pFromToValues, 1) - LBound(pFromToValues, 1)
    ReDim pMachineNbs(1 To pTableSize)
    
    ' get current machine numbers
    For i = 1 To pTableSize
        pMachineNbs(i) = pFromToValues(1, i + 1)
    Next
    
    pFromToValues() = RemoveRowCol(pFromToValues, 1)
Else
    Set pFromToRange = tableRange
    pFromToValues() = tableRange
    pTableSize = UBound(pFromToValues, 1) - LBound(pFromToValues, 1) + 1
    ReDim pMachineNbs(1 To pTableSize)
    
    ' set current machine numbers
    For i = 1 To pTableSize
        pMachineNbs(i) = i
    Next
End If

sumFromsAndTos pFromToValues
End Sub

' remove the row and col of a two dimension array based on a 1-based position
Private Function RemoveRowCol(ByRef twoDimArr As Variant, ByVal posNb As Integer)
Dim arrSize, newArrSize As Integer
Dim newArr() As Variant
arrSize = UBound(twoDimArr, 1) - LBound(twoDimArr, 1) + 1

ReDim newArr(1 To (arrSize - 1), 1 To (arrSize - 1))
Dim currPosR, currPosC, newArrR, newArrC As Integer
currPosR = 1
currPosC = 1
newArrR = 1
newArrC = 1

For i = LBound(twoDimArr, 1) To UBound(twoDimArr, 1)
    If (Not (currPosR = posNb)) Then
        For j = LBound(twoDimArr, 2) To UBound(twoDimArr, 2)
            If (Not (currPosC = posNb)) Then
                newArr(newArrR, newArrC) = twoDimArr(i, j)
                newArrC = newArrC + 1
            End If
            currPosC = currPosC + 1
        Next
        
        newArrC = 1
        newArrR = newArrR + 1
    End If
    
    currPosC = 1
    currPosR = currPosR + 1
Next

RemoveRowCol = newArr
End Function

' total To and From column sums
Private Sub sumFromsAndTos(ByRef twoDimArr As Variant)
Dim arrSize, newArrSize As Integer
arrSize = UBound(twoDimArr, 1) - LBound(twoDimArr, 1) + 1

ReDim pFromTotals(1 To arrSize)
ReDim pToTotals(1 To arrSize)
Dim currPosR, currPosC As Integer
currPosR = 1
currPosC = 1

For i = LBound(twoDimArr, 1) To UBound(twoDimArr, 1)
    For j = LBound(twoDimArr, 2) To UBound(twoDimArr, 2)
        pFromTotals(currPosR) = pFromTotals(currPosR) + twoDimArr(i, j)
        pToTotals(currPosC) = pToTotals(currPosC) + twoDimArr(i, j)
        
        currPosC = currPosC + 1
    Next
    
    currPosC = 1
    currPosR = currPosR + 1
Next

End Sub
