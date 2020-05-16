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
Private pFromTotals() As Variant
Private pToTotals() As Variant
Private pMachineNbs() As Variant
Private pSolution() As Variant
Private pTableSize As Integer

Private Sub Class_Initialize()
' Initialize in a different sub
End Sub

' Extract table information to initalize variables
Public Sub InitializeVariables(ByRef tableRange As Range, Optional ByRef MachineNbs As Variant)
If (HollierForm.CheckBox1.Value = True) Then
    Set pFromToRange = tableRange.Offset(1, 1).Resize(tableRange.Rows.count - 1, _
         tableRange.Columns.count - 1)
    pFromToValues() = tableRange
    pTableSize = UBound(pFromToValues, 1) - LBound(pFromToValues, 1)
    ReDim pMachineNbs(pTableSize - 1)
    
    ' get current machine numbers
    For i = 0 To pTableSize - 1
        pMachineNbs(i) = pFromToValues(1, i + 2)
    Next
    
    pFromToValues() = RemoveRowCol(pFromToValues, 2)
Else
    pFromToRange = tableRange
    pFromToValues() = tableRange
    pTableSize = UBound(pFromToValues, 1) - LBound(pFromToValues, 1) + 1
    
    ' set current machine numbers
    For i = 1 To pTableSize
        pMachineNbs(i) = i
    Next
End If
End Sub

' remove the row and col of a two dimension array based on a 1-based position
Function RemoveRowCol(ByRef twoDimArr As Variant, ByVal posNb As Integer)
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
