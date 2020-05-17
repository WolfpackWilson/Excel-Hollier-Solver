VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StepNode"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

' node pointers
Private nextStep As Object
Private prevStep As Object

' node Hollier values
Private pFromToValues() As Variant
Private pFromTotals() As Double
Private pToTotals() As Double
Private pMachineNbs() As Variant
Private pSolution As Variant
Private pTableSize As Integer
Private pInitialized As Boolean
Private pIn, pOut As Integer

Private Sub Class_Initialize()
' initialize in a different sub
pInitialized = False
End Sub

' extract table information to initalize variables
Public Sub InitializeVariables(ByRef tableValues As Variant, ByVal labels As Boolean, _
        Optional ByRef MachineNbs As Variant, Optional ByRef prevNode As StepNode)
If (labels) Then
    pFromToValues() = tableValues
    pTableSize = UBound(pFromToValues, 1) - LBound(pFromToValues, 1)
    ReDim pMachineNbs(1 To pTableSize)
    
    ' get current machine numbers
    For i = 1 To pTableSize
        pMachineNbs(i) = pFromToValues(1, i + 1)
    Next
    
    pFromToValues() = RemoveRowCol(pFromToValues, 1)
Else
    pFromToValues() = tableValues
    pTableSize = UBound(pFromToValues, 1) - LBound(pFromToValues, 1) + 1
    
    If IsMissing(MachineNbs) Then
        ReDim pMachineNbs(1 To pTableSize)
        ' set current machine numbers
        For i = 1 To pTableSize
            pMachineNbs(i) = i
        Next
    Else
        pMachineNbs = MachineNbs
    End If
End If

sumFromsAndTos pFromToValues
pInitialized = True
pIn = 0
pOut = 0

If Not (IsMissing(prevNode)) Then
    Set prevStep = prevNode
End If
End Sub

Public Property Get NextNode() As StepNode
Set NextNode = nextStep
End Property

Public Property Get PrevNode() As StepNode
Set PrevNode = prevStep
End Property

Public Property Get Solution() As Variant
Solution = pSolution
End Property

Public Property Get Initialized() As Boolean
Initialized = pInitialized
End Property

Public Property Get InValue() As Integer
InValue = pIn
End Property
Public Property Let InValue(value As Integer)
pIn = value
End Property

Public Property Get OutValue() As Integer
OutValue = pOut
End Property
Public Property Let OutValue(value As Integer)
pOut = value
End Property


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

' solves using hollier method 1
Public Sub SolveHollier()
Dim currPos, lowest, lowestPos As Integer
currPos = 1
lowest = LBound(pToTotals, 1)
lowestPos = currPos

' find the lowest value/position
For i = LBound(pMachineNbs, 1) To UBound(pMachineNbs, 1)
    If (pToTotals(i) < lowest) Then
        lowest = pToTotals(i)
        lowestPos = currPos
    End If
    currPos = currPos + 1
Next
pSolution = pMachineNbs(lowestPos)

If (UBound(pMachineNbs, 1) - LBound(pMachineNbs, 1) > 0) Then
    ' get machine numbers for the next step
    Dim newMachineNbs As Variant
    ReDim newMachineNbs(1 To UBound(pMachineNbs, 1) - LBound(pMachineNbs, 1))
    currPos = 1
    
    For i = LBound(pMachineNbs, 1) To UBound(pMachineNbs, 1)
        If Not ((pMachineNbs(i) = pSolution)) Then
            newMachineNbs(currPos) = pMachineNbs(i)
            currPos = currPos + 1
        End If
    Next
    
    Dim newValues As Variant
    newValues = RemoveRowCol(pFromToValues, lowestPos)
    
    Set nextStep = New StepNode
    nextStep.InitializeVariables newValues, False, newMachineNbs, Me

    nextStep.SolveHollier
End If
End Sub

Public Sub OutputSolution()
Dim nodeArr() As PointerNode
ReDim nodeArr(1 To (pTableSize * pTableSize + pTableSize)) ' table size, ins, outs, not diagonals

Dim count, currR, currC, inSeq, backtrack, totalMoves As Integer
Dim posDict, machineDict, stepNodeDict As Object
Dim tmpStep As StepNode

Set posDict = CreateObject("Scripting.Dictionary")      ' machine - position
Set machineDict = CreateObject("Scripting.Dictionary")    ' position - machine
Set stepNodeDict = CreateObject("Scripting.Dictionary") ' machine - StepNode
Set tmpStep = Me
count = 1

' fill dictionaries
While (Not (tmpStep Is Nothing))
    posDict.Add tmpStep.Solution, count
    machineDict.Add count, tmpStep.Solution
    stepNodeDict.Add tmpStep.Solution, tmpStep
    count = count + 1
    Set tmpStep = tmpStep.NextNode
Wend

' add table values as nodes
count = 1
currR = 1
currC = 1
inSeq = 0
backtrack = 0
totalMoves = 0
For i = LBound(pFromToValues, 1) To UBound(pFromToValues, 1)
    For j = LBound(pFromToValues, 2) To UBound(pFromToValues, 2)
        If (Not (currR = currC)) Then
            Dim nextSNode, prevSNode As StepNode
            Dim ptrNode As PointerNode
            Dim value, dist As Integer
            
            ' add pointer node
            Set ptrNode = New PointerNode
            Set prevSNode = stepNodeDict(pMachineNbs(currR))
            Set nextSNode = stepNodeDict(pMachineNbs(currC))
            dist = posDict(nextSNode.Solution) - posDict(prevSNode.Solution)
            value = pFromToValues(i, j)
            
            ptrNode.InitializeVariables nextSNode, prevSNode, value, dist
            Set nodeArr(count) = ptrNode
            
            ' update moves
            totalMoves = totalMoves + value
            If (dist = 1) Then
                inSeq = inSeq + value
            ElseIf (dist < 0) Then
                backtrack = backtrack + value
            End If
            
            ' update step node in/out moves
            If (Not (prevSNode Is Nothing)) Then
                prevSNode.OutValue = prevSNode.OutValue + value
            End If
            
            If (Not (nextSNode Is Nothing)) Then
                nextSNode.InValue = nextSNode.InValue + value
            End If
            
            count = count + 1
        End If
        currC = currC + 1
    Next
    currC = 1
    currR = currR + 1
Next

' TODO: add extra ins and outs

End Sub
