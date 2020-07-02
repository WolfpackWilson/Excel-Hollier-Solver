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

''=======================================================
'' Class:       StepNode
'' Desc:        Stores data and methods for the program.
''              It is also used to solve and display
''              outputs for the program.
'' Called by:   HollierProgram
'' Changes----------------------------------------------
'' Date         Programmer      Change
'' 05/25/2020   TheEric960      Written
'' 06/11/2020   TheEric960      v1.1.0
''=======================================================

Option Base 1

' ===========================================
' VARIABLE DECLARATIONS
' ===========================================

' -------------------------------------------
' node pointers
' -------------------------------------------
Private nextStep As Object
Private prevStep As Object
Private lastStep As Object

' -------------------------------------------
' node Hollier values
' -------------------------------------------
Private pFromToValues() As Variant
Private pFromTotals() As Variant
Private pToTotals() As Variant
Private pFromToRatios() As Double
Private pMachineNbs() As Variant
Private pSolution As Variant
Private pTableSize As Long
Private pIn, pOut As Long

' -------------------------------------------
' templates
' -------------------------------------------
Private Type sumsAndRatios
    fromTotals As Variant
    toTotals As Variant
    fromToRatios As Variant
End Type

' -------------------------------------------
' flags
' -------------------------------------------
Private pInitialized As Boolean
Private pHollier2 As Boolean

' -------------------------------------------
' output variables
' -------------------------------------------
Private pShape As Shape

Private Const INTMAX  As Long = 2147483647

' ===========================================
' INITIALIZE CLASS VARIABLES
' ===========================================

' -------------------------------------------
' initialize the class
' -------------------------------------------
Private Sub Class_Initialize()
    ' initialize in a different sub
    pInitialized = False
    pSolution = -1
End Sub

' -------------------------------------------
' extract table information to initalize variables
' -------------------------------------------
Public Sub InitializeVariables(ByRef tableValues As Variant, ByVal labels As Boolean, _
        Optional ByRef machineNbs As Variant, Optional ByRef prevNode As StepNode, _
        Optional ByVal hollier2 As Boolean)
    On Error GoTo EH
        
    If (labels) Then
        pFromToValues() = tableValues
        pTableSize = UBound(pFromToValues, 1) - LBound(pFromToValues, 1)
        ReDim pMachineNbs(1 To pTableSize)
        
        ''get current machine numbers
        For i = 1 To pTableSize
            pMachineNbs(i) = pFromToValues(1, i + 1)
        Next
        
        pFromToValues() = RemoveRowCol(pFromToValues, 1)
    Else
        pFromToValues() = tableValues
        pTableSize = UBound(pFromToValues, 1) - LBound(pFromToValues, 1) + 1
        
        If IsMissing(machineNbs) Then
            ReDim pMachineNbs(1 To pTableSize)
            ''set current machine numbers
            For i = 1 To pTableSize
                pMachineNbs(i) = i
            Next
        Else
            pMachineNbs = machineNbs
        End If
    End If
    
    ''make sure all diagonal values are 0
    For i = LBound(pFromToValues, 1) To pTableSize
        For j = LBound(pFromToValues, 2) To pTableSize
            If (i = j) Then
                pFromToValues(i, j) = 0
            End If
        Next
    Next
    
    ''set sums and ratios
    Dim sumsRatios As sumsAndRatios
    sumsRatios = sumFromsAndTos(pFromToValues)
    With sumsRatios
        ReDim pFromToRatios(1 To UBound(.fromToRatios))
        ReDim pFromTotals(1 To UBound(.fromTotals))
        ReDim pToTotals(1 To UBound(.toTotals))
        
        pFromToRatios = .fromToRatios
        pFromTotals = .fromTotals
        pToTotals = .toTotals
    End With
    
    pInitialized = True
    pIn = 0
    pOut = 0
    
    Set lastStep = New StepNode
    
    '' optional parameters
    If Not (IsMissing(prevNode)) Then
        Set prevStep = prevNode
    End If
    
    If Not (IsMissing(lastNode)) Then
        Set lastStep = lastNode
    End If
    
    If (IsMissing(hollier2)) Then
        pHollier2 = False
    Else
        pHollier2 = hollier2
    End If
    Exit Sub
    
EH:
    RaiseError Err.Number, Err.Source, "HollierProgram.InitializeVariables", _
            Err.Description, Erl
End Sub


' ===========================================
' GETTERS AND SETTERS
' ===========================================

Public Property Get nextNode() As StepNode
    Set nextNode = nextStep
End Property
Public Property Let nextNode(nextNode As StepNode)
    Set nextStep = nextNode
End Property

Public Property Get prevNode() As StepNode
    Set prevNode = prevStep
End Property
Public Property Let prevNode(prevNode As StepNode)
    Set prevStep = prevNode
End Property

Public Property Get lastNode() As StepNode
    Set lastNode = lastStep
End Property
Public Property Let lastNode(lastNode As StepNode)
    Set lastStep = lastNode
End Property

Public Property Get Solution() As Variant
    Solution = pSolution
End Property
Public Property Let Solution(value As Variant)
    pSolution = value
End Property

Public Property Get Initialized() As Boolean
    Initialized = pInitialized
End Property

Public Property Get InValue() As Long
    InValue = pIn
End Property
Public Property Let InValue(value As Long)
    pIn = value
End Property

Public Property Get OutValue() As Long
    OutValue = pOut
End Property
Public Property Let OutValue(value As Long)
    pOut = value
End Property

Public Property Get Shape() As Shape
    Set Shape = pShape
End Property
Public Property Let Shape(aShape As Shape)
    Set pShape = aShape
End Property


' ===========================================
' CLASS METHODS
' ===========================================

' -------------------------------------------
' remove the row and col of a two dimension array based on a 1-based position
' -------------------------------------------
Private Function RemoveRowCol(ByRef twoDimArr As Variant, ByVal posNb As Long)
    Dim arrSize, newArrSize As Long
    Dim newArr() As Variant
    arrSize = UBound(twoDimArr, 1) - LBound(twoDimArr, 1) + 1
    
    ''if array is too small
    If (arrSize - 1 <= 1) Then
        ReDim newArr(1, 1)
        If (posNb = 1) Then
            newArr(1, 1) = twoDimArr(2, 2)
        Else
            newArr(1, 1) = twoDimArr(1, 1)
        End If
        
        RemoveRowCol = newArr
        Exit Function
    End If
    
    ReDim newArr(1 To (arrSize - 1), 1 To (arrSize - 1))
    Dim currPosR, currPosC, newArrR, newArrC As Long
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

' -------------------------------------------
' total To and From column sums
' -------------------------------------------
Private Function sumFromsAndTos(ByRef twoDimArr As Variant) As sumsAndRatios
    Dim arrSize, newArrSize As Long
    arrSize = UBound(twoDimArr, 1) - LBound(twoDimArr, 1) + 1
    
    Dim fromTotals(), toTotals(), fromToRatios() As Double
    ReDim fromTotals(1 To arrSize)
    ReDim toTotals(1 To arrSize)
    Dim currPosR, currPosC As Long
    currPosR = 1
    currPosC = 1
    
    ''total froms and tos
    For i = LBound(twoDimArr, 1) To UBound(twoDimArr, 1)
        For j = LBound(twoDimArr, 2) To UBound(twoDimArr, 2)
            fromTotals(currPosR) = fromTotals(currPosR) + twoDimArr(i, j)
            toTotals(currPosC) = toTotals(currPosC) + twoDimArr(i, j)
            
            currPosC = currPosC + 1
        Next
        
        currPosC = 1
        currPosR = currPosR + 1
    Next
    
    ''total from/to ratio
    ReDim fromToRatios(1 To arrSize)
    For i = LBound(toTotals) To UBound(toTotals)
        If (toTotals(i) = 0) Then
            fromToRatios(i) = INTMAX
        Else
            fromToRatios(i) = fromTotals(i) / toTotals(i)
        End If
    Next
    
    Dim answer As sumsAndRatios
    With answer
        .fromToRatios = fromToRatios
        .fromTotals = fromTotals
        .toTotals = toTotals
    End With
    
    sumFromsAndTos = answer
End Function

' -------------------------------------------
' solves using a hollier method
' -------------------------------------------
Public Sub SolveHollier(Optional ByRef toTotals As Variant, Optional ByRef fromTotals As Variant, _
        Optional ByRef fromToValues As Variant, Optional ByRef fromToRatios As Variant, _
        Optional ByRef machineNbs As Variant, Optional ByRef skipList As Object)
    On Error GoTo EH
    
    Dim minTo, minTo2, minToPos, minToPos2 As Long
    Dim minFrom, minFrom2, minFromPos, minFromPos2 As Long
    Dim currPos, high, highPos As Long
    Dim sumsRatios As sumsAndRatios
    
    If (IsMissing(toTotals)) Then
        toTotals = pToTotals
        fromTotals = pFromTotals
        fromToValues = pFromToValues
        fromToRatios = pFromToRatios
        machineNbs = pMachineNbs
    End If
    
    If (skipList Is Nothing) Then
        Set skipList = New ArrayList
    End If
    
    ''test how many machines are left to sort
    If (UBound(machineNbs) = LBound(machineNbs)) Then
        pSolution = machineNbs(LBound(machineNbs))
        Exit Sub
    End If
    
    currPos = 1
    
    ''hollier 1 variables
    minTo = INTMAX
    minToPos = currPos
    minTo2 = INTMAX
    minToPos2 = currPos
    
    minFrom = INTMAX
    minFromPos = currPos
    minFrom2 = INTMAX
    minFromPos2 = currPos
    
    ''hollier 2 variables
    high = fromToRatios(LBound(fromToRatios, 1))
    highPos = currPos
    
    ''find the lowest values and their positions
    For i = LBound(pMachineNbs, 1) To UBound(machineNbs, 1)
        If (pHollier2) Then
            If (fromToRatios(i) > high) Then
                high = fromToRatios(i)
                highPos = currPos
            End If
        Else ''hollier 1
            If (Not skipList.Contains(machineNbs(i))) Then
                If (toTotals(i) <= minTo) Then
                    minTo2 = minTo
                    minToPos2 = minToPos
                    
                    minTo = toTotals(i)
                    minToPos = currPos
                End If
                
                If (fromTotals(i) <= minFrom) Then
                    minFrom2 = minFrom
                    minFromPos2 = minFromPos
                    
                    minFrom = fromTotals(i)
                    minFromPos = currPos
                End If
            End If
        End If
        currPos = currPos + 1
    Next
    
    If (pHollier2) Then
        pSolution = machineNbs(highPos)
    Else
        ''solve using hollier 1
        If (minTo < minFrom Or (skipList.count = UBound(machineNbs))) Then
            If (minTo = minTo2) Then
                ''tiebreaker using from-to ratios (lowest chosen)
                If (fromToRatios(minToPos) > fromToRatios(minToPos2)) Then
                    pSolution = machineNbs(minToPos2)
                Else
                    pSolution = machineNbs(minToPos)
                End If
            Else
                If (skipList.count = UBound(machineNbs)) Then
                    skipList.Remove machineNbs(minToPos)
                End If
                
                pSolution = machineNbs(minToPos)
            End If
        ElseIf (minTo >= minFrom) Then
            Dim tmpMachineNbs() As Variant
            Dim tmpFromToValues() As Variant
            tmpMachineNbs = machineNbs
            tmpFromToValues = fromToValues
        
            If (minFrom = minFrom2 And minTo <> minFrom) Then
                ''tiebreaker using from-to ratios (lowest chosen)
                If (fromToRatios(minFromPos) > fromToRatios(minFromPos2)) Then
                    minFrom = minFrom2
                    minFromPos = minFromPos2
                End If
            Else
                ''other tiebreakers
                If (minTo = minFrom) Then
                    If (minToPos <> minFromPos) Then
                        pSolution = machineNbs(minToPos)
                        tmpMachineNbs = Remove(machineNbs, minToPos)
                        tmpFromToValues = RemoveRowCol(tmpFromToValues, minToPos)
                        
                        If (UBound(tmpMachineNbs) = 1) Then GoTo SkipFromRemovals2
                    Else ''minToPos = minFromPos
                        If (Not skipList.Contains(machineNbs(minToPos))) Then
                            skipList.Add machineNbs(minToPos)
                        End If
                        
                        GoTo SkipFromRemovals
                    End If
                End If
            End If
            
            ''add the "from" machine and continue to search for "to" selection
            Dim tempStep As StepNode
            Set tempStep = New StepNode
            
            If (Not (lastStep Is Nothing)) Then
                Set tempStep = lastStep
            End If
            
            Set lastStep = New StepNode
            
            With lastStep
                If (minTo = minFrom And minFromPos > minToPos) Then
                    tmpMachineNbs = Remove(tmpMachineNbs, minFromPos - 1)
                    tmpFromToValues = RemoveRowCol(tmpFromToValues, minFromPos - 1)
                Else
                    tmpMachineNbs = Remove(tmpMachineNbs, minFromPos)
                    tmpFromToValues = RemoveRowCol(tmpFromToValues, minFromPos)
                End If
                
                .InitializeVariables tmpFromToValues, False, tmpMachineNbs, _
                        hollier2:=pHollier2
                .Solution = machineNbs(minFromPos)
                .nextNode = tempStep
            End With
            
            fromToValues = RemoveRowCol(fromToValues, minFromPos)
            machineNbs = Remove(machineNbs, minFromPos)
            sumsRatios = sumFromsAndTos(fromToValues)
            
            With sumsRatios
                fromTotals = .fromTotals
                toTotals = .toTotals
                fromToRatios = .fromToRatios
            End With
SkipFromRemovals:
            If (Not (minTo = minFrom And minToPos <> minFromPos)) Then
                If (minTo = minFrom And minToPos = minFromPos) Then
                    SolveHollier toTotals, fromTotals, fromToValues, fromToRatios, _
                            machineNbs, skipList
                    Exit Sub
                Else
                    SolveHollier toTotals, fromTotals, fromToValues, fromToRatios, _
                            machineNbs
                    Exit Sub
                End If
            End If
       End If
    End If
    
SkipFromRemovals2:
    'remove table rows/columns for the next step
    Dim sPos As Integer
    sPos = 1
    
    For i = 1 To UBound(machineNbs)
        If (machineNbs(i) = pSolution) Then
            sPos = i
            Exit For
        End If
    Next
    
    If (UBound(toTotals) <> LBound(toTotals)) Then
        fromToValues = RemoveRowCol(fromToValues, sPos)
        machineNbs = Remove(machineNbs, sPos)
        
        If (Not pHollier2) Then
            sumsRatios = sumFromsAndTos(fromToValues)
            
            With sumsRatios
                fromTotals = .fromTotals
                toTotals = .toTotals
                fromToRatios = .fromToRatios
            End With
        Else ''don't update fromToRatios for hollier 2
            fromTotals = Remove(fromTotals, sPos)
            toTotals = Remove(toTotals, sPos)
            fromToRatios = Remove(fromToRatios, sPos)
        End If
    End If
    
    ''repeat for subsequent steps
    
    Set nextStep = New StepNode
    nextStep.InitializeVariables fromToValues, False, machineNbs, Me, _
            pHollier2
    nextStep.SolveHollier toTotals, fromTotals, fromToValues, _
            fromToRatios, machineNbs
    Exit Sub
    
EH:
    RaiseError Err.Number, Err.Source, "HollierProgram.SolveHollier", Err.Description, Erl
End Sub

' -------------------------------------------
' remove a value from an array based on its position
' -------------------------------------------
Private Function Remove(ByRef arr As Variant, ByVal pos As Integer) As Variant
    Dim newArr() As Variant
    
    ''if array is too small
    If (UBound(arr) - LBound(arr) <= 1) Then
        ReDim newArr(1)
        
        If (pos = 1) Then
            newArr(1) = arr(2)
        Else
            newArr(1) = arr(1)
        End If
        
        Remove = newArr
        Exit Function
    End If
    
    ReDim newArr(1 To UBound(arr) - LBound(arr))
    Dim currPos As Integer
    
    currPos = 1
    
    For i = 1 To UBound(arr)
        If (i <> pos) Then
            newArr(currPos) = arr(i)
            currPos = currPos + 1
        End If
    Next
    
    Remove = newArr
End Function

' -------------------------------------------
' create the solution and output it into Excel
' -------------------------------------------
Public Sub OutputSolution(ByRef outputRange As Range, ByVal createFlowDia As Boolean)
    On Error GoTo EH
    
    Dim nodeArr() As PointerNode
    ReDim nodeArr(1 To (pTableSize * pTableSize)) ''table size, ins/outs, not diagonals
    
    Dim count, currR, currC, inSeq, backtrack, totalMoves As Long
    Dim posDict, machineDict, stepNodeDict As Object
    Dim tmpStep, tmpStep2 As StepNode
    Dim lastStepList, lastStepList2 As Object
    Dim outRange As Range
    
    Set lastStepList = New ArrayList
    Set posDict = New Dictionary                                ''machine - position
    Set machineDict = New Dictionary                            ''position - machine
    Set stepNodeDict = New Dictionary                           ''machine - StepNode
    Set outRange = outputRange
    Set tmpStep = Me
    count = 1
    
    ''start stitching together results
    While (Not (tmpStep Is Nothing))
        ''start adding the tail end of the sequence
        If (Not (tmpStep.lastNode Is Nothing)) Then
            Set lastStepList2 = New ArrayList
            Set tmpStep2 = tmpStep.lastNode
            
            While (Not (tmpStep2 Is Nothing))
                If (tmpStep2.Solution <> -1) Then
                    lastStepList2.Add tmpStep2
                End If
                
                Set tmpStep2 = tmpStep2.nextNode
            Wend
            
            lastStepList2.Reverse
            
            If (lastStepList2.Count > 0) Then
                For Each lStep In lastStepList2.ToArray
                    lastStepList.Add lStep
                Next
            End If
        End If
    
        Set tmpStep2 = tmpStep
        Set tmpStep = tmpStep.nextNode
    Wend
    
    lastStepList.Reverse
    
    ''finish stitching
    If (lastStepList.Count > 0) Then
        For Each lStep In lastStepList.ToArray
             tmpStep2.nextNode = lStep
             lStep.prevNode = tmpStep2  ''repair node links
             Set tmpStep2 = tmpStep2.nextNode
        Next
    End If
    
    ''fill dictionaries
    Set tmpStep = Me
    
    While (Not (tmpStep Is Nothing))
        If (tmpStep.Solution <> -1) Then
            posDict.Add tmpStep.Solution, count
            machineDict.Add count, tmpStep.Solution
            stepNodeDict.Add tmpStep.Solution, tmpStep
            count = count + 1
        End If
        
        Set tmpStep = tmpStep.nextNode
    Wend
    
    ''add table values as nodes
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
                Dim value, dist As Long
                
                ''add pointer node
                Set ptrNode = New PointerNode
                Set prevSNode = stepNodeDict.Item(pMachineNbs(currR))
                Set nextSNode = stepNodeDict.Item(pMachineNbs(currC))
                dist = posDict.Item(nextSNode.Solution) - posDict.Item(prevSNode.Solution)
                value = pFromToValues(i, j)
                
                ptrNode.InitializeVariables nextSNode, prevSNode, value, dist
                Set nodeArr(count) = ptrNode
                
                ''update moves
                totalMoves = totalMoves + value
                If (dist = 1) Then
                    inSeq = inSeq + value
                ElseIf (dist < 0) Then
                    backtrack = backtrack + value
                End If
                
                ''update step node in/out moves
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
    
    ''add extra ins and outs
    For Each machine In pMachineNbs
        Dim step As StepNode
        Dim ptrNode2 As PointerNode
        Dim nextSNode2, prevSNode2 As StepNode
        Dim value2, dist2 As Long
        
        Set step = stepNodeDict.Item(machine)
        Set ptrNode2 = New PointerNode
        value2 = step.OutValue - step.InValue
        dist2 = 0 ''this is done to separate from 1-value distances
        
        ''parts added
        If (value2 > 0) Then
            Set nextSNode2 = step
            Set prevSNode2 = New StepNode
        ''parts exit
        ElseIf (value2 < 0) Then
            value2 = -value2
            Set nextSNode2 = New StepNode
            Set prevSNode2 = step
        ''value2 = 0; no parts in/out
        Else
            Set nextSNode2 = step
            Set prevSNode2 = step
        End If
        
        ptrNode2.InitializeVariables nextSNode2, prevSNode2, value2, dist2
        
        Set nodeArr(count) = ptrNode2
        count = count + 1
    Next
    
    ''create tables -----------------------
    Dim nbRows As Long
    Dim table(), tableMoves() As Variant
    Dim tableRange As Range
    
    ''create to/from table
    nbRows = UBound(pFromToValues) - LBound(pFromToValues) + 3
    ReDim table(1 To nbRows, 1 To nbRows)
    
    For i = 1 To nbRows
        For j = 1 To nbRows
            If (i = 1 And j = 1) Then
                table(i, j) = "From/To"
            ElseIf (i = j And i <> nbRows) Then
                table(i, j) = 0
            ElseIf (i = 1) Then
                If (j <> nbRows) Then
                    table(i, j) = pMachineNbs(j - 1)
                Else
                    table(i, j) = Chr(34) & "From" & Chr(34) & " Sums"
                End If
            ElseIf (j = 1) Then
                If (i <> nbRows) Then
                    table(i, j) = pMachineNbs(i - 1)
                Else
                    table(i, j) = Chr(34) & "To" & Chr(34) & " Sums"
                End If
            ElseIf (i = nbRows And j = nbRows) Then
                table(i, j) = totalMoves
            ElseIf (i = nbRows) Then
                table(i, j) = pToTotals(j - 1)
            ElseIf (j = nbRows) Then
                table(i, j) = pFromTotals(i - 1)
            Else
                table(i, j) = pFromToValues(i - 1, j - 1)
            End If
        Next
    Next
    
    Set tableRange = outRange.Resize(nbRows, nbRows)
    tableRange.value = table
    formatTable tableRange
    
    ''hollier method 2 add on
    If (pHollier2) Then
        ReDim table(1 To nbRows, 1 To 1)
        table(1, 1) = "From/To Ratios"
        table(UBound(table, 1), 1) = ""
        
        For i = 2 To UBound(table, 1) - 1
            If (pFromToRatios(i - 1) = INTMAX) Then
                table(i, 1) = "Infinite"
            Else
                table(i, 1) = Round(pFromToRatios(i - 1), 2)
            End If
        Next
        
        Set tableRange = outRange.Resize(nbRows, 1).Offset(0, nbRows)
        tableRange.value = table
        formatTable tableRange
        tableRange.Font.Bold = False
        tableRange.Cells(1, 1).Font.Bold = True
    End If
    
    Set outRange = outRange.Offset(nbRows + 1, 0)
    
    ''create comparison table
    If (Not (pHollier2)) Then
        ReDim tableMoves(1 To 2, 1 To 8)
        tableMoves(1, 1) = "Hollier Method"
        tableMoves(1, 3) = "Machine Order"
        tableMoves(1, 5) = "In-sequence %"
        tableMoves(1, 7) = "Backtracking %"
        tableMoves(2, 1) = 1
        tableMoves(2, 5) = Round(inSeq / totalMoves * 100, 2)
        tableMoves(2, 7) = Round(backtrack / totalMoves * 100, 2)
        
        For Each machine In machineDict.Keys
            tableMoves(2, 3) = tableMoves(2, 3) & machineDict.Item(machine) & " "
        Next
        
        Set tableRange = outRange.Resize(2, 8)
        tableRange.value = tableMoves
        formatTable tableRange
    Else
        ReDim tableMoves(1 To 1, 1 To 8)
        tableMoves(1, 1) = 2
        tableMoves(1, 5) = Round(inSeq / totalMoves * 100, 2)
        tableMoves(1, 7) = Round(backtrack / totalMoves * 100, 2)
        
        For Each machine In machineDict.Keys
            tableMoves(1, 3) = tableMoves(1, 3) & machineDict.Item(machine) & " "
        Next
        
        Set tableRange = outRange.Resize(1, 8).Offset(2, 0)
        tableRange.value = tableMoves
        formatTable tableRange
        tableRange.Font.Bold = False
    End If
    
    ''more formats for the comparison table
    Dim tmpRange As Range
    
    With tableRange
        Set tmpRange = Range(.Cells(1, 1), .Cells(1, 2))
        
        tmpRange.Merge
        tmpRange.Copy
        For i = 2 To .Rows.count
            .Cells(i, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                    SkipBlanks:=False, Transpose:=False
        Next
        
        Set tmpRange = Range(.Cells(1, 1), .Cells(.Rows.count, 2))
        tmpRange.Copy
        For i = 3 To .Columns.count
            .Cells(1, i).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                    SkipBlanks:=False, Transpose:=False
            i = i + 1
        Next
        
        For i = 2 To .Rows.count
            .Rows(i).Font.Bold = False
        Next
    End With
    
    If (Not (pHollier2)) Then
        Set outRange = outRange.Offset(4, 0)
    Else
        Set outRange = outRange.Offset(5, 0)
    End If
    ''----------------------
    
    ''create the machine flow diagram
    If (createFlowDia) Then
        Dim flowChart As Variant
        Set flowChart = makeFlowChart(nodeArr, stepNodeDict, posDict)
        
        With flowChart
            .Left = outRange.Left
            .Top = outRange.Top
            
            If (pHollier2) Then
                .Top = .Top + .height
            End If
        End With
    End If
    
    Application.CutCopyMode = False
    Exit Sub
    
EH:
    RaiseError Err.Number, Err.Source, "HollierProgram.OutputSolution", Err.Description, Erl
End Sub

' -------------------------------------------
' formats the table
' -------------------------------------------
Private Sub formatTable(ByRef tableRange As Range)
    With tableRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Rows(1).Font.Bold = True
        .Columns(1).Font.Bold = True
    
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    End With
End Sub

' -------------------------------------------
' creates a flow chart of the part manufacturing movements
' -------------------------------------------
Private Function makeFlowChart(ByRef pointers() As PointerNode, ByRef stepNodeDict As Variant, _
        ByRef posDict As Variant)
    Dim shapeList As Object
    Dim currShape, currTextBox As Shape
    Dim sPosLeft, sPosTop, sPosWidth, sPosHeight, distHor, distVert As Double
    Dim inToPt, distMod, heightMod As Double
    Dim count, fontSize As Long
    
    Set shapeList = New ArrayList
    inToPt = 72
    count = 1
    
    ''format layout
    sPosLeft = 11 * 1 / 12 * inToPt
    sPosTop = 3 * inToPt
    heightMod = 10
    fontSize = 10
    
    ''format machine nodes
    sPosWidth = 0.25 * inToPt
    sPosHeight = sPosWidth
    distMod = 1
    
    ''draw and add machine nodes
    For Each machine In stepNodeDict.Keys
        Set currShape = ActiveSheet.Shapes.AddShape(msoShapeFlowchartConnector, _
                sPosLeft * (distMod + 1), sPosTop, sPosWidth, sPosHeight)
        With currShape
            .Fill.Visible = msoFalse
            
            With .Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(0, 0, 0)
                .Transparency = 0
            End With
        
            With .TextFrame2
                With .TextRange.Font.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(0, 0, 0)
                    .Transparency = 0
                    .Solid
                End With
            
                .VerticalAnchor = msoAnchorMiddle
                With .TextRange
                    .ParagraphFormat.Alignment = msoAlignCenter
                    
                    With .Characters
                        .Text = stepNodeDict.Item(machine).Solution
                    
                        With .ParagraphFormat
                            .FirstLineIndent = 0
                            .Alignment = msoAlignCenter
                        End With
                        
                        With .Font
                            .NameComplexScript = "+mn-cs"
                            .NameFarEast = "+mn-ea"
                            .Fill.Visible = msoTrue
                            .Fill.ForeColor.RGB = RGB(0, 0, 0)
                            .Fill.Transparency = 0
                            .Fill.Solid
                            .Size = fontSize
                            .Name = "+mn-lt"
                        End With
                    End With
                End With
            End With
        End With
        
        shapeList.Add currShape.Name
        stepNodeDict.Item(machine).Shape = currShape
        count = count + 1
        distMod = distMod + 1
    Next
    
    ''draw and add pointer nodes
    Dim sPosLineTop, sPosTextBoxLeft, sPosTextBoxTop As Double
    sPosLineTop = sPosTop + sPosHeight / 2
    
    For Each pointer In pointers
        If (pointer.value <> 0) Then
        
            ''curved pointer
            If (pointer.dist > 1 Or pointer.dist < 0) Then
                Set currShape = ActiveSheet.Shapes.AddConnector(msoConnectorCurve, _
                sPosLeft, sPosLineTop, sPosLeft * (distMod + 1), sPosLineTop)
                
                With currShape
                    .Line.EndArrowheadStyle = msoArrowheadTriangle
                    If (pointer.dist > 0) Then
                        .ConnectorFormat.BeginConnect pointer.prevNode.Shape, 8
                        .ConnectorFormat.EndConnect pointer.nextNode.Shape, 2
                        .Adjustments.Item(1) = heightMod * pointer.dist
                        sPosTextBoxLeft = sPosLeft * (posDict.Item(pointer.prevNode.Solution) + _
                                pointer.dist / 2 + 1) - sPosWidth / 2
                    Else
                        .ConnectorFormat.BeginConnect pointer.prevNode.Shape, 4
                        .ConnectorFormat.EndConnect pointer.nextNode.Shape, 6
                        .Adjustments.Item(1) = heightMod * (-pointer.dist)
                        sPosTextBoxLeft = sPosLeft * (posDict.Item(pointer.nextNode.Solution) + _
                                -pointer.dist / 2 + 1) - sPosWidth / 2
                    End If
                    .ShapeStyle = msoLineStylePreset1
                End With
                
                sPosTextBoxTop = sPosLineTop - heightMod * (pointer.dist + 1)
                Set currTextBox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                        sPosTextBoxLeft, sPosTextBoxTop, sPosWidth * 2, sPosHeight)
                        
            ''straight pointer
            ElseIf (pointer.dist = 1) Then
                Set currShape = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
                        sPosLeft, sPosLineTop, sPosLeft * (distMod + 1), sPosLineTop)
                        
                With currShape
                    .Line.EndArrowheadStyle = msoArrowheadTriangle
                    .ConnectorFormat.BeginConnect pointer.prevNode.Shape, 7
                    .ConnectorFormat.EndConnect pointer.nextNode.Shape, 3
                    .ShapeStyle = msoLineStylePreset1
                End With
                
                sPosTextBoxLeft = sPosLeft * (posDict.Item(pointer.prevNode.Solution) + 1) + _
                        sPosLeft / 2 - sPosWidth / 2
                Set currTextBox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                        sPosTextBoxLeft, sPosLineTop - sPosHeight / 1.25, sPosWidth * 2, sPosHeight)
            
            ''in/out pointer
            Else
                
                ''out pointer
                If (pointer.nextNode Is Nothing) Then
                    
                    ''final out
                    If (posDict.Item(pointer.prevNode.Solution) = posDict.Count) Then
                    
                        sPosTextBoxLeft = sPosLeft * (posDict.count + 1.75)
                        Set currTextBox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                sPosTextBoxLeft, sPosLineTop - sPosHeight / 2, sPosWidth * 2, sPosHeight)
                                
                        Set currShape = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
                                sPosLeft, sPosLineTop, sPosLeft * (distMod + 1), sPosLineTop)
                                
                        With currShape
                            .Line.EndArrowheadStyle = msoArrowheadTriangle
                            .ConnectorFormat.BeginConnect pointer.prevNode.Shape, 7
                            .ConnectorFormat.EndConnect currTextBox, 2
                            .ShapeStyle = msoLineStylePreset1
                        End With
                    
                    ''regular out
                    Else
                    
                        sPosTextBoxLeft = sPosLeft * (posDict.Item(pointer.prevNode.Solution) + 1) - _
                                sPosWidth / 2
                        sPosTextBoxTop = sPosLineTop + heightMod * (posDict.count)
                        Set currTextBox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                sPosTextBoxLeft + sPosLeft / 4, sPosTextBoxTop, sPosWidth * 2, sPosHeight)
                                
                        Set currShape = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
                                sPosLeft, sPosLineTop, sPosLeft * (distMod + 1), sPosLineTop)
                                
                        With currShape
                            .Line.EndArrowheadStyle = msoArrowheadTriangle
                            .ConnectorFormat.BeginConnect pointer.prevNode.Shape, 5
                            .ConnectorFormat.EndConnect currTextBox, 1
                            .ShapeStyle = msoLineStylePreset1
                        End With
                        
                    End If
                
                ''in pointer
                Else
                
                    '' first in
                    If (posDict.Item(pointer.nextNode.Solution) = 1) Then
                        Set currTextBox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                sPosLeft, sPosLineTop - sPosHeight / 2, sPosWidth * 2, sPosHeight)
                                
                        Set currShape = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
                                sPosLeft, sPosLineTop, sPosLeft * (distMod + 1), sPosLineTop)
                                
                        With currShape
                            .Line.EndArrowheadStyle = msoArrowheadTriangle
                            .ConnectorFormat.BeginConnect currTextBox, 4
                            .ConnectorFormat.EndConnect pointer.nextNode.Shape, 3
                            .ShapeStyle = msoLineStylePreset1
                        End With
                       
                    ''regular in
                    Else
                        sPosTextBoxLeft = sPosLeft * (posDict.Item(pointer.nextNode.Solution) + 1) - _
                                sPosWidth / 2
                        sPosTextBoxTop = sPosLineTop + heightMod * (posDict.count)
                        Set currTextBox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                sPosTextBoxLeft - sPosLeft / 4, sPosTextBoxTop, sPosWidth * 2, sPosHeight)
                                
                        Set currShape = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
                                sPosLeft, sPosLineTop, sPosLeft * (distMod + 1), sPosLineTop)
                                
                        With currShape
                            .Line.EndArrowheadStyle = msoArrowheadTriangle
                            .ConnectorFormat.BeginConnect currTextBox, 1
                            .ConnectorFormat.EndConnect pointer.nextNode.Shape, 5
                            .ShapeStyle = msoLineStylePreset1
                        End With
                    End If
                End If
            End If
            
            ''pretty it up
            With currTextBox
                .Fill.Visible = msoFalse
                .Line.Visible = msoFalse
                
                With .TextFrame2
                    .VerticalAnchor = msoAnchorMiddle
                    .TextRange.ParagraphFormat.Alignment = msoAlignCenter
                    
                    With .TextRange.Characters
                        .Text = pointer.value
                        
                        With .ParagraphFormat
                            .FirstLineIndent = 0
                            .Alignment = msoAlignCenter
                        End With
                        
                        With .Font
                            .NameComplexScript = "+mn-cs"
                            .NameFarEast = "+mn-ea"
                            .Fill.Visible = msoTrue
                            .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
                            .Fill.ForeColor.TintAndShade = 0
                            .Fill.ForeColor.Brightness = 0
                            .Fill.Transparency = 0
                            .Fill.Solid
                            .Size = fontSize
                            .Name = "+mn-lt"
                        End With
                    End With
                End With
            End With
            
            shapeList.Add currShape.Name
            shapeList.Add currTextBox.Name
        End If
    Next
    
    ''group shapes
    Set makeFlowChart = ActiveSheet.Shapes.Range(shapeList.ToArray).Group
End Function
