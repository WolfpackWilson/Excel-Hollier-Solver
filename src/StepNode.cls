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
Private mthd2Step As Object

' node Hollier values
Private pFromToValues() As Variant
Private pFromTotals() As Double
Private pToTotals() As Double
Private pFromToRatios() As Double
Private pMachineNbs() As Variant
Private pSolution As Variant
Private pTableSize As Long
Private pInitialized As Boolean
Private pHollier2 As Boolean

' for the output
Private pIn, pOut As Long
Private pShape As Shape

Private Sub Class_Initialize()
' initialize in a different sub
pInitialized = False
End Sub

' extract table information to initalize variables
Public Sub InitializeVariables(ByRef tableValues As Variant, ByVal labels As Boolean, _
        Optional ByRef MachineNbs As Variant, Optional ByRef prevNode As StepNode, _
        Optional ByVal hollier2 As Boolean)
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

If (IsMissing(hollier2)) Then
    pHollier2 = False
Else
    pHollier2 = hollier2
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


' remove the row and col of a two dimension array based on a 1-based position
Private Function RemoveRowCol(ByRef twoDimArr As Variant, ByVal posNb As Long)
Dim arrSize, newArrSize As Long
Dim newArr() As Variant
arrSize = UBound(twoDimArr, 1) - LBound(twoDimArr, 1) + 1

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

' total To and From column sums
Private Sub sumFromsAndTos(ByRef twoDimArr As Variant)
Dim arrSize, newArrSize As Long
arrSize = UBound(twoDimArr, 1) - LBound(twoDimArr, 1) + 1

ReDim pFromTotals(1 To arrSize)
ReDim pToTotals(1 To arrSize)
Dim currPosR, currPosC As Long
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

ReDim pFromToRatios(1 To arrSize)
For i = LBound(pToTotals) To UBound(pToTotals)
    If (pToTotals(i) = 0) Then
        pFromToRatios(i) = 2147483647
    Else
        pFromToRatios(i) = pFromTotals(i) / pToTotals(i)
    End If
Next

End Sub

' solves using a hollier method
Public Sub SolveHollier()
Dim currPos, lowest, lowestPos As Long
currPos = 1
lowest = pToTotals(LBound(pToTotals, 1))
lowestPos = currPos

If (pHollier2) Then
    lowest = pFromToRatios(LBound(pFromToRatios, 1))
End If

' find the lowest value/position
For i = LBound(pMachineNbs, 1) To UBound(pMachineNbs, 1)
    If (pHollier2) Then
        If (pFromToRatios(i) > lowest) Then ' finding highest in this method
            lowest = pFromToRatios(i)
            lowestPos = currPos
        End If
    Else
        If (pToTotals(i) < lowest) Then
            lowest = pToTotals(i)
            lowestPos = currPos
        End If
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
    nextStep.InitializeVariables newValues, False, newMachineNbs, Me, pHollier2
    nextStep.SolveHollier
End If
End Sub

' create the solution and output it into Excel
Public Sub OutputSolution(ByRef outputRange As Range, ByVal createFlowDia As Boolean)
Dim nodeArr() As PointerNode
ReDim nodeArr(1 To (pTableSize * pTableSize)) ' table size, ins/outs, not diagonals

Dim count, currR, currC, inSeq, backtrack, totalMoves As Long
Dim posDict, machineDict, stepNodeDict As Object
Dim tmpStep As StepNode
Dim outRange As Range

Set posDict = CreateObject("Scripting.Dictionary")      ' machine - position
Set machineDict = CreateObject("Scripting.Dictionary")    ' position - machine
Set stepNodeDict = CreateObject("Scripting.Dictionary") ' machine - StepNode
Set outRange = outputRange
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
            Dim value, dist As Long
            
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

' add extra ins and outs
For Each machine In pMachineNbs
    Dim step As StepNode
    Dim ptrNode2 As PointerNode
    Dim nextSNode2, prevSNode2 As StepNode
    Dim value2, dist2 As Long
    
    Set step = stepNodeDict(machine)
    Set ptrNode2 = New PointerNode
    value2 = step.OutValue - step.InValue
    dist2 = 0 ' this is done to separate from 1 distances
    
    If (value2 > 0) Then ' parts added
        Set nextSNode2 = step
        Set prevSNode2 = New StepNode
    ElseIf (value2 < 0) Then ' parts exit
        value2 = -value2
        Set nextSNode2 = New StepNode
        Set prevSNode2 = step
    Else ' value2 = 0; no parts in/out
        Set nextSNode2 = step
        Set prevSNode2 = step
    End If
    
    ptrNode2.InitializeVariables nextSNode2, prevSNode2, value2, dist2
    
    Set nodeArr(count) = ptrNode2
    count = count + 1
Next

' create tables
Dim nbRows As Long
Dim table(), tableMoves() As Variant
Dim tableRange As Range

' create to/from table
nbRows = UBound(pFromToValues) - LBound(pFromToValues) + 3
ReDim table(1 To nbRows, 1 To nbRows)

For i = 1 To nbRows
    For j = 1 To nbRows
        If (i = 1 And j = 1) Then
            table(i, j) = "From/To"
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

' hollier method 2 add on
If (pHollier2) Then
    ReDim table(1 To nbRows, 1 To 1)
    table(1, 1) = "From/To Ratios"
    table(UBound(table, 1), 1) = ""
    
    For i = 2 To UBound(table, 1) - 1
        If (pFromToRatios(i - 1) = 2147483647) Then
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

' create comparison table
If (Not (pHollier2)) Then
    ReDim tableMoves(1 To 2, 1 To 8)
    tableMoves(1, 1) = "Hollier Method"
    tableMoves(1, 3) = "Machine Order"
    tableMoves(1, 5) = "In-sequence %"
    tableMoves(1, 7) = "Backtracking %"
    tableMoves(2, 1) = 1
    tableMoves(2, 5) = Round(inSeq / totalMoves * 100, 2)
    tableMoves(2, 7) = Round(backtrack / totalMoves * 100, 2)
    
    For Each machine In machineDict
        tableMoves(2, 3) = tableMoves(2, 3) & machineDict(machine) & " "
    Next
    
    Set tableRange = outRange.Resize(2, 8)
    tableRange.value = tableMoves
    formatTable tableRange
Else
    ReDim tableMoves(1 To 1, 1 To 8)
    tableMoves(1, 1) = 2
    tableMoves(1, 5) = Round(inSeq / totalMoves * 100, 2)
    tableMoves(1, 7) = Round(backtrack / totalMoves * 100, 2)
    
    For Each machine In machineDict
        tableMoves(1, 3) = tableMoves(1, 3) & machineDict(machine) & " "
    Next
    
    Set tableRange = outRange.Resize(1, 8).Offset(2, 0)
    tableRange.value = tableMoves
    formatTable tableRange
    tableRange.Font.Bold = False
End If

With tableRange
    Range(.Cells(1, 1), .Cells(1, 2)).Select
    Selection.Merge
    Selection.Copy
    For i = 2 To .Rows.count
        .Cells(i, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
    Next
    
    Range(.Cells(1, 1), .Cells(.Rows.count, 2)).Select
    Selection.Copy
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
outRange.Offset(-(6 + nbRows), 0).Select
End Sub

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

' creates a flow chart of the part manufacturing movements
Private Function makeFlowChart(ByRef pointers() As PointerNode, ByRef stepNodeDict As Variant, _
        ByRef posDict As Variant)
Dim shapeList As Object
Dim currShape, currTextBox As Shape
Dim sPosLeft, sPosTop, sPosWidth, sPosHeight, distHor, distVert As Double
Dim inToPt, distMod As Double
Dim count, fontSize As Long

Set shapeList = CreateObject("System.Collections.ArrayList")
inToPt = 72
count = 1

' format layout
sPosLeft = 11 * 1 / 12 * inToPt
sPosTop = 3 * inToPt '8.5 * 1 / 8 * inToPt
fontSize = 10

' format machine nodes
sPosWidth = 0.25 * inToPt
sPosHeight = sPosWidth
distMod = 1

' draw and add machine nodes
For Each machine In stepNodeDict
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
                    .Text = stepNodeDict(machine).Solution
                
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
    
    shapeList.Insert shapeList.count, currShape.Name
    stepNodeDict(machine).Shape = currShape
    count = count + 1
    distMod = distMod + 1
Next

' draw and add pointer nodes
Dim sPosLineTop, sPosTextBoxLeft, sPosTextBoxTop As Double
sPosLineTop = sPosTop + sPosHeight / 2

For Each pointer In pointers
    If (pointer.value <> 0) Then
        If (pointer.dist > 1 Or pointer.dist < 0) Then ' curved pointer
            Set currShape = ActiveSheet.Shapes.AddConnector(msoConnectorCurve, _
            sPosLeft, sPosLineTop, sPosLeft * (distMod + 1), sPosLineTop)
            
            With currShape
                .Line.EndArrowheadStyle = msoArrowheadTriangle
                If (pointer.dist > 0) Then
                    .ConnectorFormat.BeginConnect pointer.prevNode.Shape, 8
                    .ConnectorFormat.EndConnect pointer.nextNode.Shape, 2
                    .Adjustments.Item(1) = 10 * pointer.dist
                    sPosTextBoxLeft = sPosLeft * (posDict(pointer.prevNode.Solution) + _
                            pointer.dist / 2 + 1) - sPosWidth / 2
                Else
                    .ConnectorFormat.BeginConnect pointer.prevNode.Shape, 4
                    .ConnectorFormat.EndConnect pointer.nextNode.Shape, 6
                    .Adjustments.Item(1) = 10 * (-pointer.dist)
                    sPosTextBoxLeft = sPosLeft * (posDict(pointer.nextNode.Solution) + _
                            -pointer.dist / 2 + 1) - sPosWidth / 2
                End If
                .ShapeStyle = msoLineStylePreset1
            End With
            
            sPosTextBoxTop = sPosLineTop - 10 * (pointer.dist + 1)
            Set currTextBox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                    sPosTextBoxLeft, sPosTextBoxTop, sPosWidth * 2, sPosHeight)
        ElseIf (pointer.dist = 1) Then ' straight pointer
            Set currShape = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
                    sPosLeft, sPosLineTop, sPosLeft * (distMod + 1), sPosLineTop)
                    
            With currShape
                .Line.EndArrowheadStyle = msoArrowheadTriangle
                .ConnectorFormat.BeginConnect pointer.prevNode.Shape, 7
                .ConnectorFormat.EndConnect pointer.nextNode.Shape, 3
                .ShapeStyle = msoLineStylePreset1
            End With
            
            sPosTextBoxLeft = sPosLeft * (posDict(pointer.prevNode.Solution) + 1) + _
                    sPosLeft / 2 - sPosWidth / 2
            Set currTextBox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                    sPosTextBoxLeft, sPosLineTop - sPosHeight / 1.25, sPosWidth * 2, sPosHeight)
        Else ' in/out pointer
            If (pointer.nextNode Is Nothing) Then ' out pointer
                If (posDict(pointer.prevNode.Solution) = posDict.count) Then ' last
                
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
                
                Else ' regular out
                
                    sPosTextBoxLeft = sPosLeft * (posDict(pointer.prevNode.Solution) + 1) - _
                            sPosWidth / 2
                    sPosTextBoxTop = sPosLineTop + 10 * (posDict.count)
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
            Else ' in pointer
            
                If (posDict(pointer.nextNode.Solution) = 1) Then ' first
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
                    
                Else ' regular in
                
                    sPosTextBoxLeft = sPosLeft * (posDict(pointer.nextNode.Solution) + 1) - _
                            sPosWidth / 2
                    sPosTextBoxTop = sPosLineTop + 10 * (posDict.count)
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
        
        ' pretty it up
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
        
        shapeList.Insert shapeList.count, currShape.Name
        shapeList.Insert shapeList.count, currTextBox.Name
    End If
Next

' group shapes
Set makeFlowChart = ActiveSheet.Shapes.Range(shapeList.ToArray).Group
End Function
