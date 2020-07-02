VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PointerNode"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

''=======================================================
'' Class:       PointerNode
'' Desc:        Represents a pointer on the machine flow
''              diagram and stores the relevant data.
'' Called by:   StepNode (HollierProgram)
'' Changes----------------------------------------------
'' Date         Programmer      Change
'' 05/25/2020   TheEric960      Written
''=======================================================


' ===========================================
' VARIABLE DECLARATIONS
' ===========================================

Private pNextNode, pPrevNode As StepNode
Private pValue, pDist As Integer


' ===========================================
' INITIALIZE CLASS VARIABLES
' ===========================================

' -------------------------------------------
' initialize the class
' -------------------------------------------
Private Sub Class_Initialize()
    ' initialize in a different sub
End Sub

' -------------------------------------------
' initialize all the the variables
' -------------------------------------------
Public Sub InitializeVariables(ByRef nextNode As Variant, ByRef prevNode As Variant, _
        ByVal value As Integer, ByVal dist As Integer)
    If (nextNode.Initialized) Then
        Set pNextNode = nextNode
    Else
        Set pNextNode = Nothing
    End If
    
    If (prevNode.Initialized) Then
        Set pPrevNode = prevNode
    Else
        Set pPrevNode = Nothing
    End If
    
    pValue = value
    pDist = dist
End Sub


' ===========================================
' GETTERS AND SETTERS
' ===========================================

Public Property Get prevNode() As StepNode
    Set prevNode = pPrevNode
End Property

Public Property Get nextNode() As StepNode
    Set nextNode = pNextNode
End Property

Public Property Get value() As Integer
    value = pValue
End Property

Public Property Get dist() As Integer
    dist = pDist
End Property
