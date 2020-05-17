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

Private pNextNode, pPrevNode As StepNode
Private pValue, pDist As Integer

Private Sub Class_Initialize()
' initialize in a different sub
End Sub

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

Public Property Get PrevNode() As StepNode
Set prevNode = pPrevNode
End Property

Public Property Get NextNode() As StepNode
Set nextNode = pNextNode
End Property

Public Property Get value() As Integer
value = pValue
End Property

Public Property Get dist() As Integer
dist = pDist
End Property
