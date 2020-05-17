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
Private pDir As Boolean

Private Sub Class_Initialize()
' initialize in a different sub
End Sub

Public Sub InitializeVariables(NextNode As StepNode, PrevNode As StepNode, _
        value As Integer, dist As Integer, dir As Boolean)
Set pNextNode = NextNode
pValue = value
pDist = dist
pDir = dir
End Sub

Public Property Get PrevNode() As StepNode
PrevNode = pPrevNode
End Property

Public Property Get NextNode() As StepNode
NextNode = pNextNode
End Property

Public Property Get value() As Integer
value = pValue
End Property

Public Property Get dist() As Integer
dist = pDist
End Property

Public Property Get dir() As Boolean
dir = pDir
End Property
