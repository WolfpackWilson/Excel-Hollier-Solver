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

Private pFromToValues() As Variant
Private pFromTotals() As Variant
Private pToTotals() As Variant
Private pMachineNbs() As Variant
Private pTableSize As Integer

Private Sub Class_Initialize()
If (HollierForm.CheckBox1.Value = True) Then
    pFromToValues() = HollierForm.InputRange
Else
    pFromToValues() = HollierForm.InputRange
End If
End Sub
