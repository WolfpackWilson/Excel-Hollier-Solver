VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HollierForm 
   OleObjectBlob   =   "HollierForm.frx":0000
   Caption         =   "Hollier Method"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7260
   StartUpPosition =   2  'CenterScreen
   TypeInfoVer     =   62
End
Attribute VB_Name = "HollierForm"
Attribute VB_Base = "0{7257B53A-2B65-417B-9C46-32037726137E}{E5D88FAC-2DAA-42FD-A293-1E6F12814AEE}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

''=======================================================
'' Form:        HollierForm
'' Desc:        Gathers user settings and data for main
''              program.
'' Called by:   HollierProgram
'' Changes----------------------------------------------
'' Date         Programmer      Change
'' 5/21/1999    John Wilson     Written
''=======================================================


' ===========================================
' VARIABLE DECLARATIONS
' ===========================================

Public InputRange As Range
Public OutputCell As Range


' ===========================================
' FORM INTERACTION METHODS
' ===========================================

' -------------------------------------------
' OK button
' -------------------------------------------
Private Sub CommandButton1_Click()
    If (getRange(RefEdit1.Text, InputRange) And Not (InputRange Is Nothing)) Then
        ''check if range option is selected, is present, and is correct
        If (OptionButton1.value = True And getRange(RefEdit2.Text, OutputCell) _
                And Not (OutputCell Is Nothing)) Then
            Dim arr1Length, arr2Length As Integer
            Dim arrRange() As Variant
            
            arrRange = InputRange
            arr1Length = UBound(arrRange, 1) - LBound(arrRange, 1)
            arr2Length = UBound(arrRange, 2) - LBound(arrRange, 2)
            
            ''check if the array is square
            If (arr1Length = arr2Length) Then
                Set OutputCell = OutputCell.Cells(1, 1)
                Me.Hide
                Main.HollierMethod
                Unload Me
            Else
                MsgBox ("Please make sure your input has the same number of rows and columns.")
            End If
        ''use a new sheet
        ElseIf (OptionButton2.value = True) Then
            Me.Hide
            Main.HollierMethod
            Unload Me
        Else
            MsgBox ("Something is wrong with you output range. Please try reselecting it.")
        End If
    Else
        MsgBox ("Something is wrong with your input range. Please try reselecting it.")
    End If
End Sub

' -------------------------------------------
' Cancel button
' -------------------------------------------
Private Sub CommandButton2_Click()
    Unload Me
End Sub

' -------------------------------------------
' Help button
' -------------------------------------------
Private Sub CommandButton3_Click()
    MsgBox ("Input Range:" & vbTab & "Select your data" & vbLf _
           & "Machine Labels:" & vbTab & "Check if machines numbers are included" & vbLf & vbLf _
           & "Output Range:" & vbTab & "Select the output cell" & vbLf _
           & "New Worksheet:" & vbTab & "Output onto a ply worksheet" & vbLf _
           & "Flow Diagram:" & vbTab & "Create a flow chart from the results")
End Sub

' -------------------------------------------
' Output Range option
' -------------------------------------------
Private Sub OptionButton1_Click()
    RefEdit2.Enabled = True
    RefEdit2.BackColor = &H80000005
End Sub

' -------------------------------------------
' New Worksheet option
' -------------------------------------------
Private Sub OptionButton2_Click()
    RefEdit2.Enabled = False
    RefEdit2.BackColor = &H80000016
End Sub


' ===========================================
' FORM SUPPORTING METHODS
' ===========================================

' -------------------------------------------
' Checks refedit and converts to range if valid
' -------------------------------------------
Function getRange(RefEditText As String, myRange As Range) As Boolean
    On Error Resume Next
    Set myRange = Range(RefEditText)
    On Error GoTo 0
    getRange = Not myRange Is Nothing
End Function
