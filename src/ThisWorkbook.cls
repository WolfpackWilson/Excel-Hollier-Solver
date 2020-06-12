VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

''=======================================================
'' Class:       ThisWorkbook
'' Desc:        Initializes the program by creating a
''              button in the add-ons tab
'' Changes----------------------------------------------
'' Date         Programmer      Change
'' 06/11/2020   TheEric960      Written
''=======================================================

Private Const Button As String = "Hollier Solver"

' -------------------------------------------
' add the program to a button on workbook start
' -------------------------------------------
Private Sub Workbook_Open()
    Dim CmdBar As CommandBar
    Dim CmdBarMenu As CommandBarControl
    Dim CmdBarMenuItem As CommandBarControl
    
    Set CmdBar = Application.CommandBars("Worksheet Menu Bar")
    Set CmdBarMenu = CmdBar.Controls("Tools")   ' Index 6
    
    ''test if the button already exists and replace it
    On Error Resume Next
        Application.DisplayAlerts = False
        CmdBarMenu.Controls(Button).Delete
        Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set CmdBarMenuItem = CmdBarMenu.Controls.Add(Type:=msoControlButton)
    With CmdBarMenuItem
        .Caption = Button
        .OnAction = "HollierProgram.HollierProgram"
    End With
End Sub

' -------------------------------------------
' remove the added button on workbook close
' -------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim CmdBar As CommandBar
    Dim CmdBarMenu As CommandBarControl
    
    Set CmdBar = Application.CommandBars("Worksheet Menu Bar")
    Set CmdBarMenu = CmdBar.Controls("Tools")   ''index 6
    
    On Error Resume Next
        Application.DisplayAlerts = False
        CmdBarMenu.Controls(Button).Delete
        Application.DisplayAlerts = True
    On Error GoTo 0

End Sub
