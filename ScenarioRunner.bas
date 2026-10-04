Attribute VB_Name = "ScenarioRunner"
Option Explicit

' ICICI Bank initiation model: scenario runner.
' Import: Alt+F11 > File > Import File > ScenarioRunner.bas, then save the workbook as .xlsm.
' Uses workbook names: ScenarioSelector, TargetPrice, Upside, Rating, PAT_FY28E, RoE_FY28E, ScenarioOut.

Public Sub RunScenarios()
    Dim original As Variant
    Dim s As Integer
    Dim outCell As Range

    original = Range("ScenarioSelector").Value
    Set outCell = Range("ScenarioOut")   ' header cell "Scenario"; rows below are Bear, Base, Bull

    Application.ScreenUpdating = False
    On Error GoTo CleanUp

    For s = 1 To 3
        Range("ScenarioSelector").Value = s
        Application.Calculate
        outCell.Offset(s, 1).Value = Range("TargetPrice").Value
        outCell.Offset(s, 2).Value = Range("Upside").Value
        outCell.Offset(s, 3).Value = Range("Rating").Value
        outCell.Offset(s, 4).Value = Range("PAT_FY28E").Value
        outCell.Offset(s, 5).Value = Range("RoE_FY28E").Value
    Next s

CleanUp:
    Range("ScenarioSelector").Value = original
    Application.Calculate
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Scenario run stopped: " & Err.Description, vbExclamation
    Else
        MsgBox "Bear, Base and Bull results refreshed on Valuation, section G.", vbInformation
    End If
End Sub

Public Sub ResetToBase()
    Range("ScenarioSelector").Value = 2
    Application.Calculate
End Sub
