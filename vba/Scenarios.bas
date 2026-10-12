Attribute VB_Name = "Scenarios"
' Scenario runner (MODELLING_STANDARDS section 7).
' Loops Scenario_Selected over every scenario in Scenario_Names (via the Scenario_Name selector),
' recalculates, runs the circularity and goal-seek solvers and writes Live_Results into the matching
' column of Stored_Results. Restores the original scenario (and re-solves it) at the end. Named ranges only.
Option Explicit

Public Sub RunAllScenarios()
    Dim original As Variant, i As Long, r As Long, n As Long
    original = Range("Scenario_Name").Value
    n = Range("Scenario_Names").Columns.Count
    For i = 1 To n
        Range("Scenario_Name").Value = Range("Scenario_Names").Cells(1, i).Value
        RunSolvers
        For r = 1 To Range("Live_Results").Rows.Count
            Range("Stored_Results").Cells(r, i).Value = Range("Live_Results").Cells(r, 1).Value
        Next r
    Next i
    Range("Scenario_Name").Value = original
    RunSolvers
End Sub

Public Sub RunSolvers()
    ' Solver.bas and GoalSeek.bas are only embedded in models with topics (named cell Solver_Topics),
    ' so they are called late-bound and only when that name exists.
    Application.Calculate
    If HasName("Solver_Topics") Then
        Application.Run "Solver.SolveAll"
        Application.Run "GoalSeek.SolveAll"
    End If
    Application.Calculate
End Sub

Private Function HasName(ByVal n As String) As Boolean
    Dim r As Range
    On Error Resume Next
    Set r = Range(n)
    HasName = Not (r Is Nothing)
    On Error GoTo 0
End Function
