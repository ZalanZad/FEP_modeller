Attribute VB_Name = "RtBProjects"
' RtB route per project (projects/bora_v5_summary SPEC section 20.3).
' RunProjects loops the Dashboard project selector Dash_Project over every project in In_Projects, recalculates
' and writes the RtB_Live column into the matching column of RtB_Stored. Restores the original project at the end
' (also after an error). RefreshAll = RunAllScenarios (Scenarios.bas) + RunProjects. Named ranges only.
Option Explicit

Public Sub RunProjects()
    Dim original As Variant, i As Long, r As Long, n As Long, errNo As Long, errMsg As String
    original = Range("Dash_Project").Value
    n = Range("In_Projects").Columns.Count
    On Error GoTo Failed
    For i = 1 To n
        Range("Dash_Project").Value = Range("In_Projects").Cells(1, i).Value
        Application.Calculate
        For r = 1 To Range("RtB_Live").Rows.Count
            Range("RtB_Stored").Cells(r, i).Value = Range("RtB_Live").Cells(r, 1).Value
        Next r
    Next i
    Range("Dash_Project").Value = original
    Application.Calculate
    Exit Sub
Failed:
    errNo = Err.Number
    errMsg = Err.Description
    On Error Resume Next
    Range("Dash_Project").Value = original
    Application.Calculate
    Application.ScreenUpdating = True
    On Error GoTo 0
    Err.Raise errNo, "RunProjects", errMsg & " (original project restored)"
End Sub

Public Sub RefreshAll()
    ' The case results are stored for the first project (Bora); the original selection is restored at the end.
    Dim original As Variant, errNo As Long, errMsg As String
    original = Range("Dash_Project").Value
    On Error GoTo Failed
    Range("Dash_Project").Value = Range("In_Projects").Cells(1, 1).Value
    RunAllScenarios
    RunProjects
    Range("Dash_Project").Value = original
    Application.Calculate
    Exit Sub
Failed:
    errNo = Err.Number
    errMsg = Err.Description
    On Error Resume Next
    Range("Dash_Project").Value = original
    Application.Calculate
    On Error GoTo 0
    Err.Raise errNo, "RefreshAll", errMsg & " (original project restored)"
End Sub
