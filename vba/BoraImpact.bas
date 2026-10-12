Attribute VB_Name = "BoraImpact"
' Bora macros (SPEC section 10). Named ranges only.
' RefreshAll: the commitment impact table, then the stored results for every key-assumption set.
' RunCommitmentImpact: on the base set, for each bid and commitment, set the lever to its zero level,
' recalculate, store base minus zeroed result in Impact_Stored and restore the lever. The last row of
' Impact_Stored holds the base; the last column pair takes the points per EUR m of RtB price PV given up.
' On any error the lever and the selected set are restored before the error is shown.
Option Explicit

Public Sub RefreshAll()
    Application.ScreenUpdating = False
    RunCommitmentImpact
    Scenarios.RunAllScenarios
    Application.ScreenUpdating = True
End Sub

Public Sub RunCommitmentImpact()
    Dim levers As Variant, zeros As Variant, live As Range, store As Range
    Dim nb As Long, nm As Long, nc As Long, i As Long, b As Long, k As Long, c As Long
    Dim orig As Variant, origSet As Variant, dpv As Variant, changed As Boolean
    Dim baseVal() As Variant, errNo As Long, errMsg As String
    levers = Array("Dash_Municipal", "Dash_Community", "Dash_Overbuild", "Dash_BESS", "Dash_BESS_Hours", _
                   "Dash_PB_Mult", "Dash_Partial", "Dash_Brownfield", "Dash_Shared", "Dash_EU_Turbine", _
                   "Dash_EU_Transformer")
    zeros = Array(0, 0, 0, 0, "keep", 1, 0, 0, 0, 0, 0)     ' "keep": no zero level (BESS hours stay)
    Set live = Range("Impact_Base_Live")
    Set store = Range("Impact_Stored")
    nb = Range("Dash_Municipal").Columns.Count
    nc = live.Columns.Count
    nm = nc \ nb
    origSet = Range("Scenario_Name").Value
    On Error GoTo Failed
    Range("Scenario_Name").Value = Range("Scenario_Names").Cells(1, 1).Value
    Application.Calculate
    ReDim baseVal(1 To nc)
    For c = 1 To nc
        baseVal(c) = live.Cells(1, c).Value
        store.Cells(store.Rows.Count, c).Value = baseVal(c)
    Next c
    For i = 0 To UBound(levers)
        For b = 1 To nb
            For k = 0 To nm
                store.Cells(i + 1, k * nb + b).ClearContents
            Next k
            If VarType(zeros(i)) <> vbString Then
                orig = Range(levers(i)).Cells(1, b).Value
                changed = True
                Range(levers(i)).Cells(1, b).Value = zeros(i)
                Application.Calculate
                For k = 0 To nm - 1
                    c = k * nb + b
                    store.Cells(i + 1, c).Value = Delta(baseVal(c), live.Cells(1, c).Value)
                Next k
                dpv = store.Cells(i + 1, nb + b).Value              ' metric 2: RtB price PV
                If VarType(dpv) = vbDouble Then
                    If dpv < -Range("Tolerance").Value Then
                        store.Cells(i + 1, nm * nb + b).Value = store.Cells(i + 1, b).Value / -dpv
                    End If
                End If
                Range(levers(i)).Cells(1, b).Value = orig
                changed = False
            End If
        Next b
    Next i
    Range("Scenario_Name").Value = origSet
    Application.Calculate
    Exit Sub
Failed:
    errNo = Err.Number
    errMsg = Err.Description
    On Error Resume Next
    If changed Then Range(levers(i)).Cells(1, b).Value = orig
    Range("Scenario_Name").Value = origSet
    Application.Calculate
    Application.ScreenUpdating = True
    On Error GoTo 0
    Err.Raise errNo, "RunCommitmentImpact", errMsg & " (levers and key-assumption set restored)"
End Sub

Private Function Delta(ByVal a As Variant, ByVal z As Variant) As Variant
    ' Base minus zeroed value; "n/a" when either side is an error or text (e.g. a failed XIRR).
    If IsError(a) Or IsError(z) Then
        Delta = "n/a"
    ElseIf VarType(a) = vbString Or VarType(z) = vbString Then
        Delta = "n/a"
    Else
        Delta = a - z
    End If
End Function
